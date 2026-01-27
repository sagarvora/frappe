# Copyright (c) 2015, Frappe Technologies Pvt. Ltd. and Contributors
# License: MIT. See LICENSE
from __future__ import annotations

import itertools
import re
from dataclasses import dataclass
from dataclasses import field as dataclass_field
from io import BytesIO
from typing import TYPE_CHECKING, Any, Literal

import xlrd
import xlsxwriter
from openpyxl import load_workbook
from openpyxl.workbook.child import INVALID_TITLE_REGEX

import frappe
from frappe.core.utils import html2text
from frappe.utils import cint
from frappe.utils.html_utils import unescape_html

if TYPE_CHECKING:
	from xlsxwriter.format import Format

ILLEGAL_CHARACTERS_RE = re.compile(
	r"[\000-\010]|[\013-\014]|[\016-\037]|\uFEFF|\uFFFE|\uFFFF|[\uD800-\uDFFF]"
)


class StyleRef:
	"""
	Immutable, hashable reference to a style dict.

	Enables automatic deduplication of identical styles via interning.
	Two StyleRefs with the same content are equal and share the same hash.
	"""

	__slots__ = ("_hash", "dict")

	def __init__(self, style: dict):
		self.dict: dict = style
		self._hash: int = hash(frozenset(style.items()))

	def __hash__(self) -> int:
		return self._hash

	def __eq__(self, other: object) -> bool:
		if not isinstance(other, StyleRef):
			return NotImplemented
		# fast path: different hashes mean definitely not equal
		if self._hash != other._hash:
			return False
		# same hash: compare dicts (handles collisions)
		return self.dict == other.dict

	def __repr__(self) -> str:
		return f"StyleRef({self.dict!r})"


### XLSX Formatter ###
@dataclass
class XLSXMetadata:
	"""
	Metadata for XLSX reports for exports.
	"""

	report_name: str = ""

	filters: dict = dataclass_field(default_factory=dict)

	row_map: dict[int, dict | list] = dataclass_field(default_factory=dict)
	column_map: dict[int, dict] = dataclass_field(default_factory=dict)

	header_index: int = 0
	last_row_index: int = 0
	max_indent_level: int = 0

	add_total_row: bool = False
	include_filters: bool = False
	ignore_visible_idx: bool = True
	include_indentation: bool = False
	include_hidden_columns: bool = False

	def get_column_index(self, fieldname: str) -> int | None:
		return next((idx for idx, col in self.column_map.items() if col.get("fieldname") == fieldname), None)

	def get_column(self, fieldname: str) -> dict | None:
		return next((col for col in self.column_map.values() if col.get("fieldname") == fieldname), None)

	def get_row(self, row_idx: int) -> dict | list | None:
		return self.row_map.get(row_idx)


class XLSXStyleBuilder:
	"""
	Builder for configuring Excel cell styles with automatic deduplication.

	Styles are stored as StyleRef objects (hashable, immutable wrappers around dicts).
	Multiple styles can be stacked on the same target - they merge in order:
	column → row → cell (later wins on conflict).

	Usage:
		builder = XLSXStyleBuilder(metadata)

		# pass dicts directly (auto-converted to StyleRef)
		builder.style_column(0, {"num_format": "#,##0.00"})
		builder.style_row(0, {"bold": True})

		# or create reusable refs for performance
		currency_fmt = builder.style({"num_format": "$#,##0.00"})
		builder.style_column(1, currency_fmt)
		builder.style_column(2, currency_fmt)

		styles = builder.build()
	"""

	def __init__(self, metadata: XLSXMetadata):
		self.metadata = metadata

		# interned style refs: StyleRef → StyleRef (for deduplication)
		self._style_cache: dict[StyleRef, StyleRef] = {}

		# sparse maps: index → list of StyleRefs (for stacking)
		self._column_styles: dict[int, list[StyleRef]] = {}
		self._row_styles: dict[int, list[StyleRef]] = {}
		self._cell_styles: dict[tuple[int, int], list[StyleRef]] = {}

		self._set_defaults()
		self._register_default_styles()

	### POST INIT METHODS ###
	def _set_defaults(self):
		self.currency_field_exists = any(
			col.get("fieldtype") == "Currency" for col in self.metadata.column_map.values()
		)

		self.currency_fields: dict[int, dict] = {}

		if self.currency_field_exists:
			for idx, col in self.metadata.column_map.items():
				if col.get("fieldtype") == "Currency":
					self.currency_fields[idx] = col

	def _register_default_styles(self):
		# highlight styles
		self._header_style = self.style({"bold": True, "font_size": 12})
		self._total_row_style = self.style({"bold": True})
		self._filter_label_style = self.style({"bold": True})

		# indent styles
		self._indent_styles: dict[int, StyleRef] = {}
		if self.metadata.max_indent_level:
			for indent in range(self.metadata.max_indent_level + 1):
				self._indent_styles[indent] = self.style({"align": "left", "indent": indent * 2})

		# fieldtype format styles
		self._float_format = self.style({"num_format": self.get_number_format("Float")})
		self._percent_format = self.style({"num_format": self.get_number_format("Percent")})
		self._date_format = self.style({"num_format": self.get_date_format()})
		self._time_format = self.style({"num_format": self.get_time_format()})
		self._datetime_format = self.style({"num_format": self.get_datetime_format()})

		# currency format cache
		self._currency_formats: dict[str, StyleRef] = {}

	### STYLE CREATION ###
	def style(self, style_dict: dict) -> StyleRef:
		"""
		Create or return an interned StyleRef for the given style dict.

		Identical dicts return the same StyleRef instance.
		"""
		# create ref first to compute hash, then use for cache lookup
		ref = StyleRef(style_dict)
		if existing := self._style_cache.get(ref):
			return existing

		self._style_cache[ref] = ref
		return ref

	def _normalize(self, style: dict | StyleRef) -> StyleRef:
		"""Convert dict to StyleRef if needed."""
		if isinstance(style, StyleRef):
			return style

		return self.style(style)

	### STYLE APPLICATION ###
	def style_column(self, col_idx: int, style: dict | StyleRef):
		"""Apply a style to an entire column. Stacks with existing styles."""
		if col_idx not in self._column_styles:
			self._column_styles[col_idx] = []

		self._column_styles[col_idx].append(self._normalize(style))
		return self

	def style_row(self, row_idx: int, style: dict | StyleRef):
		"""Apply a style to an entire row. Stacks with existing styles."""
		if row_idx not in self._row_styles:
			self._row_styles[row_idx] = []

		self._row_styles[row_idx].append(self._normalize(style))
		return self

	def style_cell(self, row_idx: int, col_idx: int, style: dict | StyleRef):
		"""Apply a style to a specific cell. Stacks with existing styles."""
		cell_key = (row_idx, col_idx)
		if cell_key not in self._cell_styles:
			self._cell_styles[cell_key] = []

		self._cell_styles[cell_key].append(self._normalize(style))
		return self

	def build(self) -> dict:
		"""
		Build the final style configuration for make_xlsx.

		Returns sparse maps with StyleRef lists. make_xlsx handles
		merging and Format object creation.
		"""
		return {
			"column_styles": self._column_styles,
			"row_styles": self._row_styles,
			"cell_styles": self._cell_styles,
		}

	### UTILITY METHODS ###
	def apply_default_styles(self, currency_formatting: bool = False, currency: str | dict | None = None):
		"""Apply standard styles: header, filters, total row, indentation, fieldtype formats."""
		self.style_header()

		if self.metadata.include_filters:
			self.style_filters()

		if self.metadata.add_total_row and self.metadata.ignore_visible_idx:
			self.style_total_row()

		if self.metadata.include_indentation:
			self.apply_indentations(0)

		self.apply_default_fieldtype_formats(currency_formatting=currency_formatting, currency=currency)

		return self

	def style_header(self):
		return self.style_row(self.metadata.header_index, self._header_style)

	def style_filters(self):
		LABEL_COLUMN_INDEX = 0
		for row_idx in range(self.metadata.header_index):
			self.style_cell(row_idx, LABEL_COLUMN_INDEX, self._filter_label_style)
		return self

	def apply_indentations(self, column: int):
		for idx, row in self.metadata.row_map.items():
			if isinstance(row, dict) and "indent" in row:
				indent = row["indent"]
				if style := self._indent_styles.get(indent):
					self.style_cell(idx, column, style)
		return self

	def style_total_row(self):
		return self.style_row(self.metadata.last_row_index, self._total_row_style)

	def apply_default_fieldtype_formats(
		self, *, currency_formatting: bool = False, currency: str | dict | None = None
	):
		fieldtype_styles = {
			"Float": self._float_format,
			"Percent": self._percent_format,
			"Date": self._date_format,
			"Time": self._time_format,
			"Datetime": self._datetime_format,
		}

		for idx, col in self.metadata.column_map.items():
			if style := fieldtype_styles.get(col.get("fieldtype")):
				self.style_column(idx, style)

		if currency_formatting:
			self.apply_currency_fieldtype_formats(currency)

		return self

	def apply_currency_fieldtype_formats(self, currency: str | dict | None = None):
		if not self.currency_field_exists:
			return self

		def _get_currency_style(currency_code: str) -> StyleRef:
			if currency_code not in self._currency_formats:
				num_format = self.get_number_format("Currency", currency_code)
				self._currency_formats[currency_code] = self.style({"num_format": num_format})
			return self._currency_formats[currency_code]

		# single currency for all currency fields
		if isinstance(currency, str):
			style = _get_currency_style(currency)
			for idx in self.currency_fields:
				self.style_column(idx, style)

		# currency mapping per field
		elif isinstance(currency, dict):
			for fieldname, code in currency.items():
				if idx := self.metadata.get_column_index(fieldname):
					self.style_column(idx, _get_currency_style(code))

		# currency per row based on metadata
		else:
			default_currency = frappe.db.get_default("currency")

			for row_idx, row in self.metadata.row_map.items():
				if not isinstance(row, dict):
					continue

				for col_idx, col in self.currency_fields.items():
					curr = self.get_field_currency(col, row) or default_currency
					self.style_cell(row_idx, col_idx, _get_currency_style(curr))

		return self

	@staticmethod
	def get_field_currency(df: dict, doc: dict) -> str | None:
		fieldname = df.get("fieldname")
		options = df.get("options")

		if not (options and fieldname and doc):
			return None

		if ":" in options:
			parts = options.split(":")
			if len(parts) == 3 and (docname := doc.get(parts[1])):
				return XLSXStyleBuilder._get_currency(parts[0], docname, parts[2])
			return None
		return doc.get(options)

	@staticmethod
	@frappe.request_cache
	def _get_currency(doctype: str, docname: str, fieldname: str) -> str | None:
		return frappe.get_value(doctype, docname, fieldname)

	### FORMAT GETTERS ###
	@staticmethod
	def get_date_format() -> str:
		return frappe.get_system_settings("date_format")

	@staticmethod
	def get_time_format() -> str:
		return frappe.get_system_settings("time_format")

	@staticmethod
	def get_datetime_format() -> str:
		return f"{XLSXStyleBuilder.get_date_format()} {XLSXStyleBuilder.get_time_format()}"

	@staticmethod
	@frappe.request_cache
	def get_number_format(
		fieldtype: Literal["Currency", "Float", "Percent"],
		currency: str | None = None,
	) -> str:
		from frappe.locale import get_number_format as _get_format

		number_format = _get_format()
		thousands_sep = number_format.thousands_separator
		decimal_sep = number_format.decimal_separator
		precision = number_format.precision

		if fieldtype == "Currency":
			precision = cint(frappe.db.get_default("currency_precision")) or precision
			format_str = XLSXStyleBuilder._build_number_format(thousands_sep, decimal_sep, precision)
			currency_symbol, symbol_on_right = XLSXStyleBuilder._get_currency_symbol_info(currency)
			return XLSXStyleBuilder._get_currency_format(format_str, currency_symbol, symbol_on_right)

		if fieldtype in ("Float", "Percent"):
			precision = cint(frappe.db.get_default("float_precision")) or precision
			format_str = XLSXStyleBuilder._build_number_format(thousands_sep, decimal_sep, precision)
			return f'{format_str}"%" ' if fieldtype == "Percent" else format_str

		return "General"

	@staticmethod
	def _build_number_format(thousands_sep: str, decimal_sep: str, precision: int = 0) -> str:
		integer_part = "#,##0" if thousands_sep else "#0"
		decimal_part = (decimal_sep + "0" * precision) if precision > 0 else ""
		return f"{integer_part}{decimal_part}"

	@staticmethod
	def _get_currency_symbol_info(currency: str | None) -> tuple[str, bool]:
		if not currency or frappe.db.get_default("hide_currency_symbol") == "Yes":
			return "", False

		symbol, on_right = frappe.db.get_value("Currency", currency, ["symbol", "symbol_on_right"])
		return frappe._(symbol or currency), bool(on_right)

	@staticmethod
	def _get_currency_format(
		format_string: str,
		currency_symbol: str | None = None,
		symbol_on_right: bool = False,
	) -> str:
		if not currency_symbol:
			return format_string

		if symbol_on_right:
			return f'{format_string}" {currency_symbol}";-{format_string}" {currency_symbol}"'

		return f'"{currency_symbol} "{format_string};"{currency_symbol} "-{format_string}'


### Excel Creation ###
def make_xlsx(
	data: list[list[Any]],
	sheet_name: str,
	wb: xlsxwriter.Workbook | None = None,
	column_widths: list[int] | None = None,
	styles: dict | None = None,
) -> BytesIO:
	"""
	Create an Excel file with the given data and formatting options.

	Args:
		data: List of rows, where each row is a list of cell values
		sheet_name: Name of the Excel sheet
		wb: Existing workbook to add sheet to. If None, creates new workbook
		column_widths: List of column widths in Excel units. If None, auto-sized
		styles: Dictionary from XLSXStyleBuilder.build() containing:
			- column_styles: dict[int, list[StyleRef]]
			- row_styles: dict[int, list[StyleRef]]
			- cell_styles: dict[tuple[int, int], list[StyleRef]]

	Returns:
		BytesIO: object containing the Excel file data
	"""
	column_widths = column_widths or []
	styles = styles or {}

	xlsx_file = BytesIO()
	created_wb = wb is None

	if created_wb:
		wb = xlsxwriter.Workbook(xlsx_file, {"in_memory": True})

	sheet_name_sanitized = INVALID_TITLE_REGEX.sub(" ", sheet_name)
	ws = wb.add_worksheet(sheet_name_sanitized[:31])

	for i, column_width in enumerate(column_widths):
		if column_width:
			ws.set_column(i, i, column_width)

	col_styles: dict[int, list[StyleRef]] = styles.get("column_styles") or {}
	row_styles: dict[int, list[StyleRef]] = styles.get("row_styles") or {}
	cell_styles: dict[tuple[int, int], list[StyleRef]] = styles.get("cell_styles") or {}

	styling_enabled = bool(col_styles or row_styles or cell_styles)

	if not styling_enabled:
		ws.set_row(0, cell_format=wb.add_format({"bold": True}))

	# format cache: tuple of StyleRefs → Format object
	# this avoids creating duplicate Format objects for identical style combinations
	format_cache: dict[tuple[StyleRef, ...], Format] = {}

	def get_format(refs: tuple[StyleRef, ...]) -> Format:
		"""Get or create a Format for a tuple of StyleRefs."""
		format_obj = format_cache.get(refs)
		if format_obj is None:
			if len(refs) == 1:
				merged = refs[0].dict
			else:
				# merge all style dicts in order (later wins on conflict)
				merged = {}
				for ref in refs:
					merged.update(ref.dict)
			format_obj = wb.add_format(merged)
			format_cache[refs] = format_obj

		return format_obj

	# local references for hot loop
	write = ws.write
	illegal_chars_sub = ILLEGAL_CHARACTERS_RE.sub
	handle_html_content = sheet_name not in {"Data Import Template", "Data Export"}
	col_styles_get = col_styles.get
	row_styles_get = row_styles.get
	cell_styles_get = cell_styles.get
	itertools_chain = itertools.chain

	for row_idx, row in enumerate(data):
		row_refs = row_styles_get(row_idx)

		for col_idx, value in enumerate(row):
			if isinstance(value, str):
				if handle_html_content:
					value = handle_html(value)
				value = illegal_chars_sub("", value)

			cell_format = None
			if styling_enabled:
				col_refs = col_styles_get(col_idx)
				cell_refs = cell_styles_get((row_idx, col_idx))

				key = tuple(itertools_chain(col_refs or (), row_refs or (), cell_refs or ()))
				if key:
					cell_format = get_format(key)

			write(row_idx, col_idx, value, cell_format)

	if created_wb:
		wb.close()

	xlsx_file.seek(0)
	return xlsx_file


### Utilities ###
def handle_html(data: str) -> str:
	# return if no html tags found
	if "<" not in data or ">" not in data:
		return data

	h = unescape_html(data or "")

	try:
		value = html2text(h, strip_links=True, wrap=False)
	except Exception:
		# unable to parse html, send it raw
		return data

	return value.replace("  \n", ", ").replace("\n", " ").replace("# ", ", ")


def read_xlsx_file_from_attached_file(file_url=None, fcontent=None, filepath=None):
	if file_url:
		_file = frappe.get_doc("File", {"file_url": file_url})
		filename = _file.get_full_path()
	elif fcontent:
		filename = BytesIO(fcontent)
	elif filepath:
		filename = filepath
	else:
		return

	rows = []
	wb1 = load_workbook(filename=filename, data_only=True)
	ws1 = wb1.active
	for row in ws1.iter_rows():
		rows.append([cell.value for cell in row])
	return rows


def read_xls_file_from_attached_file(content):
	book = xlrd.open_workbook(file_contents=content)
	sheets = book.sheets()
	sheet = sheets[0]
	return [sheet.row_values(i) for i in range(sheet.nrows)]


def build_xlsx_response(data, filename):
	from frappe.desk.utils import provide_binary_file

	provide_binary_file(filename, "xlsx", make_xlsx(data, filename).getvalue())
