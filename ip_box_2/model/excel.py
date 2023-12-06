import typing

from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color


class BaseSheet:

    _header = None

    def __init__(self, sheet, title, header : typing.Mapping[str, str] = None, header_font=None, header_fill=None):
        self._sheet = sheet
        sheet.title = title
        self._header_font = header_font or Font(color='FF000000', bold=True)
        self._header_fill = header_fill or PatternFill("solid", fgColor=Color(indexed=22))
        if header:
            self._header = header
        for cell, title in (self._header or {}).items():
            self.set_header(cell, title)

    def set_header(self, cell, value):
        self.set_value(cell, value, font=self._header_font, fill=self._header_fill)

    def set_value(self, cell, value, font=None, fill=None):
        first_cell = cell.partition(":")[0] if ':' in cell else cell
        self._sheet[first_cell] = value
        if font:
            self[first_cell].font = font
        if fill:
            self[first_cell].fill = fill
        if ':' in cell:
            self._sheet.merge_cells(cell)

    def __setitem__(self, key, value):
        self.set_value(key, value)

    def __getitem__(self, item):
        return self._sheet[item]