from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color


class BaseSheet:

    def __init__(self, sheet, title, header_font=None, header_fill=None):
        self._sheet = sheet
        sheet.title = title
        self._header_font = header_font or Font(color='FF000000', bold=True)
        self._header_fill = header_fill or PatternFill("solid", fgColor=Color(indexed=22))

    def set_header(self, cell, value):
        self[cell] = value
        self[cell].font = self._header_font
        self[cell].fill = self._header_fill

    def __setitem__(self, key, value):
        self._sheet[key] = value

    def __getitem__(self, item):
        return self._sheet[item]