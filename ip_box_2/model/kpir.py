#!/usr/bin/env python3
import abc
import argparse
import datetime
import decimal
import os
import csv
import re

import dateutil.parser
from decimal import Decimal
import yaml

from openpyxl.workbook import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter


class KPiRRow:

    def __init__(self, row):
        if not len(row) == 16:
            raise ValueError("wrong KPiR row")
        self._row = row

    @property
    def number(self) -> int:
        return int(self._row[0])

    @property
    def date(self) -> datetime.date:
        return dateutil.parser.isoparse(self._row[1]).date()

    @property
    def invoice_number(self) -> str:
        return self._row[2]

    @property
    def company_name(self) -> str:
        return self._row[3]

    @property
    def company_address(self) -> str:
        return self._row[4]

    @property
    def description(self) -> str:
        return self._row[5]

    @property
    def sales_income(self) -> decimal.Decimal:
        return as_decimal(self._row[6])

    @property
    def other_income(self) -> decimal.Decimal:
        return as_decimal(self._row[7])

    @property
    def income(self) -> decimal.Decimal:
        return as_decimal(self._row[8])

    @property
    def is_income(self) -> bool:
        return self.income > 0

    @property
    def other_cost(self) -> decimal.Decimal:
        return as_decimal(self._row[12])

    @property
    def cost(self) -> decimal.Decimal:
        return as_decimal(self._row[13])

    @property
    def is_cost(self) -> bool:
        return self.cost > 0


def as_decimal(value: str) -> decimal.Decimal:
    return Decimal(value.replace(",", "."))


class KPiR:

    def __init__(self, year: str, rows):
        self._year = year
        self._rows = [KPiRRow(row) for row in rows]

    @property
    def costs(self):
        return self.filter(lambda row: row.is_cost)

    @property
    def incomes(self):
        return self.filter(lambda row: row.is_income)

    def filter(self, condition):
        return [row for row in self._rows if condition(row)]

    def filter_year(self, year: int):
        return self.filter(lambda row: row.date.year == year)

    @property
    def rows(self):
        return self._rows

    def compute_totals(self, year):
        incomes = Decimal("0")
        costs = Decimal("0")
        for row in self.filter_year(year):
            if row.is_cost:
                costs += row.cost
            if row.is_income:
                incomes += row.income
        return incomes, costs


class KPiRCsvReader:

    def __init__(self, file_path, year, with_header=True):
        self._file_path = file_path
        self._year = str(year)
        self._with_header = with_header

    def __enter__(self):
        self._file = open(self._file_path, 'r')
        self._reader = csv.reader(self._file)
        return self

    def __exit__(self, *args):
        self._file.close()
        return True

    def _is_applicable(self, row):
        return len(row) == 16 and row[1].startswith(self._year)

    def _read_rows(self):
        row_number = 0
        for row in self._reader:
            row_number += 1
            if row_number == 1 and self._with_header:
                continue
            if self._is_applicable(row):
                yield row

    def read(self):
        return KPiR(self._year, [row for row in self._read_rows()])
