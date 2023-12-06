import datetime
import dateutil.parser
import decimal
from common import CsvReader, CsvWriter
import typing


class KpirRow:

    def __init__(self, row):
        if not len(row) == 16:
            raise ValueError("wrong KPiR row")
        self._row = row

    @property
    def raw(self):
        return self._row

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
    return decimal.Decimal(value.replace(",", "."))


class KpirCsvReader(CsvReader):

    def __init__(self, path, header_lines=2, skip_footer=True):
        super().__init__(path, header_lines=header_lines)
        self._skip_footer = skip_footer

    def __iter__(self):
        iterator = super().__iter__()
        last = next(iterator)
        for row in iterator:
            yield last
            last = row
        if not self._skip_footer:
            yield last


class Kpir:

    _rows: typing.List[KpirRow]

    def __init__(self, rows: typing.Iterable[KpirRow], year: int = None, month: int = None):
        self._rows = [row for row in rows]
        for row in self._rows:
            if year and row.date.year != year:
                raise ValueError(f'Rows within {year} year expected, but got {row.date.year} for row number {row.number}')
            if month and row.date.month != month:
                raise ValueError(
                    f'Rows within {month} month expected, but got {row.date.month} for row number {row.number}')
        self._rows.sort(key=lambda item: item.number)
        self._year = year
        self._month = month

    def __iter__(self):
        return self._rows.__iter__()

    def filter(self, year, month=None, record_type=None):
        def filtered_rows_generator():
            for row in self:
                if ((row.date.year == year)
                        and (not month or row.date.month == month) and record_type):
                        yield row
        return Kpir(filtered_rows_generator(), year, month)
