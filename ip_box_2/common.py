import calendar
import typing
import csv
import datetime


class MonthRange:

    def __init__(self, start, end):
        self._start = MonthRange.parse(start)
        self._end = MonthRange.parse(end)

    def __iter__(self):
        start_month = self._start.month
        end_month = self._end.month if self._start.year == self._end.year else 12
        for year in range(self._start.year, self._end.year + 1):
            for month in range(start_month, end_month + 1):
                yield year, month
            start_month = 1
            end_month = self._end.month if year == self._end.year - 1 else 12

    @staticmethod
    def parse(date_str) -> datetime.date:
        return datetime.datetime.strptime(date_str, "%Y-%m").date()


def end_of_month(date: datetime.date) -> datetime.date:
    num_days = calendar.monthrange(date.year, date.month)[1]
    return datetime.date(date.year, date.month, num_days)


class DaysRange:

    def __init__(self, start_date, end_date):
        if start_date > end_date:
            raise ValueError(f'start date ({start_date.strftime("%Y-%m-%d")}) '
                             f'is after end date ({end_date.strftime("%Y-%m-%d")}')
        self._start = start_date
        self._end = end_date

    def __iter__(self):
        delta = self._end - self._start
        for i in range(delta.days + 1):
            yield self._start + datetime.timedelta(days=i)

    @classmethod
    def till_end_of_month(cls, date: datetime.date):
        return cls(date, end_of_month(date))

    @classmethod
    def single(cls, date):
        return cls(date, date)


class TimeEntry:

    def __init__(self, date: datetime.date, task_type: str, notes: str, hours: float):
        self._date = date
        self._task_type = task_type
        self._notes = notes
        self._hours = hours

    @property
    def date(self) -> datetime.date:
        return self._date

    @property
    def task_type(self) -> str:
        return self._task_type

    @property
    def notes(self) -> str:
        return self._notes

    @property
    def hours(self) -> float:
        return self._hours

    @classmethod
    def from_row(cls, row):
        return cls(datetime.datetime.strptime(row[0], '%Y-%m-%d'), row[1], row[2], float(row[3]))


class ClassifiedTimeEntry(TimeEntry):

    def __init__(self, date: datetime.date, task_type: str, notes: str, hours: float, is_ip: bool):
        super().__init__(date, task_type, notes, hours)
        self._is_ip = is_ip

    @property
    def is_ip(self) -> bool:
        return self._is_ip

    @classmethod
    def from_time_entry(cls, entry: TimeEntry, is_ip: bool):
        return cls(entry.date, entry.task_type, entry.notes, entry.hours, is_ip)


class AggregatedTimeEntry:

    def __init__(self, date: datetime.date, ip_hours: float, other_hours: float, notes: typing.Set = None):
        self._date = date
        self._ip_hours = ip_hours
        self._other_hours = other_hours
        self._notes = notes or set()

    def merge(self, other):
        return self.__class__(self.date, self.ip_hours + other.ip_hours, self.other_hours + other.other_hours,
                              notes=self.notes.union(other.notes))

    @property
    def date(self) -> datetime.date:
        return self._date

    @property
    def ip_hours(self) -> float:
        return self._ip_hours

    @property
    def other_hours(self) -> float:
        return self._other_hours

    @property
    def total_hours(self) -> float:
        return self.ip_hours + self.other_hours

    @property
    def notes(self) -> typing.Set:
        return self._notes

    @classmethod
    def from_row(cls, row):
        return cls(datetime.datetime.strptime(row[0], '%Y-%m-%d').date(), float(row[1]), float(row[2]), row[4].split('; '))


class CsvReader:

    def __init__(self, path, delimiter=',', skip_header=True, header_lines=1):
        self._path = path
        self._delimiter = delimiter
        self._skip_header = skip_header
        self._header_lines = header_lines

    def __enter__(self):
        self._file = open(self._path, newline='').__enter__()
        self._reader = csv.reader(self._file, delimiter=self._delimiter)
        if self._skip_header:
            for _ in range(0, self._header_lines):
                next(self._reader)
        return self

    def __iter__(self):
        for row in self._reader:
            yield row

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._file.__exit__(exc_type, exc_val, exc_tb)


class CsvWriter:

    def __init__(self, filepath, header=None):
        self._filepath = filepath
        self._header = header

    def __enter__(self):
        self._file = open(self._filepath, 'w+')
        self._writer = csv.writer(self._file)
        if self._header:
            self._writer.writerow(self._header)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._file.close()

    def write(self, row: list[str]):
        self._writer.writerow(row)


class IpClassificationDetailedWriter(CsvWriter):

    def __init__(self, filepath, header=None):
        super().__init__(filepath, header=[
            'date',
            'task',
            'notes',
            'hours',
            'is_ip'
        ])

    def write(self, entry: ClassifiedTimeEntry):
        super().write([
            entry.date.strftime("%Y-%m-%d"),
            entry.task_type,
            entry.notes,
            f"{float(f'{entry.hours:.2f}'):g}",
            entry.is_ip
        ])


class IpClassificationAggregatedWriter(CsvWriter):

    def __init__(self, filepath, format: str = "%Y-%m-%d"):
        super().__init__(filepath, header=[
            'day',
            'ip',
            'other',
            'total',
            'notes'
        ])
        self._format= format

    def write(self, entry: AggregatedTimeEntry):
        super().write([
            entry.date.strftime(self._format),
            f"{float(f'{entry.ip_hours:.2f}'):g}",
            f"{float(f'{entry.other_hours:.2f}'):g}",
            f"{float(f'{entry.total_hours:.2f}'):g}",
            '; '.join(entry.notes)
        ])


class IpClassificationAggregator:

    def __init__(self, filepath, dates_range: DaysRange, pattern = "%Y-%m-%d"):
        self._filepath = filepath
        self._aggregation_dict = {}
        self._range = dates_range
        self._pattern = pattern

    def append(self, classified_time_entry: ClassifiedTimeEntry):
        key = self.make_aggregation_key(classified_time_entry.date)
        converted_entry = AggregatedTimeEntry(date=classified_time_entry.date,
                                              ip_hours=classified_time_entry.hours if classified_time_entry.is_ip else 0,
                                              other_hours=0 if classified_time_entry.is_ip else classified_time_entry.hours,
                                              notes={classified_time_entry.notes} if classified_time_entry.is_ip else None)
        merged_entry = self._aggregation_dict[key].merge(converted_entry) \
            if key in self._aggregation_dict \
            else converted_entry
        self._aggregation_dict[key] = merged_entry

    def make_aggregation_key(self, date: datetime.date) -> str:
        return date.strftime(self._pattern)

    def __enter__(self):
        return self

    def flush(self) -> typing.Iterable[AggregatedTimeEntry]:
        for date in self._range:
            key = self.make_aggregation_key(date)
            if key in self._aggregation_dict:
                yield self._aggregation_dict[key]
            else:
                yield AggregatedTimeEntry(date, ip_hours=0, other_hours=0)

    def __exit__(self, exc_type, exc_val, exc_tb):
        with IpClassificationAggregatedWriter(self._filepath, self._pattern) as writer:
            for entry in self.flush():
                writer.write(entry)
