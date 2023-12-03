import datetime
import calendar
import typing


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
