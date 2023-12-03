import typing

import calendar
import click
import os
import csv
from pathlib import Path
import datetime
from context import pass_ip_box, IpBoxContext
from common import MonthRange, DaysRange, ClassifiedTimeEntry, TimeEntry, AggregatedTimeEntry


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


@click.option('--start-date', '-s', help="Start date in YYYY-mm format", required=True, default='2019-01', prompt="Provide start date in format YYYY-mm")
@click.option('--end-date', '-e', help="End date in YYYY-mm format", required=True, default=datetime.datetime.now().strftime("%Y-%m"), prompt="Provide end date in format YYYY-mm")
@click.option('--harvest-reports', '-r', help='Harvest reports path', required=True, default='reports/harvest', type=click.Path(exists=True, file_okay=False, dir_okay=True), prompt="Harvest reports directory")
@click.option('--output', '-o', help='Reports Output path', required=True, default='reports/timesheet', type=click.Path(exists=False, file_okay=False, dir_okay=True), prompt="Output directory")
@click.command()
@pass_ip_box
def classify(ip_box: IpBoxContext, start_date, end_date, harvest_reports, output):
    for (year, month) in MonthRange(start_date, end_date):
        reports_dir = f'{harvest_reports}/{year}/{month:02d}'
        for project in ip_box.get_active_projects(year, month):
            base_path = f'{output}/{year}/{month:02d}/{project.id}'
            os.makedirs(base_path, exist_ok=True)
            range_start = datetime.date(year, month, 1)
            with (IpClassificationDetailedWriter(f'{base_path}/detailed.csv') as writer,
                  IpClassificationAggregator(f'{base_path}/agg-daily.csv',
                                             DaysRange.till_end_of_month(range_start)) as days_aggregator,
                  IpClassificationAggregator(f'{base_path}/agg-monthly.csv',
                                            DaysRange.single(range_start), '%Y-%m') as month_aggregator):
                for harvest_selector in project.harvest_selectors:
                    report = Path(f'{reports_dir}/{harvest_selector.code}.csv')
                    if report.is_file():
                        with open(report, newline='') as csvfile:
                            reader = csv.reader(csvfile, delimiter=',')
                            next(reader)  # skip header
                            for row in reader:
                                time_entry = TimeEntry.from_row(row)
                                is_ip = harvest_selector.classify(time_entry)
                                classified_time_entry = ClassifiedTimeEntry.from_time_entry(time_entry, is_ip)
                                writer.write(classified_time_entry)
                                days_aggregator.append(classified_time_entry)
                                month_aggregator.append(classified_time_entry)
