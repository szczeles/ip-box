import click

import os
from context import pass_ip_box, IpBoxContext, Project
from calendar import monthrange
from pathlib import Path
from openpyxl.workbook import Workbook
from common import CsvReader, AggregatedTimeEntry, IpClassificationAggregator
from model.excel import BaseSheet
import datetime


class SummarySheet(BaseSheet):
    _title = 'Sumy'

    _header = {
        'B2': 'Pracownik',
        'B3': 'Identyfikator KPWI',
        'F2': 'Godzin pracy',
        'G2': 'W tym KPWI',
        'H2': 'Procent',
        'E15': 'SUMA'
    }
    _columns_width = {
        'B': 30,
        'C': 30,
        'E': 15,
        'F': 15,
        'G': 15,
        'H': 15
    }

    def __init__(self, sheet, project: Project, months, **kwargs):
        super().__init__(sheet, **kwargs)
        self['C2'] = project.employee
        self['C3'] = project.id
        for idx, month in enumerate(months):
            row = idx + 3
            self[f'E{row}'] = month
            self[f'F{row}'] = f"='{month}'!H2"
            self[f'G{row}'] = f"='{month}'!H3"
            self[f'H{row}'] = f"=G{row}/F{row}"
            self[f'H{row}'].number_format = '0.00%'
        self['F15'] = '=SUM(F3:F14)'
        self['G15'] = '=SUM(G3:G14)'
        self['H15'] = '=G15/F15'
        self[f'H15'].number_format = '0.00%'


class MonthSheet(BaseSheet):

    _header = {
        'A1': 'Dzień',
        'B1': 'KWIP',
        'C1': 'Inne',
        'D1': 'Łącznie',
        'E1': 'Podstawa obliczenia',
        'G2': 'Suma w miesiącu',
        'G3': 'w tym KPWI',
        'G4': 'procentowo KPWI'
    }
    _columns_width = {
        'A': 10,
        'B': 10,
        'C': 10,
        'D': 10,
        'E': 50,
        'F': 20,
        'G': 30
    }
    _wrap = True

    def __init__(self, sheet, year, month, *args, **kwargs):
        super().__init__(sheet, *args, **kwargs)
        _, days_count = monthrange(year, month)
        for day in range(1, days_count + 1):
            row = day + 1
            self[f'A{row}'] = day
            self[f'D{row}'] = f'=B{row}+C{row}'
        self['H2'] = f"=SUM(D2:D{days_count + 1})"
        self['H3'] = f"=SUM(B2:B{days_count + 1})"
        self['H4'] = "=H3/H2"
        self['H4'].number_format = '0.00%'


class IpBoxTimesheetSummaryWriter:

    _months = ['Styczeń', 'Luty', 'Marzec', 'Kwiecień', 'Maj', 'Czerwiec', 'Lipiec', 'Sierpień', 'Wrzesień',
              'Październik', 'Listopad', 'Grudzień']

    def __init__(self, project: Project, year: int, path):
        self._project = project
        self._year = year
        self._path = path

    def __enter__(self):
        self._workbook = Workbook()
        self._summary_sheet = SummarySheet(self._workbook.active, self._project, self._months)
        self._months_sheets_dict = dict(
            (idx + 1, MonthSheet(self._workbook.create_sheet(), self._year, idx + 1, title=month))
            for idx, month in enumerate(self._months))
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._workbook.save(filename=self._path)

    def write(self, entry: AggregatedTimeEntry):
        sheet = self._months_sheets_dict[entry.date.month]
        day = entry.date.day
        row = day + 1
        sheet[f'A{row}'] = day
        sheet[f'B{row}'] = entry.ip_hours
        sheet[f'C{row}'] = entry.other_hours
        sheet[f'D{row}'] = f'=B{row}+C{row}'
        sheet[f'E{row}'] = ', '.join(entry.notes)


def first_days_of_the_month(year: int):
    for month in range(1, 13):
        yield datetime.date(year, month, 1)


@click.option('--year', '-y', 'years',
              help="Year in format YYYY",
              type=int,
              required=True,
              multiple=True,
              prompt="Provide year in format YYYY", )
@click.option('--reports', '-r', 'reports_dir',
              help='Timesheet reports path',
              required=True,
              default='reports/timesheet',
              type=click.Path(exists=True, file_okay=False, dir_okay=True),
              prompt="Timesheet reports directory")
@click.option('--output', '-o',
              help='Summary output path',
              required=True,
              default='reports/timesheet',
              type=click.Path(exists=False, file_okay=False, dir_okay=True),
              prompt="Output directory")
@click.command()
@pass_ip_box
def summary(ip_box: IpBoxContext, years, reports_dir, output):
    for year in years:
        os.makedirs(f'{output}/{year}', exist_ok=True)
        for project in ip_box.get_active_projects(year):
            with (IpBoxTimesheetSummaryWriter(project, year, f'{output}/{year}/{project.id}.xlsx') as writer,
                IpClassificationAggregator(f'{output}/{year}/{project.id}.csv',
                                           first_days_of_the_month(year), '%Y-%m') as month_aggregator):
                for month in range(1, 13):
                    path = Path(f'{reports_dir}/{year}/{month:02d}/{project.id}/agg-daily.csv')
                    if path.is_file():
                        with CsvReader(path) as reader:
                            for row in reader:
                                entry = AggregatedTimeEntry.from_row(row)
                                writer.write(entry)
                                month_aggregator.append_aggregated(entry)
                    else:
                        click.echo(f'report for {path} not found')
