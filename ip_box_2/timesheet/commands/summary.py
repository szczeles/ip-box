import click

import os
from context import pass_ip_box, IpBoxContext, Project
from calendar import monthrange
from pathlib import Path
from openpyxl.workbook import Workbook
from openpyxl.utils import get_column_letter
from common import CsvReader, AggregatedTimeEntry
from model.excel import BaseSheet


class SummarySheet(BaseSheet):

    def __init__(self, sheet, project: Project, months, **kwargs):
        super().__init__(sheet, 'Sumy', **kwargs)
        self.set_header('B2', 'Pracownik')
        self.set_header('B3', 'Identyfikator KPWI')
        self.set_header('F2', 'Godzin pracy')
        self.set_header('G2', 'W tym KPWI')
        self.set_header('H2', 'Procent')
        sheet['C2'] = project.employee
        sheet['C3'] = project.id
        for letter in ('B', 'C'):
            sheet.column_dimensions[letter].width = 30
        for letter in ('E', 'F', 'G', 'H'):
            sheet.column_dimensions[letter].width = 15
        for idx, month in enumerate(months):
            row = idx + 3
            sheet[f'E{row}'] = month
            sheet[f'F{row}'] = f"='{month}'!H2"
            sheet[f'G{row}'] = f"='{month}'!H3"
            sheet[f'H{row}'] = f"=G{row}/F{row}"
            sheet[f'H{row}'].number_format = '0.00%'
        sheet['E15'] = 'SUMA'
        sheet['F15'] = '=SUM(F3:F14)'
        sheet['G15'] = '=SUM(G3:G14)'
        sheet['H15'] = '=G15/F15'
        sheet[f'H15'].number_format = '0.00%'


class MonthSheet(BaseSheet):

    _header = ['Dzień',	'KWIP',	'Inne',	'Łącznie', 'Podstawa obliczenia']

    def __init__(self, sheet, year, month, *args, **kwargs):
        super().__init__(sheet, *args, **kwargs)
        for idx, header in enumerate(self._header):
            cell = f"{get_column_letter(idx + 1)}1"
            self.set_header(cell, header)
        sheet.column_dimensions[get_column_letter(len(self._header))].width = 20
        sheet.column_dimensions['E'].width = 50
        sheet.column_dimensions['G'].width = 15
        _, days_count = monthrange(year, month)
        for day in range(1, days_count + 1):
            row = day + 1
            sheet[f'A{row}'] = day
            sheet[f'D{row}'] = f'=B{row}+C{row}'
        sheet['G2'] = 'Suma w miesiącu'
        sheet['H2'] = f"=SUM(D2:D{days_count + 1})"
        sheet['G3'] = 'w tym KPWI'
        sheet['H3'] = f"=SUM(B2:B{days_count + 1})"
        sheet['G4'] = 'procentowo KPWI'
        sheet['H4'] = "=H3/H2"
        sheet['H4'].number_format = '0.00%'


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
            (idx + 1, MonthSheet(self._workbook.create_sheet(), self._year, idx + 1, month))
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
            with IpBoxTimesheetSummaryWriter(project, year, f'{output}/{year}/{project.id}.xlsx') as writer:
                for month in range(1, 13):
                    path = Path(f'{reports_dir}/{year}/{month:02d}/{project.id}/agg-daily.csv')
                    if path.is_file():
                        with CsvReader(path) as reader:
                            for row in reader:
                                entry = AggregatedTimeEntry.from_row(row)
                                writer.write(entry)
                    else:
                        click.echo(f'report for {path} not found')
