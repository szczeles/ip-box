import os
import typing

import click
from openpyxl.workbook import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

from context import pass_ip_box, IpBoxContext, Project
from model.excel import BaseSheet
from model.kpir import Kpir, KpirCsvReader
from model.classification import ClassifiedKpirRow


class Legacy:
    _headers = dict(
        projects=dict(
            title='Projekty',
            header={
                'A1': 'Projekt',
                'B1': 'Identyfikator KPWI',
                'C1': 'Rodzaj',
                'D1': 'Numer Interpretacji Indywidualnej',
                'E1': 'Opis Projektu',
                'F1': 'Data Rozpoczęcia',
                'G1': 'Data Zakończenia',
                'H1': 'Prace Stworzone',
                'I1': 'Pracownicy zaangażowany w projekt'
            }),
        sales_records=dict(
            title='Dokumenty księgowe',
            header={
                'A1:A2': 'Data zdarzenia',
                'B1:B2': 'Nr dowodu księgowego',
                'C1:C2': 'Numer w KPiR',
                'D1:D2': 'Opis zdarzenia gospodarczego',
                'E1:E2': 'Identyfikator KPWI',
                'F1:H1': 'Przychód',
                'F2': 'Zbycie KPWI',
                'G2': 'Pozostałe',
                'H2': 'SUMA',
                'I1:M1': 'Koszty działalności badawczo - rozwojowej',
                'I2': 'Kategoria A',
                'J2': 'Kategoria B',
                'K2': 'Kategoria C',
                'L2': 'Kategoria D',
                'M2': 'SUMA'
            }),
        results=dict(
            title='Wyniki',
            header={
                'A1': 'Identyfikator KPWI',
                'B1': 'Pole PIT IP',
                'C1': 'Wartość',
                'A2': 'Przychody z kwalifikowanych praw',
                'B2': '16',
                'A3': 'Koszty Kwalifikowane',
                'A4': 'Koszty pośrednie',
                'A5': 'Koszty uzyskania przychodu związane z kwalifikowanymi prawami',
                'B5': '17',
                'A6': 'Nexus niezaokrąglony',
                'A7': 'Nexus',
                'A8': 'Kwalifikowany dochód (dochody) z kwalifikowanych praw obliczony (obliczone) zgodnie z art. 30ca ust. 4 ustawy',
                'B8': '18 / 42',
                'A9': 'Dochód z kwalifikowanych praw niepodlegający opodatkowaniu stawką 5%, o której mowa w art. 30ca ust. 1 ustawy',
                'B9': '19 / 20 (jeżeli ujemny)',
                'A10': 'Podstawa opodatkowania',
                'B10': '31',
                'A11': 'Podstawa opodatkowania',
                'B11': '42',
                'A12': 'Podatek',
                'B12': '44',
                'A13': 'Podatek',
                'B13': '47',
            })
    )

    def __init__(self, kpir, projects):
        self._kpir = kpir
        self._projects = projects

    def generate(self, year, path, report_name):
        projects = self._get_projects(year)
        wb = Workbook()
        self._generate_projects(wb.active, projects)
        self._generate_records(wb.create_sheet(), year, projects)
        self._generate_results(wb.create_sheet(), year, projects)
        print(f"saving to {path}/{report_name}")
        wb.save(filename=f'{path}/{report_name}')

    def _get_projects(self, year):
        return [project for project in self._projects if project.start_date.year <= year <= project.end_date.year]

    def _generate_header(self, sheet, header):
        sheet.title = header['title']
        for (cell, value) in header['header'].items():
            if ':' in cell:
                self._set_header(sheet, cell.partition(":")[0], value)
                sheet.merge_cells(cell)
            else:
                self._set_header(sheet, cell, value)

    def _set_header(self, sheet, cell, value):
        sheet[cell] = value
        sheet[cell].font = Font(color='FF000000', bold=True)
        sheet[cell].fill = PatternFill("solid", fgColor=Color(indexed=22))

    def _generate_projects(self, sheet, projects):
        self._generate_header(sheet, self._headers['projects'])
        index = 2
        for project in projects:
            sheet[f'A{index}'] = project.name
            sheet[f'B{index}'] = project.kpwi_id
            sheet[f'C{index}'] = project.kind
            sheet[f'D{index}'] = project.interpretation_number
            sheet[f'E{index}'] = project.description
            sheet[f'F{index}'] = project.start_date.isoformat()
            sheet[f'G{index}'] = project.end_date.isoformat() if project.end_date else ""
            sheet[f'H{index}'] = ", ".join(project.results)
            sheet[f'I{index}'] = ", ".join(project.employees)
            index += 1

    def _generate_records(self, sheet, year, projects):
        self._generate_header(sheet, self._headers['sales_records'])
        index = 3
        for record in self._get_records(year, projects):
            sheet[f'A{index}'] = record.date
            sheet[f'B{index}'] = record.invoice_number
            sheet[f'C{index}'] = record.kpir_number
            sheet[f'D{index}'] = record.description
            sheet[f'E{index}'] = ", ".join(record.projects_kpwi_ids)
            sheet[f'F{index}'] = record.income_kpwi
            sheet[f'G{index}'] = record.income_other
            sheet[f'H{index}'] = f'=F{index}+G{index}'
            sheet[f'I{index}'] = record.cost_a
            sheet[f'J{index}'] = record.cost_b
            sheet[f'K{index}'] = record.cost_c
            sheet[f'L{index}'] = record.cost_d
            sheet[f'M{index}'] = f'=I{index}+J{index}+K{index}+L{index}'
            index += 1

    def _get_records(self, year, projects):
        for row in self._kpir.filter_year(year):
            related_projects = []
            for project in projects:
                if project.is_related(row):
                    related_projects.append(project)
            if related_projects:
                pass
                # yield KPWIRecord(row, related_projects)


    def _generate_results(self, sheet, year, projects):
        self._generate_header(sheet, self._headers['results'])
        total_incomes, total_costs = self._kpir.compute_totals(year)
        index = 4
        min_column = None
        max_column = None
        for project in projects:
            column = get_column_letter(index)
            self._set_header(sheet, f'{column}1', project.kpwi_id)
            sheet[f'{column}2'] = f"={self._sales_record_for_kpwi(column + '1', 'F')}"
            sheet[f'{column}3'] = f"={self._sales_record_for_kpwi(column + '1', 'M')}"
            sheet[f'{column}4'] = f"=({column}2/{total_incomes})*({total_costs}-C3)"
            sheet[f'{column}5'] = f'={column}3 + {column}4'
            sheet[f'{column}6'] = self._nexus_formula(column + '1')
            sheet[f'{column}7'] = f'=IF({column}6>1,1,{column}6)'
            sheet[f'{column}8'] = f'=({column}2-{column}5)*{column}7'
            sheet[f'{column}9'] = f'=({column}2-{column}5)*(1-{column}7)'
            sheet[f'{column}10'] = f'={column}8'
            sheet[f'{column}11'] = f'={column}8'
            sheet[f'{column}12'] = f'={column}8*0.05'
            sheet[f'{column}13'] = f'={column}8*0.05'
            if min_column is None:
                min_column = column
            max_column = column
            index += 1
        for row in [2, 3, 4, 5, 8, 9, 10, 11, 12, 13]:
            sheet[f'C{row}'] = f'=SUM({min_column}{row}:{max_column}{row})'

    def _nexus_formula(self, kpwi_cell):
        nominator = '+'.join([self._sales_record_for_kpwi(kpwi_cell, cost_column) for cost_column in ['I', 'J', 'K']])
        return f"=1.3*({nominator})/{self._sales_record_for_kpwi(kpwi_cell, 'M')}"

    def _sales_record_for_kpwi(self, kpwi_cell, column):
        sales_records_tab = self._headers['sales_records']['title']
        return f"SUMIF('{sales_records_tab}'!E:E,{kpwi_cell},'{sales_records_tab}'!{column}:{column})"


class ProjectsSummarySheet(BaseSheet):
    _title = 'Zestawienie projektów'
    _header = {
        'A1': 'Projekt',
        'B1': 'Identyfikator KPWI',
        'C1': 'Rodzaj',
        'D1': 'Numer Interpretacji Indywidualnej',
        'E1': 'Opis Projektu',
        'F1': 'Data Rozpoczęcia',
        'G1': 'Data Zakończenia',
        'H1': 'Prace Stworzone',
        'I1': 'Pracownicy zaangażowany w projekt'
    }

    def __init__(self, sheet, projects, **kwargs):
        super().__init__(sheet, self._title, **kwargs)
        self._projects = projects
        for idx, project in enumerate(projects):
            row = idx + 2
            for (letter, value) in (
                ('A', project.name),
                ('B', project.id),
                ('C', 'Autorskie prawo do programu komputerowego'),
                ('D', project.interpretation_number),
                ('E', project.description),
                ('F', project.start_date.isoformat()),
                ('G', project.end_date.isoformat() if project.end_date else ''),
                ('H', ', '.join([result.name for result in project.results])),
                ('I', project.employee)
            ):
                self[f'{letter}{row}'] = value


class SalesRecordsSummarySheet(BaseSheet):
    _title = 'Zestawienie ksiąg'
    _header = {
        'A1:A2': 'Data zdarzenia',
        'B1:B2': 'Nr dowodu księgowego',
        'C1:C2': 'Numer w KPiR',
        'D1:D2': 'Opis zdarzenia gospodarczego',
        'E1:E2': 'Identyfikator KPWI',
        'F1:H1': 'Przychód',
        'F2': 'Zbycie KPWI',
        'G2': 'Pozostałe',
        'H2': 'SUMA',
        'I1:M1': 'Koszty działalności badawczo - rozwojowej',
        'I2': 'Kategoria A',
        'J2': 'Kategoria B',
        'K2': 'Kategoria C',
        'L2': 'Kategoria D',
        'M2': 'SUMA'
    }

    def __init__(self, sheet, **kwargs):
        super().__init__(sheet, self._title, **kwargs)
        self._row_idx = 0

    def add(self, row: ClassifiedKpirRow):
        if row.is_kpiw:
            self._row_idx += 1
            idx = self._row_idx + 2
            self[f'A{idx}'] = row.date.isoformat()
            self[f'B{idx}'] = row.invoice_number
            self[f'C{idx}'] = row.number
            self[f'D{idx}'] = row.description
            self[f'E{idx}'] = ", ".join(row.projects_ids)
            self[f'F{idx}'] = row.income_kpwi
            self[f'G{idx}'] = row.income_other
            self[f'H{idx}'] = f'=F{idx}+G{idx}'
            self[f'I{idx}'] = row.cost_a
            self[f'J{idx}'] = row.cost_b
            self[f'K{idx}'] = row.cost_c
            self[f'L{idx}'] = row.cost_d
            self[f'M{idx}'] = f'=I{idx}+J{idx}+K{idx}+L{idx}'



class ResultsSheet(BaseSheet):
    _title = 'Wyniki'

    _header = {
        'A1': 'Identyfikator KPWI',
        'B1': 'Pole PIT IP',
        'C1': 'Wartość',
        'A2': 'Przychody z kwalifikowanych praw',
        'B2': '16',
        'A3': 'Koszty Kwalifikowane',
        'A4': 'Koszty pośrednie',
        'A5': 'Koszty uzyskania przychodu związane z kwalifikowanymi prawami',
        'B5': '17',
        'A6': 'Nexus niezaokrąglony',
        'A7': 'Nexus',
        'A8': 'Kwalifikowany dochód (dochody) z kwalifikowanych praw obliczony (obliczone) zgodnie z art. 30ca ust. 4 ustawy',
        'B8': '18 / 42',
        'A9': 'Dochód z kwalifikowanych praw niepodlegający opodatkowaniu stawką 5%, o której mowa w art. 30ca ust. 1 ustawy',
        'B9': '19 / 20 (jeżeli ujemny)',
        'A10': 'Podstawa opodatkowania',
        'B10': '31',
        'A11': 'Podstawa opodatkowania',
        'B11': '42',
        'A12': 'Podatek',
        'B12': '44',
        'A13': 'Podatek',
        'B13': '47',
    }

    def __init__(self, sheet, projects, **kwargs):
        super().__init__(sheet, self._title, **kwargs)
        self._projects = projects
        self._total_incomes = 0
        self._total_costs = 0

    def add(self, row: ClassifiedKpirRow):
        if row.is_income:
            self._total_incomes += row.income
        elif row.is_cost:
            self._total_costs += row.cost
        else:
            raise ValueError(f'strange row: {row}')

    def flush(self):
        index = 4
        min_column = None
        max_column = None
        for project in self._projects:
            column = get_column_letter(index)
            self.set_header(f'{column}1', project.id)
            self[f'{column}2'] = f"={self._sales_record_for_kpwi(column + '1', 'F')}"
            self[f'{column}3'] = f"={self._sales_record_for_kpwi(column + '1', 'M')}"
            self[f'{column}4'] = f"=({column}2/{self._total_incomes})*({self._total_costs}-C3)"
            self[f'{column}5'] = f'={column}3 + {column}4'
            self[f'{column}6'] = self._nexus_formula(column + '1')
            self[f'{column}7'] = f'=IF({column}6>1,1,{column}6)'
            self[f'{column}8'] = f'=({column}2-{column}5)*{column}7'
            self[f'{column}9'] = f'=({column}2-{column}5)*(1-{column}7)'
            self[f'{column}10'] = f'={column}8'
            self[f'{column}11'] = f'={column}8'
            self[f'{column}12'] = f'={column}8*0.05'
            self[f'{column}13'] = f'={column}8*0.05'
            if min_column is None:
                min_column = column
            max_column = column
            index += 1
        for row in [2, 3, 4, 5, 8, 9, 10, 11, 12, 13]:
            self[f'C{row}'] = f'=SUM({min_column}{row}:{max_column}{row})'

    def _nexus_formula(self, kpwi_cell):
        nominator = '+'.join([self._sales_record_for_kpwi(kpwi_cell, cost_column) for cost_column in ['I', 'J', 'K']])
        return f"=1.3*({nominator})/{self._sales_record_for_kpwi(kpwi_cell, 'M')}"

    def _sales_record_for_kpwi(self, kpwi_cell, column):
        sales_records_tab = SalesRecordsSummarySheet._title
        return f"SUMIF('{sales_records_tab}'!E:E,{kpwi_cell},'{sales_records_tab}'!{column}:{column})"


class SalesRecordsSummaryWriter:

    def __init__(self, projects: list[Project], year: int, path):
        self._projects = projects
        self._year = year
        self._path = path

    def __enter__(self):
        self._workbook = Workbook()
        self._projects_sheet = ProjectsSummarySheet(self._workbook.active, self._projects)
        self._sales_records_sheet = SalesRecordsSummarySheet(self._workbook.create_sheet())
        self._results_sheet = ResultsSheet(self._workbook.create_sheet(), self._projects)
        return self

    def write_classified_sales_record(self, row: ClassifiedKpirRow):
        self._sales_records_sheet.add(row)
        self._results_sheet.add(row)

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._results_sheet.flush()
        self._workbook.save(filename=self._path)


@click.option('--year', '-y', 'years',
              help="Year in format YYYY",
              type=int,
              required=True,
              multiple=True,
              prompt="Provide year in format YYYY", )
@click.option('--timesheet', '-t', 'timesheet_dir',
              help='Timesheet reports path',
              required=True,
              default='reports/timesheet',
              type=click.Path(exists=True, file_okay=False, dir_okay=True),
              prompt="Timesheet reports directory")
@click.option('--kpir', '-k', 'kpir_dir',
              help='KPiR directory path',
              required=True,
              default='reports/record',
              type=click.Path(exists=True, file_okay=False, dir_okay=True),
              prompt="KPiR reports directory")
@click.option('--output', '-o',
              help='Record output path',
              required=True,
              default='reports/record',
              type=click.Path(exists=False, file_okay=False, dir_okay=True),
              prompt="Output directory")
@click.command()
@pass_ip_box
def summary(ip_box: IpBoxContext, years, timesheet_dir, kpir_dir, output):
    os.makedirs(output, exist_ok=True)
    for year in years:
        projects = ip_box.get_active_projects(year)
        with (SalesRecordsSummaryWriter(projects, year,f'{output}/{year}.xlsx') as writer,
              KpirCsvReader(f'{kpir_dir}/{year}-classified.csv') as kpir_reader):
            for row in (ClassifiedKpirRow(row) for row in kpir_reader):
                writer.write_classified_sales_record(row)
