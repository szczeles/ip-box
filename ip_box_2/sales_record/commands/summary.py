import os
import click
from openpyxl.workbook import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

from context import pass_ip_box, IpBoxContext, Project
from model.excel import BaseSheet


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
        super().__init__(sheet, 'Zestawienie projektów', **kwargs)
        self._projects = projects
        for cell, title in self._header.items():
            self.set_header(cell, title)
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

    def __init__(self, sheet, **kwargs):
        super().__init__(sheet, 'Zestawienie ksiąg', **kwargs)


class ResultsSheet(BaseSheet):

    def __init__(self, sheet, **kwargs):
        super().__init__(sheet, 'Wyniki', **kwargs)


class SalesRecordsSummaryWriter:

    def __init__(self, projects: list[Project], year: int, path):
        self._projects = projects
        self._year = year
        self._path = path

    def __enter__(self):
        self._workbook = Workbook()
        self._projects_sheet = ProjectsSummarySheet(self._workbook.active, self._projects)
        self._slaes_records_sheet = SalesRecordsSummarySheet(self._workbook.create_sheet())
        self._results_sheet = ResultsSheet(self._workbook.create_sheet())
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
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
@click.option('--output', '-o',
              help='Record output path',
              required=True,
              default='reports/record',
              type=click.Path(exists=False, file_okay=False, dir_okay=True),
              prompt="Output directory")
@click.command()
@pass_ip_box
def summary(ip_box: IpBoxContext, years, timesheet_dir, output):
    os.makedirs(output, exist_ok=True)
    for year in years:
        projects = ip_box.get_active_projects(year)
        with SalesRecordsSummaryWriter(projects, year,f'{output}/{year}.xlsx') as writer:
            pass

