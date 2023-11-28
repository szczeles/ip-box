import click
import os
from datetime import datetime
from .api import HarvestAPI
from .writers import MultiProjectReportWriter

class MonthRange:

    def __init__(self, start, end):
        self._start = MonthRange.parse(start)
        self._end = MonthRange.parse(end)

    def __iter__(self):
        months = (self._start.year - self._end.year) * 12 + self._start.month - self._end.month
        start_month = self._start.month
        end_month = self._end.month if self._start.year == self._end.year else 12
        for year in range(self._start.year, self._end.year + 1):
            for month in range(start_month, end_month + 1):
                yield (year, month)
            start_month = 1
            end_month = self._end.month if year == self._end.year - 1 else 12

    @staticmethod
    def parse(date_str) -> datetime.date:
        return datetime.strptime(date_str, "%Y-%m").date()


@click.option('--account-id', help="Harvest Account ID", required=True, envvar='HARVEST_ACCOUNT_ID')
@click.option('--token', help="Harvest Token", required=True, envvar='HARVEST_TOKEN')
@click.option('--start-date', '-s', help="Start date in YYYY-mm format", required=True, default='2019-01', prompt="Provide start date in format YYYY-mm")
@click.option('--end-date', '-e', help="End date in YYYY-mm format", required=True, default=datetime.now().strftime("%Y-%m"), prompt="Provide end date in format YYYY-mm")
@click.option('--output', '-o', help='Reports Output path', required=True, default='reports/harvest', type=click.Path(exists=False, file_okay=False, dir_okay=True), prompt="Output directory")
@click.command()
def fetch(account_id, token, start_date, end_date, output):
    click.echo(f"fetching harvest entries for range: [{start_date}; {end_date}]")
    api = HarvestAPI(account_id, token)
    for (year, month) in MonthRange(start_date, end_date):
        base_path = f'{output}/{year}/{month:02d}'
        os.makedirs(base_path, exist_ok=True)
        with MultiProjectReportWriter(base_path) as writer:
            click.echo(f'fetching: {year}-{month:02d}')
            for time_entry in api.iterate_time_entries(f'{year}-{month:02d}'):
                writer.write(time_entry)
    else:
        click.echo('fetching completed')


@click.group()
def harvest():
    pass


harvest.add_command(fetch)
