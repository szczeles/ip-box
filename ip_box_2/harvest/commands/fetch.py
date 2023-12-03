import click
import os
from datetime import datetime
from harvest.writers import MultiProjectReportWriter
from .context import pass_harvest_api
from common import MonthRange


@click.option('--start-date', '-s', help="Start date in YYYY-mm format", required=True, default='2019-01', prompt="Provide start date in format YYYY-mm")
@click.option('--end-date', '-e', help="End date in YYYY-mm format", required=True, default=datetime.now().strftime("%Y-%m"), prompt="Provide end date in format YYYY-mm")
@click.option('--output', '-o', help='Reports Output path', required=True, default='reports/harvest', type=click.Path(exists=False, file_okay=False, dir_okay=True), prompt="Output directory")
@click.command()
@pass_harvest_api
def fetch(api, start_date, end_date, output):
    click.echo(f"fetching harvest entries for range: [{start_date}; {end_date}]")
    for (year, month) in MonthRange(start_date, end_date):
        base_path = f'{output}/{year}/{month:02d}'
        os.makedirs(base_path, exist_ok=True)
        with MultiProjectReportWriter(base_path) as writer:
            click.echo(f'fetching: {year}-{month:02d}')
            for time_entry in api.iterate_time_entries(f'{year}-{month:02d}'):
                writer.write(time_entry)
    else:
        click.echo('fetching completed')
