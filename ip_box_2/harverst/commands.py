import click
import os
from datetime import datetime
from .api import HarvestAPI
from .matchers import IPBoxMatcher
from .writers import MultiProjectReportWriter

@click.option('--account-id', help="Harvest Account ID", required=True, envvar='HARVEST_ACCOUNT_ID')
@click.option('--token', help="Harvest Token", required=True, envvar='HARVEST_TOKEN')
@click.option('--year', '-y', 'years', help="Year in YYYY format", multiple=True, required=True, type=click.IntRange(2019, datetime.now().year), envvar='IP_BOX_YEAR')
@click.option('--month', '-m', 'months', help='list of month in MM format', multiple=True, required=False, type=click.IntRange(min=1, max=12), default=(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12))
@click.option('--output', '-o', help='Reports Output path', required=True, default='reports', type=click.Path(exists=False, file_okay=False, dir_okay=True))
@click.option('--task-type', '-t', 'task_type', help='Time entry type considered as IP Box compatible', required=False, multiple=True, default=(),)
@click.command()
def fetch(account_id, token, years, months, output, task_type):
    api = HarvestAPI(account_id, token)

    ipbox_matcher = IPBoxMatcher(task_type)

    for year in years:
        for month in months:
            base_path = f'{output}/{year}/{month:02d}/harvest'
            os.makedirs(base_path, exist_ok=True)
            with MultiProjectReportWriter(base_path, ipbox_matcher) as writer:
                click.echo(f'processing: {year}-{month:02d}')
                for time_entry in api.iterate_time_entries(f'{year}-{month:02d}'):
                    writer.write(time_entry)
                writer.write_summary()


@click.group()
def harvest():
    pass


harvest.add_command(fetch)
