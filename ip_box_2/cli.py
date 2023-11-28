import click

from harverst import commands as harvest_cli
from timesheet import commands as timesheet_cli


@click.group()
def entry_point():
    pass


entry_point.add_command(harvest_cli.harvest)
entry_point.add_command(timesheet_cli.timesheet)

if __name__ == '__main__':
    entry_point(auto_envvar_prefix='IP_BOX')
