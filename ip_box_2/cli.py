import click

from harvest.commands import harvest
from timesheet.commands import timesheet


@click.group()
def entry_point():
    pass


entry_point.add_command(harvest)
entry_point.add_command(timesheet)

if __name__ == '__main__':
    entry_point(auto_envvar_prefix='IP_BOX')
