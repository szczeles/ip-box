import click
from .classify import classify
from .summary import summary


@click.group()
def timesheet():
    pass


timesheet.add_command(classify)
timesheet.add_command(summary)
