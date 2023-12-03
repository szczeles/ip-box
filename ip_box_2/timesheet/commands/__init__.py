import click
from harvest.api import HarvestAPI
from .classify import classify

@click.group()
def timesheet():
    pass


timesheet.add_command(classify)
