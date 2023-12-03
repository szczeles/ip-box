import click
from harvest.api import HarvestAPI

pass_harvest_api = click.make_pass_decorator(HarvestAPI)
