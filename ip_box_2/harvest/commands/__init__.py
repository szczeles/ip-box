import click
from harvest.api import HarvestAPI
from .fetch import fetch

@click.group()
@click.option('--account-id', help="Harvest Account ID", required=True, envvar='HARVEST_ACCOUNT_ID')
@click.option('--token', help="Harvest Token", required=True, envvar='HARVEST_TOKEN')
@click.pass_context
def harvest(ctx, account_id, token):
    ctx.obj = HarvestAPI(account_id, token)


harvest.add_command(fetch)
