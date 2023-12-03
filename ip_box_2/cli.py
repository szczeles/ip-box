import os

import click
from yaml import load
try:
    from yaml import CLoader as Loader
except ImportError:
    from yaml import Loader

from harvest.commands import harvest
from timesheet.commands import timesheet
from sales_record.commands import sales_record
from context import IpBoxContext


@click.group(context_settings={'auto_envvar_prefix': 'IP_BOX'})
@click.option('--config', default='config.yaml', type=click.Path())
@click.pass_context
def entry_point(ctx, config):
    if os.path.exists(config):
        with open(config, 'r') as f:
            config = load(f.read(), Loader=Loader)
        ctx.default_map = config
        ctx.obj = IpBoxContext(config)


entry_point.add_command(harvest)
entry_point.add_command(timesheet)
entry_point.add_command(sales_record)


if __name__ == '__main__':
    entry_point()
