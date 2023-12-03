import click
from .classify import classify
from .summary import summary


@click.group()
def sales_record():
    pass


sales_record.add_command(classify)
sales_record.add_command(summary)
