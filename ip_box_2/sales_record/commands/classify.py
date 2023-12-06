import os
import typing

import click

from context import pass_ip_box, IpBoxContext
from model.kpir import Kpir, KpirCsvReader, KpirRow
from model.classification import KpirClassificationCsvWriter


@click.option('--year', '-y', 'years',
              help="Year in format YYYY",
              type=int,
              required=True,
              multiple=True,
              prompt="Provide year in format YYYY", )
@click.option('--kpir', '-k', 'kpir_dir',
              help='KPiR directory path',
              required=True,
              default='reports/kpir',
              type=click.Path(exists=True, file_okay=False, dir_okay=True),
              prompt="KPiR reports directory")
@click.option('--output', '-o',
              help='Record classification output path',
              required=True,
              default='reports/record',
              type=click.Path(exists=False, file_okay=False, dir_okay=True),
              prompt="Record classification output directory")
@click.command()
@pass_ip_box
def classify(ip_box: IpBoxContext, years: typing.List[int], kpir_dir: click.Path, output: click.Path):
    os.makedirs(output, exist_ok=True)
    for year in years:
        with (KpirCsvReader(f'{kpir_dir}/{year}.csv') as kpir_reader,
              KpirClassificationCsvWriter(f'{output}/{year}-classified.csv') as writer):
            for row in Kpir([KpirRow(row) for row in kpir_reader], year=year):
                result = ip_box.classify_sales_record(row)
                writer.write_classification_result(row, result)
