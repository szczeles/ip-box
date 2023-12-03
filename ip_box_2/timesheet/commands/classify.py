import click
import os
from pathlib import Path
import datetime
from context import pass_ip_box, IpBoxContext
from common import (MonthRange, DaysRange, ClassifiedTimeEntry, TimeEntry, IpClassificationDetailedWriter,
                    IpClassificationAggregator, CsvReader)


@click.option('--start-date', '-s', help="Start date in YYYY-mm format", required=True, default='2019-01', prompt="Provide start date in format YYYY-mm")
@click.option('--end-date', '-e', help="End date in YYYY-mm format", required=True, default=datetime.datetime.now().strftime("%Y-%m"), prompt="Provide end date in format YYYY-mm")
@click.option('--harvest-reports', '-r', help='Harvest reports path', required=True, default='reports/harvest', type=click.Path(exists=True, file_okay=False, dir_okay=True), prompt="Harvest reports directory")
@click.option('--output', '-o', help='Reports Output path', required=True, default='reports/timesheet', type=click.Path(exists=False, file_okay=False, dir_okay=True), prompt="Output directory")
@click.command()
@pass_ip_box
def classify(ip_box: IpBoxContext, start_date, end_date, harvest_reports, output):
    for (year, month) in MonthRange(start_date, end_date):
        reports_dir = f'{harvest_reports}/{year}/{month:02d}'
        for project in ip_box.get_active_projects(year, month):
            base_path = f'{output}/{year}/{month:02d}/{project.id}'
            os.makedirs(base_path, exist_ok=True)
            range_start = datetime.date(year, month, 1)
            with (IpClassificationDetailedWriter(f'{base_path}/detailed.csv') as writer,
                  IpClassificationAggregator(f'{base_path}/agg-daily.csv',
                                             DaysRange.till_end_of_month(range_start)) as days_aggregator,
                  IpClassificationAggregator(f'{base_path}/agg-monthly.csv',
                                            DaysRange.single(range_start), '%Y-%m') as month_aggregator):
                for harvest_selector in project.harvest_selectors:
                    report = Path(f'{reports_dir}/{harvest_selector.code}.csv')
                    if report.is_file():
                        with CsvReader(report) as reader:
                            for row in reader:
                                time_entry = TimeEntry.from_row(row)
                                is_ip = harvest_selector.classify(time_entry)
                                classified_time_entry = ClassifiedTimeEntry.from_time_entry(time_entry, is_ip)
                                writer.write(classified_time_entry)
                                days_aggregator.append(classified_time_entry)
                                month_aggregator.append(classified_time_entry)
