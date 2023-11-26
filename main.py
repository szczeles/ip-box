import argparse
import os

from ip_box.sales_record import KPiRCsvReader, ProjectsDescriptor, SalesRecordsReport
# from ip_box.timesheet import IPBoxTimeSheet
# from ip_box.harvest import HarvestAPI, IPBoxMatcher, MultiProjectReportWriter


def main():
    args_parser = argparse.ArgumentParser()

    args_parser.add_argument('--year', help='report year', type=int, required=True)
    args_parser.add_argument('--kpir', help='KPiR CSV file path', required=True)
    args_parser.add_argument('--projects', help="path to projects descriptor file", required=True)
    args_parser.add_argument('--path', help="Reports path", required=False, default="out")
    args_parser.add_argument('--sales-records-report', help="Sales Records report file name", required=False,
                             default='Ewidencja-IP-Box.xlsx')

    ipbox_args = args_parser.parse_args()

    report_path = f'{ipbox_args.path}/{ipbox_args.year}'

    os.makedirs(report_path, exist_ok=True)

    with KPiRCsvReader(ipbox_args.kpir, ipbox_args.year) as reader:
        with ProjectsDescriptor(ipbox_args.projects) as descriptor:
            SalesRecordsReport(reader.read(), descriptor.load()).\
                generate(ipbox_args.year, report_path, ipbox_args.sales_records_report)


if __name__ == '__main__':
    main()


# if __name__ == '__main__':
#     parser = argparse.ArgumentParser()
#     parser.add_argument('--year', help='YYYY format', required=True)
#     parser.add_argument('--path', help='Reports path', required=True)
#     parser.add_argument('--projects', help='IP Box compatible projects name', required=False, nargs='+')
#     parser.add_argument('--prefixes', help='Reports files prefixes', required=False, nargs='*',
#                         default=['Harvest ', 'Upwork '])
#     args = parser.parse_args()
#
#     os.makedirs(f'{args.path}/{args.year}', exist_ok=True)
#
#     IPBoxTimeSheet(args.path, int(args.year), args.projects, args.prefixes).generate()
#
#
# if __name__ == '__main__':
#     parser = argparse.ArgumentParser()
#     parser.add_argument('--account', help='Account ID', required=True)
#     parser.add_argument('--token', help='Personal token', required=True)
#     parser.add_argument('--year', help="YYYY format", required=True)
#     parser.add_argument('--months', help='list of months in MM format', required=False,
#                         nargs='+', default=['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12'])
#     parser.add_argument('--path', help='Reports path', required=True)
#     parser.add_argument('--types', help='Time entry types considered as IP Box compatible',
#                         nargs='+', required=False, default=['Development'])
#     args = parser.parse_args()
#
#     api = HarvestAPI(args.account, args.token)
#
#     ipbox_matcher = IPBoxMatcher(args.types)
#
#     for month in args.months:
#         os.makedirs(f'{args.path}/{args.year}/{month}', exist_ok=True)
#         with MultiProjectReportWriter(f'{args.path}/{args.year}/{month}', ipbox_matcher) as writer:
#             for time_entry in api.iterate_time_entries(f'{args.year}-{month}'):
#                 writer.write(time_entry)
#             print(writer.ipbox_hours)