#!/usr/bin/env python3

import argparse
import datetime
import calendar
from dataclasses import dataclass
import os
import csv

import upwork
from config import map_project

@dataclass
class Invoice:
    date: str
    id: str
    client_id: str
    project: str
    subtype: str
    notes: str
    amount: float

    def date_map(self, date):
        m, d, y = date.split('/')
        return f'{y}-{m}-{d}'

    def get_billed_period(self):
        *_, start, end = self.notes.split(' - ')
        return self.date_map(start), self.date_map(end)

@dataclass
class TimeEntry:
    date: str
    invoice: Invoice
    hours: float
    notes: str
    task: str


class UpworkTimeEntriesProvider:

    def __init__(self, client):
        self.client = client
        me = client.hr.get_user_me()
        self.user_id = me['reference']
        self.user_name = me['id']

    def get_invoices(self, start_date, end_date):
        query = f"SELECT date, reference,buyer_company__id,assignment_name,buyer_company_name,subtype,description,amount WHERE date >= '{start_date}' AND date <= '{end_date}' AND type = 'APInvoice'"

        results = self.client.finreport.get_provider_earnings(self.user_id, query)
        for row in sorted(results['table']['rows'], key=lambda r: r['c'][0]['v']):
            assignment = row['c'][3]['v']
            client = row['c'][4]['v']
            yield Invoice(
                date=f"{row['c'][0]['v'][:4]}-{row['c'][0]['v'][4:6]}-{row['c'][0]['v'][6:8]}",
                id=row['c'][1]['v'],
                client_id=row['c'][2]['v'],
                project=map_project(assignment, client),
                subtype=row['c'][5]['v'],
                notes=row['c'][6]['v'],
                amount=float(row['c'][7]['v'])
            )


    def iterate_time_entries(self, month):
        start_date = f'{month}-01'
        end_date = f'{month}-{calendar.monthrange(int(month[:4]), int(month[5:7]))[1]}'

        for invoice in self.get_invoices(start_date, end_date):
            if invoice.subtype in ('Fixed Price', 'Miscellaneous'):
                print("Fixed price invoice ", invoice)
                continue

            yield from self.get_time_entries(invoice)


    def get_time_entries(self, invoice):
        start_date, end_date = invoice.get_billed_period()
        query = f"SELECT worked_on, hours, task, memo WHERE company_id = '{invoice.client_id}' and worked_on >= '{start_date}' AND worked_on <= '{end_date}'"
        results = client.timereport.get_provider_report(self.user_name, query)
        for row in sorted(results['table']['rows'], key=lambda r: r['c'][0]['v']):
            yield TimeEntry(
                date=f"{row['c'][0]['v'][:4]}-{row['c'][0]['v'][4:6]}-{row['c'][0]['v'][6:8]}",
                invoice=invoice,
                hours=float(row['c'][1]['v']),
                task=row['c'][2]['v'],
                notes=row['c'][3]['v']
            )
        
class ReportWriter:
    def __init__(self, filepath):
        self.file = open(filepath, 'w+')
        self.writer = csv.writer(self.file)
        self.writer.writerow([
            'date',
            'notes',
            'hours',
            'invoice date',
            'invoice id',
            'invoice notes',
            'invoice amount'
        ])

    def write(self, entry):
        self.writer.writerow([
            entry.date,
            entry.notes,
            entry.hours,
            entry.invoice.date,
            entry.invoice.id,
            entry.invoice.notes,
            entry.invoice.amount,
        ])

    def close(self):
        self.file.close()
        

class MultiProjectReportWriter:
    def __init__(self, base_path):
        self.base_path = base_path
        self.writers = {}

    def get_path(self, project):
        return f'{self.base_path}/Upwork {project}.csv'

    def write(self, entry):
        if not entry.invoice.project in self.writers:
            self.writers[entry.invoice.project] = ReportWriter(self.get_path(entry.invoice.project))
        self.writers[entry.invoice.project].write(entry)

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        assert type == None
        for project in self.writers.keys():
            self.writers[project].close()
        return True

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--api-key', help='API Key', required=True)
    parser.add_argument('--api-secret', help='API Secret', required=True)
    parser.add_argument('--month', help='YYYY-MM format', required=True)
    parser.add_argument('--path', help='Reports path', required=True)
    parser.add_argument('--access-token', help='Access token', required=False)
    parser.add_argument('--access-token-secret', help='Access token secret', required=False)
    args = parser.parse_args()

    if args.access_token is not None and args.access_token_secret is not None:
        client = upwork.Client(args.api_key, args.api_secret,
            oauth_access_token=args.access_token,
            oauth_access_token_secret=args.access_token_secret)
    else:
        client = upwork.Client(args.api_key, args.api_secret)
        verifier = input('Please enter the verification code you get '
                         'following this link:\n{0}\n\n> '.format( client.auth.get_authorize_url()))
        access_token, access_token_secret = client.auth.get_access_token(verifier)
        print(f"Generated tokens: {access_token}, {access_token_secret}")
        client = upwork.Client(args.api_key, args.api_secret,
            oauth_access_token=access_token,
            oauth_access_token_secret=access_token_secret)

    time_entries_provider = UpworkTimeEntriesProvider(client)
    os.makedirs(f'{args.path}/{args.month}', exist_ok=True)
    with MultiProjectReportWriter(f'{args.path}/{args.month}') as writer:
        for time_entry in time_entries_provider.iterate_time_entries(args.month):
            writer.write(time_entry)
