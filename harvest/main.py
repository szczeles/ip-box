#!/usr/bin/env python3

import argparse
import datetime
import calendar
from dataclasses import dataclass
from collections import defaultdict
import os
import csv

import requests

@dataclass
class TimeEntry:
    date: str
    project: str
    task: str
    notes: str
    hours: float


def matches_ipbox(time_entry: TimeEntry):
    if time_entry.task == 'Development':
        return True
    return False
    

class HarvestAPI:

    def __init__(self, account_id, token):
        self.account_id = account_id
        self.token = token

    def call_api(self, url, params):
        return requests.get(f'https://api.harvestapp.com/v2{url}',
            params=params,
            headers={
                'Authorization':  f'Bearer {self.token}',
                'Harvest-Account-Id': f'{self.account_id}',
                'User-Agent': 'IP Box time entries scraper (https://github.com/szczeles/ip-box)'
            }).json()

    def iterate_time_entries(self, month):
        start_date = f'{month}-01'
        end_date = f'{month}-{calendar.monthrange(int(month[:4]), int(month[5:7]))[1]}'

        entries = []
        page = 1
        while page is not None:
            response = self.call_api('/time_entries', 
                {'from': start_date, 'to': end_date, 'page': page})
            entries.extend(response['time_entries'])
            page = response['next_page']

        for entry in sorted(entries, key=lambda entry: entry['spent_date']):
            yield TimeEntry(
                date=entry['spent_date'],
                project=entry['project']['name'],
                task=entry['task']['name'],
                notes=entry['notes'],
                hours=entry['hours']
            )
        
class ReportWriter:
    def __init__(self, filepath):
        self.file = open(filepath, 'w+')
        self.writer = csv.writer(self.file)
        self.writer.writerow([
            'date',
            'task',
            'notes',
            'hours',
            'ip_box_applicable'
        ])

    def write(self, entry):
        self.writer.writerow([
            entry.date,
            entry.task,
            entry.notes,
            entry.hours,
            matches_ipbox(entry)
        ])

    def close(self):
        self.file.close()
        

class MultiProjectReportWriter:
    def __init__(self, base_path):
        self.base_path = base_path
        self.writers = {}
        self.ipbox_hours = defaultdict(float)

    def get_path(self, project):
        return f'{self.base_path}/{project}.csv'

    def write(self, entry):
        if not entry.project in self.writers:
            self.writers[entry.project] = ReportWriter(self.get_path(entry.project))
        self.writers[entry.project].write(entry)
        if matches_ipbox(entry):
            self.ipbox_hours[entry.project] += entry.hours

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        assert type == None
        for project in self.writers.keys():
            self.writers[project].close()
        return True

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--account', help='Account ID', required=True)
    parser.add_argument('--token', help='Personal token', required=True)
    parser.add_argument('--month', help='YYYY-MM format', required=True)
    parser.add_argument('--path', help='Reports path', required=True)
    args = parser.parse_args()

    api = HarvestAPI(args.account, args.token)
    os.makedirs(f'{args.path}/{args.month}', exist_ok=True)
    with MultiProjectReportWriter(f'{args.path}/{args.month}') as writer:
        for time_entry in api.iterate_time_entries(args.month):
            writer.write(time_entry)

        print(writer.ipbox_hours)
