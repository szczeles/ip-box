#!/usr/bin/env python3

import argparse
import calendar
from dataclasses import dataclass
from collections import defaultdict
import os
import csv
import requests
import re


@dataclass
class TimeEntry:
    date: str
    project: str
    task: str
    notes: str
    hours: float


class IPBoxMatcher:

    def __init__(self, task_types):
        self._task_types = set(task_types)

    def __call__(self, time_entry: TimeEntry):
        return time_entry.task in self._task_types


class HarvestAPI:

    def __init__(self, account_id, token):
        self.account_id = account_id
        self.token = token

    def call_api(self, url, params):
        return requests.get(f'https://api.harvestapp.com/v2{url}',
                            params=params,
                            headers={
                                'Authorization': f'Bearer {self.token}',
                                'Harvest-Account-Id': f'{self.account_id}',
                                'User-Agent': 'IP Box time entries scraper (https://github.com/szczeles/ip-box)'
                            }).json()

    def iterate_time_entries(self, month):
        start_date = f'{month}-01'
        end_date = f'{month}-{calendar.monthrange(int(month[:4]), int(month[5:7]))[1]}'
        allowed_chars_regex = "[^A-Za-z0-9- -]"

        entries = []
        page = 1
        while page is not None:
            response = self.call_api('/time_entries', {'from': start_date, 'to': end_date, 'page': page})
            entries.extend(response['time_entries'])
            page = response['next_page']

        for entry in sorted(entries, key=lambda elem: elem['spent_date']):
            yield TimeEntry(
                date=entry['spent_date'],
                project=re.sub(allowed_chars_regex, "", entry['project']['name']),
                task=entry['task']['name'],
                notes=entry['notes'],
                hours=entry['hours']
            )


class ReportWriter:
    def __init__(self, filepath, ipbox_matcher):
        self.file = open(filepath, 'w+')
        self.writer = csv.writer(self.file)
        self.writer.writerow([
            'date',
            'task',
            'notes',
            'hours',
            'ip_box_applicable'
        ])
        self._ipbox_matcher = ipbox_matcher

    def write(self, entry):
        self.writer.writerow([
            entry.date,
            entry.task,
            entry.notes,
            entry.hours,
            self._ipbox_matcher(entry)
        ])

    def close(self):
        self.file.close()


class MultiProjectReportWriter:
    def __init__(self, base_path, ipbox_matcher):
        self.base_path = base_path
        self.writers = {}
        self.ipbox_hours = defaultdict(float)
        self.total_hours = 0
        self._ipbox_matcher = ipbox_matcher

    def get_path(self, project):
        return f'{self.base_path}/Harvest {project}.csv'

    def write(self, entry):
        if not entry.project in self.writers:
            self.writers[entry.project] = ReportWriter(self.get_path(entry.project), self._ipbox_matcher)
        self.writers[entry.project].write(entry)
        self.total_hours += entry.hours
        if self._ipbox_matcher(entry):
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
    parser.add_argument('--types', help='Time entry types considered as IP Box compatible',
                        nargs='+', required=False, default=['Development'])
    args = parser.parse_args()

    api = HarvestAPI(args.account, args.token)
    os.makedirs(f'{args.path}/{args.month}', exist_ok=True)
    ipbox_matcher = IPBoxMatcher(args.types)
    with MultiProjectReportWriter(f'{args.path}/{args.month}', ipbox_matcher) as writer:
        for time_entry in api.iterate_time_entries(args.month):
            writer.write(time_entry)

        print(writer.ipbox_hours)
        print("Total hours", "{:0.2f}".format(writer.total_hours))
