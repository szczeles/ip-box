import calendar
import requests
from .model import TimeEntry


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

        entries = []
        page = 1
        while page is not None:
            response = self.call_api('/time_entries', {'from': start_date, 'to': end_date, 'page': page})
            entries.extend(response['time_entries'])
            page = response['next_page']

        for entry in sorted(entries, key=lambda elem: elem['spent_date']):
            yield TimeEntry(
                date=entry['spent_date'],
                project=entry['project']['name'],
                task=entry['task']['name'],
                notes=entry['notes'],
                hours=entry['hours']
            )
