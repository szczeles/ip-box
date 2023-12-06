import datetime

import click
from common import TimeEntry
from model.kpir import KpirRow
from model.classification import SalesRecordsClassifier, KpiwClassificationResult


class KpiwHarvestProjectSelector:

    def __init__(self, config):
        self._config = config

    @property
    def code(self) -> str:
        return self._config['code']

    @property
    def types(self) -> list[str]:
        return self._config['types']

    def classify(self, entry: TimeEntry) -> bool:
        return entry.task_type in self.types


class ProjectResult:

    def __init__(self, config):
        self._config = config

    @property
    def name(self) -> str:
        return self._config['name']

    @property
    def start_date(self):
        return self._config['startDate']

    @property
    def end_date(self):
        return self._config['endDate'] if 'endDate' in self._config else None


class Project:

    def __init__(self, config):
        self._config = config

    @property
    def id(self) -> str:
        return self._config['id']

    @property
    def name(self) -> str:
        return self._config['name']

    @property
    def description(self) -> str:
        return self._config['description']

    @property
    def start_date(self) -> datetime:
        return self._config['startDate']

    @property
    def end_date(self) -> datetime or None:
        if 'endDate' in self._config:
            return self._config['endDate']

    @property
    def employee(self) -> str:
        return self._config['employee']

    @property
    def interpretation_number(self) -> str:
        return self._config['interpretationNumber']

    @property
    def results(self) -> list[ProjectResult]:
        return [ProjectResult(config) for config in self._config['results']]

    @property
    def harvest_selectors(self) -> list[KpiwHarvestProjectSelector]:
        return [KpiwHarvestProjectSelector(config) for config in self._config.get('harvest', [])]

    def has_started(self, year: int, month: int = None) -> bool:
        return datetime.date(self.start_date.year, self.start_date.month, 1) <= datetime.date(year, month or 12, 1)

    def has_finished(self, year: int, month: int = None) -> bool:
        return (self.end_date is not None and
                datetime.date(self.end_date.year, self.end_date.month, 1) < datetime.date(year, month or 1, 1))

    def is_active(self, year: int, month: int = None) -> bool:
        return self.has_started(year, month) and not self.has_finished(year, month)


class IpBoxContext:

    def __init__(self, config):
        self._config = config
        self._projects = [Project(project) for project in config['projects']]
        self._wages = [(wage['from'], wage['wage']) for wage in config['wages']]
        self._record_classifier = SalesRecordsClassifier(config['ip_box_records'])

    @property
    def projects(self) -> list[Project]:
        return self._projects

    @property
    def ignore_codes(self) -> list[str]:
        return self._config['harvest'].get('ignore', [])

    def get_active_projects(self, year: int, month: int = None) -> list[Project]:
        return [project for project in self.projects if project.is_active(year, month)]

    def get_wage(self, date: datetime.date):
        result = None
        for (start_date, wage) in self._wages:
            if date >= start_date:
                result = wage
            else:
                break
        if not result:
            raise Exception('date is before employment date')
        return result

    def classify_sales_record(self, row: KpirRow) -> KpiwClassificationResult:
        result = self._record_classifier.matches(row)
        return KpiwClassificationResult(self.get_active_projects(row.date.year, row.date.month), result)


pass_ip_box = click.make_pass_decorator(IpBoxContext)