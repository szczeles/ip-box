from dataclasses import dataclass

import click


@dataclass
class TimeEntry:
    date: str
    project: str
    task: str
    notes: str
    hours: float


@dataclass
class ReportSummary:
    entries: int
    total_hours: float
    ip_box_hours: float

    def __str__(self):
        return (f'total entries: {self.entries}\n'
                f'ip box hours:  {self.ip_box_hours}\n'
                f'total hours:   {self.total_hours}\n'
                f'percentage:    {self.ip_box_hours/self.total_hours:.1%}')

    def __add__(self, other):
        return ReportSummary(entries=self.entries + other.entries, total_hours=self.total_hours + other.total_hours, ip_box_hours=self.ip_box_hours + other.ip_box_hours)
