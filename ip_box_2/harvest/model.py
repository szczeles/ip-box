from dataclasses import dataclass


@dataclass
class TimeEntry:
    date: str
    project: str
    task: str
    notes: str
    hours: float
