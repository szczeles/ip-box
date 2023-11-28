from .model import TimeEntry


class IPBoxMatcher:

    def __init__(self, task_types):
        self._task_types = set(task_types)

    def __call__(self, time_entry: TimeEntry):
        return time_entry.task in self._task_types
