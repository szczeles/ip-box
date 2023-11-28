import csv


from .model import TimeEntry


class ReportWriter:

    def __init__(self, filepath):
        self.file = open(filepath, 'w+')
        self.writer = csv.writer(self.file)
        self.writer.writerow([
            'date',
            'task',
            'notes',
            'hours',
        ])
        self._entries_count = 0
        self._ip_box_hours = 0
        self._total_hours = 0

    def write(self, entry: TimeEntry):
        self.writer.writerow([
            entry.date,
            entry.task,
            entry.notes,
            entry.hours,
        ])

    def close(self):
        self.file.close()


class MultiProjectReportWriter:

    def __init__(self, base_path):
        self.base_path = base_path
        self.writers = {}

    def get_path(self, project, extension='csv'):
        return f'{self.base_path}/{project.replace("/", "-")}.{extension}'

    def write(self, entry):
        if not entry.project in self.writers:
            self.writers[entry.project] = ReportWriter(self.get_path(entry.project))
        self.writers[entry.project].write(entry)

    def write_summary(self):
        for (project, summary) in self.get_summary().items():
            with open(self.get_path(project, 'txt'), 'w') as f:
                f.write(str(summary))

    def get_summary(self):
        return {project:writter.get_summary() for (project, writter) in self.writers.items() if project}

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        for project in self.writers.keys():
            self.writers[project].close()
        return True
