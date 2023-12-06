import abc
import decimal
import typing
from model.kpir import KpirRow
from common import CsvWriter


class ClassificationResult:

    def __init__(self, matches: bool = False, category: str = None):
        if not matches and category:
            raise ValueError('cannot set cattegory when there is no match')
        if category:
            if category not in ('A', 'B', 'C', 'D'):
                raise ValueError(f'unknown category: {category}, please provide one of (A, B, C, D)')
        self._matches = matches
        self._category = category

    def __bool__(self):
        return self._matches

    def __or__(self, other):
        cat = self._merge_categories(self._category, other.category)
        if self.matches or other.matches:
            return ClassificationResult(True, cat)
        return ClassificationResult()

    def __and__(self, other):
        cat = self._merge_categories(self._category, other.category)
        if self.matches and other.matches:
            return ClassificationResult(True, cat)
        return ClassificationResult()

    @property
    def category(self):
        return self._category

    @property
    def matches(self):
        return self._matches

    @staticmethod
    def _merge_categories(cat1, cat2):
        if cat1 == cat2:
            return cat1
        if cat1 is None:
            return cat2
        elif cat2 is None:
            return cat1
        raise ValueError(f'categories mismatch: [{cat1}] vs [{cat2}]')


class KpiwClassificationResult:

    def __init__(self, projects, classification: ClassificationResult):
        self._projects = projects
        self._classification = classification

    def __bool__(self):
        return bool(self._classification) and bool(self._projects)

    @property
    def projects(self):
        return self._projects if self else None

    @property
    def category(self):
        return self._classification.category if self else None


class ClassificationCondition:

    @abc.abstractmethod
    def matches(self, row: KpirRow) -> ClassificationResult:
        return ClassificationResult()


class AnyOfCondition(ClassificationCondition):

    def __init__(self, conditions: typing.Iterable[ClassificationCondition]):
        self._conditions = conditions

    def matches(self, row: KpirRow) -> ClassificationResult:
        result = None
        for condition in self._conditions:
            current = condition.matches(row)
            if result is not None:
                result = result or current
            else:
                result = current
        return result or ClassificationResult()


class AllOfCondition(ClassificationCondition):

    def __init__(self, conditions: typing.Iterable[ClassificationCondition]):
        self._conditions = conditions

    def matches(self, row: KpirRow) -> ClassificationResult:
        result = None
        for condition in self._conditions:
            current = condition.matches(row)
            if not current:
                return ClassificationResult()
            if result is not None:
                result = result and current
            else:
                result = current
        return result or ClassificationResult()


class SalesRecordTypeCondition(ClassificationCondition):

    def __init__(self, income=None, cost=None):
        self._income = income
        self._cost = cost

    def matches(self, row: KpirRow) -> ClassificationResult:
        if row.is_income:
            return self._income and self._income.matches(row)
        return self._cost and self._cost.matches(row)


class PropertyMatcher:

    def __init__(self, config, extractor):
        self._config = config
        self._extractor = extractor
        self._matchers = []
        if 'startsWith' in config:
            starts_with = config.pop('startsWith')
            self._matchers.append(lambda val: val.startswith(starts_with))
        if 'contains' in config:
            contains = config.pop('contains')
            self._matchers.append(lambda val: contains in val)
        if 'equals' in config:
            equals = config.pop('equals')
            self._matchers.append(equals.__eq__)
        if config:
            raise ValueError(f'unsupported property matcher: {config.keys()}')

    def matches(self, row) -> bool:
        value = self._extractor(row)
        return self._matchers and all(matcher(value) for matcher in self._matchers)


class ConfigurableCondition(ClassificationCondition):

    def __init__(self, config):
        self._config = config
        self._conditions = []
        self._category = config.pop('category', None)
        if 'description' in config:
            description = config.pop('description')
            self._conditions.append(PropertyMatcher(description, lambda row: row.description))
        if 'companyName' in config:
            company_name = config.pop('companyName')
            self._conditions.append(PropertyMatcher(company_name, lambda row: row.company_name))
        if config:
            raise ValueError(f'unsupported keys: {config.keys()}')


    def matches(self, row: KpirRow) -> ClassificationResult:
        matches = all(condition.matches(row) for condition in self._conditions)
        if matches:
            return ClassificationResult(True, self.category if row.is_cost else None)
        return ClassificationResult()

    @property
    def category(self):
        if self._category:
            return self._category
        raise ValueError('missing category key in configuration')


class ConfigurableCompositeCondition(AnyOfCondition):

    def __init__(self, config):
        super().__init__([ConfigurableCondition(c) for c in config])
        self._config = config


class SalesRecordsClassifier(ClassificationCondition):

    def __init__(self, config):
        self._config = config
        self._income_condition = self._create_condition('incomes')
        self._cost_condition = self._create_condition('costs')
        self._condition = SalesRecordTypeCondition(income=self._income_condition, cost=self._cost_condition)

    def _create_condition(self, key):
        config = self._config.get(key, [])
        return ConfigurableCompositeCondition(config)

    def matches(self, row: KpirRow) -> ClassificationResult:
        return self._condition.matches(row)


class ClassifiedKpirRow(KpirRow):

    def __init__(self, row):
        super().__init__(row[0:16])
        self._row = row

    @property
    def projects_ids(self) -> list[str]:
        return self._row[16].split(', ')

    @property
    def kpiw_cost_category(self):
        return self._row[17]

    @property
    def is_kpiw(self):
        return bool(self._row[16])

    @property
    def income_kpwi(self):
        if self.is_income and self.is_kpiw:
            # TODO: compute rate based on timesheet
            return self.income * decimal.Decimal(0.8)
        return 0

    @property
    def income_other(self):
        if self.is_income and self.is_kpiw:
            return self.income - self.income_kpwi
        return 0

    @property
    def cost_a(self):
        return self._get_kpiw_cost('A')

    @property
    def cost_b(self):
        return self._get_kpiw_cost('B')

    @property
    def cost_c(self):
        return self._get_kpiw_cost('C')

    @property
    def cost_d(self):
        return self._get_kpiw_cost('D')

    def _get_kpiw_cost(self, category):
        if self.is_cost and self.is_kpiw and self.kpiw_cost_category == category:
            return self.cost
        return 0

    @classmethod
    def from_kpir_row(cls, row: KpirRow, classification: KpiwClassificationResult):
        result = []
        result.extend(row.raw)
        if classification:
            result.append(', '.join(project.id for project in classification.projects))
            if row.is_cost:
                result.append(classification.category)
            else:
                result.append('')
        else:
            result.extend(['', ''])
        return cls(result)


class KpirClassificationCsvWriter(CsvWriter):

    def __init__(self, filepath):
        super().__init__(filepath, header=['TODO'])

    def write_classification_result(self, row: KpirRow, classification: KpiwClassificationResult):
        self.write_classified_row(ClassifiedKpirRow.from_kpir_row(row, classification))

    def write_classified_row(self, row: ClassifiedKpirRow):
        self.write(row.raw)