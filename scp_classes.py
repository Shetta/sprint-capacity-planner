# Sprint Capacity Planner class definitions
import datetime
from dataclasses import dataclass, field
from decimal import Decimal
import pandas as pd


@dataclass
class Default:
    name: str
    Value: int


@dataclass
class BankHoliday:
    dates: list  # list of dicts: {'date': current_date_date, 'countries': countries[] }

    def is_holiday(self, datetime_date):
        if isinstance(datetime_date, datetime.datetime):
            datetime_date = datetime_date.date()
        return next((item for item in self.dates if item['date'] == datetime_date), None)

    def is_holiday_in_country(self, datetime_date, country):
        if isinstance(datetime_date, datetime.datetime):
            datetime_date = datetime_date.date()
        holiday_bool = False
        for item in self.dates:
            if item['date'] == datetime_date:
                for holiday_country in item['countries']:
                    if holiday_country == country:
                        holiday_bool = True
        return holiday_bool


@dataclass
class Developer:
    name: str
    Country: str
    FTE: float
    Start_date_on_project: datetime.datetime
    End_date_on_project: datetime.datetime


@dataclass
class Vacation:
    EMPLOYEE: str
    EMPLOYEE_ID: str
    EMPLOYEE_UID: str
    APPROVER: str
    COUNTRY: str
    CITY: str
    PROJECT_CODE: str
    PROJECT_ROLE: str
    VACATION_TYPE: str
    CREATION_DATE: datetime.datetime
    CREATED_BY: str
    LAST_MODIFIED_DATE: datetime.datetime
    LAST_MODIFIED_BY: str
    START_DATE: datetime.datetime
    END_DATE: datetime.datetime
    DURATION: int


@dataclass
class EmployeeVacation:
    name: str
    vacation_days: list
    fte: float

    def add_vacations_from_range(self, range_start, range_end):
        date_range = pd.bdate_range(range_start, range_end)
        for single_date in date_range:
            self.vacation_days.append(single_date.date())
        self.vacation_days.sort()

    def add_vacation_date(self, single_date):
        if isinstance(single_date, datetime.datetime):
            single_date = single_date.date()
        self.vacation_days.append(single_date)
        self.vacation_days.sort()

    def print_all_vacations(self):
        print(self.name)
        for vacation in self.vacation_days:
            print(vacation)

    def is_on_holiday(self, datetime_date):
        if isinstance(datetime_date, datetime.datetime):
            datetime_date = datetime_date.date()
        if datetime_date in self.vacation_days:
            return True
        else:
            return False


@dataclass
class Sprint:
    Sprint: str
    Start_date: datetime.datetime
    End_date: datetime.datetime
    Workdays_in_sprint: int
    Dev_team_size_UK: float
    Dev_team_size_HU: float
    Dev_team_size_total: float

    def define_workday_dates(self):
        date_range = pd.bdate_range(self.Start_date, self.End_date)
        for single_date in date_range:
            print(single_date.strftime("%Y-%m-%d"))


@dataclass
class SprintDetails:
    sprint: str
    start_date: datetime.datetime
    end_date: datetime.datetime
    dev_team_size_UK: float
    dev_team_size_HU: float
    dev_team_size_total: float = 0
    workdays_in_sprint: int = 0
    members_on_holiday: list = field(init=False, repr=False)
    members_available: list = field(init=False, repr=False)
    fte_on_holiday: list = field(init=False, repr=False)
    fte_available: list = field(init=False, repr=False)

    def __post_init__(self):
        self.dev_team_size_total = self.dev_team_size_HU + self.dev_team_size_UK
        workdays_range = pd.bdate_range(start=self.start_date, end=self.end_date)
        self.workdays_in_sprint = len(workdays_range)
        self.members_on_holiday = []
        self.members_available = []
        self.fte_on_holiday = []
        self.fte_available = []
        for date in workdays_range:
            if isinstance(date, pd.Timestamp):
                date = date.date()
            self.fte_on_holiday.append({'date': date, 'fte': 0})
            self.fte_available.append({'date': date, 'fte': 0})

    def get_dev_team_size_total(self):
        self.dev_team_size_total = self.dev_team_size_HU + self.dev_team_size_UK
        return self.dev_team_size_total

    def add_employee_on_holiday(self, date, name, fte):
        if isinstance(date, pd.Timestamp):
            date = date.date()
        self.members_on_holiday.append({'date': date, 'name': name})
        index = 0
        for item in self.fte_on_holiday:
            if item['date'] == date:
                self.fte_on_holiday[index] = {'date': date, 'fte': float(format(Decimal.from_float(item['fte'] + fte), '.2f'))}
            index += 1

    def add_employee_available(self, date, name, fte):
        if isinstance(date, pd.Timestamp):
            date = date.date()
        self.members_available.append({'date': date, 'name': name})
        index = 0
        for item in self.fte_available:
            if item['date'] == date:
                self.fte_available[index] = {'date': date, 'fte': float(format(Decimal.from_float(item['fte'] + fte), '.2f'))}
            index += 1

    def get_total_fte_available(self):
        total_fte = 0.0
        for item in self.fte_available:
            total_fte = float(format(Decimal.from_float(total_fte + item['fte']), '.2f'))
        return total_fte

    def get_total_fte_on_holiday(self):
        total_fte = 0.0
        for item in self.fte_on_holiday:
            total_fte = float(format(Decimal.from_float(total_fte + item['fte']), '.2f'))
        return total_fte

    def get_sprint_capacity(self):
        total_fte_available = self.get_total_fte_available()
        sprint_capacity = total_fte_available / (total_fte_available + self.get_total_fte_on_holiday())
        sprint_capacity = float(format(Decimal.from_float(sprint_capacity), '.2f'))
        return sprint_capacity
