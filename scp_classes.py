# Sprint Capacity Planner class definitions
import datetime
from dataclasses import dataclass, field
from decimal import Decimal
import pandas as pd

from scp_mapping import EMPLOYEE_AVAILABLE, EMPLOYEE_ON_VACATION, EMPLOYEE_ON_SICK_LEAVE, EMPLOYEE_NOT_ON_PROJECT, EMPLOYEE_ON_WEEKEND, EMPLOYEE_ON_BANK_HOLIDAY


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
    fte_on_holiday: list = field(init=False)
    fte_available: list = field(init=False)

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


@dataclass
class Employee:
    name: str
    country: str
    fte: float
    start_date_on_project: datetime.datetime
    end_date_on_project: datetime.datetime
    vacations: list = field(init=False)
    sick_leaves: list = field(init=False)
    extra_working_days: list = field(init=False)
    bank_holidays: list = field(init=False)

    def __post_init__(self):
        self.vacations = []
        self.sick_leaves = []
        self.extra_working_days = []
        self.bank_holidays = []

    def add_vacation(self, single_date):
        if isinstance(single_date, datetime.datetime):
            single_date = single_date.date()
        self.vacations.append(single_date)
        self.vacations.sort()

    def add_vacations_from_range(self, range_start, range_end):
        date_range = pd.bdate_range(range_start, range_end)
        for single_date in date_range:
            self.add_vacation(single_date.date())

    def add_sick_leave(self, single_date):
        if isinstance(single_date, datetime.datetime):
            single_date = single_date.date()
        self.sick_leaves.append(single_date)
        self.sick_leaves.sort()

    def add_sick_leaves_from_range(self, range_start, range_end):
        date_range = pd.bdate_range(range_start, range_end)
        for single_date in date_range:
            self.add_sick_leave(single_date.date())

    def add_bank_holiday(self, single_date):
        if isinstance(single_date, datetime.datetime):
            single_date = single_date.date()
        self.bank_holidays.append(single_date)
        self.bank_holidays.sort()

    def add_extra_working_day(self, single_date):
        if isinstance(single_date, datetime.datetime):
            single_date = single_date.date()
        self.extra_working_days.append(single_date)
        self.extra_working_days.sort()

    def is_available(self, single_date):
        employee_available = False
        if isinstance(single_date, datetime.datetime):
            single_date = single_date.date()
        if single_date >= self.start_date_on_project.date() and \
                (self.end_date_on_project is None or self.end_date_on_project.date() >= single_date):
            if single_date in self.vacations or single_date in self.sick_leaves or single_date in self.bank_holidays:
                employee_available = False
            else:
                if single_date.weekday() >= 5:
                    if single_date in self.extra_working_days:
                        employee_available = True
                    else:
                        employee_available = False
                else:
                    employee_available = True
        return employee_available

    def status(self, single_date):
        if isinstance(single_date, datetime.datetime):
            single_date = single_date.date()
        if single_date >= self.start_date_on_project.date():
            if self.end_date_on_project is None or self.end_date_on_project.date() >= single_date:
                if single_date in self.vacations:
                    employee_status = EMPLOYEE_ON_VACATION
                elif single_date in self.bank_holidays:
                    employee_status = EMPLOYEE_ON_BANK_HOLIDAY
                elif single_date in self.sick_leaves:
                    employee_status = EMPLOYEE_ON_SICK_LEAVE
                else:
                    if single_date.weekday() < 5 or single_date in self.extra_working_days:
                        employee_status = EMPLOYEE_AVAILABLE
                    else:
                        employee_status = EMPLOYEE_ON_WEEKEND
            else:
                employee_status = EMPLOYEE_NOT_ON_PROJECT
        else:
            employee_status = EMPLOYEE_NOT_ON_PROJECT
        return employee_status

