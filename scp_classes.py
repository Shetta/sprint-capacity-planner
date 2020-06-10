# Sprint Capacity Planner class definitions
import datetime
from dataclasses import dataclass


@dataclass
class Default:
    Name: str
    Value: int


@dataclass
class BankHoliday:
    Country: str
    Date: datetime.datetime
    Comment: str


@dataclass
class Developer:
    Name: str
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
class Sprint:
    Sprint: str
    Start_date: datetime.datetime
    End_date: datetime.datetime
    Workdays_in_sprint: int
    Dev_team_size_UK: float
    Dev_team_size_HU: float
    Dev_team_size_total: float


