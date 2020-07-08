# Mapping values to meaningful names

# Employee statuses
EMPLOYEE_UNKNOWN = 99
EMPLOYEE_AVAILABLE = 1
EMPLOYEE_ON_VACATION = 2
EMPLOYEE_ON_SICK_LEAVE = 3
EMPLOYEE_NOT_ON_PROJECT = 4
EMPLOYEE_ON_WEEKEND = 5
EMPLOYEE_ON_BANK_HOLIDAY = 6

EMPLOYEE_STATUS_TEXT = {EMPLOYEE_AVAILABLE: 'Available',
                        EMPLOYEE_ON_VACATION: 'On vacation',
                        EMPLOYEE_ON_SICK_LEAVE: 'On sick leave',
                        EMPLOYEE_NOT_ON_PROJECT: 'Not on project',
                        EMPLOYEE_ON_WEEKEND: 'On weekend',
                        EMPLOYEE_ON_BANK_HOLIDAY: 'On bank holiday',
                        EMPLOYEE_UNKNOWN: 'Unknown'}

VACATION_TYPE_SICK_LEAVE = 'Sick leave'
