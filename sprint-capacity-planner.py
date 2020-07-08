# library modules
import datetime
import time
import openpyxl
import sys
from dataclasses import fields
import pandas as pd
from decimal import Decimal

# local include files
import credentials as cr
from scp_classes import Default, BankHoliday, Developer, Sprint, EmployeeVacation, SprintDetails, Employee
from scp_mapping import VACATION_TYPE_SICK_LEAVE, EMPLOYEE_STATUS_TEXT

# global variables
vacations_list = []
extra_sick_leaves_list = []
bank_holidays_list = []
developers_list = []
sprints_list = []
defaults_list = []
extra_working_days_list = []
employee_obj_list = []
vacations_obj = []
bank_holidays_obj = []
developers_obj = []
sprints_obj = []
defaults_obj = []
employee_leaves_list = []

start_time = time.time()


def load_to_dictionary(full_file_name, sheet_name, load_max_col, load_max_row=50):
    xlsfile = full_file_name['path'] + '/' + full_file_name['filename']
    all_rows = []
    if load_max_col < 1:
        all_rows[0] = 'load_max_col must be greater than 0'
        return all_rows
    if load_max_row < 1:
        all_rows[0] = 'load_max_row must be greater than 0'
        return all_rows
    try:
        wb = openpyxl.load_workbook(xlsfile, read_only=True, data_only=True)
    except Exception:
        print("Unexpected error:", sys.exc_info()[0])
        return all_rows
    keys = []
    try:
        sheet = wb[sheet_name]
    except Exception:
        print("Unexpected error:", sys.exc_info()[0])
        return all_rows
    for first_row in sheet.iter_rows(min_row=1, max_row=1, min_col=0, max_col=load_max_col, values_only=True):
        for key in first_row:
            keys.append(key)
    for row in sheet.iter_rows(min_row=2, max_row=load_max_row, min_col=0, max_col=load_max_col, values_only=True):
        if row[0] is None:
            continue
        one_row = {}
        for i in range(0, load_max_col):
            one_row[first_row[i]] = row[i]
        all_rows.append(one_row)
    return all_rows


def load_to_objects(object_class, full_file_name, sheet_name, load_max_col, load_max_row=50, ):
    xlsfile = full_file_name['path'] + '/' + full_file_name['filename']
    all_rows = []
    if load_max_col < 1:
        all_rows[0] = 'load_max_col must be greater than 0'
        return all_rows
    if load_max_row < 1:
        all_rows[0] = 'load_max_row must be greater than 0'
        return all_rows
    try:
        wb = openpyxl.load_workbook(xlsfile, read_only=True, data_only=True)
    except:
        print("Unexpected error:", sys.exc_info()[0])
        return all_rows
    keys = []
    try:
        sheet = wb[sheet_name]
    except:
        print("Unexpected error:", sys.exc_info()[0])
        return all_rows
    for row in sheet.iter_rows(min_row=2, max_row=load_max_row, min_col=0, max_col=load_max_col, values_only=True):
        one_object = object.__new__(object_class)
        i = 0
        for field, field_value in zip(fields(object_class), row):
            one_object.__setattr__(field.name, field_value)
            i = i + 1

        all_rows.append(one_object)

    return all_rows


def vacations_data_process(list_of_vacations, list_of_bank_holidays, list_of_employees):
    list_of_employee_leaves = []
    list_of_vacations = sorted(list_of_vacations, key=lambda x: (x['EMPLOYEE'], x['START DATE']))
    list_of_bank_holidays = sorted(list_of_bank_holidays, key=lambda x: (x['Date']))
    list_of_employees = sorted(list_of_employees, key=lambda x: (x['Name']))
    for employee in list_of_employees:
        employee_bank_holidays = get_employee_bank_holidays(employee['Name'], list_of_employees, list_of_bank_holidays)
        employee_obj = EmployeeVacation(employee['Name'], employee_bank_holidays, employee['FTE'])
        filtered_vacations = [vac for vac in list_of_vacations if vac['EMPLOYEE'] == employee['Name']]
        for vacation_record in filtered_vacations:
            start_date_date = vacation_record['START DATE']
            if isinstance(start_date_date, datetime.datetime):
                start_date_date = start_date_date.date()
            end_date_date = vacation_record['END DATE']
            if isinstance(end_date_date, datetime.datetime):
                end_date_date = end_date_date.date()
            employee_obj.add_vacations_from_range(start_date_date, end_date_date)
        list_of_employee_leaves.append(employee_obj)
    return list_of_employee_leaves


def bank_holidays_data_process(list_of_bank_holidays):
    all_holidays = []
    countries = []
    list_of_bank_holidays = sorted(list_of_bank_holidays, key=lambda x: (x['Date']))
    current_date = ""
    current_date_date = ""
    for item in list_of_bank_holidays:
        if item['Date'] != current_date:
            if current_date != "":
                all_holidays.append({'date': current_date_date, 'countries': countries})
                countries = []
            current_date = item['Date']
            if isinstance(current_date, datetime.datetime):
                current_date_date = current_date.date()
            else:
                current_date_date = current_date
        countries.append(item['Country'])
    all_holidays.append({'date': current_date_date, 'countries': countries})
    holidays_obj = BankHoliday(all_holidays)
    return holidays_obj


def who_is_on_holiday(datetime_date):
    on_holiday = []
    global employee_leaves_list
    for employee in employee_leaves_list:
        if employee.is_on_holiday(datetime_date):
            on_holiday.append(employee.name)
    return on_holiday


def get_employee_bank_holidays(employee, list_of_employees, list_of_bank_holidays):
    employee_bank_holidays = []
    employee_dict = next((item for item in list_of_employees if item['Name'] == employee), None)
    if type(employee_dict) == dict:
        country_code = employee_dict['Country']
        for bank_holiday in list_of_bank_holidays:
            if bank_holiday['Country'] == country_code:
                employee_bank_holidays.append(bank_holiday['Date'].date())

    return employee_bank_holidays


def load_all_to_dict():
    global bank_holidays_list
    global vacations_list
    global developers_list
    global sprints_list
    global defaults_list
    global extra_sick_leaves_list
    global extra_working_days_list
    bank_holidays_list = load_to_dictionary(cr.testexcelfile, 'Bank holidays', 3, 100)
    vacations_list = load_to_dictionary(cr.testexcelfile, 'Vacations', 16, 200)
    developers_list = load_to_dictionary(cr.testexcelfile, 'Developers', 5)
    sprints_list = load_to_dictionary(cr.testexcelfile, 'Sprints', 7, 100)
    defaults_list = load_to_dictionary(cr.testexcelfile, 'Defaults', 2)
    extra_sick_leaves_list = load_to_dictionary(cr.testexcelfile, 'Extra sick leave', 3)
    extra_working_days_list = load_to_dictionary(cr.testexcelfile, 'Extra working days', 3)


def load_all_to_object():
    # global developers_obj
    global sprints_obj
    global defaults_obj
    # developers_obj = load_to_objects(Developer, cr.testexcelfile, 'Developers', 5)
    sprints_obj = load_to_objects(Sprint, cr.testexcelfile, 'Sprints', 7, 100)
    defaults_obj = load_to_objects(Default, cr.testexcelfile, 'Defaults', 2)


def sprints_data_process(list_of_sprints, list_of_employee_vacations):
    list_of_sprint_details = []
    for sprint in list_of_sprints:
        sprint_details_obj = SprintDetails(sprint['Sprint'], sprint['Start date'], sprint['End date'],
                                           sprint['Dev team size UK'], sprint['Dev team size HU'])
        date_range = pd.bdate_range(sprint['Start date'], sprint['End date'])
        for single_date in date_range:
            for vacation_obj in list_of_employee_vacations:
                if vacation_obj.is_on_holiday(single_date):
                    sprint_details_obj.add_employee_on_holiday(single_date, vacation_obj.name, vacation_obj.fte)
                else:
                    sprint_details_obj.add_employee_available(single_date, vacation_obj.name, vacation_obj.fte)
        list_of_sprint_details.append(sprint_details_obj)
    return list_of_sprint_details


def validate_employees_list(list_of_employees):
    data_error = False
    for employee in list_of_employees:
        if employee['Start date on project'] is None:
            print('ERROR!', employee['Name'], ': start date cannot be empty!')
            data_error = True
        if employee['End date on project'] is not None and (
                employee['End date on project'] < employee['Start date on project']):
            print('ERROR!', employee['Name'], ': end date(', employee['End date on project'].date(),
                  ') cannot be earlier than start date (', employee['Start date on project'].date(), ')!')
            data_error = True
    return data_error


def test1():
    global employee_leaves_list
    for employee_obj in employee_leaves_list:
        # employee_obj.print_all_vacations()
        print(employee_obj)


def test2():
    global sprint_details_list
    for sprint_detail in sprint_details_list:
        # print(sprint_detail)
        print(sprint_detail.sprint, sprint_detail.start_date.date(), sprint_detail.end_date.date(),
              sprint_detail.get_total_fte_available(), sprint_detail.get_total_fte_on_holiday(),
              sprint_detail.get_sprint_capacity())


def test3():
    global developers_list, bank_holidays_list, employee_leaves_list
    print(get_employee_bank_holidays('Garra Peters', developers_list, bank_holidays_list))
    print(employee_leaves_list)


def test4():
    emp1 = Employee('Bela', 'HU', 1, datetime.datetime(2019, 12, 25, 10, 11), datetime.datetime(2020, 8, 1, 10, 11))
    emp1.add_vacation(datetime.datetime(2020, 7, 25, 10, 11))
    emp1.add_sick_leave(datetime.datetime(2020, 6, 25, 10, 11))
    emp1.add_vacation(datetime.datetime(2020, 6, 24, 10, 11))
    emp1.add_vacation(datetime.datetime(2020, 8, 3, 10, 11))
    emp1.add_extra_working_day(datetime.date(2020, 6, 13))
    emp1.add_extra_working_day(datetime.date(2020, 6, 27))
    emp1.add_extra_working_day(datetime.date(2020, 6, 20))
    x_date = datetime.date(2020, 6, 23)
    print(x_date, emp1.is_available(x_date), emp1.status(x_date))
    x_date = datetime.date(2020, 6, 24)
    print(x_date, emp1.is_available(x_date), emp1.status(x_date))
    x_date = datetime.date(2020, 6, 25)
    print(x_date, emp1.is_available(x_date), emp1.status(x_date))
    x_date = datetime.date(2020, 6, 27)
    print(x_date, emp1.is_available(x_date), emp1.status(x_date))
    x_date = datetime.date(2020, 6, 28)
    print(x_date, emp1.is_available(x_date), emp1.status(x_date))
    x_date = datetime.date(2020, 8, 2)
    print(x_date, emp1.is_available(x_date), emp1.status(x_date))


def test5():
    global employee_obj_list
    x_date = datetime.date(2020, 12, 12)
    for item in employee_obj_list:
        print(item.name, x_date, 'available:', item.is_available(x_date), 'status:',
              EMPLOYEE_STATUS_TEXT[item.status(x_date)])


def employee_data_process(list_of_employees, list_of_vacations, list_of_bank_holidays, list_of_extra_sick_leaves,
                          list_of_extra_working_days):
    list_of_employee_objects = []
    for employee in list_of_employees:
        employee_obj = Employee(employee['Name'], employee['Country'], employee['FTE'],
                                employee['Start date on project'], employee['End date on project'])
        filtered_vacations = [vac for vac in list_of_vacations if vac['EMPLOYEE'] == employee['Name']]
        for vacation_record in filtered_vacations:
            start_date_date = vacation_record['START DATE']
            if isinstance(start_date_date, datetime.datetime):
                start_date_date = start_date_date.date()
            end_date_date = vacation_record['END DATE']
            if isinstance(end_date_date, datetime.datetime):
                end_date_date = end_date_date.date()
            if vacation_record['VACATION TYPE'] == VACATION_TYPE_SICK_LEAVE:
                employee_obj.add_sick_leaves_from_range(start_date_date, end_date_date)
            else:
                employee_obj.add_vacations_from_range(start_date_date, end_date_date)
        filtered_sick_leaves = [leave for leave in list_of_extra_sick_leaves if leave['Name'] == employee['Name']]
        for sick_leave_record in filtered_sick_leaves:
            start_date_date = sick_leave_record['Start date']
            if isinstance(start_date_date, datetime.datetime):
                start_date_date = start_date_date.date()
            end_date_date = sick_leave_record['End date']
            if isinstance(end_date_date, datetime.datetime):
                end_date_date = end_date_date.date()
            employee_obj.add_sick_leaves_from_range(start_date_date, end_date_date)
        filtered_bank_holidays = [bank_holiday for bank_holiday in list_of_bank_holidays if
                                  bank_holiday['Country'] == employee['Country']]
        for bank_holiday_record in filtered_bank_holidays:
            employee_obj.add_bank_holiday(bank_holiday_record['Date'])
        filtered_extra_working_days = [extra_working_day for extra_working_day in list_of_extra_working_days if
                                       extra_working_day['Country'] == employee['Country']]
        for extra_working_day_record in filtered_extra_working_days:
            employee_obj.add_extra_working_day(extra_working_day_record['Date'])
        list_of_employee_objects.append(employee_obj)
    return list_of_employee_objects


if __name__ == '__main__':
    now = datetime.datetime.now()
    print("Started at:", now.strftime("%Y-%m-%d %H:%M:%S"))
    print("#####----------------------------------#####")
    load_all_to_dict()
    if validate_employees_list(developers_list):
        exit(-1)
    employee_leaves_list = vacations_data_process(vacations_list, bank_holidays_list, developers_list)
    # test1()
    bank_holidays_obj = bank_holidays_data_process(bank_holidays_list)
    x_date = datetime.date(2020, 12, 25)
    sprint_details_list = sprints_data_process(sprints_list, employee_leaves_list)
    employee_obj_list = employee_data_process(developers_list, vacations_list, bank_holidays_list,
                                              extra_sick_leaves_list, extra_working_days_list)
    test5()

    # print(bank_holidays_obj)
    # print(bank_holidays_obj.is_holiday(datetime.date(2020, 12, 25)))
    # load_all_to_object()

    # sp1 = sprints[0]
    # sp1.define_workday_dates()
    print("#####----------------------------------#####")
    print("--- Execution time: %s seconds ---" % (time.time() - start_time))
