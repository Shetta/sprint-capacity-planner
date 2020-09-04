# config.py - Configuration parameters

bank_holidays_excel = {'path': '.',
                       'filename': 'capacity-planner-test.xlsx',
                       'sheet': 'Bank holidays',
                       'columns': ['Country', 'Date', 'Comment']}

vacation_requests_excel = {'path': '.',
                           'filename': 'vacation_requests.xlsx',
                           'sheet': 'vacation_requests',
                           'columns': ['Name', 'Start date', 'End date']}

team_members_excel = {'path': '.',
                      'filename': 'capacity-planner-test.xlsx',
                      'sheet': 'Developers',
                      'columns': ['Name', 'Country', 'FTE', 'Start date on project', 'End date on project']}

sprints_excel = {'path': '.',
                 'filename': 'capacity-planner-test.xlsx',
                 'sheet': 'Sprints',
                 'columns': ['Sprint ID', 'Start date', 'End date']}

extra_sick_leaves_excel = {'path': '.',
                           'filename': 'capacity-planner-test.xlsx',
                           'sheet': 'Extra sick leave',
                           'columns': ['Name', 'Start date', 'End date']}

extra_working_days_excel = {'path': '.',
                            'filename': 'capacity-planner-test.xlsx',
                            'sheet': 'Extra working days',
                            'columns': ['Country', 'Date', 'Comment']}

defaults_excel = {'path': '.',
                  'filename': 'capacity-planner-test.xlsx',
                  'sheet': 'Defaults',
                  'columns': ['Name', 'Value']
                  }

overview_excel = {'path': '.',
                  'filename': 'capacity_overview.xlsx',
                  'sheet': 'Overview'}
