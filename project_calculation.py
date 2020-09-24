import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials

CREDENTIALS_FILE = 'projectcalculation-19370cfff173.json'  # имя файла с закрытым ключом

credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE,
                                                               ['https://www.googleapis.com/auth/spreadsheets',
                                                                'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)

spreadsheetId_1 = '1C6djSWlQzaMIeZX8ZLSY3UTSDkPEzdKdZVueVoYOTQQ'  # ProjectCalculation
spreadsheetId_2 = '1OyigdnJbCdKHVXZEz3mwrG0aJEqPnka814f9xJ9OW8M'  # ProjectCalculationFinance
# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------

# get staff info
# get staff projects
staff_projects_data = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId_1,
                                                               ranges=["Сотрудники!K1:AZ1"],
                                                               valueRenderOption='UNFORMATTED_VALUE',
                                                               dateTimeRenderOption='FORMATTED_STRING').execute()

# get staff info
staff_info_data = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId_1,
                                                           ranges=["Сотрудники!A3:J200"],
                                                           valueRenderOption='UNFORMATTED_VALUE',
                                                           dateTimeRenderOption='FORMATTED_STRING').execute()

# get staff loading
staff_loading_data = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId_1,
                                                              ranges=["Сотрудники!K3:AZ200"],
                                                              valueRenderOption='UNFORMATTED_VALUE',
                                                              dateTimeRenderOption='FORMATTED_STRING').execute()
# get staff salary
staff_salary_data = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId_2,
                                                             ranges=["Сотрудники!A2:C200"],
                                                             valueRenderOption='UNFORMATTED_VALUE',
                                                             dateTimeRenderOption='FORMATTED_STRING').execute()

# parsing staff salary
staff_salary = {}
for row in staff_salary_data['valueRanges'][0]['values']:
    staff_salary[row[0]] = row[2]

# parsing staff info
staff_array = []
for row1, row2 in zip(staff_info_data['valueRanges'][0]['values'], staff_loading_data['valueRanges'][0]['values']):
    staff_projects = {}
    for _row2, project in zip(row2, staff_projects_data['valueRanges'][0]['values'][0]):
        if type(_row2) is float or type(_row2) is int:
            if _row2 > 0:
                staff_projects[project] = _row2
    # set staff salary
    salary = 0
    if row1[0] in staff_salary.keys():
        salary = staff_salary[row1[0]]
    staff = {"full name": row1[0], "division": row1[1], "salary": salary, "projects": staff_projects,
             "parameter 1": row1[6], "parameter 2": row1[7], "parameter 3": row1[8]}
    staff_array.append(staff)
# ----------------------------------------------------------------------------------------------------------------------


# get project info
project_info_data = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId_1,
                                                             ranges=["Проекты!A2:E100"],
                                                             valueRenderOption='UNFORMATTED_VALUE',
                                                             dateTimeRenderOption='FORMATTED_STRING').execute()
# parsing project data
projects_array = []
for row in project_info_data['valueRanges'][0]['values']:
    projects_array.append(
        {"name": row[0], "code": row[1], "holder": row[2], "customer": row[3], "priority": row[4], "loading": 0.0,
         "cost": 0.0, "divisions": []})

# calculate project info
for project in projects_array:
    project_name = project["name"]
    parameter_1 = ''
    parameter_2 = ''
    parameter_3 = ''
    for staff in staff_array:
        if project_name in staff["projects"].keys():
            project["loading"] += staff["projects"][project_name]
            project["cost"] += staff["projects"][project_name] * staff["salary"]
            project["divisions"].append(staff["division"])
            if staff['parameter 1'] != '':
                parameter_1 += staff['parameter 1']
                parameter_1 += '; '
            if staff['parameter 2'] != '':
                parameter_2 += staff['parameter 2']
                parameter_2 += '; '
            if staff['parameter 3'] != '':
                parameter_3 += staff['parameter 3']
                parameter_3 += '; '
    project["divisions"] = set(project["divisions"])
    project['parameter 1'] = parameter_1
    project['parameter 2'] = parameter_2
    project['parameter 3'] = parameter_3
    divisions = ''
    for division in project["divisions"]:
        divisions += division
        divisions += '; '
    project["divisions"] = divisions

# ----------------------------------------------------------------------------------------------------------------------


# get division info
division_info_data = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId_1,
                                                              ranges=["Подразделения!A2:A100"],
                                                              valueRenderOption='UNFORMATTED_VALUE',
                                                              dateTimeRenderOption='FORMATTED_STRING').execute()
# parsing division info
divisions_array = []
for row in division_info_data['valueRanges'][0]['values']:
    divisions_array.append({"name": row[0], "staff": 0, "loading": 0, "cost": 0})

# calculate division info
for division in divisions_array:
    division_name = division['name']
    for staff in staff_array:
        if division_name == staff['division']:
            division['staff'] += 1
            division['cost'] += staff['salary']

    for project in projects_array:
        if division_name == project['holder']:
            division['loading'] += project['loading']

    parameter_1 = ''
    parameter_2 = ''
    parameter_3 = ''
    for staff in staff_array:
        if division_name == staff['division']:
            if staff['parameter 1'] != '':
                parameter_1 += staff['parameter 1']
                parameter_1 += '; '
            if staff['parameter 2'] != '':
                parameter_2 += staff['parameter 2']
                parameter_2 += '; '
            if staff['parameter 3'] != '':
                parameter_3 += staff['parameter 3']
                parameter_3 += '; '
    division['parameter 1'] = parameter_1
    division['parameter 2'] = parameter_2
    division['parameter 3'] = parameter_3
# ----------------------------------------------------------------------------------------------------------------------


# get customer info
customer_info_data = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheetId_1,
                                                              ranges=["Заказчики!A2:A100"],
                                                              valueRenderOption='UNFORMATTED_VALUE',
                                                              dateTimeRenderOption='FORMATTED_STRING').execute()
# parsing customer info
customers_array = []
for row in customer_info_data['valueRanges'][0]['values']:
    customers_array.append({'name': row[0], 'projects': '', 'cost': 0})

# calculate customer info
for customer in customers_array:
    customer_name = customer['name']
    for project in projects_array:
        if customer_name == project['customer']:
            customer['projects'] += project['name']
            customer['projects'] += '; '
            customer['cost'] += project['cost']
# ----------------------------------------------------------------------------------------------------------------------


# set project result ProjectCalculation

results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId_1, body={
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "Проекты!F2:J100",
         "majorDimension": "COLUMNS",
         # сначала заполнять столбцы, затем ряды (т.е. самые внутренние списки в values - это столбцы)
         "values": [[project["loading"] for project in projects_array],
                    [project['divisions'] for project in projects_array],
                    [project['parameter 1'] for project in projects_array],
                    [project['parameter 2'] for project in projects_array],
                    [project['parameter 3'] for project in projects_array]]}
    ]
}).execute()

# set project result ProjectCalculationFinance
results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId_2, body={
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "Проекты!A2:K100",
         "majorDimension": "COLUMNS",
         # сначала заполнять столбцы, затем ряды (т.е. самые внутренние списки в values - это столбцы)
         "values": [[project["name"] for project in projects_array],
                    [project['code'] for project in projects_array],
                    [project['holder'] for project in projects_array],
                    [project['customer'] for project in projects_array],
                    [project['priority'] for project in projects_array],
                    [project['loading'] for project in projects_array],
                    [project['divisions'] for project in projects_array],
                    [project['parameter 1'] for project in projects_array],
                    [project['parameter 2'] for project in projects_array],
                    [project['parameter 3'] for project in projects_array],
                    [project["cost"] for project in projects_array]]}
    ]
}).execute()
# ----------------------------------------------------------------------------------------------------------------------


# set division result ProjectCalculation
results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId_1, body={
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "Подразделения!B2:F100",
         "majorDimension": "COLUMNS",
         # сначала заполнять столбцы, затем ряды (т.е. самые внутренние списки в values - это столбцы)
         "values": [[division["staff"] for division in divisions_array],
                    [division["loading"] for division in divisions_array],
                    [division["parameter 1"] for division in divisions_array],
                    [division["parameter 2"] for division in divisions_array],
                    [division["parameter 3"] for division in divisions_array]]}
    ]
}).execute()

# set division result ProjectCalculationFinance
results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId_2, body={
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "Подразделения!A2:G100",
         "majorDimension": "COLUMNS",
         # сначала заполнять столбцы, затем ряды (т.е. самые внутренние списки в values - это столбцы)
         "values": [[division["name"] for division in divisions_array],
                    [division["staff"] for division in divisions_array],
                    [division["loading"] for division in divisions_array],
                    [division["parameter 1"] for division in divisions_array],
                    [division["parameter 2"] for division in divisions_array],
                    [division["parameter 3"] for division in divisions_array],
                    [division["cost"] for division in divisions_array]]}
    ]
}).execute()

# results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId_2, body={
#     "valueInputOption": "USER_ENTERED",
#     "data": [
#         {"range": "Подразделения!G2:G100",
#          "majorDimension": "COLUMNS",
#          # сначала заполнять столбцы, затем ряды (т.е. самые внутренние списки в values - это столбцы)
#          "values": [[division["cost"] for division in divisions_array]]}
#     ]
# }).execute()
# ----------------------------------------------------------------------------------------------------------------------


# set customer result ProjectCalculation
results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId_1, body={
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "Заказчики!B2:B100",
         "majorDimension": "COLUMNS",
         # сначала заполнять столбцы, затем ряды (т.е. самые внутренние списки в values - это столбцы)
         "values": [[customer['projects'] for customer in customers_array]]}
    ]
}).execute()

# set customer result ProjectCalculationFinance
results = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheetId_2, body={
    "valueInputOption": "USER_ENTERED",
    "data": [
        {"range": "Заказчики!A2:C100",
         "majorDimension": "COLUMNS",
         # сначала заполнять столбцы, затем ряды (т.е. самые внутренние списки в values - это столбцы)
         "values": [[customer['name'] for customer in customers_array],
                    [customer['projects'] for customer in customers_array],
                    [customer['cost'] for customer in customers_array]]}
    ]
}).execute()
print()
