import asana
import json
import os
import datetime
from calendar import monthrange
from Google import Create_Service

CLIENT_SECRET_SERVICE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

personal_access_token = "1/1135096210698629:718f78c0f27970b442cd307018d87f8d"
spreadsheet_id = '1wyOwn_21U_EtdHJTi7HJ-lGAc0_BsQQi40HelATS0lU'


service = Create_Service(CLIENT_SECRET_SERVICE, API_NAME, API_VERSION, SCOPES)
client = asana.Client.access_token(personal_access_token)
now = datetime.datetime.now()


def tasks_json():
    output_list = []
    project_gid = ['1198913045061016', '1198913045061019',
                   '1198913121709013', '1198913121709020', '1198913121709027', '1198913121709034', '1198913121709041', '1199104176594285']
    # pylint: disable=maybe-no-member
    for i in project_gid:
        result = client.tasks.get_tasks_for_project(i, {'completed_since': f'{now.year}-{now.month:02d}-01T02:00:00.000Z'}, opt_pretty=True, opt_fields=[
            'name', 'assignee.name', 'completed', 'completed_at',  'tags.name'])
        output_list += list(result)
    with open("meta.json", "w", encoding="utf-8") as output:
        json.dump(output_list, output, ensure_ascii=False)


def cell_updater(cell_range_insert, value_range_body):
    # pylint: disable=maybe-no-member
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        range=worksheet_name + "!" + cell_range_insert,
        body=value_range_body).execute()


def cell_appender(cell_range_insert, value_range_body):
    # pylint: disable=maybe-no-member
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        range=worksheet_name + "!" + cell_range_insert,
        body=value_range_body).execute()


def cell_names():
    alphabet = ['B', 'D', 'F', 'H', 'J', 'L',
                'N', 'P', 'R', 'T', 'V', 'X', 'Z']
    with open("meta1.json", 'r', encoding="utf-8") as f:
        data = json.load(f)
        data = [k for k in data if not (k['completed_at'] == None)]
        data = sorted(data, key=lambda k: k['completed_at'])

    names_list = []
    for i in range(len(data)):
        if data[i]['assignee'] is not None:
            name = data[i]['assignee']['name']
            names_list.append(name)
    names_list = sorted(list(set(names_list)))
    i = 0
    for a in names_list:
        values = (
            (f'{a}', ''),
            ('час', 'задачі')
        )
        value_range_body = {
            'majorDimension': 'ROWS',
            'values': values
        }
        cell_updater(f"{alphabet[i]}1", value_range_body)
        data_new = [k for k in data if (k['assignee']['name'] == a)]
        for d in range(len(data_new)):
            if (data_new[d]['completed_at'] is not None and data_new[d]['tags'] is not None):
                values_tasks = (
                    ('-',
                     f"{data_new[d]['tags'][0]['name']} - {data_new[d]['name']}"),
                    ('', '')
                )
                value_range_body_tasks = {
                    'majorDimension': 'ROWS',
                    'values': values_tasks
                }
                date = data_new[d]['completed_at']
                date = list(date)
                if date[8] == '0':
                    day = int(date[9])
                else:
                    day = int(date[8]+date[9])
                cell_updater(f'{alphabet[i]}{day+2}',
                             value_range_body_tasks)
        i += 1


def cell_days():
    # pylint: disable=maybe-no-member
    k = 28
    i = 31
    values_time = (
        ('1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13', '14', '15',
         '16', '17', '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28'),
        ('', '')
    )
    value_range_body_time = {
        'majorDimension': 'COLUMNS',
        'values': values_time
    }
    cell_updater(f'A3', value_range_body_time)
    q_days = monthrange(now.year, now.month)
    while i <= (q_days[1]+2):
        k += 1
        values = (
            (f'{k}', ''),
            ('', '')
        )
        value_range_body = {
            'majorDimension': 'ROWS',
            'values': values
        }
        cell_updater(f'A{i}', value_range_body)
        i += 1
    if i > (q_days[1]+2):
        values_sum = (
            ('', ''),
            ('Сума', '')
        )
        value_range_body_sum = {
            'majorDimension': 'ROWS',
            'values': values_sum
        }
        cell_appender(f'A{i-1}', value_range_body_sum)


request_body = {
    'requests': [
        {
            'addSheet': {
                'properties': {
                    'title': now.strftime("%B %Y"),
                    'hidden': False
                }
            }
        }
    ]
}

# pylint: disable=maybe-no-member
service.spreadsheets().batchUpdate(
    spreadsheetId=spreadsheet_id, body=request_body).execute()

worksheet_name = now.strftime("%B %Y")


# cell_days()

cell_names()
