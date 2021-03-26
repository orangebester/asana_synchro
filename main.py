import time
import json
import os
import datetime
from calendar import monthrange

import asana
from Google import Create_Service

from formatting import create_table, update_table, merge_columns, format_name, format_scnd_header, format_bold, format_days, format_tasks, format_time_tasks
from personal_data import personal_access_token, spreadsheet_id


# pylint: disable=maybe-no-member

CLIENT_SECRET_SERVICE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


service = Create_Service(CLIENT_SECRET_SERVICE, API_NAME, API_VERSION, SCOPES)
client = asana.Client.access_token(personal_access_token)
now = datetime.datetime.now()
worksheet_name = now.strftime("%B %Y")


def tasks_json():
    """Creates json file and downloads there data about tasks and subtasks in Asana"""

    output_list = []
    output_list_subtasks = []
    count = 0
    project_gid = ['1198913045061016', '1198913045061019',
                   '1198913121709013', '1198913121709020', '1198913121709027', '1198913121709034', '1198913121709041', '1199104176594285']
    # pylint: disable=maybe-no-member

    # download tasks
    for i in project_gid:
        result = client.tasks.get_tasks_for_project(i, {'completed_since': f'{now.year}-{now.month:02d}-01T02:00:00.000Z'}, opt_pretty=True, opt_fields=[
            'name', 'assignee.name', 'completed', 'completed_at',  'tags.name', 'subtasks'])
        output_list += list(result)

    # download subtasks
    for i in output_list:
        if len(i['subtasks']) != 0:
            result = client.tasks.get_subtasks_for_task(i['gid'], {'completed_since': f'{now.year}-{now.month:02d}-01T02:00:00.000Z'}, opt_pretty=True, opt_fields=[
                'name', 'assignee.name', 'completed', 'completed_at', 'tags.name'])
            output_list_subtasks += list(result)
            count += 1
            if count % 149 == 0:
                time.sleep(60.0)

    output_list.extend(output_list_subtasks)
    with open("meta.json", "w", encoding="utf-8") as output:
        json.dump(output_list, output, ensure_ascii=False)


def cell_updater(cell_range_insert, value_range_body):
    """Updates data in cells

    Args: 
        cell_range_insert (str): starting cell range
        value_range_body (dict): values to download
    """

    # pylint: disable=maybe-no-member
    time.sleep(1.1)
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        range=worksheet_name + "!" + cell_range_insert,
        body=value_range_body).execute()


def cell_appender(cell_range_insert, value_range_body):
    """Appends data to other cells

    Args:
        cell_range_insert (str): starting point cell
        value_range_body (dict): values to download
    """

    # pylint: disable=maybe-no-member
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        range=worksheet_name + "!" + cell_range_insert,
        body=value_range_body).execute()


def tasks_spreadsheet():
    """Updates cells in table according to data in json file"""

    startrow_index = 1
    n = cell_days()
    alphabet = ['B', 'D', 'F', 'H', 'J', 'L',
                'N', 'P', 'R', 'T', 'V', 'X', 'Z']
    # downloading data from json
    with open("meta.json", 'r', encoding="utf-8") as f:
        data = json.load(f)
        data = [k for k in data if not (k['completed_at'] == None)]
        data = sorted(data, key=lambda k: k['completed_at'])

    # downloading and formatting names of assignee
    names_list = []
    for i in range(len(data)):
        if data[i]['assignee'] is not None:
            name = data[i]['assignee']['name']
            names_list.append(name)
    names_list = sorted(list(set(names_list)))
    i = 0
    format_name(sheet_id, service, spreadsheet_id)
    format_scnd_header(sheet_id, service, spreadsheet_id)

    # downloading tasks by names of assignee
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

        # formatting cells
        merge_columns(startrow_index, sheet_id, service, spreadsheet_id)
        format_bold(startrow_index, sheet_id, service, spreadsheet_id)
        format_tasks(startrow_index, sheet_id, service, spreadsheet_id)
        format_time_tasks(startrow_index, sheet_id, service, spreadsheet_id)

        startrow_index += 2
        # sorting tasks by name and checking whether task is completed
        data_new = [k for k in data if (k['assignee'] is not None and k['assignee']
                                        ['name'] == a and k['completed_at'] is not None)]
        for d in range(len(data_new)):
            data_new[d]['tags'] = [k for k in data_new[d]['tags']
                                   if float_tasks(k) != None]
            date = data_new[d]['completed_at']
            date = list(date)
            if date[8] == '0':
                day = int(date[9])
            else:
                day = int(date[8] + date[9])
            # adding new tag "day"
            data_new[d]['day'] = day

        # checking whether task has a tag
        data_new = [k for k in data_new if len(k['tags']) != 0]
        while len(data_new) >= 1:
            data_so_new = [k for k in data_new if (
                k['day'] == data_new[0]['day'])]
            data_new = [k for k in data_new if k not in data_so_new]
            var1 = float(data_so_new[0]['tags'][0]['name'].replace(',', '.'))
            var2 = f"{data_so_new[0]['tags'][0]['name'].replace('.',',')} - {data_so_new[0]['name']}\n"

            if len(data_so_new) > 1:
                # updating task if there are more than one tag
                for c in range(len(data_so_new[1:])):
                    var1 += float(data_so_new[1:][c]['tags']
                                  [0]['name'].replace(',', '.'))
                    var2 += f"{data_so_new[1:][c]['tags'][0]['name'].replace('.',',')} - {data_so_new[1:][c]['name']}\n"

            var1 = str(var1).replace('.', ',')
            values_tasks = (
                (f'{var1}',
                    f"{var2}"),
                ('', '')
            )
            value_range_body_tasks = {
                'majorDimension': 'ROWS',
                'values': values_tasks
            }
            cell_updater(f"{alphabet[i]}{data_so_new[0]['day']+2}",
                         value_range_body_tasks)

        # summing values of all tasks for one assignee
        values = (
            (f'=СУММ({alphabet[i]}3:{alphabet[i]}{n-2})', ''),
            ('', '')
        )
        value_range_body = {
            'majorDimension': 'ROWS',
            'values': values
        }
        cell_updater(f'{alphabet[i]}{n}', value_range_body)
        i += 1


def float_tasks(a):
    '''Checks whether type of tag is not str'''

    try:
        return float(a['name'].replace(',', '.'))
    except ValueError:
        pass


def cell_days():
    """Creates days in table referring to number of days of current month

    Returns: 
        k: number of days 
    """

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

    # adding more days whether number of days of current month > 28
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

    # adding "Сумма" in the of the list
    if i > (q_days[1] + 2):
        i = i-1
        values_sum = (
            ('', ''),
            ('Сума', '')
        )
        value_range_body_sum = {
            'majorDimension': 'ROWS',
            'values': values_sum
        }
        cell_appender(f'A{i}', value_range_body_sum)
        k = k + 4
        format_days(k, sheet_id, service, spreadsheet_id)
        return k


tasks_json()

create_table(now, service, spreadsheet_id)

spreadsheet = service.spreadsheets().get(
    spreadsheetId=spreadsheet_id).execute()
for sheet in spreadsheet['sheets']:
    if sheet['properties']['title'] == f'{worksheet_name}':
        sheet_id = sheet['properties']['sheetId']


update_table(sheet_id, service, spreadsheet_id)

tasks_spreadsheet()
