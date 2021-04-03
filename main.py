import time
import json
import os
import datetime
import string
from calendar import monthrange

import asana
from Google import Create_Service

from formatting import create_table, update_table, merge_columns, format_name, format_scnd_header, format_bold, format_days, format_tasks, format_time_tasks
from personal_data import personal_access_token, spreadsheet_id, project_ids


# pylint: disable=maybe-no-member

CLIENT_SECRET_SERVICE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


service = Create_Service(CLIENT_SECRET_SERVICE, API_NAME, API_VERSION, SCOPES)
client = asana.Client.access_token(personal_access_token)


def get_sheet_id():
    worksheet_name = datetime.datetime.now().strftime("%B %Y")
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id).execute()
    for sheet in spreadsheet['sheets']:
        if sheet['properties']['title'] == f'{worksheet_name}':
            return sheet['properties']['sheetId']


def download_project_tasks(project_id):
    now = datetime.datetime.now()
    fields = ['name', 'assignee.name', 'completed',
              'completed_at', 'tags.name', 'subtasks']
    tasks = client.tasks.get_tasks_for_project(
        project_id,
        {'completed_since': f'{now.year}-{now.month:02d}-01T02:00:00.000Z'},
        opt_pretty=True,
        opt_fields=fields
    )

    return list(tasks)


def download_project_subtasks(task):
    now = datetime.datetime.now()
    fields = [
        'name', 'assignee.name', 'completed', 'completed_at', 'tags.name']
    subtasks = client.tasks.get_subtasks_for_task(
        task,
        {'completed_since': f'{now.year}-{now.month:02d}-01T02:00:00.000Z'},
        opt_pretty=True,
        opt_fields=fields
    )

    return list(subtasks)


def dump_tasks(tasks, file_path):
    with open(file_path, "w", encoding="utf-8") as output:
        json.dump(tasks, output, ensure_ascii=False)


def download_tasks(project_ids):
    return [
        task
        for project_id in project_ids
        for task in download_project_tasks(project_id)
    ]


def download_subtasks(tasks):
    subtasks = [
        subtask
        for task in tasks
        for subtask in download_project_subtasks(tasks)
    ]

    for i, task in enumerate(tasks):
        if task['subtasks']:
            subtasks += download_project_subtasks(task['gid'])

            if not i % 149:
                time.sleep(60.0)

    return subtasks


def get_tasks(project_ids):
    """Creates json file and downloads there data about tasks and subtasks in Asana. """
    # pylint: disable=maybe-no-member
    tasks = download_tasks(project_ids)
    tasks += download_subtasks(tasks)
    dump_tasks(tasks, 'meta.json')


def update_cells(cell_range_insert, value_range_body):
    """Updates data in cells

    Args:
        cell_range_insert (str): starting cell range
        value_range_body (dict): values to download
    """

    # pylint: disable=maybe-no-member
    worksheet_name = datetime.datetime.now().strftime("%B %Y")
    time.sleep(1.1)
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption='USER_ENTERED',
        range=worksheet_name + "!" + cell_range_insert,
        body=value_range_body).execute()


def create_28_days():
    values_time = (
        tuple(map(str, range(1, 29))),
        ('', '')
    )
    value_range_body_time = {
        'majorDimension': 'COLUMNS',
        'values': values_time
    }

    update_cells(f'A3', value_range_body_time)


def create_remaining_days(q_days, min_days, last_day):
    while min_days <= q_days:
        last_day += 1
        values = (
            (f'{last_day}', ''),
            ('', '')
        )
        value_range_body = {
            'majorDimension': 'ROWS',
            'values': values
        }
        update_cells(f'A{min_days}', value_range_body)
        min_days += 1
    last_day += 4
    return last_day


def add_sum_cell(position):
    values_sum = (
        ('Сума', ''),
        ('', '')
    )
    value_range_body_sum = {
        'majorDimension': 'ROWS',
        'values': values_sum
    }

    update_cells(f'A{position}', value_range_body_sum)


def create_days_of_month():
    """Creates days in table referring to number of days of current month

    Returns:
        last_day: position in sheet of the last day of month
    """
    # pylint: disable=maybe-no-member

    create_28_days()
    now = datetime.datetime.now()
    q_days = monthrange(now.year, now.month)[1]+2
    last_day = 28
    min_days = 31

    last_day = create_remaining_days(q_days, min_days, last_day)

    add_sum_cell(last_day)
    format_days(last_day, sheet_id, service, spreadsheet_id)
    return last_day


def download_from_json(file_path):
    # downloading data from json
    with open(file_path, 'r', encoding="utf-8") as f:
        data = json.load(f)
        data = [k for k in data if not (k['completed_at'] is None)]
        data = sorted(data, key=lambda k: k['completed_at'])
        return data


def download_names(data):
    names_list = []
    for item in data:
        if data[item]['assignee'] is not None:
            name = data[item]['assignee']['name']
            names_list.append(name)
    names_list = sorted(list(set(names_list)))
    return names_list


def update_positions_for_assignee(assignee):

    letters = ['B', 'D', 'F', 'H', 'J', 'L',
               'N', 'P', 'R', 'T', 'V', 'X', 'Z']
    alphabet = [string.ascii_uppercase]

    if len(assignee) <= len(letters):
        return letters
    else:
        new_letters = []
        number_iterations = list(range(len(assignee) // len(letters)))

        for element in number_iterations:
            for letter in letters:
                new_element = f'{alphabet[element]}'+f'{letter}'
                new_letters.append(new_element)
        letters += new_letters
        return letters


def format_table(startrow_index, sheet_id, service, spreadsheet_id):
    merge_columns(startrow_index, sheet_id, service, spreadsheet_id)
    format_bold(startrow_index, sheet_id, service, spreadsheet_id)
    format_tasks(startrow_index, sheet_id, service, spreadsheet_id)
    format_time_tasks(startrow_index, sheet_id, service, spreadsheet_id)


def float_tasks(a):
    '''Checks whether type of tag is not str'''

    try:
        return float(a['name'].replace(',', '.'))
    except ValueError:
        pass


def add_day(data_sorted, task):
    date = data_sorted[task]['completed_at']
    date = list(date)
    if date[8] == '0':
        day = int(date[9])
    else:
        day = int(date[8] + date[9])
    # adding new tag "day"
    data_sorted[task]['day'] = day


def sort_task(data, name):
    data_sorted = [k for k in data if (
        k['assignee'] is not None
        and k['assignee']['name'] == name
        and k['completed_at'] is not None)]

    for task in data_sorted:
        data_sorted[task]['tags'] = [k for k in data_sorted[task]['tags']
                                     if float_tasks(k) is not None]
        add_day(data_sorted, task)
    return data_sorted


def update_task_day(task_name, description, position_letter, task_index, data_day):
    task_name = str(task_name).replace('.', ',')
    values_tasks = (
        (f'{task_name}',
            f"{description}"),
        ('', '')
    )
    value_range_body_tasks = {
        'majorDimension': 'ROWS',
        'values': values_tasks
    }
    update_cells(f"{position_letter[task_index]}{data_day[0]['day']+2}",
                 value_range_body_tasks)


def check_tag(data_sorted, task_index, position_letter):
    data_sorted = [k for k in data_sorted if len(k['tags']) != 0]
    while len(data_sorted) >= 1:
        data_day = [k for k in data_sorted if (
            k['day'] == data_sorted[0]['day'])]
        data_sorted = [k for k in data_sorted if k not in data_day]
        task_name = float(data_day[0]['tags'][0]['name'].replace(',', '.'))
        description = f"{data_day[0]['tags'][0]['name'].replace('.',',')} - {data_day[0]['name']}\n"

        if len(data_day) > 1:
            # updating task if there are more than one tag
            for count in range(len(data_day[1:])):
                task_name += float(data_day[1:][count]['tags']
                                   [0]['name'].replace(',', '.'))
                description += f"{data_day[1:][count]['tags'][0]['name'].replace('.',',')} - {data_day[1:][count]['name']}\n"

        update_task_day(task_name, description,
                        position_letter, task_index, data_day)


def sum_assignee_tasks(task_index, position_letter, position_number):
    values = (
        (f'=СУММ({position_letter[task_index]}3:{position_letter[task_index]}{position_number-2})', ''),
        ('', '')
    )
    value_range_body = {
        'majorDimension': 'ROWS',
        'values': values
    }
    update_cells(
        f'{position_letter[task_index]}{position_number}', value_range_body)


def preparing_header(name, position_letter, task_index):
    values = (
        (f'{name}', ''),
        ('час', 'задачі')
    )
    value_range_body = {
        'majorDimension': 'ROWS',
        'values': values
    }
    update_cells(f"{position_letter[task_index]}1", value_range_body)


def update_by_name(data, names_list, position_letter, position_number):
    task_index = 0
    startrow_index = 1
    for name in names_list:
        preparing_header(name, position_letter, task_index)
        format_table(startrow_index, sheet_id, service, spreadsheet_id)

        startrow_index += 2

        data_sorted = sort_task(data, name)
        check_tag(data_sorted, task_index, position_letter)
        sum_assignee_tasks(task_index, position_letter, position_number)

        task_index += 1


def update_tasks_table():
    """Updates cells in table according to data in json file"""
    position_number = create_days_of_month()

    data = download_from_json('meta.json')

    # downloading and formatting names of assignee
    names_list = download_names(data)
    position_letter = update_positions_for_assignee(names_list)

    format_name(sheet_id, service, spreadsheet_id)
    format_scnd_header(sheet_id, service, spreadsheet_id)

    update_by_name(data, names_list, position_letter, position_number)


if __name__ == "__main__":
    get_tasks(project_ids)
    create_table(service, spreadsheet_id)
    sheet_id = get_sheet_id()
    update_table(sheet_id, service, spreadsheet_id)
    update_tasks_table()
