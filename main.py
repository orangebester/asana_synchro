import asana
import time
import json
import os
import datetime
from calendar import monthrange
from Google import Create_Service
# pylint: disable=maybe-no-member

CLIENT_SECRET_SERVICE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

personal_access_token = "1/1135096210698629:718f78c0f27970b442cd307018d87f8d"
spreadsheet_id = '1wyOwn_21U_EtdHJTi7HJ-lGAc0_BsQQi40HelATS0lU'


service = Create_Service(CLIENT_SECRET_SERVICE, API_NAME, API_VERSION, SCOPES)
client = asana.Client.access_token(personal_access_token)
now = datetime.datetime.now()
worksheet_name = now.strftime("%B %Y")


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
    time.sleep(1.1)
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


def tasks_spreadsheet():
    startrow_index = 1
    n = cell_days()
    alphabet = ['B', 'D', 'F', 'H', 'J', 'L',
                'N', 'P', 'R', 'T', 'V', 'X', 'Z']
    with open("meta.json", 'r', encoding="utf-8") as f:
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
    format_name()
    format_scnd_header()
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
        merge_columns(startrow_index)
        format_bold(startrow_index)
        format_tasks(startrow_index)
        format_time_tasks(startrow_index)
        startrow_index += 2
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
            data_new[d]['day'] = day
        data_new = [k for k in data_new if len(k['tags']) != 0]
        while len(data_new) >= 1:
            data_so_new = [k for k in data_new if (
                k['day'] == data_new[0]['day'])]
            data_new = [k for k in data_new if k not in data_so_new]
            var1 = float(data_so_new[0]['tags'][0]['name'].replace(',', '.'))
            var2 = f"{data_so_new[0]['tags'][0]['name'].replace('.',',')} - {data_so_new[0]['name']}\n"
            if len(data_so_new) > 1:
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
    try:
        return float(a['name'].replace(',', '.'))
    except ValueError:
        pass


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
        format_days(k)
        return k


def create_table():
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


def update_table():
    request_body = {
        'requests': [
            {
                'updateSheetProperties': {
                    'properties': {
                        'sheetId': sheet_id,
                        'gridProperties': {
                            'frozenRowCount': 2
                        }
                    },
                    'fields': 'gridProperties.frozenRowCount'
                }
            }
        ]

    }

    # pylint: disable=maybe-no-member
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=request_body).execute()


def merge_columns(i):
    request_body = {
        "requests": [
            {
                "mergeCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": i,
                        "endColumnIndex": i+2
                    },
                    "mergeType": "MERGE_ALL"
                }
            }
        ]
    }
    # pylint: disable=maybe-no-member
    time.sleep(1.1)
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=request_body).execute()


def format_name():
    request_body = {"requests": [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.271,
                            "green": 0.506,
                            "blue": 0.557
                        },
                        "horizontalAlignment": "CENTER",
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1.0,
                                "green": 1.0,
                                "blue": 1.0
                            },
                            "fontSize": 10,
                            "bold": True
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            }
        }
    ]
    }
    # pylint: disable=maybe-no-member
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=request_body).execute()


def format_scnd_header():
    request_body = {"requests": [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 0.898,
                            "blue": 0.6
                        },
                        "horizontalAlignment": "CENTER",
                        "textFormat": {
                            "foregroundColor": {
                                "red": 0.0,
                                "green": 0.0,
                                "blue": 0.0
                            },
                            "fontSize": 10
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            }
        }
    ]
    }
    # pylint: disable=maybe-no-member
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=request_body).execute()


def format_bold(i):
    request_body = {"requests": [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 1,
                    "endRowIndex": 2,
                    "startColumnIndex": i,
                    "endColumnIndex": i+1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 1.0,
                            "green": 0.898,
                            "blue": 0.6
                        },
                        "horizontalAlignment": "CENTER",
                        "textFormat": {
                            "foregroundColor": {
                                "red": 0.0,
                                "green": 0.0,
                                "blue": 0.0
                            },
                            "fontSize": 10,
                            "bold": True
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
            }
        }
    ]
    }
    # pylint: disable=maybe-no-member
    time.sleep(1.1)
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=request_body).execute()


def format_days(i):
    request_body = {"requests": [
        {"updateDimensionProperties": {
            "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 1
                    },
            "properties": {
                "pixelSize": 43
            },
            "fields": "pixelSize"}
         },
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": 0,
                    "endRowIndex": i,
                    "startColumnIndex": 0,
                    "endColumnIndex": 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "CENTER",
                        "textFormat": {
                            "foregroundColor": {
                                "red": 0.6,
                                "green": 0.6,
                                "blue": 0.6
                            },
                            "fontSize": 10
                        }
                    }
                },
                "fields": "userEnteredFormat(textFormat,horizontalAlignment)"
            }
        }]
    }
    # pylint: disable=maybe-no-member
    time.sleep(1.1)
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=request_body).execute()


def format_tasks(j):
    request_body = {
        "requests": [
            {"updateDimensionProperties": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "COLUMNS",
                    "startIndex": j+1,
                    "endIndex": j+2
                },
                "properties": {
                    "pixelSize": 274
                },
                "fields": "pixelSize"}
             },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 2,
                        "endRowIndex": 35,
                        "startColumnIndex": j,
                        "endColumnIndex": j+2
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "LEFT",
                            "verticalAlignment": "MIDDLE",
                            "wrapStrategy": "WRAP",
                            "textFormat": {
                                "foregroundColor": {
                                    "red": 0.0,
                                    "green": 0.0,
                                    "blue": 0.0
                                },
                                "fontSize": 10
                            }
                        }
                    },
                    "fields": "userEnteredFormat(textFormat,horizontalAlignment, verticalAlignment, wrapStrategy)"
                }
            }
        ]
    }
    # pylint: disable=maybe-no-member
    time.sleep(1.1)
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=request_body).execute()


def format_time_tasks(j):
    request_body = {
        "requests": [
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": j,
                        "endIndex": j+1
                    },
                    "properties": {
                        "pixelSize": 54
                    },
                    "fields": "pixelSize"}
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 2,
                        "endRowIndex": 35,
                        "startColumnIndex": j,
                        "endColumnIndex": j+1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "verticalAlignment": "MIDDLE",
                            "wrapStrategy": "WRAP",
                            "textFormat": {
                                "foregroundColor": {
                                    "red": 0.0,
                                    "green": 0.0,
                                    "blue": 0.0
                                },
                                "fontSize": 10,
                                "bold": True
                            }
                        }
                    },
                    "fields": "userEnteredFormat(textFormat,horizontalAlignment, verticalAlignment, wrapStrategy)"
                }
            }
        ]
    }
    # pylint: disable=maybe-no-member
    time.sleep(1.1)
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body=request_body).execute()


create_table()

spreadsheet = service.spreadsheets().get(
    spreadsheetId=spreadsheet_id).execute()
for sheet in spreadsheet['sheets']:
    if sheet['properties']['title'] == f'{worksheet_name}':
        sheet_id = sheet['properties']['sheetId']

tasks_json()

update_table()

tasks_spreadsheet()
