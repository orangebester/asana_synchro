import asana
import time
import datetime
from calendar import monthrange


def create_table(service, spreadsheet_id):
    now = datetime.datetime.now()
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


def update_table(sheet_id, service, spreadsheet_id):
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


def merge_columns(i, sheet_id, service, spreadsheet_id):
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


def format_name(sheet_id, service, spreadsheet_id):
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


def format_scnd_header(sheet_id, service, spreadsheet_id):
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


def format_bold(i, sheet_id, service, spreadsheet_id):
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


def format_days(i, sheet_id, service, spreadsheet_id):
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


def format_tasks(j, sheet_id, service, spreadsheet_id):
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


def format_time_tasks(j, sheet_id, service, spreadsheet_id):
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
