import asana
import json
import os
from Google import Create_Service

CLIENT_SECRET_SERVICE = 'client_secret.json'
API_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
personal_access_token = "1/1135096210698629:718f78c0f27970b442cd307018d87f8d"

#service = Create_Service(CLIENT_SECRET_SERVICE, API_NAME, API_VERSION, SCOPES)
client = asana.Client.access_token(personal_access_token)


def tasks_json():
    output_list = []
    project_gid = ['1198913045061016', '1198913045061019',
                   '1198913121709013', '1198913121709020', '1198913121709027', '1198913121709034', '1198913121709041', '1199104176594285']
    # pylint: disable=maybe-no-member
    for i in project_gid:
        result = client.tasks.get_tasks_for_project(i, opt_pretty=True, opt_fields=[
            'name', 'assignee.name', 'completed', 'completed_at',  'tags.name'])
        output_list += list(result)
    with open("meta.json", "w", encoding="utf-8") as output:
        json.dump(output_list, output, ensure_ascii=False)


tasks_json()
