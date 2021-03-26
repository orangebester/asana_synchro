import time
import json
import datetime

import asana
from Google import Create_Service

from formatting import create_table, update_table, merge_columns, format_name, format_scnd_header, format_bold, \
    format_days, format_tasks, format_time_tasks


def download_project_tasks(project_id):
    now = datetime.datetime.now()
    fields = ['name', 'assignee.name', 'completed', 'completed_at', 'tags.name', 'subtasks']
    tasks = client.tasks.get_tasks_for_project(
        project_id,
        {'completed_since': f'{now.year}-{now.month:02d}-01T02:00:00.000Z'},
        opt_pretty=True,
        opt_fields=fields
    )

    return list(tasks)


def download_project_subtasks(task_id):
    return []


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
    subtasks = []

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
    dump_tasks(task, file_path)
