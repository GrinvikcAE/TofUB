import io
import os.path
import pprint as pp
from time import sleep
import asyncio

import apiclient.discovery
import gspread
import httplib2
from bestconfig import Config
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from oauth2client.service_account import ServiceAccountCredentials

from result_table import write_to_table_result, write_to_table_individual, write_to_table_tasks


def create_service_account():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

    credentials = None
    config = Config('./keys/settings.json')
    CREDENTIALS_FILE = config.get('CREDENTIALS_FILE')
    TOKEN = config.get('TOKEN')

    if os.path.exists(f'./keys/{TOKEN}'):
        credentials = Credentials.from_authorized_user_file(f'./keys/{TOKEN}', SCOPES)

    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            credentials = ServiceAccountCredentials.from_json_keyfile_name(f'./keys/{CREDENTIALS_FILE}', SCOPES)
            with open('./keys/token.json', 'w') as token:
                token.write(credentials.to_json())
    try:
        httpAuth = credentials.authorize(httplib2.Http())
        return httpAuth
    except HttpError as error:
        print(F'An error with service account: {error}')
        return None


def create_new_table(http_auth):
    try:
        service = apiclient.discovery.build('sheets', 'v4', http=http_auth)
        spreadsheet = service.spreadsheets().create(body={
            'properties': {'title': 'Played_tasks', 'locale': 'ru_RU'},
            'sheets': [{'properties': {'sheetType': 'GRID',
                                       'sheetId': 0,
                                       'title': 'Played_tasks',
                                       'gridProperties': {'rowCount': 150, 'columnCount': 15}}}],
        }).execute()
        spreadsheetId = spreadsheet['spreadsheetId']

        config = Config('./keys/settings.json')
        EMAILS = config.get('EMAILS')
        for email in EMAILS:
            add_permission(http_auth, email, spreadsheetId)

        driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)

        print('https://docs.google.com/spreadsheets/d/' + spreadsheetId)

        ROOT_FOLDER_ID = config.get('ROOT_FOLDER_ID')

        driveService.files().update(
            fileId=spreadsheetId,
            addParents=ROOT_FOLDER_ID,
            removeParents=spreadsheet.get("parents"),
            fields="id, parents",
        ).execute()
    except HttpError as error:
        print(F'An error occurred: {error}')
        return None


def create_folder(http_auth):
    try:
        driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)
        file_metadata = {
            'name': 'TofUB',
            'mimeType': 'application/vnd.google-apps.folder'
        }

        file = driveService.files().create(body=file_metadata, fields='id'
                                           ).execute()
        pp.pprint(F'Folder ID: "{file.get("id")}".')

        config = Config('./keys/settings.json')
        EMAILS = config.get('EMAILS')
        for email in EMAILS:
            add_permission(http_auth, email, file.get("id"))

        print('https://drive.google.com/drive/folders/' + file.get('id'))
        # return file.get('id')
    except HttpError as error:
        print(F'An error with create folder: {error}')
        return None


def get_list_of_files(http_auth, q=None):
    try:
        if q is not None:
            q = f"name contains '{q}'"
        driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)
        results = driveService.files().list(pageSize=20, fields="nextPageToken, files(id, name, mimeType, permissions)",
                                            q=q).execute()
        return results
    except HttpError as error:
        print(F'An error with get list of files: {error}')
        return None


def delete_file(http_auth, file_id):
    try:
        driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)
        driveService.files().delete(fileId=file_id).execute()
    except HttpError as error:
        print(F'An error occurred: {error}')
        return None


def copy_table(http_auth, new_name):
    try:
        config = Config('./keys/settings.json')
        ORIGINAL_TABLE_ID = config.get('ORIGINAL_TABLE_ID')
        WORK_FOLDER_ID = config.get('WORK_FOLDER_ID')
        driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)
        copy = (driveService.files().copy(fileId=ORIGINAL_TABLE_ID,
                                          body={'name': new_name}, ).execute())
        driveService.files().update(
            fileId=copy.get("id"),
            addParents=WORK_FOLDER_ID,
            removeParents=copy.get("parents"),
            fields="id, parents",
        ).execute()
        return copy.get("id")
    except HttpError as error:
        print(F'An error occurred: {error}')
        return None


def add_permission(http_auth, email, file_id):
    try:
        driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)
        driveService.permissions().create(
            fileId=file_id,
            body={'type': 'user', 'role': 'writer', 'emailAddress': email},
            fields='id'
        ).execute()
    except HttpError as error:
        print(F'An error occurred: {error}')
        return None


def delete_permission(http_auth, email=None, q=None):
    try:
        driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)
        list_if_file = get_list_of_files(http_auth, q=q)
        for i in range(len(list_if_file['files'])):
            if list_if_file['files'][i]['mimeType'] == 'application/vnd.google-apps.spreadsheet':
                for j in range(len(list_if_file['files'][i]['permissions'])):
                    if list_if_file['files'][i]['permissions'][j]['emailAddress'] == email:
                        driveService.permissions().delete(fileId=list_if_file['files'][i]['id'],
                                                          permissionId=list_if_file['files'][i]['permissions'][j][
                                                              'id']).execute()
    except HttpError as error:
        print(F'An error occurred: {error}')
        return None


def sstart(http_auth, names_of_files, emails):
    try:
        for i in range(len(names_of_files)):
            spreadsheetId = copy_table(http_auth, names_of_files[i])
            add_permission(http_auth, emails[i], spreadsheetId)
            print(f'{emails[i]}: ' + 'https://docs.google.com/spreadsheets/d/' + spreadsheetId)
            print()
    except HttpError as error:
        print(F'An error occurred: {error}')
        return None


def check_cells(sheet):
    ready_cells = 0
    for cell in range(len(sheet)):
        if sheet[cell] != '0,00':
            ready_cells += 1
    return ready_cells


def write_score(sheet_values, action, score):
    config = Config('./keys/settings.json')
    STOP = config.get('STOP')
    match len(sheet_values):
        case 12:
            for k in range(1, len(sheet_values[3])):
                if sheet_values[3][k] not in STOP:
                    if action == 1:
                        score[f'{sheet_values[10][0]}'].append([sheet_values[3][k], 'Д',
                                                                float(sheet_values[10][action].replace(',', '.')) - 90])
                    elif action == 2:
                        score[f'{sheet_values[10][0]}'].append([sheet_values[3][k], 'О',
                                                                float(sheet_values[10][action].replace(',', '.')) - 60])
            for k in range(1, len(sheet_values[4])):
                if sheet_values[4][k] not in STOP:
                    if action == 1:
                        score[f'{sheet_values[11][0]}'].append([sheet_values[4][k], 'О',
                                                                float(sheet_values[11][action].replace(',', '.')) - 60])
                    elif action == 2:
                        score[f'{sheet_values[11][0]}'].append([sheet_values[4][k], 'Д',
                                                                float(sheet_values[11][action].replace(',', '.')) - 90])
        case 13:
            for k in range(1, len(sheet_values[3])):
                if sheet_values[3][k] not in STOP:
                    if action == 1:
                        score[f'{sheet_values[10][0]}'].append([sheet_values[3][k], 'Д',
                                                                float(sheet_values[10][action].replace(',', '.')) - 90])
                    elif action == 2:
                        score[f'{sheet_values[10][0]}'].append([sheet_values[3][k], 'Р',
                                                                float(sheet_values[10][action].replace(',', '.')) - 30])
                    elif action == 3:
                        score[f'{sheet_values[10][0]}'].append([sheet_values[3][k], 'О',
                                                                float(sheet_values[10][action].replace(',', '.')) - 60])
            for k in range(1, len(sheet_values[4])):
                if sheet_values[4][k] not in STOP:
                    if action == 1:
                        score[f'{sheet_values[11][0]}'].append([sheet_values[4][k], 'О',
                                                                float(sheet_values[11][action].replace(',', '.')) - 60])
                    elif action == 2:
                        score[f'{sheet_values[11][0]}'].append([sheet_values[4][k], 'Д',
                                                                float(sheet_values[11][action].replace(',', '.')) - 90])
                    elif action == 3:
                        score[f'{sheet_values[11][0]}'].append([sheet_values[4][k], 'Р',
                                                                float(sheet_values[11][action].replace(',', '.')) - 30])
            for k in range(1, len(sheet_values[5])):
                if sheet_values[5][k] not in STOP:
                    if action == 1:
                        score[f'{sheet_values[12][0]}'].append([sheet_values[5][k], 'Р',
                                                                float(sheet_values[12][action].replace(',', '.')) - 30])
                    elif action == 2:
                        score[f'{sheet_values[12][0]}'].append([sheet_values[5][k], 'О',
                                                                float(sheet_values[12][action].replace(',', '.')) - 60])
                    elif action == 3:
                        score[f'{sheet_values[12][0]}'].append([sheet_values[5][k], 'Д',
                                                                float(sheet_values[12][action].replace(',', '.')) - 90])
        case 14:
            for k in range(1, len(sheet_values[3])):
                if sheet_values[3][k] not in STOP:
                    if action == 1:
                        score[f'{sheet_values[10][0]}'].append([sheet_values[3][k], 'Д',
                                                                float(sheet_values[10][action].replace(',', '.')) - 90])
                    elif action == 3:
                        score[f'{sheet_values[10][0]}'].append([sheet_values[3][k], 'Р',
                                                                float(sheet_values[10][action].replace(',', '.')) - 30])
                    elif action == 4:
                        score[f'{sheet_values[10][0]}'].append([sheet_values[3][k], 'О',
                                                                float(sheet_values[10][action].replace(',', '.')) - 60])
            for k in range(1, len(sheet_values[4])):
                if sheet_values[4][k] not in STOP:
                    if action == 1:
                        score[f'{sheet_values[11][0]}'].append([sheet_values[4][k], 'О',
                                                                float(sheet_values[11][action].replace(',', '.')) - 60])
                    elif action == 2:
                        score[f'{sheet_values[11][0]}'].append([sheet_values[4][k], 'Д',
                                                                float(sheet_values[11][action].replace(',', '.')) - 90])
                    elif action == 4:
                        score[f'{sheet_values[11][0]}'].append([sheet_values[4][k], 'Р',
                                                                float(sheet_values[11][action].replace(',', '.')) - 30])
            for k in range(1, len(sheet_values[5])):
                if sheet_values[5][k] not in STOP:
                    if action == 1:
                        score[f'{sheet_values[12][0]}'].append([sheet_values[5][k], 'Р',
                                                                float(sheet_values[12][action].replace(',', '.')) - 30])
                    elif action == 2:
                        score[f'{sheet_values[12][0]}'].append([sheet_values[5][k], 'О',
                                                                float(sheet_values[12][action].replace(',', '.')) - 60])
                    elif action == 3:
                        score[f'{sheet_values[12][0]}'].append([sheet_values[5][k], 'Д',
                                                                float(sheet_values[12][action].replace(',', '.')) - 90])
            for k in range(1, len(sheet_values[6])):
                if sheet_values[6][k] not in STOP:
                    if action == 2:
                        score[f'{sheet_values[13][0]}'].append([sheet_values[6][k], 'Р',
                                                                float(sheet_values[13][action].replace(',', '.')) - 30])
                    elif action == 3:
                        score[f'{sheet_values[13][0]}'].append([sheet_values[6][k], 'О',
                                                                float(sheet_values[13][action].replace(',', '.')) - 60])
                    elif action == 4:
                        score[f'{sheet_values[13][0]}'].append([sheet_values[6][k], 'Д',
                                                                float(sheet_values[13][action].replace(',', '.')) - 90])

    return score


def check_file(http_auth, q=None, name_files=None):
    try:
        service = apiclient.discovery.build('sheets', 'v4', http=http_auth)

        list_if_file = get_list_of_files(http_auth, q=q)

        config = Config('./keys/settings.json')
        EMAILS = config.get('EMAILS')
        for i in range(len(list_if_file['files'])):
            if (list_if_file['files'][i]['mimeType'] == 'application/vnd.google-apps.spreadsheet'
                    and list_if_file['files'][i]['name'] not in name_files):

                print(f"Checking file: {list_if_file['files'][i]['name']}")

                ranges = ["Действие_1!B24:I37"]
                result = service.spreadsheets().values().batchGet(spreadsheetId=list_if_file['files'][i]['id'],
                                                                  ranges=ranges,
                                                                  valueRenderOption='FORMATTED_VALUE',
                                                                  dateTimeRenderOption='FORMATTED_STRING').execute()
                sheet_values = result['valueRanges'][0]['values']
                score = {}
                if sheet_values[2][7] == 'TRUE':
                    print(f"Need to stop {list_if_file['files'][i]['name']}")
                    sleep(10)
                    if sheet_values[1][7] == '2':
                        results = [sheet_values[10][1], sheet_values[10][2], sheet_values[11][1], sheet_values[11][2]]
                        if check_cells(results) == 4:
                            for j in range(len(list_if_file['files'][i]['permissions'])):
                                if list_if_file['files'][i]['permissions'][j]['emailAddress'] not in EMAILS:
                                    delete_permission(http_auth,
                                                      email=list_if_file['files'][i]['permissions'][j]['emailAddress'],
                                                      q=list_if_file['files'][i]['name'])
                            score = {f'{sheet_values[10][0]}': [float(sheet_values[10][-3].replace(',', '.')),
                                                                int(sheet_values[10][-1].replace(',', '.'))],
                                     f'{sheet_values[11][0]}': [float(sheet_values[11][-3].replace(',', '.')),
                                                                int(sheet_values[11][-1].replace(',', '.'))]}
                            score = write_score(sheet_values, 1, score)

                            ranges = ["Действие_2!B24:I37"]
                            result = service.spreadsheets().values().batchGet(
                                spreadsheetId=list_if_file['files'][i]['id'],
                                ranges=ranges,
                                valueRenderOption='FORMATTED_VALUE',
                                dateTimeRenderOption='FORMATTED_STRING').execute()
                            sheet_values = result['valueRanges'][0]['values']

                            score = write_score(sheet_values, 2, score)

                        else:
                            print(f"Not time to stop {list_if_file['files'][i]['name']}")

                    elif sheet_values[1][7] == '3':
                        results = [sheet_values[10][1], sheet_values[10][2], sheet_values[10][3],
                                   sheet_values[11][1], sheet_values[11][2], sheet_values[11][3],
                                   sheet_values[12][1], sheet_values[12][2], sheet_values[12][3]]
                        if check_cells(results) == 9:
                            for j in range(len(list_if_file['files'][i]['permissions'])):
                                if list_if_file['files'][i]['permissions'][j]['emailAddress'] not in EMAILS:
                                    delete_permission(http_auth,
                                                      email=list_if_file['files'][i]['permissions'][j]['emailAddress'],
                                                      q=list_if_file['files'][i]['name'])
                            score = {f'{sheet_values[10][0]}': [float(sheet_values[10][-3].replace(',', '.')),
                                                                int(sheet_values[10][-1].replace(',', '.'))],
                                     f'{sheet_values[11][0]}': [float(sheet_values[11][-3].replace(',', '.')),
                                                                int(sheet_values[11][-1].replace(',', '.'))],
                                     f'{sheet_values[12][0]}': [float(sheet_values[12][-3].replace(',', '.')),
                                                                int(sheet_values[12][-1].replace(',', '.'))]
                                     }
                            score = write_score(sheet_values, 1, score)

                            ranges = ["Действие_2!B24:I37"]
                            result = service.spreadsheets().values().batchGet(
                                spreadsheetId=list_if_file['files'][i]['id'],
                                ranges=ranges,
                                valueRenderOption='FORMATTED_VALUE',
                                dateTimeRenderOption='FORMATTED_STRING').execute()
                            sheet_values = result['valueRanges'][0]['values']
                            score = write_score(sheet_values, 2, score)

                            ranges = ["Действие_3!B24:I37"]
                            result = service.spreadsheets().values().batchGet(
                                spreadsheetId=list_if_file['files'][i]['id'],
                                ranges=ranges,
                                valueRenderOption='FORMATTED_VALUE',
                                dateTimeRenderOption='FORMATTED_STRING').execute()
                            sheet_values = result['valueRanges'][0]['values']
                            score = write_score(sheet_values, 3, score)

                    elif sheet_values[1][7] == '4':
                        results = [sheet_values[10][1], sheet_values[10][3], sheet_values[10][4],
                                   sheet_values[11][1], sheet_values[11][2], sheet_values[11][4],
                                   sheet_values[12][1], sheet_values[12][2], sheet_values[12][3],
                                   sheet_values[13][2], sheet_values[13][3], sheet_values[13][4]]
                        if check_cells(results) == 12:
                            for j in range(len(list_if_file['files'][i]['permissions'])):
                                if list_if_file['files'][i]['permissions'][j]['emailAddress'] not in EMAILS:
                                    delete_permission(http_auth,
                                                      email=list_if_file['files'][i]['permissions'][j]['emailAddress'],
                                                      q=list_if_file['files'][i]['name'])
                            score = {f'{sheet_values[10][0]}': [float(sheet_values[10][-3].replace(',', '.')),
                                                                int(sheet_values[10][-1].replace(',', '.'))],
                                     f'{sheet_values[11][0]}': [float(sheet_values[11][-3].replace(',', '.')),
                                                                int(sheet_values[11][-1].replace(',', '.'))],
                                     f'{sheet_values[12][0]}': [float(sheet_values[12][-3].replace(',', '.')),
                                                                int(sheet_values[12][-1].replace(',', '.'))],
                                     f'{sheet_values[13][0]}': [float(sheet_values[13][-3].replace(',', '.')),
                                                                int(sheet_values[13][-1].replace(',', '.'))]
                                     }

                            score = write_score(sheet_values, 1, score)

                            ranges = ["Действие_2!B24:I37"]
                            result = service.spreadsheets().values().batchGet(
                                spreadsheetId=list_if_file['files'][i]['id'],
                                ranges=ranges,
                                valueRenderOption='FORMATTED_VALUE',
                                dateTimeRenderOption='FORMATTED_STRING').execute()
                            sheet_values = result['valueRanges'][0]['values']
                            score = write_score(sheet_values, 2, score)

                            ranges = ["Действие_3!B24:I37"]
                            result = service.spreadsheets().values().batchGet(
                                spreadsheetId=list_if_file['files'][i]['id'],
                                ranges=ranges,
                                valueRenderOption='FORMATTED_VALUE',
                                dateTimeRenderOption='FORMATTED_STRING').execute()
                            sheet_values = result['valueRanges'][0]['values']
                            score = write_score(sheet_values, 3, score)

                            ranges = ["Действие_4!B24:I37"]
                            result = service.spreadsheets().values().batchGet(
                                spreadsheetId=list_if_file['files'][i]['id'],
                                ranges=ranges,
                                valueRenderOption='FORMATTED_VALUE',
                                dateTimeRenderOption='FORMATTED_STRING').execute()
                            sheet_values = result['valueRanges'][0]['values']
                            score = write_score(sheet_values, 4, score)

                    if len(score) > 0:
                        for command in score:
                            for result in range(2, len(score[command])):
                                if score[command][result][-1] < 0:
                                    score[command][result][-1] = 0

                        print('Upload results')

                        RESULT_ID = config.get('RESULT_ID')

                        file_id = RESULT_ID
                        driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)
                        request = driveService.files().export_media(fileId=file_id,
                                                                    mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                        filename = './data/Result.xlsx'
                        fh = io.FileIO(filename, 'wb')
                        downloader = MediaIoBaseDownload(fh, request)
                        done = False
                        while done is False:
                            status, done = downloader.next_chunk()
                            print("Download %d%%." % int(status.progress() * 100))

                        write_to_table_result(score, list_if_file['files'][i]['name'][4])

                        SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
                                  'https://www.googleapis.com/auth/drive']

                        CREDENTIALS_FILE = config.get('CREDENTIALS_FILE')
                        credentials = ServiceAccountCredentials.from_json_keyfile_name(f'./keys/{CREDENTIALS_FILE}',
                                                                                       SCOPES)
                        gc = gspread.authorize(credentials)
                        content = open('./data/Result.csv', 'r').read()
                        gc.import_csv(RESULT_ID, content)

                        print('Upload individual')

                        INDIVIDUAL_RESULTS_ID = config.get('INDIVIDUAL_RESULTS_ID')

                        file_id = INDIVIDUAL_RESULTS_ID
                        driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)
                        request = driveService.files().export_media(fileId=file_id,
                                                                    mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                        filename = './data/Individual_Results.xlsx'
                        fh = io.FileIO(filename, 'wb')
                        downloader = MediaIoBaseDownload(fh, request)
                        done = False
                        while done is False:
                            status, done = downloader.next_chunk()
                            print("Download %d%%." % int(status.progress() * 100))

                        write_to_table_individual(score, list_if_file['files'][i]['name'][4])

                        content = open('./data/Individual_Results.csv', 'r').read()
                        gc.import_csv(INDIVIDUAL_RESULTS_ID, content)

                        name_files.append(list_if_file['files'][i]['name'])

        return name_files

    except HttpError as error:
        print(F'An error occurred: {error}')
        return None


def check_tasks(http_auth, q=None):
    service = apiclient.discovery.build('sheets', 'v4', http=http_auth)
    list_if_file = get_list_of_files(http_auth, q=q)

    score = {}

    for i in range(len(list_if_file['files'])):
        if list_if_file['files'][i]['mimeType'] == 'application/vnd.google-apps.spreadsheet':

            ranges = ["Действие_1!I25"]
            result = service.spreadsheets().values().batchGet(spreadsheetId=list_if_file['files'][i]['id'],
                                                              ranges=ranges,
                                                              valueRenderOption='FORMATTED_VALUE',
                                                              dateTimeRenderOption='FORMATTED_STRING').execute()
            action = int(result['valueRanges'][0]['values'][0][0])

            for c in range(1, action + 1):
                ranges = [f"Действие_{c}!B27:D30"]
                result = service.spreadsheets().values().batchGet(spreadsheetId=list_if_file['files'][i]['id'],
                                                                  ranges=ranges,
                                                                  valueRenderOption='FORMATTED_VALUE',
                                                                  dateTimeRenderOption='FORMATTED_STRING').execute()
                sheet_values = result['valueRanges'][0]['values']

                ranges = [f"Действие_{c}!J9:L9"]
                result = service.spreadsheets().values().batchGet(spreadsheetId=list_if_file['files'][i]['id'],
                                                                  ranges=ranges,
                                                                  valueRenderOption='FORMATTED_VALUE',
                                                                  dateTimeRenderOption='FORMATTED_STRING').execute()
                try:
                    # print(result['valueRanges'][0]['values'])

                    refusal = ', '.join(result['valueRanges'][0]['values'][0])
                except KeyError as error:
                    # print(F'An error occurred: {error}')
                    refusal = ''

                ranges = [f"Действие_{c}!J5:L5"]
                result = service.spreadsheets().values().batchGet(spreadsheetId=list_if_file['files'][i]['id'],
                                                                  ranges=ranges,
                                                                  valueRenderOption='FORMATTED_VALUE',
                                                                  dateTimeRenderOption='FORMATTED_STRING').execute()
                task = int(result['valueRanges'][0]['values'][0][0])

                for j in range(len(sheet_values)):
                    if len(sheet_values[j]) == 2:
                        score[f'Действие_{c}_Д'] = [sheet_values[j][0], task, refusal]
                    elif len(sheet_values[j]) == 3:
                        score[f'Действие_{c}_О'] = [sheet_values[j][0], task]

            config = Config('./keys/settings.json')
            PLAYED_TASKS_ID = config.get('PLAYED_TASKS_ID')
            file_id = PLAYED_TASKS_ID
            driveService = apiclient.discovery.build('drive', 'v3', http=http_auth)
            request = driveService.files().export_media(fileId=file_id,
                                                        mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            filename = './data/Played_tasks.xlsx'
            fh = io.FileIO(filename, 'wb')
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while done is False:
                status, done = downloader.next_chunk()
                print("Download %d%%." % int(status.progress() * 100))

            write_to_table_tasks(score, q[-1])

            SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

            CREDENTIALS_FILE = config.get('CREDENTIALS_FILE')
            credentials = ServiceAccountCredentials.from_json_keyfile_name(f'./keys/{CREDENTIALS_FILE}', SCOPES)
            gc = gspread.authorize(credentials)

            content = open('./data/Played_tasks.csv', 'r').read()

            gc.import_csv(PLAYED_TASKS_ID, content)
    print('Check completed')

# http_auth = create_service_account()
# check_tasks(http_auth, q='Бой_3')
