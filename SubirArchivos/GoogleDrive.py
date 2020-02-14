import os
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from apiclient.http import MediaFileUpload

SCOPES = ['https://www.googleapis.com/auth/drive']

def authentication():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('drive', 'v3', credentials=creds)
    return service

def update_file(service, nombre, tipo, idGD):
    drive_service = service
    file_metadata = {'name': nombre}
    media = MediaFileUpload(nombre , mimetype=tipo, resumable = True)
    file = drive_service.files().update(fileId= idGD, body=file_metadata, media_body=media, fields='id').execute()