import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


def get_service():
    """Создаёт и возвращает сервисный объект Google Sheets API."""
    creds_path = os.path.join(os.path.dirname(__file__), '..', 'google_sheets_key.json')

    creds = Credentials.from_service_account_file(creds_path, scopes=["https://www.googleapis.com/auth/spreadsheets"])

    service = build("sheets", "v4", credentials=creds)
    return service
