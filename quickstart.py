import os.path
from typing import Tuple
from enum import Enum
import mimetypes
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.http import MediaFileUpload


class Drive:
    class GoogleWorkspaceMimetypes(Enum):
        DOCS = "application/vnd.google-apps.document"
        FILE = "application/vnd.google-apps.file"
        FOLDER = "application/vnd.google-apps.folder"
        SLIDES = "application/vnd.google-apps.presentation"
        SHEETS = "application/vnd.google-apps.spreadsheet"

    # If modifying these scopes, delete the file token.json.
    SCOPES = list(
        map(
            lambda x: "https://www.googleapis.com" + x,
            [
                "/auth/drive",
                "/auth/drive.appdata",
                "/auth/drive.file",
                "/auth/drive.install",
            ],
        )
    )
    BASE_FOLDER_NAME = "GDrive Conversions"

    def __init__(self):
        self.creds = self.__get_creds()
        self.drive = build("drive", "v3", credentials=self.creds)

    def __get_creds(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", Drive.SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", Drive.SCOPES
                )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())
        return creds

    def close(self):
        self.drive.close()

    def get_recent_files(self, num: int, fields: Tuple[str, ...] = ("id", "name")):
        """
        fields options https://developers.google.com/drive/api/v3/reference/files
        """
        fields_string = ", ".join(fields)
        results: dict = (
            self.drive.files()
            .list(pageSize=num, fields=f"nextPageToken, files({fields_string})")
            .execute()
        )
        files = results.get("files", [])
        return files

    def __search_file(self, query: str, fields: Tuple[str, ...] = ("id", "name")):
        """
        fields options https://developers.google.com/drive/api/v3/reference/files
        """
        fields_string = ", ".join(fields)
        page_token = None
        files = []
        while True:
            response = (
                self.drive.files()
                .list(
                    q=query,
                    spaces="drive",
                    fields=f"nextPageToken, files({fields_string})",
                    pageToken=page_token,
                )
                .execute()
            )
            for found_file in response.get("files", []):
                # Process change
                files.append(found_file)
            page_token = response.get("nextPageToken", None)
            if page_token is None:
                break
        return files

    def __get_base_folder(self):
        query = f"name = '{Drive.BASE_FOLDER_NAME}' and mimeType = '{Drive.GoogleWorkspaceMimetypes.FOLDER.value}'"
        result = self.__search_file(query)
        if result:
            folder: dict = result[0]
            return folder.get("id")
        else:
            file_metadata = {
                "name": Drive.BASE_FOLDER_NAME,
                "mimeType": Drive.GoogleWorkspaceMimetypes.FOLDER.value,
            }
            folder: dict = (
                self.drive.files().create(body=file_metadata, fields="id").execute()
            )
            return folder.get("id")

    def upload_file(self, name: str):
        folder_id = self.__get_base_folder()
        file_metadata = {"name": name, "parents": [folder_id]}
        media = MediaFileUpload(name, mimetype=mimetypes.guess_type(name)[0])
        uploaded_file: dict = (
            self.drive.files()
            .create(body=file_metadata, media_body=media, fields="id")
            .execute()
        )
        print(uploaded_file.get("id"))

    def delete_file(self, file_id: str):
        self.drive.files().delete(fileId=file_id).execute()


def main():
    drive = Drive()
    drive.upload_file("DS Syllabus 2021.doc")

    drive.close()


if __name__ == "__main__":
    main()
