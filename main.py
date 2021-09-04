import argparse
import glob
import mimetypes
import os
import time
from enum import Enum
from threading import Timer
from typing import Tuple

from tqdm import tqdm
from google.auth.transport.requests import Request
from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import HttpRequest, MediaFileUpload, MediaIoBaseDownload
from watchdog.events import FileSystemEvent, PatternMatchingEventHandler
from watchdog.observers import Observer


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
            ],
        )
    )
    _BASE_FOLDER_NAME = "GDrive Conversions"

    def __init__(self):
        self._creds = self.__get_creds()
        self._drive = build("drive", "v3", credentials=self._creds)

    def __get_creds(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        token_path = os.path.join(os.path.dirname(__file__), "token.json")
        credentials_path = os.path.join(os.path.dirname(__file__), "credentials.json")
        if os.path.exists(token_path):
            creds = Credentials.from_authorized_user_file(token_path, Drive.SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                # cause creds.refresh doesn't work for some reason.
                # faster to just delete the token and request another
                try:
                    creds.refresh(Request())
                except RefreshError:
                    os.remove(token_path)
                    return self.__get_creds()
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    credentials_path, Drive.SCOPES
                )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(token_path, "w") as token:
                token.write(creds.to_json())
        return creds

    def close(self):
        self._drive.close()

    def get_recent_files(self, num: int, fields: Tuple[str, ...] = ("id", "name")):
        """
        fields options https://developers.google.com/drive/api/v3/reference/files
        """
        fields_string = ", ".join(fields)
        results: dict = (
            self._drive.files()
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
                self._drive.files()
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
        query = f"name = '{Drive._BASE_FOLDER_NAME}' and mimeType = '{Drive.GoogleWorkspaceMimetypes.FOLDER.value}'"
        result = self.__search_file(query)
        if result:
            folder: dict = result[0]
            return folder.get("id")
        else:
            file_metadata = {
                "name": Drive._BASE_FOLDER_NAME,
                "mimeType": Drive.GoogleWorkspaceMimetypes.FOLDER.value,
            }
            folder: dict = (
                self._drive.files().create(body=file_metadata, fields="id").execute()
            )
            return folder.get("id")

    def upload_file(self, filepath: str, mimetype: str = None):
        folder_id = self.__get_base_folder()
        _, filename = os.path.split(filepath)
        file_metadata = {"name": filename, "parents": [folder_id]}
        if mimetype is not None:
            file_metadata["mimeType"] = mimetype
        base_mimetype = mimetypes.guess_type(filepath)[0]
        media = MediaFileUpload(filepath, mimetype=base_mimetype)
        uploaded_file: dict = (
            self._drive.files()
            .create(body=file_metadata, media_body=media, fields="id")
            .execute()
        )
        return uploaded_file.get("id")

    def delete_file(self, file_id: str):
        self._drive.files().delete(fileId=file_id).execute()

    def convert_file(self, filepath: str, delete: bool = True):
        basename, ext = os.path.splitext(filepath)
        mimetype = None
        mimetype_conversion = {
            (".doc", ".docx"): Drive.GoogleWorkspaceMimetypes.DOCS.value,
            (".ppt", ".pptx"): Drive.GoogleWorkspaceMimetypes.SLIDES.value,
        }
        for k, v in mimetype_conversion.items():
            if ext in k:
                mimetype = v
                break
        uploaded_file_id = self.upload_file(filepath, mimetype)

        request: HttpRequest = self._drive.files().export_media(
            fileId=uploaded_file_id, mimeType="application/pdf"
        )
        with open(f"{basename}.pdf", "wb") as f:
            with tqdm(total=1.0, leave=False) as pbar:
                downloader = MediaIoBaseDownload(f, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                    pbar.update(status.progress())

        if delete:
            self.delete_file(uploaded_file_id)


class GDriveEventHandler(PatternMatchingEventHandler):
    TIMER_INTERVAL = 60

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.__drive = None
        self.__drive_timer = None

    def __close_drive(self):
        if self.__drive:
            self.__drive.close()
            self.__drive = None

    def __stop_timer(self):
        if self.__drive_timer is not None:
            # Kill timer thread if it is still ongoing
            if self.__drive_timer.is_alive():
                self.__drive_timer.cancel()
            # Wait for thread to terminate
            self.__drive_timer.join()
            self.__drive_timer = None

    def __restart_drive_timer(self):
        self.__stop_timer()
        self.__drive_timer = Timer(
            GDriveEventHandler.TIMER_INTERVAL, self.__close_drive
        )
        self.__drive_timer.start()

    @property
    def drive(self):
        self.__restart_drive_timer()
        if not self.__drive:
            self.__drive = Drive()
        return self.__drive

    def on_created(self, event: FileSystemEvent):
        print(event.src_path)
        self.drive.convert_file(event.src_path)

    def shutdown(self):
        self.__stop_timer()
        self.__close_drive()


class Watcher:
    def __init__(self, event_handler, path=".") -> None:
        self._event_handler = event_handler
        self.observer = Observer()

        self.path = path
        self.observer.schedule(self._event_handler, self.path, True)

    def start(self):
        self.observer.start()
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self._event_handler.shutdown()
            self.observer.stop()
        self.observer.join()


def watch_dir(dir_path):
    event_handler = GDriveEventHandler(patterns=["*.doc", "*.docx", "*.ppt", "*.pptx"])
    watcher = Watcher(event_handler, path=dir_path)
    watcher.start()


def convert_files(files):
    FILE_TYPES = (".doc", ".docx", ".ppt", ".pptx")
    drive = Drive()
    for f in files:
        filepath = os.path.abspath(os.path.expanduser(f))
        if os.path.isdir(filepath):
            search_globs = list(map(lambda x: f"{filepath}/**/*{x}", FILE_TYPES))
            filepaths = [
                result
                for search_glob in search_globs
                for result in glob.glob(search_glob, recursive=True)
            ]
            if not len(filepaths):
                return
            confirm = input(
                f"There are {len(filepaths)} files to convert, are you sure? [y/n] \n"
            )
            if confirm in ("y", "Y"):
                for i in tqdm(filepaths):
                    drive.convert_file(i)

        else:
            if os.path.splitext(filepath)[1] not in FILE_TYPES:
                print(f"{filepath} has an invalid extension. Skipping file ...")
            else:
                drive.convert_file(filepath)
    drive.close()


def create_parser():
    parser = argparse.ArgumentParser(
        description="Use Google Drive to convert certain filetypes to pdf."
    )
    parser.add_argument("-w", "--w", dest="watch", action="store_true")
    parser.add_argument(
        "files",
        metavar="file",
        nargs="*",
        help="list of files to convert or a dir to watch",
    )
    return parser


if __name__ == "__main__":
    # TODO Consider some form of logging
    parser = create_parser()
    args = parser.parse_args()
    if args.watch:
        if len(args.files) == 1 and os.path.isdir(os.path.expanduser(args.files[0])):
            watch_dir(os.path.expanduser(args.files[0]))
        elif len(args.files) == 0:
            watch_dir(os.getcwd())
        else:
            print("Invalid args. Only accepts 1 directory to watch for new files.")
    else:
        if args.files:
            convert_files(args.files)
        else:
            print("No files given.")
