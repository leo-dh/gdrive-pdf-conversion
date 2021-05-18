import time
import os
from threading import Timer

from watchdog.events import FileSystemEvent, PatternMatchingEventHandler
from watchdog.observers import Observer

from drive import Drive


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


def main():
    # TODO Consider some form of logging?
    # TODO add argparse to make it a CLI

    event_handler = GDriveEventHandler(patterns=["*.doc", "*.docx", "*.ppt", "*.pptx"])
    watcher = Watcher(event_handler, path=os.path.expanduser("."))
    watcher.start()


if __name__ == "__main__":
    main()
