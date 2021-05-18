from drive import Drive


def main():
    # TODO add watchdog to watch for new files and automatically convert them
    # TODO add argparse to make it a CLI
    drive = Drive()
    drive.convert_file("files/Lecture 1.pptx")
    drive.close()


if __name__ == "__main__":
    main()
