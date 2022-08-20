from watchdog.observers import Observer
from watchdog.events import *
import time

import os
import re
import shutil

import print
import log
import datetime


# def printer_loading(filename):
#     open(filename, "r")
#     win32api.ShellExecute(
#         0,
#         "logger.debug",
#         filename,
#         '/d:"%s"' % win32print.GetDefaultPrinter(),
#         ".",
#         0
#     )


logger = log.__init__()


class FileEventHandler(FileSystemEventHandler):

    def __init__(self):
        FileSystemEventHandler.__init__(self)
        self.update_print_mode()
        self.update_printing_mode(0)
        self.update_printing_file("None.None")

    def update_print_mode(self):
        self.print_mode = os.path.exists("D:\\FTP\\双.打印模式")
        logger.debug("current print mode: {0}".format(self.print_mode))

    def update_printing_mode(self, to):
        self.printing_mode = to  # idle printing,single printing,double
        logger.debug("current printing mode: {0}".format(self.printing_mode))

    def update_printing_file(self, to):
        self.printing_file = to
        logger.debug("current printing file: {0}".format(self.printing_file))
    
    def on_moved(self, event):
        if event.is_directory:
            logger.debug("directory moved from {0} to {1}".format(
                event.src_path, event.dest_path))
        else:
            logger.debug("file moved from {0} to {1}".format(
                event.src_path, event.dest_path))
            if os.path.split(event.src_path)[1].split(".")[1] == "打印模式" and os.path.split(event.dest_path)[1].split(".")[1] == "打印模式":
                self.update_print_mode()
                logger.info("print mode changed to {0}".format(
                    os.path.split(event.dest_path)[1].split(".")[0]))
            if os.path.split(event.dest_path)[1] == "继续.双面打印提示" and self.printing_mode == 2:
                logger.info("continue printing double: {0}".format(self.printing_file))
                print.print_double_continue(self.printing_file)
                logger.info("file printed, try to move file")
                file_name = os.path.split(self.printing_file)[1]
                shutil.move(self.printing_file, "D:\\FTP\\printed\\" +
                                "[" + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "]" + file_name)
                logger.info("file moved")
                self.update_printing_mode(0)
                self.update_printing_file("None.None")

    def on_created(self, event):
        if event.is_directory:
            logger.debug("directory created: {0}".format(event.src_path))
        else:
            logger.debug("file created: {0}".format(event.src_path))
            logger.info(
                "starting to check and print file {0}".format(event.src_path))
            file_path = event.src_path
            if not os.path.split(file_path)[0] == "D:\\FTP":
                logger.info(
                    "{0} not in required file path, terminate printing".format(event.src_path))
            elif "$" in file_path:
                logger.info("banned character found in filename {0}, terminate printing".format(
                    event.src_path))
            else:
                logger.info("printing file {0}".format(event.src_path))
                self.update_printing_mode(self.print_mode + 1)
                self.update_printing_file(file_path)
                if self.print_mode == 1:
                    for w in os.listdir("D:\\FTP"):
                        try:
                            if w.split(".")[1] == "双面打印提示":
                                os.rename("D:\\FTP\\" + w, "D:\\FTP\\" +
                                    "完成操作后将此处改为继续" + ".双面打印提示")
                                logger.info("hint file reset to {0}".format("D:\\FTP\\" + "完成操作后将此处改为继续" + ".双面打印提示"))
                                break
                        except:
                            pass
                print.printer_loading(file_path, self.print_mode)
                if self.printing_mode == 1:
                    logger.info("file printed, try to move file")
                    file_name = os.path.split(file_path)[1]
                    shutil.move(file_path, "D:\\FTP\\printed\\" +
                                "[" + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "]" + file_name)
                    logger.info("file moved")
                    self.update_printing_mode(0)
                    self.update_printing_file("None.None")

    def on_deleted(self, event):
        if event.is_directory:
            logger.debug("directory deleted: {0}".format(event.src_path))
        else:
            logger.debug("file deleted: {0}".format(event.src_path))

    def on_modified(self, event):
        if event.is_directory:
            logger.debug("directory modified: {0}".format(event.src_path))
        else:
            logger.debug("file modified: {0}".format(event.src_path))


if __name__ == "__main__":
    observer = Observer()
    event_handler = FileEventHandler()
    observer.schedule(event_handler, r"D:\FTP", True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    logger.debug("FTP AutoPrint Started.")
