import PyPDF2
import win32api
import win32print

import os
import word2pdf

import time
import log
import shutil
import datetime

from PyPDF2 import PdfFileReader, PdfFileWriter

logger = log.__init__()

def printer_loading(filename, mode):
    if not mode:
        print_single(filename)
    else:
        print_double(filename)


def print_single(filename):
    print_file(filename)


def print_double(filename):
    # word2pdf.convert_word_to_pdf(filename, os.path.split(filename)[0] + "temp\\" + os.path.split(filename)[1].split(".")[0] + ".pdf")
    if '.pdf' not in filename:
        for w in os.listdir("D:\\FTP"):
            try:
                if w.split(".")[1] == "错误警告":
                    os.rename("D:\\FTP\\" + w, "D:\\FTP\\" +
                          "双面打印目前仅支持pdf格式" + ".错误警告")
                    logger.error("not pdf file {0} detected in double printing process, terminate printing and display warn".format("D:\\FTP\\" + w))
                    break
            except:
                pass
    else:
        time.sleep(1)
        pdf_input = PdfFileReader(filename)
        pageCnt = pdf_input.getNumPages()
        logger.debug("file {0} with {1} pages".format(filename, pageCnt))
        logger.info("start writing into 2 pdfs")
        pdf_output_0 = PdfFileWriter()
        pdf_output_1 = PdfFileWriter()
        for i in range(0, pageCnt):
            if i == pageCnt - 1:
                pdf_output_1.addPage(pdf_input.getPage(i))
            else:
                if i % 2 == 0:
                    pdf_output_0.addPage(pdf_input.getPage(i))
                if i % 2 == 1:
                    pdf_output_1.addPage(pdf_input.getPage(i))
        logger.info("start generating 2 pdfs")
        with open(os.path.split(filename)[0] + "\\temp\\" + os.path.split(filename)[1].split(".")[0] + "_0.pdf", 'wb') as pdf_out:
            pdf_output_0.write(pdf_out)
        logger.debug("pdf 0 generated")
        with open(os.path.split(filename)[0] + "\\temp\\" + os.path.split(filename)[1].split(".")[0] + "_1.pdf", 'wb') as pdf_out:
            pdf_output_1.write(pdf_out)
        logger.debug("pdf 1 generated")
        time.sleep(1)
        logger.info("start printing pdf 1")
        print_file(os.path.split(filename)[0] + "\\temp\\" + os.path.split(filename)[1].split(".")[0] + "_1.pdf")
        logger.info("pdf 1 printed, reset hint file")
        for w in os.listdir("D:\\FTP"):
            try:
                if w.split(".")[1] == "双面打印提示":
                    os.rename("D:\\FTP\\" + w, "D:\\FTP\\" +
                          "完成操作后将此处改为继续" + ".双面打印提示")
                    logger.info("hint file reset to {0}".format("D:\\FTP\\" + "完成操作后将此处改为继续" + ".双面打印提示"))
                    break
            except:
                pass
        
def print_double_continue(filename):
    logger.info("start printing pdf 0")
    print_file(os.path.split(filename)[0] + "\\temp\\" + os.path.split(filename)[1].split(".")[0] + "_0.pdf")
    logger.info("pdf 0 printed")
    logger.info("all file printed, try to move all temp files")
    file_path = os.path.split(filename)[0] + "\\temp\\" + os.path.split(filename)[1].split(".")[0] + "_0.pdf"
    file_name = os.path.split(file_path)[1]
    shutil.move(file_path, "D:\\FTP\\printed\\" + "[" + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "]" + file_name)
    file_path = os.path.split(filename)[0] + "\\temp\\" + os.path.split(filename)[1].split(".")[0] + "_1.pdf"
    file_name = os.path.split(file_path)[1]
    shutil.move(file_path, "D:\\FTP\\printed\\" + "[" + datetime.datetime.now().strftime("%Y%m%d%H%M%S") + "]" + file_name)


def print_file(filename):
    open(filename, "r")
    win32api.ShellExecute(
        0,
        "logger.debug",
        filename,
        '/d:"%s"' % win32print.GetDefaultPrinter(),
        ".",
        0
    )
