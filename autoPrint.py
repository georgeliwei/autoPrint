import win32print as wp
import win32api as wapi
import win32ui as wu
from win32com import client
from PIL import Image, ImageWin
import time
import os


def print_picture_file(fname):
    default_printer_name = wp.GetDefaultPrinter()
    if default_printer_name is None or default_printer_name == "":
        return

    hDC = wu.CreateDC()
    hDC.CreatePrinterDC(default_printer_name)
    printable_area = hDC.GetDeviceCaps(8), hDC.GetDeviceCaps(10)
    printer_size = hDC.GetDeviceCaps(110), hDC.GetDeviceCaps(111)
    printer_margins = hDC.GetDeviceCaps(112), hDC.GetDeviceCaps(113)

    img = Image.open(fname)
    if img.size[0] > img.size[1]:
        img = img.transpose(Image.ROTATE_90)
    ratios = [1.0 * printable_area[0] / img.size[0], 1.0 * printable_area[1] / img.size[1]]
    scale = min(ratios)
    hDC.StartDoc(fname)
    hDC.StartPage()
    dib = ImageWin.Dib(img)
    scaled_width, scaled_height = [int(scale * i) for i in img.size]
    x1 = int((printer_size[0] - scaled_width) / 2)
    y1 = int((printer_size[1] - scaled_height) / 2)
    x2 = x1 + scaled_width
    y2 = y1 + scaled_height
    dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))
    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()


def print_pdf_file(fname):
    default_printer_name = wp.GetDefaultPrinter()
    if default_printer_name is None or default_printer_name == "":
        return
    GHOSTSCRIPT_PATH = "D:\\software\\gs\\bin\\gswin64.exe"
    GSPRINT_PATH = "D:\\software\\gsprint\\gsprint.exe"
    cmd_parameter = '-dNOSAFER -dFitPage -ghostscript "%s" -printer "%s" %s' % (GHOSTSCRIPT_PATH, default_printer_name, fname)
    wapi.ShellExecute(
        0,
        'open',
        GSPRINT_PATH,
        cmd_parameter,
        ".",
        0)


def print_word_file(fname):
    dot_pos = fname.find(".")
    pdf_name = fname[0:dot_pos] + ".pdf"
    word_app = client.Dispatch("word.Application")
    word_app.Documents.Open(fname)
    word_app.ActiveDocument.SaveAs(pdf_name, FileFormat=client.constants.wdExportFormatPDF)
    word_app.ActiveDocument.Close()
    word_app.Quit()
    print_pdf_file(pdf_name)


def wait_printer_finish_work(wait_time):
    alread_wait_time = 0
    while True:
        time.sleep(5)
        alread_wait_time += 5
        if alread_wait_time >= wait_time:
            break


def find_print_file_list(printdir):
    print_log_f_name = "%s\\autoprint.log" % printdir
    print_log_f = open(print_log_f_name, "w")

    printed_log_f_name = "%s\\printedfile.log" % printdir

    g_printed_file_list = []
    try:
        printed_file_log_f = open(printed_log_f_name, "r")
        lines = printed_file_log_f.readlines()
        printed_file_log_f.close()
        g_printed_file_list = [item.strip() for item in lines]
    except:
        g_printed_file_list = []

    cur_time = time.localtime(time.time())
    cur_year = cur_time.tm_year
    cur_month = cur_time.tm_mon
    cur_day = cur_time.tm_mday

    print_log_f.write("[%s]Begin check print dir [%s]\n" % (time.asctime( time.localtime(time.time())), printdir))

    all_file_list = os.listdir(printdir)

    if len(all_file_list) <= 0:
        print_log_f.write("[%s]No file need print\n" % (time.asctime( time.localtime(time.time()))))
        print_log_f.close()
        return

    wait_print_file_list = []
    for file in all_file_list:
        format_file_name = file.lower()
        if not(".pdf" in format_file_name
               or ".jpg" in format_file_name
               or ".doc" in format_file_name
               or ".docx" in format_file_name
               or ".bmp" in format_file_name):
            continue
        cur_print_time = "%s_%s_%s" % (cur_year, cur_month, cur_day)
        if cur_print_time not in format_file_name:
            continue
        if file in g_printed_file_list:
            continue
        wait_print_file_list.append(file)

    if len(wait_print_file_list) > 5:
        print_log_f.write("[%s]Too much file need print, manual check\n" % (time.asctime(time.localtime(time.time()))))
        print_log_f.close()
        return

    for file in wait_print_file_list:
        print_log_f.write("[%s]Begin print file [%s]\n" % (time.asctime(time.localtime(time.time())), file))
        full_file_path = "%s\\%s" % (printdir, file)
        format_file_name = file.lower()
        if ".pdf" in format_file_name:
            print_pdf_file(full_file_path)
            g_printed_file_list.append(file)
            continue
        if ".jpg" in format_file_name or ".bmp" in format_file_name:
            print_picture_file(full_file_path)
            g_printed_file_list.append(file)
            continue
        if ".doc" in format_file_name or ".docx" in format_file_name:
            print_word_file(full_file_path)
            g_printed_file_list.append(file)
            continue

    printed_file_log_f = open(printed_log_f_name, "w")
    for item in g_printed_file_list:
        printed_file_log_f.write("%s\n" % item)
    printed_file_log_f.close()
    print_log_f.write("[%s]Finish printed file [%s]\n" % (time.asctime(time.localtime(time.time())),
                                                      "; ".join(wait_print_file_list)))
    print_log_f.close()


if __name__ == "__main__":
    print_dir = "E:\\OneDriveSpace\\OneDrive\\print"
    find_print_file_list(print_dir)
    time.sleep(60)



