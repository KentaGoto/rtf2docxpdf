import os
import shutil
import win32com
from win32com.client import *
from multiprocessing import Pool, freeze_support
from time import time


def all_files(directory):
    for root, dirs, files in os.walk(directory):
        for file in files:
            yield os.path.join(root, file)


def rtf2docxpdf(rtf_fullpath):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    rtf_fullpath = rtf_fullpath.replace("/", "\\")
    dirname = os.path.dirname(rtf_fullpath)
    current_file = os.path.basename(rtf_fullpath)
    fname, ext = os.path.splitext(current_file)
    os.chdir(dirname)
    doc = word.Documents.Open(rtf_fullpath)
    docx = dirname + '/' + fname + '.docx'
    pdf = dirname + '/' + fname + '.pdf'
    doc.SaveAs(docx, FileFormat=16)
    print(docx)
    doc.SaveAs(pdf, FileFormat=17)
    print(pdf)
    doc.Close()


def genarate_x(path):
    dirname = os.path.dirname(path)
    current_file = os.path.basename(path)
    fname, ext = os.path.splitext(current_file)
    os.chdir(dirname)
    if ext == '.rtf':
        rtf2docxpdf(dirname + '/' + current_file)
        os.remove(dirname + '/' + current_file)


if __name__ == '__main__':
    freeze_support()

    s = input("Dir: ")
    root_dir = s.strip('\"')
    root_dir_copy = root_dir + '__copy'
    shutil.copytree(root_dir, root_dir_copy)

    start = time()
    print('Processing...')

    files = list()
    for i in all_files(root_dir_copy):
        files.append(i)

    # multiprocessing
    with Pool(processes=None) as pool:
        pool.map(genarate_x, files)
        pool.close()

    print('Done!\n')
    print('{}s'.format(time() - start))
