import sys
import os
import subprocess
import pdfkit
import thread

PYTHON_INTERP = "D:\windows\programmes\Python27\python.exe"

options = {
    'encoding': "UTF-8",
    'quiet': ''
}


def convert_doc_to_pdf(in_file):
    print(in_file)
    subprocess.Popen([PYTHON_INTERP, './unoconv.py', "-f", "pdf", in_file], stdout=subprocess.PIPE)

def convert_chm(in_file):
    directory = os.path.splitext(in_file)[0]
    subprocess.call(["hh.exe", "-decompile", directory, in_file])

def change_ext_to_pdf(in_file):
    return os.path.splitext(in_file)[0]+".pdf"

def convert_web_to_pdf(in_file):
    pdfkit.from_file(in_file, change_ext_to_pdf(in_file), options=options)

def nope(in_file):
    pass

def main(convert):
    for dirname, dirnames, filenames in os.walk('.'):
        for filename in filenames:
            
            abs = os.path.abspath(dirname + "\\" + filename)
            try:
                convert.get(os.path.splitext(filename)[1], nope)(abs)
            except Exception, e:
                print("[ERROR] " + abs + " :" + e.message)

if __name__ == '__main__':

    first = {
        ".doc": convert_doc_to_pdf,
        ".docx": convert_doc_to_pdf,
        ".odt": convert_doc_to_pdf,
        ".pps": convert_doc_to_pdf,
        ".ppt": convert_doc_to_pdf,

        ".chm": convert_chm,
    }
    second = {
        ".htm": convert_web_to_pdf,
       # ".hhc": convert_web_to_pdf,
        ".html": convert_web_to_pdf,
    }
    main(first)
    main(second)
    print ("all work is done")
