# coding: utf8
import os
import subprocess
import pdfkit
PYTHON_INTERP = "python"
				#"D:\windows\programmes\Python27\python.exe"

_7ZIP_PATH = "7za.exe"

PATH_CONV = r'D:\Telechargements\Mega Folder\unzip'
# r"/home/baptiste/Téléchargements/Mega Folder/unzip/"

error_list = []

verbose = False

options = {'encoding': "UTF-8", 'quiet': '', '--load-error-handling': 'ignore'} if not verbose else {'encoding': "UTF-8"}

def convert_ps_to_pdf(in_file):
    subprocess.call(["ps2pdf", "-dEPSCrop", in_file], stdout=subprocess.PIPE)

def convert_doc_to_pdf(in_file):
    subprocess.call([PYTHON_INTERP, './unoconv.py', "-f", "pdf", in_file], stdout=subprocess.PIPE)#I am not even going to read unoconv.py to make this more efficient :)

def convert_chm(in_file):
    directory = os.path.splitext(in_file)[0]
  #  subprocess.call([_7ZIP_PATH, "x",  in_file ])
    subprocess.call(["hh.exe", "-decompile", directory, in_file], stdout=subprocess.PIPE) # in fact 7zip is doing a far better job at decompressing chm files

def change_ext_to_pdf(in_file):
    return os.path.splitext(in_file)[0]+".pdf"


#TMP_FILE = open("tmp", "rw")
def convert_web_to_pdf(in_file):
	#TMP_FILE.write(in_file + " " + change_ext_to_pdf(in_file) + "\n")
    pdfkit.from_file(in_file, change_ext_to_pdf(in_file), options=options)


def get_clean_ext(filename):
    return os.path.splitext(filename)[1].lower()

def main(convert):

    if verbose:
        count = 0
        print("initializing")
        for dirname, dirnames, filenames in os.walk(PATH_CONV):
            for filename in filenames:
                if get_clean_ext(filename) in convert and not os.path.isfile(change_ext_to_pdf(dirname + os.sep + filename)):
                    count += 1
        print(str(count) + " files to process")


    current = 0
    print(PATH_CONV)
    for dirname, dirnames, filenames in os.walk(PATH_CONV):
        if verbose:
            print(str(current) + "/" + str(count) + " (" + str(100.0 * current / count) + "%)")
        for filename in filenames:
            if filename == "index.html":
                options["--cache-dir"] = os.path.abspath(dirname)
            key = get_clean_ext(filename)
            abs = os.path.abspath(dirname + os.sep + filename)
            try:
                if key in convert:
                    if not os.path.isfile(change_ext_to_pdf(abs)):
                        if verbose:
                            print(abs)
                            current += 1
                        convert.get(key)(abs)


            except Exception as e:
                error_list.append(abs)
                if verbose:
                    print("[ERROR] " + abs + " :" + e.message)

if __name__ == '__main__':


    first = {
        ".doc": convert_doc_to_pdf,
        ".wri": convert_doc_to_pdf,
        ".wps": convert_doc_to_pdf,
        ".odt": convert_doc_to_pdf,
        ".pps": convert_doc_to_pdf,
        ".wpd": convert_doc_to_pdf,
        ".ppt": convert_doc_to_pdf,
        ".rtf": convert_doc_to_pdf,
        ".xls": convert_doc_to_pdf,
        ".eps": convert_doc_to_pdf,
        ".psd": convert_doc_to_pdf,
        ".pcx": convert_doc_to_pdf,
        ".xlsx": convert_doc_to_pdf,
        ".docx": convert_doc_to_pdf,

        ".ps": convert_ps_to_pdf,

        ".chm": convert_chm,
    }
    second = {
        #".mht": convert_doc_to_pdf,
        ".htm": convert_web_to_pdf,
        ".html": convert_web_to_pdf,
    }

    #we are doing two passes because "chm" files are adding new directories containing html files, and i'm lazy
    main(first)
    main(second)

    print ("all work is done")

    if len(error_list) > 0:
        print("errors: ")
        for error in error_list:
            print(error)
