# What it does
Recursively converts all kind of documents in a directory into pdf files.
Supported formats are:
 - .doc
 - .wri
 - .wps
 - .odt
 - .pps
 - .wpd
 - .ppt
 - .rtf
 - .xls
 - .eps
 - .psd
 - .pcx
 - .xlsx
 - .docx
 - .ps
 - .chm
 - .htm
 - .html

# How to use it 
Replace the placeholders for the variables `PYTHON_INTERP` and `PATH_CONV` with respectively your python interpreter and the path of the directory you want to process, then run the file main.py.
You can turn on error messages while processing with the variable `verbose`.

If you have chm files in the directory, it will raise an error on linux/mac, so you will have to adapt the line:
```python
  subprocess.call(["hh.exe", "-decompile", directory, in_file], stdout=subprocess.PIPE) 
```
with your chm decompiler (7zip works fine, surely a bunch of others too)
