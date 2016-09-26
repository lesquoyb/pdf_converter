# pdf_converter
Converts recursively all the office and html/chm files in a directory into pdf files

#how to use it 
Replace the placeholders for the variables PYTHON_INTERP and PATH_CONV with respectively your python interpreter and the path of the directory you want to process, then run the file main.py.
You can turn on error messages while processing with the variable 'verbose'.
If you have chm files, it will raise an error on linux, so you will have to adapt the line:

  subprocess.call(["hh.exe", "-decompile", directory, in_file], stdout=subprocess.PIPE) 

with your chm decompiler (7zip works fine, surely a buch of others too)
