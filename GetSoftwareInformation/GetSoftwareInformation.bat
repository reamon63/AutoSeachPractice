echo off
.\Python37-32\Python.exe .\get_software_information.py
copy result.html history\result_%date:~0,4%%date:~5,2%%date:~8,2%.html
