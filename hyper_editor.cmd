@echo off
set /p p="Start page? "
set /p s="Inctement? "
python.exe hyper_editor.py -f %1 -p %p% -s %s%

