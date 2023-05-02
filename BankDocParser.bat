@echo off

set PYTHON_SCRIPT=gui.py  :: Replace with the name of your Python script
set PYTHON_EXECUTABLE=C:\Python310\python.exe  :: Replace with the path to your Python executable

echo Running %PYTHON_SCRIPT%...

%PYTHON_EXECUTABLE% %PYTHON_SCRIPT%

echo Script execution completed.
