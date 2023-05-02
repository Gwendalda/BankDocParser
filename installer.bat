@echo off

set PYTHON_VERSION=3.10.0  :: Replace with the desired Python version
set PYTHON_INSTALL_PATH=C:\Python310  :: Replace with the desired installation path

:: Download the Python installer
echo Downloading Python %PYTHON_VERSION%...
curl -o python_installer.exe https://www.python.org/ftp/python/%PYTHON_VERSION%/python-%PYTHON_VERSION%-amd64.exe

:: Install Python
echo Installing Python %PYTHON_VERSION%...
python_installer.exe /quiet InstallAllUsers=1 PrependPath=1 TargetDir=%PYTHON_INSTALL_PATH%

:: Update pip
echo Updating pip...
%PYTHON_INSTALL_PATH%\python.exe -m pip install --upgrade pip

:: Clean up
echo Cleaning up...
del python_installer.exe

echo Python installation completed.
