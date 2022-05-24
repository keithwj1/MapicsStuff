from cx_Freeze import setup, Executable
import xlrd
from jira import JIRA
import getpass
import datetime
import sys
import os.path
import json
from os import path

base = None    

#executables = [Executable("C:\Users\kjones\source\repos\ImportExcelToJira\ImportExcelToJira\ImportExcelToJira.py", base=base)]
executables = [Executable("C:/Users/kjones/source/repos/ImportExcelToJira/ImportExcelToJira/ImportExcelToJira.py", base=base)]
packages = ["idna"]
options = {
    'build_exe': {    
        'packages':packages,
    },    
}

setup(
    name = "Import CA From Excel To Jira",
    options = options,
    version = "1.0",
    description = 'This Program will allow the automatic import of data from excel into Jira',
    executables = executables
)