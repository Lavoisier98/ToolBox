from cx_Freeze import setup, Executable # type: ignore
import pandas as pd

#IMPORTANT:
#Step 1: Changing the name of the script at 'executables' variable
#Step 2: Changing the names of the libraries at 'packages' variable, remembering 'idna' has to be always in the item 0 of the list object
#Step 3: Changing the 'name' parameter in the subprocedure 'setup'
#Step 4: Optionally change the description to appear at exe file properties
#Step 5: At folder with this file, right click holding the shift button and execute the Windows Power Shell
#Step 6: Type 'python ', the name of this .py file (currently PyToExe.py) and ' build'

base = None

executables = [Executable("Materials.py", base=base)]
packages = ["idna","pyautogui","pyperclip","time","openpyxl"]

options = {
    'build_exe': {    
        'packages':packages,
    },    
}
setup(
    name = "Materials",
    options = options,
    version = "1.0",
    description = 'File to consult massively sales\' scenarios at CX Data Simulator',
    executables = executables
)