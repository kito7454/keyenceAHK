import numpy as np
import ahk
from ahk import AHK
import pandas as pd
from numpy.f2py.auxfuncs import throw_error


class ahk_routine:
    def __init__(self,file_name = r"C:\Users\USER1\PycharmProjects\keyenceAHK\VK_automation.xlsx"):
        self.file_name = file_name
        self.ahk = ahk.ahk()
        self.df = pd.read_excel(self.file_name)

    #     check file format for correct headers
        valid_headers = all(
            self.df.columns ==
            ['name', 'ClassNN', 'Type', 'Action', 'Argument', 'Window', 'Mandatory', 'Notes']
        )


        if not valid_headers:
            throw_error("invalid file headers")

    def run(self):
        for i in range(self.df):
            row = self.df.iloc[i]
            self.run_row(row)

    def run_row(self,row):
        # find window
        curr_win = ahk.find_window(title=row.Window)
        if row.Mandatory & (curr_win == None):
            throw_error("windowNotFound")
        # get control


    def get_control_from_window(self,window,classNN):
        # finds a control in a window  with control_class == classNN
        # find classNN by hoveringover button in ahk window spy
        all_controls = window.list_controls()
        result = next((obj for obj in all_controls if obj.control_class == classNN), None)

        if result:
            return result
        else:
            return None

def check_for_variable(term,variableDictionary):
    if 'var ' in term:
        key = term.split()[1]
        print(key)
        return variableDictionary[key]
    else:
        return term
