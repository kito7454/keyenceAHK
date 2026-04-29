from gc import get_count
from idlelib.zoomheight import get_window_geometry
import time
import numpy as np
import ahk
from ahk import AHK
import pandas as pd
from numpy.f2py.auxfuncs import throw_error


class ahk_routine:
    def __init__(self,file_name = r"C:\Users\USER1\PycharmProjects\keyenceAHK\VK_automation.xlsx",variableDictionary = {}):
        self.file_name = file_name
        self.ahk = ahk.AHK()
        self.df = pd.read_excel(self.file_name)
        self.variableDictionary = variableDictionary

    #     check file format for correct headers
        valid_headers = all(
            self.df.columns ==
            ['name', 'ClassNN', 'Type', 'Action', 'Argument', 'Window', 'Mandatory', 'Notes']
        )


        if not valid_headers:
            throw_error("invalid file headers")

    def run_all(self):
        for i in range(len(self.df)):
            row = self.df.iloc[i]
            self.run_row(row)

    def run_selected(self,indices):
        for i in indices:
            row = self.df.iloc[i]
            self.run_row(row)


    # todo: make run individual row by reading row and executing actions
    def run_row(self,row):
        time.sleep(1)
        # takes a single row from excel and runs it in ahk
        print(row.key)
        # find window
        win_name = check_for_variable(row.Window,variableDictionary=self.variableDictionary)

        if "text " in row.Window:
                curr_win = self.get_window_by_substring(row.Notes)

        else:#normal case
            curr_win = ahk.find_window(title=row.Window)

        if bool(row.Mandatory == 1) and bool(curr_win is None):
            throw_error("windowNotFound")
        if curr_win is not None:
            # get control
            curr_ctrl = self.get_control_from_window(curr_win, row.ClassNN)

            if (row.Type == 'Button') & (row.Action == 'Click'):
                curr_ctrl.click(blocking=True)
                time.sleep(0.1)
                curr_ctrl.click(blocking=True)

            if (row.Type == 'Field') & (row.Action == 'Send'):
                if row.Argument is not None:
                    curr_ctrl.send(f"{{Text}}{row.Argument}",blocking=True)



    def get_control_from_window(self,window,classNN):
        # finds a control in a window  with control_class == classNN
        # find classNN by hoveringover button in ahk window spy
        all_controls = window.list_controls()
        result = next((obj for obj in all_controls if obj.control_class == classNN), None)

        if result:
            return result
        else:
            return None

    def get_window_by_substring(self,substring):
        # finds a window with text that contains "substring"
        all_windows = self.ahk.list_windows()
        result = None
        for w in all_windows:
            if substring in w.text:
                result = w

        return result

def check_for_variable(term,variableDictionary):
    # if "var varName" is inputted in excel, returns varName
    # otherwise just returns the input
    if 'var ' in term:
        key = term.split()[1]
        print(key)
        return variableDictionary[key]
    else:
        return term

if __name__ == "__main__":
    ar = ahk_routine(file_name=r"C:\Users\USER1\PycharmProjects\keyenceAHK\VK_automation_noLoad.xlsx")
    ar.run_selected([0,1,2,3,4,5])