import numpy as np
import ahk
from ahk import AHK

# --------------------
# handles all the backend for autohotkeying things
#run as administrator or it wont work
#---------------------

# takes the path of an image and finds the coordinates of it on the screen
class AHKHelper():

    def __init__(self):
        super().__init__()
        self.ahk = AHK()

    def findButtonPic(self,imagePath):
    #         "C:\\Users\\TeamD\\Desktop\\kyle\\ahk\\browseButton.png"
        self.ahk = AHK()
        return self.ahk.image_search(imagePath)

    def getWindow(self,windowName):
        self.win = ahk.find_window(title=windowName)
        self.win.activate()
        return self.win

    def findButtonText(self,buttonName):
        buttonList = self.win.list_controls()
        button = [x for x in buttonList if x.get_text()==buttonName]

        return button[0]

    def clickButton(self,buttonName,windowName):
        self.ahk.control_click(
            button='L',  # Left mouse button
            click_count=1,
            control=buttonName,  # Control name (e.g., Button1)
            title=windowName,  # Window title
        )

    def get_control_from_window(self,window,classNN):
        # finds a control in a window  with control_class == classNN
        # find classNN by hoveringover button in ahk window spy
        all_controls = window.list_controls()
        result = next((obj for obj in all_controls if obj.control_class == classNN), None)

        if result:
            return result
        else:
            return None


# if __name__ == '__main__':
#
#     ah = AHKHelper()
#     ah.getWindow('Smart Processing Commander')
#
#     button = ah.findButtonText('Browse')
#     control_pos= button.get_position()
#
#     ah.clickButton('Browse','Smart Processing Commander')
#     ah.ahk.type(r'C:\Users\TeamD\Desktop\kyle\exampleImages\2.jpeg')
#
#     ah.ahk.send("{enter}")