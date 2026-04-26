import ahkHelper
import time

ah = ahkHelper.AHKHelper()
ls = ah.ahk.list_windows()




for wi in ls:
    print(wi.title)

viewerWindow = ah.ahk.find_window(title="Viewer Application")
# viewerWindow.activate()
time.sleep(0.5)
viewerControls= viewerWindow.list_controls()
# for c in viewerControls:
#     print(c)

enter_window = ah.ahk.find_window(title = "Load Teaching Settings File")
enter_controls = enter_window.list_controls()
enter_field = ah.get_control_from_window(window=enter_window,classNN="Edit1")
print(enter_field)
# viewerWindow.minimize()
