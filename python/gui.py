from warnings import catch_warnings
import pyautogui
import time as tim
from datetime import datetime, time
from ctypes import *

pyautogui.FAILSAFE = False


class LASTINPUTINFO(Structure):
    _fields_ = [
        ('cbSize', c_uint),
        ('dwTime', c_int)
    ]


def get_duration():
    lastInputlnfo = LASTINPUTINFO()
    lastInputlnfo.cbSize = sizeof(lastInputlnfo)

    if windll.user32.GetLastInputInfo(byref(lastInputlnfo)):
        millis = windll.kernel32.GetTickCount() - lastInputlnfo.dwTime
        return millis / 1000.0
    else:
        return 0


while True:
    now = datetime.now().time()
    fourPm = time(16, 00)

    if now >= fourPm:
        print("0")
        break

    dur = get_duration()
    try:
        if (dur > 4 * 60):  # 4 mins
            print(f"duration = {dur}")
            pyautogui.press("shift")
            pyautogui.press("f15")
    except Exception as e:
        print(f"Error encountered {e}")

    tim.sleep(5)  # 5 seconds

    # print(f"{now} loop)
