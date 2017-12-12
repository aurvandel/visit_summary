import win32gui
import ctypes
from pywinauto.findwindows    import find_window

EnumWindows = ctypes.windll.user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
GetWindowText = ctypes.windll.user32.GetWindowTextW
GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
IsWindowVisible = ctypes.windll.user32.IsWindowVisible

titles = []
def foreach_window(hwnd, lParam):
    if IsWindowVisible(hwnd):
        length = GetWindowTextLength(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        GetWindowText(hwnd, buff, length + 1)
        titles.append((hwnd, buff.value))
    return True
EnumWindows(EnumWindowsProc(foreach_window), 0)
icentra = []
for i in range(len(titles)):
    if "PowerChart" in titles[i][1]:
        icentra += titles[i]

handle = icentra[0]
title = icentra[1]

win32gui.SetForegroundWindow(find_window(title=title))