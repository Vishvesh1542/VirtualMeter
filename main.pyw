import configparser
import time
import os
import ctypes
import subprocess
import time, traceback

from typing import List
import pythoncom
import pywintypes
import win32gui
from win32com.shell import shell, shellcon

user32 = ctypes.windll.user32

def _make_filter(class_name: str, title: str):
    """https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-enumwindows"""

    def enum_windows(handle: int, h_list: list):
        if not (class_name or title):
            h_list.append(handle)
        if class_name and class_name not in win32gui.GetClassName(handle):
            return True  # continue enumeration
        if title and title not in win32gui.GetWindowText(handle):
            return True  # continue enumeration
        h_list.append(handle)

    return enum_windows

def find_window_handles(parent: int = None, window_class: str = None, title: str = None) -> List[int]:
    cb = _make_filter(window_class, title)
    try:
        handle_list = []
        if parent:
            win32gui.EnumChildWindows(parent, cb, handle_list)
        else:
            win32gui.EnumWindows(cb, handle_list)
        return handle_list
    except pywintypes.error:
        return []

def force_refresh():
    user32.UpdatePerUserSystemParameters(1)

def enable_activedesktop():
    """https://stackoverflow.com/a/16351170"""
    try:
        progman = find_window_handles(window_class='Progman')[0]
        cryptic_params = (0x52c, 0, 0, 0, 500, None)
        user32.SendMessageTimeoutW(progman, *cryptic_params)
    except IndexError as e:
        raise WindowsError('Cannot enable Active Desktop') from e

def set_wallpaper(image_path: str, use_activedesktop: bool = True):
    if use_activedesktop:
        try:
            enable_activedesktop()
        finally:
            pass
    pythoncom.CoInitialize()
    iad = pythoncom.CoCreateInstance(shell.CLSID_ActiveDesktop,
                                     None,
                                     pythoncom.CLSCTX_INPROC_SERVER,
                                     shell.IID_IActiveDesktop)
    iad.SetWallpaper(str(image_path), 0)
    iad.ApplyChanges(shellcon.AD_APPLY_ALL)
    force_refresh()

def get_config():
    MyConfig = configparser.ConfigParser()
    MyConfig.read("config.ini")
    return MyConfig

def main(Myconfig):
    vda = ctypes.WinDLL(os.getcwd() + "//VirtualDesktopAccessor.dll")
    prev_window = 0
    while True:
        window = vda.GetCurrentDesktopNumber()

        if window != prev_window:
            try: 
                r_layout = Myconfig.get('layouts' ,f"{window}")
                r_wallpaper = Myconfig.get('wallpapers', f"{window}")
                subprocess.call(["C:\\Program Files\\Rainmeter\\Rainmeter.exe", "!LoadLayout", r_layout])

                set_wallpaper(r_wallpaper, bool(Myconfig.get('preferences', 'wallpaper_anim')))

            except configparser.NoSectionError as e:
                print(e)
            
            prev_window = window

        time.sleep(0.2)


config = get_config()

try:    
    main(config)
except Exception as e:
    with open(os.getcwd() + "//error.log", "w+") as file:
        file.write(f"{time.strftime("%Y-%m-%d %H:%M:%S")}".ljust(25) + f"{traceback.format_exc()}")