import pystray
from PIL import Image
import win32api
import win32con
import os
import sys
import ctypes
from pathlib import Path
import traceback
import win32gui
import win32process
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import threading
import time
from win32com.shell import shell, shellcon

try:
    from ctypes import cast, POINTER
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
except ImportError as e:
    print(f"Ошибка импорта: {e}")
    print("Убедитесь, что установлены все зависимости:")
    print("pip install -r requirements.txt")
    sys.exit(1)

def is_dark_theme():
    try:
        # Проверяем системную тему
        key = win32api.RegOpenKey(win32con.HKEY_CURRENT_USER, 
                                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize", 0, 
                                win32con.KEY_READ)
        value, _ = win32api.RegQueryValueEx(key, "AppsUseLightTheme")
        win32api.RegCloseKey(key)
        return value == 0
    except:
        return False

def get_icon_path(icon_name):
    try:
        theme = "dark theme" if is_dark_theme() else "light theme"
        scale = "2x" if ctypes.windll.shcore.GetScaleFactorForDevice(0) > 100 else "1x"
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_path, "Icons", theme, scale, icon_name)
    except Exception as e:
        print(f"Ошибка при получении пути к иконке: {e}")
        return os.path.join("Icons", "light theme", "1x", icon_name)

def get_microphone():
    try:
        devices = AudioUtilities.GetMicrophone()
        if not devices:
            print("Микрофон не найден")
            return None
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        return cast(interface, POINTER(IAudioEndpointVolume))
    except Exception as e:
        print(f"Ошибка при получении микрофона: {e}")
        return None

def get_microphone_state():
    try:
        microphone = get_microphone()
        if microphone:
            return not microphone.GetMute()
    except Exception as e:
        print(f"Ошибка при получении состояния микрофона: {e}")
    return False

def toggle_microphone():
    try:
        microphone = get_microphone()
        if microphone:
            current_state = microphone.GetMute()
            microphone.SetMute(not current_state, None)
            return not current_state
    except Exception as e:
        print(f"Ошибка при переключении микрофона: {e}")
        traceback.print_exc()
    return False

def on_click(icon, item=None):
    try:
        is_enabled = toggle_microphone()
        update_icon(icon, is_enabled)
    except Exception as e:
        print(f"Ошибка при обработке клика: {e}")
        traceback.print_exc()

def update_icon(icon, is_enabled=None):
    try:
        if is_enabled is None:
            is_enabled = get_microphone_state()
        icon.icon = Image.open(get_icon_path("ic_mic.png" if is_enabled else "ic_mic_muted.png"))
    except Exception as e:
        print(f"Ошибка при обновлении иконки: {e}")
        traceback.print_exc()

def create_menu(icon):
    return pystray.Menu(
        pystray.MenuItem("Включить/Выключить микрофон", lambda: on_click(icon)),
        pystray.MenuItem("Выход", lambda: icon.stop())
    )

def main():
    # Создаем иконки
    mic_on_icon = Image.open(get_icon_path("ic_mic.png"))
    mic_off_icon = Image.open(get_icon_path("ic_mic_muted.png"))
    
    # Создаем меню
    menu = pystray.Menu(
        pystray.MenuItem('Включить/Выключить микрофон', toggle_microphone),
        pystray.MenuItem('Выход', lambda: icon.stop())
    )
    
    # Создаем иконку в трее
    icon = pystray.Icon(
        "mic_control",
        mic_on_icon if get_microphone_state() else mic_off_icon,
        "Mic Control",
        menu
    )
    
    # Запускаем иконку в трее
    icon.run()

if __name__ == "__main__":
    try:
        # Устанавливаем высокое DPI осознание
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        main()
    except Exception as e:
        print(f"Ошибка при запуске программы: {e}")
        traceback.print_exc()
        sys.exit(1) 