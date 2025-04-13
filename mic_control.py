import pystray
from PIL import Image, ImageDraw
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

# Глобальная переменная для хранения иконки
icon = None

def is_dark_theme():
    """Проверяем, используется ли темная тема Windows"""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return value == 0
    except:
        return False

def get_icon_path(filename):
    """Получаем путь к иконке с учетом темы Windows и масштабирования экрана"""
    base_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Icons")
    theme = "dark theme" if is_dark_theme() else "light theme"
    scale = "2x" if ctypes.windll.shcore.GetScaleFactorForDevice(0) > 100 else "1x"
    return os.path.join(base_path, theme, scale, filename)

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
    """Переключаем состояние микрофона"""
    try:
        devices = AudioUtilities.GetMicrophone()
        if not devices:
            return
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        current_state = volume.GetMute()
        volume.SetMute(1 - current_state, None)
        
        # Обновляем иконку
        if icon:
            icon.icon = Image.open(get_icon_path("ic_mic.png" if not current_state else "ic_mic_muted.png"))
    except Exception as e:
        print(f"Ошибка при переключении микрофона: {e}")

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
    global icon
    
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