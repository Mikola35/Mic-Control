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

# Глобальные переменные
icon = None
theme_check_thread = None
stop_theme_check = False

def get_scale_factor():
    """Получаем масштаб экрана"""
    try:
        scale_factor = ctypes.windll.shcore.GetScaleFactorForDevice(0)
        return "2x" if scale_factor > 100 else "1x"
    except:
        return "1x"

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
    """Получаем путь к иконке с учетом темы и масштаба"""
    theme = "dark theme" if is_dark_theme() else "light theme"
    scale = get_scale_factor()
    path = os.path.join("Icons", theme, scale, filename)
    print(f"Путь к иконке: {path}")
    return path

def update_icon():
    """Обновляем иконку с учетом текущей темы"""
    global icon
    if icon:
        try:
            microphone = get_microphone()
            if microphone:
                is_muted = microphone.GetMute()
                icon_path = get_icon_path("ic_mic.png" if not is_muted else "ic_mic_muted.png")
                icon.icon = Image.open(icon_path)
        except Exception as e:
            print(f"Ошибка при обновлении иконки: {e}")

def theme_check_loop():
    """Проверяем изменения темы Windows"""
    global stop_theme_check
    last_theme = is_dark_theme()
    
    while not stop_theme_check:
        current_theme = is_dark_theme()
        if current_theme != last_theme:
            print("Обнаружено изменение темы Windows")
            update_icon()
            last_theme = current_theme
        time.sleep(1)

def get_microphone():
    """Получаем доступ к микрофону"""
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
        microphone = get_microphone()
        if microphone:
            current_state = microphone.GetMute()
            new_state = 1 - current_state
            microphone.SetMute(new_state, None)
            update_icon()
            return new_state == 0
    except Exception as e:
        print(f"Ошибка при переключении микрофона: {e}")
    return False

def on_click(icon, item=None):
    """Обработчик клика по иконке"""
    try:
        if item is None:  # Это клик по иконке
            print("Клик по иконке")
            toggle_microphone()
    except Exception as e:
        print(f"Ошибка при обработке клика: {e}")

def create_menu(icon):
    return pystray.Menu(
        pystray.MenuItem("Включить/Выключить микрофон", lambda: on_click(icon)),
        pystray.MenuItem("Выход", lambda: icon.stop())
    )

def main():
    global icon, theme_check_thread, stop_theme_check
    try:
        print("Запуск программы...")
        
        # Проверяем микрофон
        microphone = get_microphone()
        if not microphone:
            print("Ошибка: микрофон не найден")
            return
            
        # Загружаем иконки
        try:
            mic_on_path = get_icon_path("ic_mic.png")
            mic_off_path = get_icon_path("ic_mic_muted.png")
            print(f"Путь к иконке включенного микрофона: {mic_on_path}")
            print(f"Путь к иконке выключенного микрофона: {mic_off_path}")
            
            mic_on_icon = Image.open(mic_on_path)
            mic_off_icon = Image.open(mic_off_path)
        except Exception as e:
            print(f"Ошибка при загрузке иконок: {e}")
            return
        
        # Создаем меню
        menu = pystray.Menu(
            pystray.MenuItem('Включить/Выключить микрофон', toggle_microphone),
            pystray.MenuItem('Выход', lambda: icon.stop())
        )
        
        # Создаем иконку в трее
        icon = pystray.Icon(
            "mic_control",
            mic_on_icon if microphone.GetMute() == 0 else mic_off_icon,
            "Mic Control",
            menu
        )
        
        # Устанавливаем обработчик клика
        icon.on_click = on_click
        
        # Запускаем поток проверки темы
        theme_check_thread = threading.Thread(target=theme_check_loop)
        theme_check_thread.daemon = True
        theme_check_thread.start()
        
        print("Программа запущена")
        # Запускаем иконку в трее
        icon.run()
    except Exception as e:
        print(f"Критическая ошибка: {e}")
        stop_theme_check = True
        if theme_check_thread:
            theme_check_thread.join()
        sys.exit(1)

if __name__ == "__main__":
    try:
        # Устанавливаем высокое DPI осознание
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        main()
    except Exception as e:
        print(f"Ошибка при запуске программы: {e}")
        traceback.print_exc()
        sys.exit(1) 