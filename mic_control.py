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
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, IAudioMeterInformation
import threading
import time
from win32com.shell import shell, shellcon
import keyboard
import re

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
hotkey_thread = None
stop_hotkey_check = False
volume_check_thread = None
stop_volume_check = False

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
            return True  # Возвращаем True, чтобы показать, что событие обработано
    except Exception as e:
        print(f"Ошибка при обработке клика: {e}")
    return False

def create_menu(icon):
    return pystray.Menu(
        pystray.MenuItem("Включить/Выключить микрофон", lambda: on_click(icon)),
        pystray.MenuItem("Выход", lambda: icon.stop())
    )

def parse_hotkey(hotkey_str):
    """Парсим строку хоткея в формат для keyboard"""
    try:
        # Читаем хоткей из файла
        with open("hotkey.txt", "r") as f:
            hotkey_str = f.read().strip()
        
        # Преобразуем строку в формат keyboard
        hotkey = hotkey_str.replace("+", "+").lower()
        return hotkey
    except Exception as e:
        print(f"Ошибка при чтении хоткея: {e}")
        return "ctrl+alt+m"  # Возвращаем хоткей по умолчанию

def hotkey_check_loop():
    """Проверяем нажатие хоткея"""
    global stop_hotkey_check
    hotkey = parse_hotkey("")
    
    while not stop_hotkey_check:
        try:
            if keyboard.is_pressed(hotkey):
                print(f"Нажат хоткей: {hotkey}")
                toggle_microphone()
                time.sleep(0.5)  # Задержка для предотвращения множественных срабатываний
        except Exception as e:
            print(f"Ошибка при проверке хоткея: {e}")
        time.sleep(0.1)

def get_microphone_peak():
    """Получаем текущий уровень входного сигнала микрофона"""
    try:
        devices = AudioUtilities.GetMicrophone()
        if not devices:
            return 0
        interface = devices.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None)
        meter = cast(interface, POINTER(IAudioMeterInformation))
        return int(meter.GetPeakValue() * 100)
    except:
        return 0

def volume_check_loop():
    """Проверяем уровень входного сигнала и обновляем тултип"""
    global stop_volume_check, icon
    last_peak = -1
    
    while not stop_volume_check:
        try:
            current_peak = get_microphone_peak()
            if current_peak != last_peak and icon:
                is_muted = get_microphone().GetMute() if get_microphone() else False
                status = "Выключен" if is_muted else f"Включен (Уровень: {current_peak}%)"
                icon.title = f"Mic Control\n{status}"
                last_peak = current_peak
        except:
            pass
        time.sleep(0.1)

def main():
    global icon, theme_check_thread, stop_theme_check, hotkey_thread, stop_hotkey_check, volume_check_thread, stop_volume_check
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
        
        # Запускаем поток проверки хоткея
        hotkey_thread = threading.Thread(target=hotkey_check_loop)
        hotkey_thread.daemon = True
        hotkey_thread.start()
        
        # Запускаем поток проверки уровня сигнала
        volume_check_thread = threading.Thread(target=volume_check_loop)
        volume_check_thread.daemon = True
        volume_check_thread.start()
        
        print("Программа запущена")
        # Запускаем иконку в трее
        icon.run()
    except Exception as e:
        print(f"Критическая ошибка: {e}")
        stop_theme_check = True
        stop_hotkey_check = True
        stop_volume_check = True
        if theme_check_thread:
            theme_check_thread.join()
        if hotkey_thread:
            hotkey_thread.join()
        if volume_check_thread:
            volume_check_thread.join()
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