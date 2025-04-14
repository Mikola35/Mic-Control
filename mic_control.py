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
microphone = None
meter = None

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
    theme = "dark theme" if is_dark_theme() else "light theme"
    path = os.path.join("Icons", theme, filename.replace('.png', '.ico'))
    print(f"Путь к иконке: {path}")
    return path

def get_volume_icon_path(is_muted, peak):
    theme = "dark theme" if is_dark_theme() else "light theme"
    if is_muted:
        icon_name = "ic_mic_muted.ico"
    else:
        if peak <= 0:
            icon_name = "ic_mic.ico"
        elif peak > 15:
            icon_name = "ic_mic_vol-10.ico"
        else:
            # 1-15% делим на 9 ступеней
            level = min(9, max(1, int((peak - 1) / (15 / 9)) + 1))
            icon_name = f"ic_mic_vol-{level:02}.ico"
    path = os.path.join("Icons", theme, icon_name)
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
                icon_path = get_icon_path("ic_mic.ico" if not is_muted else "ic_mic_muted.ico")
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
    global microphone, meter
    try:
        if microphone is None:
            devices = AudioUtilities.GetMicrophone()
            if not devices:
                print("Микрофон не найден")
                return None
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            microphone = cast(interface, POINTER(IAudioEndpointVolume))
            
            # Получаем доступ к измерителю уровня
            interface = devices.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None)
            meter = cast(interface, POINTER(IAudioMeterInformation))
        return microphone
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

def create_menu():
    """Создаем контекстное меню"""
    # Получаем текущий хоткей
    hotkey = parse_hotkey("")
    hotkey_text = hotkey.upper().replace("+", " + ")
    
    def on_toggle():
        toggle_microphone()
    
    def on_settings():
        os.system("control mmsys.cpl")
    
    def on_exit():
        icon.stop()
    
    return pystray.Menu(
        pystray.MenuItem(f'Включить/Выключить микрофон ({hotkey_text})', on_toggle),
        pystray.MenuItem('Настройки микрофона', on_settings),
        pystray.MenuItem('Выход', on_exit)
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
    global meter
    try:
        if meter:
            return int(meter.GetPeakValue() * 100)
    except:
        pass
    return 0

def volume_check_loop():
    """Проверяем уровень входного сигнала и обновляем тултип"""
    global stop_volume_check, icon
    last_peak = -1
    last_muted = None
    current_level = None  # для плавной смены иконки
    
    while not stop_volume_check:
        try:
            current_peak = get_microphone_peak()
            microphone = get_microphone()
            is_muted = microphone.GetMute() if microphone else True

            # Определяем целевой уровень для иконки
            if is_muted:
                target_level = 'muted'
            elif current_peak <= 0:
                target_level = 'zero'
            elif current_peak > 7:
                target_level = 10
            else:
                # 1-7% делим на 9 ступеней
                level = min(9, max(1, int((current_peak - 1) / (7 / 9)) + 1))
                target_level = level

            # Инициализация
            if current_level is None:
                current_level = target_level

            # Плавное движение к целевому уровню
            if isinstance(target_level, int) and isinstance(current_level, int):
                if current_level < target_level:
                    current_level += 1
                elif current_level > target_level:
                    current_level -= 1
            else:
                current_level = target_level

            # Формируем имя иконки
            if is_muted:
                icon_name = "ic_mic_muted.ico"
            elif current_level == 'zero':
                icon_name = "ic_mic.ico"
            elif isinstance(current_level, int):
                if current_level == 10:
                    icon_name = "ic_mic_vol-10.ico"
                else:
                    icon_name = f"ic_mic_vol-{current_level:02}.ico"
            else:
                icon_name = "ic_mic.ico"

            theme = "dark theme" if is_dark_theme() else "light theme"
            icon_path = os.path.join("Icons", theme, icon_name)

            # Обновляем тултип и иконку только если что-то поменялось
            if (current_peak != last_peak or is_muted != last_muted) and icon:
                status = "Выключен" if is_muted else f"Включен (Уровень: {current_peak}%)"
                icon.title = f"Mic Control\n{status}"
            if icon:
                icon.icon = Image.open(icon_path)
            last_peak = current_peak
            last_muted = is_muted
        except:
            pass
        time.sleep(0.3)

def cleanup():
    """Очищаем ресурсы"""
    global microphone, meter, stop_theme_check, stop_hotkey_check, stop_volume_check
    
    stop_theme_check = True
    stop_hotkey_check = True
    stop_volume_check = True
    
    if theme_check_thread:
        theme_check_thread.join()
    if hotkey_thread:
        hotkey_thread.join()
    if volume_check_thread:
        volume_check_thread.join()
    
    # Освобождаем COM-объекты
    microphone = None
    meter = None

def main():
    global icon, theme_check_thread, stop_theme_check, hotkey_thread, stop_hotkey_check, volume_check_thread, stop_volume_check
    try:
        print("Запуск программы...")
        
        # Проверяем микрофон
        if not get_microphone():
            print("Ошибка: микрофон не найден")
            return
            
        # Загружаем иконки
        try:
            mic_on_path = get_icon_path("ic_mic.ico")
            mic_off_path = get_icon_path("ic_mic_muted.ico")
            print(f"Путь к иконке включенного микрофона: {mic_on_path}")
            print(f"Путь к иконке выключенного микрофона: {mic_off_path}")
            
            mic_on_icon = Image.open(mic_on_path)
            mic_off_icon = Image.open(mic_off_path)
        except Exception as e:
            print(f"Ошибка при загрузке иконок: {e}")
            return
        
        # Создаем меню
        menu = create_menu()
        
        # Создаем иконку в трее
        icon = pystray.Icon(
            "mic_control",
            mic_on_icon if microphone.GetMute() == 0 else mic_off_icon,
            "Mic Control",
            menu
        )
        
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
        cleanup()
        sys.exit(1)
    finally:
        cleanup()

if __name__ == "__main__":
    try:
        # Устанавливаем высокое DPI осознание
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        main()
    except Exception as e:
        print(f"Ошибка при запуске программы: {e}")
        traceback.print_exc()
        cleanup()
        sys.exit(1) 