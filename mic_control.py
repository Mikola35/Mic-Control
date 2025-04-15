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
import keyboard
import re
import pythoncom
import winreg
import datetime
import win32com.client

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

LOG_PATH = "mic_control.log"

def log(msg):
    now = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{now}] {msg}"
    print(line)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception as e:
        print(f"[MicControl] Ошибка записи в лог: {e}")

# Очищаем лог при запуске
with open(LOG_PATH, "w", encoding="utf-8") as f:
    f.write("")

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

def get_default_mic_id_and_name():
    try:
        MMDeviceEnumerator = win32com.client.Dispatch("MMDeviceEnumerator")
        eRender = 0
        eCapture = 1
        eAll = 2
        eMultimedia = 1
        device = MMDeviceEnumerator.GetDefaultAudioEndpoint(eCapture, eMultimedia)
        mic_id = device.GetId()
        mic_name = device.Properties["{a45c254e-df1c-4efd-8020-67d146a850e0},2"] # PKEY_Device_FriendlyName
        return mic_id, mic_name
    except Exception as e:
        print(f"[MicControl] Не удалось получить id и имя микрофона через win32com: {e}")
        return None, None

def get_default_mic_id_comtypes():
    try:
        from comtypes import CLSCTX_ALL
        from comtypes.client import CreateObject
        import comtypes
        IID_IMMDeviceEnumerator = comtypes.GUID('{A95664D2-9614-4F35-A746-DE8DB63617E6}')
        eCapture = 1
        eMultimedia = 1
        MMDeviceEnumerator = CreateObject('MMDeviceEnumerator.MMDeviceEnumerator', interface=None)
        enum = MMDeviceEnumerator.QueryInterface(IID_IMMDeviceEnumerator)
        device = enum.GetDefaultAudioEndpoint(eCapture, eMultimedia)
        return device.GetId()
    except Exception as e:
        print(f"[MicControl] Не удалось получить id микрофона через comtypes: {e}")
        return None

def get_default_mic_id_pycaw():
    try:
        from pycaw.pycaw import AudioUtilities
        dev = AudioUtilities.GetDefaultDevice('Capture')
        if dev:
            return getattr(dev, 'id', None), getattr(dev, 'FriendlyName', None)
        return None, None
    except Exception as e:
        print(f"[MicControl] Не удалось получить id и имя микрофона через pycaw: {e}")
        return None, None

def get_default_mic_id_coreaudio():
    try:
        from pycaw.utils import AudioUtilities
        # Получаем IMMDeviceEnumerator через pycaw
        enumerator = AudioUtilities.GetDeviceEnumerator()
        eCapture = 1
        eMultimedia = 1
        device = enumerator.GetDefaultAudioEndpoint(eCapture, eMultimedia)
        return device.GetId()
    except Exception as e:
        log(f"[MicControl] Не удалось получить id микрофона через CoreAudio/pycaw: {e}")
        return None

def get_friendly_name_by_id_pycaw(device_id):
    try:
        from pycaw.pycaw import AudioUtilities
        devices = AudioUtilities.GetAllDevices()
        for d in devices:
            if getattr(d, 'id', None) == device_id:
                return getattr(d, 'FriendlyName', None)
        return None
    except Exception as e:
        log(f"[MicControl] Не удалось найти FriendlyName по id через pycaw: {e}")
        return None

def default_mic_monitor_loop():
    last_id = None
    while True:
        mic_id = get_default_mic_id_coreaudio()
        if mic_id != last_id:
            mic_name = get_friendly_name_by_id_pycaw(mic_id)
            log(f"[CoreAudio] Текущий дефолтный микрофон: id={mic_id}, name={mic_name}")
            last_id = mic_id
        time.sleep(1)

def toggle_microphone():
    """Переключаем состояние микрофона"""
    try:
        microphone = get_microphone()
        if microphone:
            mic_id, mic_name = get_default_mic_id_pycaw()
            print(f"[MicControl] (pycaw) Мьютим устройство: id={mic_id}, name={mic_name}")
            current_state = microphone.GetMute()
            log(f"[MicControl] До: Микрофон {'выключен' if current_state else 'включён'} (Mute={current_state})")
            new_state = 1 - current_state
            log(f"[MicControl] Действие: {'Включаем микрофон...' if new_state == 0 else 'Мьютим микрофон...'}")
            microphone.SetMute(new_state, None)
            after_state = microphone.GetMute()
            log(f"[MicControl] После: Микрофон {'выключен' if after_state else 'включён'} (Mute={after_state})")
            update_icon()
            return new_state == 0
    except Exception as e:
        log(f"Ошибка при переключении микрофона: {e}")
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

def get_microphone_peak_db():
    global meter
    peak = 0
    try:
        if meter:
            peak = meter.GetPeakValue()
    except:
        pass
    return peak

def get_default_mic_name():
    mic_id = get_default_mic_id_coreaudio()
    return get_friendly_name_by_id_pycaw(mic_id)

def volume_check_loop():
    global stop_volume_check, icon
    last_peak = -1
    last_muted = None
    last_state = None
    last_mic_name = None
    current_level = None
    zero_start = None  # время, когда начался 0
    idle_mode = False
    while not stop_volume_check:
        try:
            peak = get_microphone_peak_db()
            peak_percent = int(peak * 100)
            microphone = get_microphone()
            is_muted = microphone.GetMute() if microphone else True
            now = time.time()

            # Показываем если мьют изменился не по хоткею
            if last_muted is not None and is_muted != last_muted:
                log(f"[MicControl] Состояние микрофона изменилось системой: Было={int(last_muted)}, Стало={int(is_muted)}")

            # Таймер для простоя
            if not is_muted and peak_percent == 0:
                if zero_start is None:
                    zero_start = now
                elif now - zero_start > 5:
                    idle_mode = True
            else:
                zero_start = None
                idle_mode = False

            # Определяем состояние
            if is_muted:
                state = 'Заглушён'
                icon_name = "ic_mic_muted.ico"
            elif idle_mode:
                state = 'Не используется'
                icon_name = "ic_mic_idle.ico"
            elif peak_percent == 0:
                state = 'Активен'
                icon_name = "ic_mic.ico"
            else:
                state = 'Активен'
                if peak_percent > 7:
                    icon_name = "ic_mic_vol-10.ico"
                else:
                    level = min(9, max(1, int((peak_percent - 1) / (7 / 9)) + 1))
                    icon_name = f"ic_mic_vol-{level:02}.ico"

            theme = "dark theme" if is_dark_theme() else "light theme"
            icon_path = os.path.join("Icons", theme, icon_name)

            # Формируем тултип
            mic_name = get_default_mic_name()
            tooltip = f"🎤 {mic_name}"

            if (peak_percent != last_peak or is_muted != last_muted or state != last_state or mic_name != last_mic_name) and icon:
                icon.title = tooltip
                icon.icon = Image.open(icon_path)
                last_peak = peak_percent
                last_state = state
                last_mic_name = mic_name
            last_muted = is_muted
        except Exception as e:
            log(f"[MicControl] Ошибка в volume_check_loop: {e}")
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

def print_all_audio_devices():
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    devices = AudioUtilities.GetAllDevices()
    # Получаем id дефолтного микрофона
    default_mic_id = get_default_mic_id_coreaudio()
    inputs = [d for d in devices if str(getattr(d, 'id', '')).startswith('{0.0.1.')]
    outputs = [d for d in devices if str(getattr(d, 'id', '')).startswith('{0.0.0.')]
    # Сортировка по FriendlyName
    inputs = sorted(inputs, key=lambda d: str(getattr(d, 'FriendlyName', '')))
    outputs = sorted(outputs, key=lambda d: str(getattr(d, 'FriendlyName', '')))
    # INPUT
    with open("audio_devices_input.txt", "w", encoding="utf-8") as f:
        f.write("INPUT (микрофоны):\n")
        f.write(f"{'N':<3} {'id':<60} {'FriendlyName':<90} {'Default':<7} {'Mute':<6} {'state':<30} {'data_flow':<20} {'type':<40}\n")
        for i, d in enumerate(inputs):
            mute = ""
            if str(getattr(d, 'state', '')) == 'AudioDeviceState.Active':
                try:
                    endpoint = d._ctl.QueryInterface(IAudioEndpointVolume)
                    mute = "Yes" if endpoint.GetMute() else "No"
                except:
                    mute = ""
            is_default = "✓" if getattr(d, 'id', None) == default_mic_id else ""
            f.write(f"{i+1:<3} {getattr(d, 'id', ''):<60} {getattr(d, 'FriendlyName', ''):<90} {is_default:<7} {mute:<6} {str(getattr(d, 'state', '')):<30} {str(getattr(d, 'data_flow', '')):<20} {str(type(d)):<40}\n")
        f.write("\n---\n")
    # OUTPUT
    with open("audio_devices_output.txt", "w", encoding="utf-8") as f:
        f.write("OUTPUT (динамики):\n")
        f.write(f"{'N':<3} {'id':<60} {'FriendlyName':<90} {'state':<30} {'data_flow':<20} {'type':<40}\n")
        for i, d in enumerate(outputs):
            f.write(f"{i+1:<3} {getattr(d, 'id', ''):<60} {getattr(d, 'FriendlyName', ''):<90} {str(getattr(d, 'state', '')):<30} {str(getattr(d, 'data_flow', '')):<20} {str(type(d)):<40}\n")
        f.write("\n---\n")

def audio_devices_update_loop():
    pythoncom.CoInitialize()
    while True:
        print_all_audio_devices()
        time.sleep(1)

def write_registry_audio_tables():
    import winreg
    root = r"SYSTEM\\CurrentControlSet\\Enum\\SWD\\MMDEVAPI"
    columns = [
        "Key", "FriendlyName", "DeviceDesc", "HardwareID", "Driver",
        "ParentIdPrefix", "Class", "Service", "Manufacturer", "LocationInformation",
        "Capabilities", "ConfigFlags", "ContainerID", "InstallDate"
    ]
    mic_rows = []
    spk_rows = []
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, root)
        for i in range(winreg.QueryInfoKey(key)[0]):
            subkey_name = winreg.EnumKey(key, i)
            subkey = winreg.OpenKey(key, subkey_name)
            data = {"Key": subkey_name}
            for col in columns[1:]:
                try:
                    val = winreg.QueryValueEx(subkey, col)[0]
                except Exception as e:
                    val = ""
                data[col] = str(val)
            if subkey_name.startswith('{0.0.1.'):
                mic_rows.append(data)
            elif subkey_name.startswith('{0.0.0.'):
                spk_rows.append(data)
    except Exception as e:
        mic_rows.append({"Key": "Ошибка", "FriendlyName": str(e), "DeviceDesc": "", "HardwareID": "", "Driver": "", "ParentIdPrefix": "", "Class": "", "Service": "", "Manufacturer": "", "LocationInformation": "", "Capabilities": "", "ConfigFlags": "", "ContainerID": "", "InstallDate": ""})
    # Сортировка по FriendlyName
    mic_rows = sorted(mic_rows, key=lambda row: row['FriendlyName'])
    spk_rows = sorted(spk_rows, key=lambda row: row['FriendlyName'])
    # Определяем ширину колонок для каждого файла
    def get_col_widths(rows):
        return {col: max(len(col), *(len(row[col]) for row in rows)) for col in columns} if rows else {col: len(col) for col in columns}
    # Микрофоны
    mic_col_widths = get_col_widths(mic_rows)
    with open("registry_microphones.txt", "w", encoding="utf-8") as f:
        f.write("Микрофоны из реестра Windows (MMDEVAPI):\n")
        f.write(" ".join(f"{col:<{mic_col_widths[col]}}" for col in columns) + "\n")
        for row in mic_rows:
            f.write(" ".join(f"{row[col]:<{mic_col_widths[col]}}" for col in columns) + "\n")
        f.write("\n---\n")
    # Динамики
    spk_col_widths = get_col_widths(spk_rows)
    with open("registry_speakers.txt", "w", encoding="utf-8") as f:
        f.write("Динамики из реестра Windows (MMDEVAPI):\n")
        f.write(" ".join(f"{col:<{spk_col_widths[col]}}" for col in columns) + "\n")
        for row in spk_rows:
            f.write(" ".join(f"{row[col]:<{spk_col_widths[col]}}" for col in columns) + "\n")
        f.write("\n---\n")

def print_default_mic_on_start():
    mic_id, mic_name = get_default_mic_id_pycaw()
    print(f"[MicControl] (pycaw) Дефолтный микрофон при старте: id={mic_id}, name={mic_name}")

def main():
    global icon, theme_check_thread, stop_theme_check, hotkey_thread, stop_hotkey_check, volume_check_thread, stop_volume_check
    try:
        threading.Thread(target=default_mic_monitor_loop, daemon=True).start()
        print_default_mic_on_start()
        print('Вызов write_registry_audio_tables()')
        write_registry_audio_tables()
        threading.Thread(target=audio_devices_update_loop, daemon=True).start()
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

# Запускаю функцию при старте
write_registry_audio_tables() 