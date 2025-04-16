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
import tkinter as tk
from tkinter import simpledialog, messagebox
import json
import customtkinter as ctk

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
last_mic_id = None

LOG_PATH = "mic_control.log"
SETTINGS_PATH = "settings.json"

def log(msg):
    pass  # Отключено по просьбе пользователя, ничего не пишем в лог

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

def get_volume_icon_name(peak_percent):
    if peak_percent == 0:
        return "ic_mic.ico"
    elif peak_percent <= 3:
        return "ic_mic_vol-03.ico"
    elif peak_percent <= 7:
        return "ic_mic_vol-04.ico"
    elif peak_percent <= 15:
        return "ic_mic_vol-05.ico"
    elif peak_percent <= 25:
        return "ic_mic_vol-06.ico"
    elif peak_percent <= 40:
        return "ic_mic_vol-07.ico"
    elif peak_percent <= 60:
        return "ic_mic_vol-08.ico"
    elif peak_percent <= 80:
        return "ic_mic_vol-09.ico"
    else:
        return "ic_mic_vol-10.ico"

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
    global microphone, meter, last_mic_id
    try:
        current_id = get_default_mic_id_coreaudio()
        # Если id сменился или объекты невалидны — пересоздаём
        if (
            microphone is None or
            meter is None or
            getattr(microphone, 'GetId', lambda: None)() != current_id or
            last_mic_id != current_id
        ):
            devices = AudioUtilities.GetMicrophone()
            if not devices:
                print("Микрофон не найден")
                microphone = None
                meter = None
                last_mic_id = None
                return None
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            microphone = cast(interface, POINTER(IAudioEndpointVolume))
            # Получаем доступ к измерителю уровня
            interface = devices.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None)
            meter = cast(interface, POINTER(IAudioMeterInformation))
            last_mic_id = current_id
        return microphone
    except Exception as e:
        print(f"Ошибка при получении микрофона: {e}")
        microphone = None
        meter = None
        last_mic_id = None
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

def load_settings():
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {"show_volume": True, "hotkey": "CTRL+ALT+M"}

def save_settings(settings):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f)

def open_settings_window():
    try:
        import sys
        import os
        import customtkinter as ctk
        import datetime
        global stop_hotkey_check, hotkey_thread

        HOTKEY_LOG = "hotkey_capture.log"

        # --- Пауза хоткей-ловушки ---
        stop_hotkey_check = True
        if hotkey_thread:
            hotkey_thread.join()
        # ---------------------------

        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        def center(win):
            win.update_idletasks()
            w = win.winfo_width()
            h = win.winfo_height()
            ws = win.winfo_screenwidth()
            hs = win.winfo_screenheight()
            x = (ws // 2) - (w // 2)
            y = (hs // 2) - (h // 2)
            win.geometry(f'+{x}+{y}')

        last_logged_keys = {"set": None}

        def log_hotkey_event(event_type, key_set):
            key_tuple = tuple(sorted(key_set))
            if key_tuple != last_logged_keys["set"]:
                now = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
                with open(HOTKEY_LOG, "a", encoding="utf-8") as f:
                    f.write(f"[{now}] {event_type}: {'+'.join(key_tuple) if key_tuple else '[NONE]'}\n")
                last_logged_keys["set"] = key_tuple

        def save_hotkey(new_hotkey):
            settings["hotkey"] = new_hotkey
            save_settings(settings)
            btn_hotkey.configure(text=new_hotkey)

        def norm_key(key):
            if key in ["CONTROL", "CTRL"]:
                return "CTRL"
            if key == "ALT":
                return "ALT"
            if key == "SHIFT":
                return "SHIFT"
            if len(key) == 1:
                return key.upper()
            return key.upper()

        def on_hotkey_click(event=None):
            btn_hotkey.configure(text="...нажмите сочетание...")
            btn_hotkey.unbind('<Button-1>')
            root.bind_all('<KeyPress>', on_key_press)
            root.bind_all('<KeyRelease>', on_key_release)
            pressed_keys.clear()
            last_hotkey[0] = btn_hotkey.cget("text")  # сохраняем старый хоткей
            last_logged_keys["set"] = None
            with open(HOTKEY_LOG, "w", encoding="utf-8") as f:
                f.write("")

        pressed_keys = set()
        last_hotkey = [""]
        capturing = {"active": False}

        def on_key_press(event):
            capturing["active"] = True
            key = event.keysym.upper()
            if key in ["SHIFT_L", "SHIFT_R", "CONTROL_L", "CONTROL_R", "ALT_L", "ALT_R"]:
                key = key.replace("_L", "").replace("_R", "")
            key = norm_key(key)
            pressed_keys.add(key)
            display = "+".join(sorted(pressed_keys))
            btn_hotkey.configure(text=display)
            last_hotkey.append(display)
            log_hotkey_event("PRESS", pressed_keys)

        def on_key_release(event):
            key = event.keysym.upper()
            if key in ["SHIFT_L", "SHIFT_R", "CONTROL_L", "CONTROL_R", "ALT_L", "ALT_R"]:
                key = key.replace("_L", "").replace("_R", "")
            key = norm_key(key)
            if key in pressed_keys:
                pressed_keys.remove(key)
            log_hotkey_event("RELEASE", pressed_keys)
            if capturing["active"] and not pressed_keys:
                capturing["active"] = False
                hotkey = last_hotkey[-1].replace("...нажмите сочетание...", "").strip()
                mods = {"CTRL", "ALT", "SHIFT"}
                parts = [k for k in hotkey.split("+") if k]
                if any(k not in mods for k in parts):
                    save_hotkey(hotkey)
                else:
                    btn_hotkey.configure(text=last_hotkey[0])
                btn_hotkey.icursor("end") if hasattr(btn_hotkey, 'icursor') else None
                root.unbind_all('<KeyPress>')
                root.unbind_all('<KeyRelease>')
                btn_hotkey.bind('<Button-1>', on_hotkey_click)

        def on_check():
            settings["show_volume"] = var_show_volume.get()
            save_settings(settings)

        settings = load_settings()

        root = ctk.CTk()
        root.title("Настройки Mic Control")
        root.resizable(False, False)
        root.geometry("840x280")

        icon_path = os.path.abspath("app_icon.ico")
        if os.path.exists(icon_path):
            try:
                root.iconbitmap(icon_path)
            except Exception:
                pass

        current = settings.get("hotkey", "CTRL+ALT+M")

        frame = ctk.CTkFrame(root)
        frame.pack(pady=40, padx=40, fill="x")
        label1 = ctk.CTkLabel(frame, text="Горячая клавиша для мьюта - ", font=("Segoe UI", 22))
        label1.pack(side="left", padx=(0, 10))
        btn_hotkey = ctk.CTkButton(frame, text=current, width=250, height=48, font=("Segoe UI", 22, "bold"), fg_color="#e0e0e0", text_color="#222", hover_color="#d0d0d0")
        btn_hotkey.pack(side="left")
        btn_hotkey.bind('<Button-1>', on_hotkey_click)

        var_show_volume = ctk.BooleanVar(value=settings.get("show_volume", True))
        chk = ctk.CTkCheckBox(root, text="Отображать громкость в иконке", variable=var_show_volume, command=on_check, font=("Segoe UI", 20))
        chk.pack(pady=10)

        btn_close = ctk.CTkButton(root, text="Закрыть", command=root.destroy, width=180, height=48, font=("Segoe UI", 20))
        btn_close.pack(pady=30)

        root.after(10, lambda: center(root))
        root.mainloop()

        # --- Возобновляем хоткей-ловушку после закрытия окна ---
        stop_hotkey_check = False
        def start_hotkey_thread():
            global hotkey_thread
            hotkey_thread = threading.Thread(target=hotkey_check_loop)
            hotkey_thread.daemon = True
            hotkey_thread.start()
        start_hotkey_thread()
        # ------------------------------------------------------
    except Exception as e:
        print(f"[MicControl] Ошибка в окне настроек: {e}")
        import traceback
        traceback.print_exc()

def create_menu():
    """Создаем контекстное меню"""
    # Получаем текущий хоткей
    hotkey = parse_hotkey("")
    hotkey_text = hotkey.upper().replace("+", " + ")

    def on_toggle():
        toggle_microphone()

    def on_settings():
        threading.Thread(target=open_settings_window, daemon=True).start()

    def on_exit():
        icon.stop()

    return pystray.Menu(
        pystray.MenuItem(f'Включить/Выключить микрофон ({hotkey_text})', on_toggle),
        pystray.MenuItem('Настройки', on_settings),
        pystray.MenuItem('Выход', on_exit)
    )

def parse_hotkey(hotkey_str):
    """Парсим строку хоткея в формат для keyboard"""
    try:
        # Читаем хоткей из settings.json
        settings = load_settings()
        hotkey_str = settings.get("hotkey", "CTRL+ALT+M")
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
    global stop_volume_check, icon, microphone, meter, last_mic_id
    last_peak = -1
    last_muted = None
    last_state = None
    last_mic_name = None
    last_mic_id = None
    current_level = None
    zero_start = None  # время, когда начался 0
    idle_mode = False
    force_meter_reinit = False
    while not stop_volume_check:
        try:
            # Проверяем id дефолтного микрофона
            mic_id = get_default_mic_id_coreaudio()
            if mic_id != last_mic_id:
                microphone = None
                meter = None
                last_mic_id = mic_id
                force_meter_reinit = True

            # Если после смены микрофона уровень всегда 0 — пробуем пересоздать meter
            if force_meter_reinit:
                devices = AudioUtilities.GetMicrophone()
                if devices:
                    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                    microphone = cast(interface, POINTER(IAudioEndpointVolume))
                    interface = devices.Activate(IAudioMeterInformation._iid_, CLSCTX_ALL, None)
                    meter = cast(interface, POINTER(IAudioMeterInformation))
                # Проверяем, появился ли уровень
                peak = get_microphone_peak_db()
                if peak > 0:
                    force_meter_reinit = False
            else:
                peak = get_microphone_peak_db()

            peak_percent = int(peak * 100)
            microphone_obj = get_microphone()
            is_muted = microphone_obj.GetMute() if microphone_obj else True
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
            else:
                state = 'Активен'
                icon_name = get_volume_icon_name(peak_percent)

            theme = "dark theme" if is_dark_theme() else "light theme"
            icon_path = os.path.join("Icons", theme, icon_name)

            # Формируем тултип
            mic_name = get_default_mic_name()
            tooltip = f"{mic_name}"

            if (peak_percent != last_peak or is_muted != last_muted or state != last_state or mic_name != last_mic_name) and icon:
                icon.title = tooltip
                icon.icon = Image.open(icon_path)
                last_peak = peak_percent
                last_state = state
                last_mic_name = mic_name
            last_muted = is_muted
        except Exception as e:
            log(f"[MicControl] Ошибка в volume_check_loop: {e}")
        time.sleep(0.1)

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
        print_default_mic_on_start()
        print('Вызов write_registry_audio_tables()')
        write_registry_audio_tables()
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