import customtkinter as ctk
import json
import os
import threading
import keyboard

SETTINGS_PATH = "settings.json"

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

    def save_hotkey(new_hotkey):
        settings["hotkey"] = new_hotkey
        save_settings(settings)
        btn_hotkey.configure(text=new_hotkey)
        print(f"[Settings] Сохранён хоткей: {new_hotkey}")

    def on_hotkey_click(event=None):
        btn_hotkey.configure(text="...нажмите сочетание...")
        btn_hotkey.unbind('<Button-1>')
        btn_hotkey.configure(state="disabled")
        print("[Settings] Жду хоткей (нажми нужное сочетание)...")
        threading.Thread(target=read_hotkey_thread, daemon=True).start()

    def read_hotkey_thread():
        try:
            hotkey = keyboard.read_hotkey(suppress=False)
            parts = [k.upper() if len(k) == 1 else k.upper() for k in hotkey.split('+')]
            pretty = '+'.join(parts)
            root.after(0, lambda: save_hotkey(pretty))
        except Exception as e:
            print(f"[Settings] Ошибка при захвате хоткея: {e}")
        finally:
            root.after(0, lambda: btn_hotkey.configure(state="normal"))
            root.after(0, lambda: btn_hotkey.bind('<Button-1>', on_hotkey_click))

    def on_check():
        settings["show_volume"] = var_show_volume.get()
        save_settings(settings)
        print(f"[Settings] show_volume: {settings['show_volume']}")

    settings = load_settings()
    root = ctk.CTk()
    root.title("Настройки Mic Control")
    root.resizable(False, False)
    root.geometry("600x420")

    icon_path = os.path.abspath("icons/app_icon.ico")
    if os.path.exists(icon_path):
        try:
            root.iconbitmap(icon_path)
        except Exception:
            pass

    current = settings.get("hotkey", "CTRL+ALT+M")

    frame = ctk.CTkFrame(root)
    frame.pack(pady=30, padx=30, fill="x")
    label1 = ctk.CTkLabel(frame, text="Горячая клавиша для мьюта", font=("Segoe UI", 22))
    label1.pack(side="left", padx=(0, 10))
    btn_hotkey = ctk.CTkButton(frame, text=current, width=250, height=48, font=("Segoe UI", 22, "bold"), fg_color="#e0e0e0", text_color="#222", hover_color="#d0d0d0")
    btn_hotkey.pack(side="left")
    btn_hotkey.bind('<Button-1>', on_hotkey_click)

    var_show_volume = ctk.BooleanVar(value=settings.get("show_volume", True))
    chk = ctk.CTkCheckBox(root, text="Отображать громкость в иконке", variable=var_show_volume, command=on_check, font=("Segoe UI", 20))
    chk.pack(pady=10)

    btn_close = ctk.CTkButton(root, text="Закрыть", command=root.destroy, width=180, height=48, font=("Segoe UI", 20))
    btn_close.pack(pady=20)

    root.after(10, lambda: center(root))
    root.mainloop()

if __name__ == "__main__":
    open_settings_window() 