import os
from PIL import Image

def png_to_ico(png_path, ico_path, size=64):
    img = Image.open(png_path).convert("RGBA")
    img = img.resize((size, size), Image.LANCZOS)
    img.save(ico_path, format="ICO", sizes=[(size, size)])

def convert_all_png_to_ico():
    themes = ["light theme", "dark theme"]
    base_dir = "Icons"
    for theme in themes:
        theme_dir = os.path.join(base_dir, theme)
        if not os.path.isdir(theme_dir):
            continue
        for fname in os.listdir(theme_dir):
            if fname.endswith('.png') and ("ic_mic" in fname):
                src_path = os.path.join(theme_dir, fname)
                base_name = fname.split('.')[0]
                ico_path = os.path.join(theme_dir, base_name + ".ico")
                print(f"PNG → ICO: {src_path} -> {ico_path}")
                png_to_ico(src_path, ico_path)

if __name__ == "__main__":
    convert_all_png_to_ico()
    print("Готово! Все .ico для трея сгенерированы.") 