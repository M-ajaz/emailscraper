from PIL import Image
import os

assets = os.path.dirname(__file__)
png_path = os.path.join(assets, "icon.png")
img = Image.open(png_path)

# Windows ICO — multiple sizes
ico_path = os.path.join(assets, "icon.ico")
img.save(ico_path, format="ICO", sizes=[(16,16),(32,32),(48,48),(256,256)])
print(f"Saved {ico_path}")

# macOS — save as large PNG (electron-builder handles ICNS conversion)
icns_png = os.path.join(assets, "icon.icns.png")
img.save(icns_png, "PNG")
print(f"Saved {icns_png}")
