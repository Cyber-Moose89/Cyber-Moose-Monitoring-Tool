from PIL import Image
import pystray
from pystray import MenuItem as item
import threading

def create_tray_icon(app):
    image = Image.open("assets/icon.png")
    menu = (
        item('Show', app.show_window),
        item('Exit', app.exit_app)
    )
    app.tray_icon = pystray.Icon("CyberMoose Watch", image, "CyberMoose Watch", menu)
    threading.Thread(target=app.tray_icon.run, daemon=True).start()
