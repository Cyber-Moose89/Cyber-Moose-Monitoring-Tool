from app.gui import ServiceMonitorApp
import tkinter as tk

if __name__ == "__main__":
    root = tk.Tk()
    app = ServiceMonitorApp(root)
    root.mainloop()
