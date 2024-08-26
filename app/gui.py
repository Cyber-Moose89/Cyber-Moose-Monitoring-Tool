import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import psutil
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import threading
import time
import configparser
from PIL import Image, ImageTk
import pystray
from pystray import MenuItem as item
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import win32serviceutil
import socket
from flask import Flask, jsonify, render_template


class ServiceMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CyberMoose Watch")
        self.root.geometry("1050x900")  # Set a default window size

        # Set window and icon
        self.icon_image = ImageTk.PhotoImage(file="assets/icon.png")
        self.root.iconphoto(True, self.icon_image)  # Taskbar icon

        # Load configuration
        self.config = configparser.ConfigParser()
        self.config.read('config.ini')

        # Initialize
        self.services = []
        self.processes = []
        self.selected_services = self.config.get('MONITORING', 'Services', fallback='').split(',')
        self.selected_processes = self.config.get('MONITORING', 'Processes', fallback='').split(',')
        self.disk_thresholds = {}
        self.cpu_threshold = tk.IntVar(value=self.config.getint('HARDWARE', 'CPU_Threshold', fallback=80))
        self.ram_threshold = tk.IntVar(value=self.config.getint('HARDWARE', 'RAM_Threshold', fallback=80))
        self.max_restart_attempts = tk.IntVar(value=self.config.getint('HARDWARE', 'Max_Restart_Attempts', fallback=3))
        self.auto_restart_service = tk.BooleanVar(
            value=self.config.getboolean('HARDWARE', 'Auto_Restart_Service', fallback=False))

        self.email_frequency = tk.IntVar(value=self.config.getint('EMAIL', 'Frequency', fallback=30))
        self.send_repeat_email = tk.BooleanVar(value=self.config.getboolean('EMAIL', 'SendRepeatEmail', fallback=False))

        self.daily_report = tk.BooleanVar(value=self.config.getboolean('REPORTS', 'DailyReport', fallback=False))
        self.weekly_report = tk.BooleanVar(value=self.config.getboolean('REPORTS', 'WeeklyReport', fallback=False))
        self.monthly_report = tk.BooleanVar(value=self.config.getboolean('REPORTS', 'MonthlyReport', fallback=False))

        self.report_active = tk.BooleanVar(value=self.config.getboolean('REPORTS', 'ReportActive', fallback=False))

        self.last_service_status = {}
        self.last_process_status = {}
        self.last_email_sent = {}

        # Email configuration
        self.smtp_server = tk.StringVar(value=self.config.get('EMAIL', 'SMTP_Server', fallback='smtp.gmail.com'))
        self.smtp_port = tk.StringVar(value=self.config.get('EMAIL', 'SMTP_Port', fallback='587'))
        self.email_from = tk.StringVar(value=self.config.get('EMAIL', 'From', fallback=''))
        self.email_password = tk.StringVar(value=self.config.get('EMAIL', 'Password', fallback=''))
        self.email_to = tk.StringVar(value=self.config.get('EMAIL', 'To', fallback=''))
        self.email_subject = tk.StringVar(value=self.config.get('EMAIL', 'Subject', fallback='Service Alert'))

        # Server configuration
        hostname = socket.gethostname()
        default_ip = socket.gethostbyname(hostname)
        self.server_ip = tk.StringVar(value=self.config.get('SERVER', 'IP', fallback=default_ip))
        self.server_port = tk.IntVar(value=self.config.getint('SERVER', 'Port', fallback=5000))
        self.enable_remote_monitoring = tk.BooleanVar(
            value=self.config.getboolean('SERVER', 'EnableRemoteMonitoring', fallback=False))

        # Tabbed interface
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=1, fill="both")

        # Frames for different tabs
        self.system_status_frame = ttk.Frame(self.notebook)
        self.manage_services_frame = ttk.Frame(self.notebook)
        self.manage_processes_frame = ttk.Frame(self.notebook)
        self.settings_frame = ttk.Frame(self.notebook)
        self.about_frame = ttk.Frame(self.notebook)  # New frame for the About tab

        self.notebook.add(self.system_status_frame, text="System Status")
        self.notebook.add(self.manage_services_frame, text="Manage Services")
        self.notebook.add(self.manage_processes_frame, text="Manage Processes")
        self.notebook.add(self.settings_frame, text="Settings")
        self.notebook.add(self.about_frame, text="About")  # Add the About tab

        # GUI components for each frame
        self.create_system_status_widgets()
        self.create_manage_services_widgets()
        self.create_manage_processes_widgets()
        self.create_settings_widgets()
        self.create_about_widgets()  # Create widgets for the About tab

        # System tray icon
        self.create_tray_icon()

        # Override the close window protocol to minimize to tray
        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        # Start the monitoring process
        self.start_monitoring()

        # Setup Flask server if enabled
        if self.enable_remote_monitoring.get():
            self.setup_flask_routes()
            threading.Thread(target=self.run_flask_server, daemon=True).start()

    def create_about_widgets(self):
        """Create widgets for the About tab."""

        # label frame to hold the about text and make it responsive
        self.about_label_frame = tk.Frame(self.about_frame)
        self.about_label_frame.pack(expand=True, fill="both", padx=10, pady=10)

        about_text = ("CyberMoose Watch\n\n"
                      "Version: 1.0\n\n"
                      "Welcome to CyberMoose Watch, a monitoring application designed to keep your systems \n"
                      "running smoothly. This application offers real-time insights into system services and processes, \n"
                      "automated alerts, and performance reports.\n\n"
                      "With CyberMoose Watch, you can set customizable thresholds for CPU, RAM, and disk usage, ensuring \n"
                      "that you're always aware of how your system resources are being utilized. The tool also supports \n"
                      "scheduled reporting, helping you maintain optimal performance over time.\n\n"
                      "Key Features:\n"
                      "- Real-time monitoring of system services and processes\n"
                      "- Automated alerts and recovery options\n"
                      "- Customizable thresholds for resource usage (CPU, RAM, Disk)\n"
                      "- Regular performance reporting (daily, weekly, monthly)\n\n"
                      "About Me:\n"
                      "I'm Alex Losev, a Python developer with a focus on IT management and system monitoring. My expertise \n"
                      "in system administration is complemented by a strong background in cybersecurity. I strive to create \n"
                      "applications that help you maintain the health and performance of your systems with ease.\n\n"
                      "If you have any questions or need support, feel free to reach out.\n\n"
                      "Contact: +972 52 993 50 37\n\n"
                      "© 2024 Cyber Moose (Alex Losev). All rights reserved.")

        # label with responsive wrapping
        self.about_label = tk.Label(
            self.about_label_frame,
            text=about_text,
            font=('Helvetica', 12),
            justify=tk.LEFT,
            wraplength=self.root.winfo_width() - 40  # Dynamic wraplength based on window size
        )
        self.about_label.pack(expand=True, fill="both", padx=10, pady=10)

        # dynamically adjust the wraplength
        self.about_label_frame.bind("<Configure>", self.update_wraplength)

    def update_wraplength(self, event):
        """Update the wraplength of the about text when the window size changes."""
        new_wraplength = event.width - 40  # wraplength based on new width
        self.about_label.config(wraplength=new_wraplength)

    def create_system_status_widgets(self):
        # Status section
        tk.Label(self.system_status_frame, text="System Status", font=('Helvetica', 16, 'bold')).grid(row=0, column=0,
                                                                                                      columnspan=2, pady=10)

        tk.Label(self.system_status_frame, text="CPU Usage:", font=('Helvetica', 12)).grid(row=1, column=0, sticky='w',
                                                                                           padx=10, pady=5)
        self.cpu_label = tk.Label(self.system_status_frame, text="0%", font=('Helvetica', 12))
        self.cpu_label.grid(row=1, column=1, sticky='w', padx=10, pady=5)

        tk.Label(self.system_status_frame, text="RAM Usage:", font=('Helvetica', 12)).grid(row=2, column=0, sticky='w',
                                                                                           padx=10, pady=5)
        self.ram_label = tk.Label(self.system_status_frame, text="0%", font=('Helvetica', 12))
        self.ram_label.grid(row=2, column=1, sticky='w', padx=10, pady=5)

        self.disk_labels = []
        row = 3
        for partition in psutil.disk_partitions():
            if partition.fstype:
                label = tk.Label(self.system_status_frame, text=f"Disk {partition.device}:", font=('Helvetica', 12))
                label.grid(row=row, column=0, sticky='w', padx=10, pady=5)
                usage_label = tk.Label(self.system_status_frame, text="0% used", font=('Helvetica', 12))
                usage_label.grid(row=row, column=1, sticky='w', padx=10, pady=5)
                self.disk_labels.append((label, usage_label))
                row += 1

        # frames for services and processes
        self.services_scroll_frame = tk.Frame(self.system_status_frame)
        self.services_scroll_frame.grid(row=row, column=0, columnspan=2, pady=(10, 0), sticky='nsew')
        row += 1

        self.processes_scroll_frame = tk.Frame(self.system_status_frame)
        self.processes_scroll_frame.grid(row=row, column=0, columnspan=2, pady=(10, 0), sticky='nsew')

        # canvas to hold the scrollable frame
        self.services_canvas = tk.Canvas(self.services_scroll_frame)
        self.processes_canvas = tk.Canvas(self.processes_scroll_frame)

        # scrollbar to the canvas
        self.services_scrollbar = tk.Scrollbar(self.services_scroll_frame, orient="vertical",
                                               command=self.services_canvas.yview)
        self.services_canvas.configure(yscrollcommand=self.services_scrollbar.set)

        self.processes_scrollbar = tk.Scrollbar(self.processes_scroll_frame, orient="vertical",
                                                command=self.processes_canvas.yview)
        self.processes_canvas.configure(yscrollcommand=self.processes_scrollbar.set)

        # frames inside the canvas
        self.services_frame = tk.Frame(self.services_canvas)
        self.processes_frame = tk.Frame(self.processes_canvas)

        self.services_canvas.create_window((0, 0), window=self.services_frame, anchor="nw")
        self.processes_canvas.create_window((0, 0), window=self.processes_frame, anchor="nw")

        self.services_canvas.grid(row=0, column=0, sticky="nsew")
        self.services_scrollbar.grid(row=0, column=1, sticky="ns")

        self.processes_canvas.grid(row=0, column=0, sticky="nsew")
        self.processes_scrollbar.grid(row=0, column=1, sticky="ns")

        self.services_frame.bind("<Configure>", lambda e: self.services_canvas.configure(
            scrollregion=self.services_canvas.bbox("all")))
        self.processes_frame.bind("<Configure>", lambda e: self.processes_canvas.configure(
            scrollregion=self.processes_canvas.bbox("all")))

        # Control buttons
        self.controls_frame = tk.Frame(self.system_status_frame)
        self.controls_frame.grid(row=row + 1, column=0, columnspan=2, padx=10, pady=10, sticky='ew')

        tk.Button(self.controls_frame, text="Show CPU/RAM Graph", command=self.show_graphical_monitoring,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=0, column=2, sticky='ew', padx=5, pady=5)

        # Bottom label
        self.footer_label = tk.Label(self.system_status_frame, text="Created by Alex Losev | +972 52 993 50 37",
                                     font=('Helvetica', 10, 'italic'))
        self.footer_label.grid(row=row + 2, column=0, columnspan=2, pady=10, sticky='e')

        self.system_status_frame.grid_rowconfigure(row + 1, weight=1)
        self.system_status_frame.grid_columnconfigure(0, weight=1)

    def create_manage_services_widgets(self):
        tk.Label(self.manage_services_frame, text="Manage Services", font=('Helvetica', 16, 'bold')).grid(row=0, column=0,
                                                                                                          columnspan=2,
                                                                                                          pady=10)

        tk.Button(self.manage_services_frame, text="Scan Services", command=self.scan_services, font=('Helvetica', 12),
                  bg="#E0E0E0", fg="black").grid(row=1, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

        tk.Label(self.manage_services_frame, text="Search Services:", font=('Helvetica', 12)).grid(row=2, column=0,
                                                                                                   sticky='w', padx=10,
                                                                                                   pady=5)
        self.service_search_var = tk.StringVar()
        tk.Entry(self.manage_services_frame, textvariable=self.service_search_var, font=('Helvetica', 12)).grid(row=2,
                                                                                                                column=1,
                                                                                                                sticky='ew',
                                                                                                                padx=10,
                                                                                                                pady=5)
        self.service_search_var.trace("w", lambda *args: self.update_filtered_items(self.service_search_var.get(),
                                                                                    "services"))

        self.lst_scanned_services = tk.Listbox(self.manage_services_frame, selectmode=tk.MULTIPLE, bg='#f0f0f0',
                                               fg='#000', font=('Helvetica', 12))
        self.lst_scanned_services.grid(row=3, column=0, sticky='nsew', padx=10, pady=5)

        self.lst_monitored_services = tk.Listbox(self.manage_services_frame, bg='#f0f0f0', fg='#000',
                                                 font=('Helvetica', 12))
        self.lst_monitored_services.grid(row=3, column=1, sticky='nsew', padx=10, pady=5)

        tk.Button(self.manage_services_frame, text="Add to Monitor List", command=self.add_to_monitor_list,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=4, column=0, columnspan=2, sticky='ew',
                                                                         padx=10, pady=5)

        tk.Button(self.manage_services_frame, text="Remove from Monitor List", command=self.remove_from_monitor_list,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=5, column=0, columnspan=2, sticky='ew',
                                                                         padx=10, pady=5)

        tk.Button(self.manage_services_frame, text="Start Service", command=self.start_selected_service,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=6, column=0, columnspan=1, sticky='ew',
                                                                         padx=10, pady=5)
        tk.Button(self.manage_services_frame, text="Stop Service", command=self.stop_selected_service,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=6, column=1, columnspan=1, sticky='ew',
                                                                         padx=10, pady=5)
        tk.Button(self.manage_services_frame, text="Restart Service", command=self.restart_selected_service,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=7, column=0, columnspan=2, sticky='ew',
                                                                         padx=10, pady=5)

        tk.Button(self.manage_services_frame, text="Load Monitored Items", command=self.load_monitored_items,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=8, column=0, columnspan=2, sticky='ew',
                                                                         padx=10, pady=5)

        tk.Button(self.manage_services_frame, text="Save Settings", command=self.save_monitoring_settings,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=9, column=0, columnspan=2, sticky='ew',
                                                                         padx=10, pady=5)

        self.manage_services_frame.grid_rowconfigure(3, weight=1)
        self.manage_services_frame.grid_columnconfigure(1, weight=1)

    def create_manage_processes_widgets(self):
        tk.Label(self.manage_processes_frame, text="Manage Processes", font=('Helvetica', 16, 'bold')).grid(row=0, column=0,
                                                                                                            columnspan=2,
                                                                                                            pady=10)

        tk.Button(self.manage_processes_frame, text="Scan Processes", command=self.scan_processes, font=('Helvetica', 12),
                  bg="#E0E0E0", fg="black").grid(row=1, column=0, columnspan=2, sticky='ew', padx=10, pady=5)

        tk.Label(self.manage_processes_frame, text="Search Processes:", font=('Helvetica', 12)).grid(row=2, column=0,
                                                                                                     sticky='w', padx=10,
                                                                                                     pady=5)
        self.process_search_var = tk.StringVar()
        tk.Entry(self.manage_processes_frame, textvariable=self.process_search_var, font=('Helvetica', 12)).grid(row=2,
                                                                                                                 column=1,
                                                                                                                 sticky='ew',
                                                                                                                 padx=10,
                                                                                                                 pady=5)
        self.process_search_var.trace("w", lambda *args: self.update_filtered_items(self.process_search_var.get(),
                                                                                    "processes"))

        self.lst_scanned_processes = tk.Listbox(self.manage_processes_frame, selectmode=tk.MULTIPLE, bg='#f0f0f0',
                                                fg='#000', font=('Helvetica', 12))
        self.lst_scanned_processes.grid(row=3, column=0, sticky='nsew', padx=10, pady=5)

        self.lst_monitored_processes = tk.Listbox(self.manage_processes_frame, bg='#f0f0f0', fg='#000',
                                                  font=('Helvetica', 12))
        self.lst_monitored_processes.grid(row=3, column=1, sticky='nsew', padx=10, pady=5)

        tk.Button(self.manage_processes_frame, text="Add to Monitor List", command=self.add_to_monitor_list,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=4, column=0, columnspan=2, sticky='ew',
                                                                         padx=10, pady=5)

        tk.Button(self.manage_processes_frame, text="Remove from Monitor List", command=self.remove_from_monitor_list,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=5, column=0, columnspan=2, sticky='ew',
                                                                         padx=10, pady=5)

        tk.Button(self.manage_processes_frame, text="Load Processes", command=self.load_monitored_items,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=6, column=0, columnspan=2, sticky='ew',
                                                                         padx=10, pady=5)

        # Add the Save Settings button
        tk.Button(self.manage_processes_frame, text="Save Settings", command=self.save_monitoring_settings,
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=7, column=0, columnspan=2, sticky='ew',
                                                                         padx=10, pady=5)

        self.manage_processes_frame.grid_rowconfigure(3, weight=1)
        self.manage_processes_frame.grid_columnconfigure(1, weight=1)

    def create_settings_widgets(self):
        tk.Label(self.settings_frame, text="Settings", font=('Helvetica', 16, 'bold')).grid(row=0, column=0, columnspan=4,
                                                                                            pady=10)

        # Column 1
        tk.Label(self.settings_frame, text="SMTP Server:", font=('Helvetica', 12)).grid(row=1, column=0, sticky='w',
                                                                                        padx=10, pady=5)
        tk.Entry(self.settings_frame, textvariable=self.smtp_server, font=('Helvetica', 12)).grid(row=1, column=1,
                                                                                                  sticky='ew', padx=10,
                                                                                                  pady=5)

        tk.Label(self.settings_frame, text="SMTP Port:", font=('Helvetica', 12)).grid(row=2, column=0, sticky='w', padx=10,
                                                                                      pady=5)
        tk.Entry(self.settings_frame, textvariable=self.smtp_port, font=('Helvetica', 12)).grid(row=2, column=1,
                                                                                                sticky='ew', padx=10,
                                                                                                pady=5)

        tk.Label(self.settings_frame, text="Email From:", font=('Helvetica', 12)).grid(row=3, column=0, sticky='w', padx=10,
                                                                                       pady=5)
        tk.Entry(self.settings_frame, textvariable=self.email_from, font=('Helvetica', 12)).grid(row=3, column=1,
                                                                                                 sticky='ew', padx=10,
                                                                                                 pady=5)

        tk.Label(self.settings_frame, text="Email Password:", font=('Helvetica', 12)).grid(row=4, column=0, sticky='w',
                                                                                           padx=10, pady=5)
        tk.Entry(self.settings_frame, textvariable=self.email_password, show="*", font=('Helvetica', 12)).grid(row=4,
                                                                                                               column=1,
                                                                                                               sticky='ew',
                                                                                                               padx=10,
                                                                                                               pady=5)

        tk.Label(self.settings_frame, text="Email To (comma-separated):", font=('Helvetica', 12)).grid(row=5, column=0,
                                                                                                       sticky='w', padx=10,
                                                                                                       pady=5)
        tk.Entry(self.settings_frame, textvariable=self.email_to, font=('Helvetica', 12)).grid(row=5, column=1, sticky='ew',
                                                                                               padx=10, pady=5)

        tk.Label(self.settings_frame, text="Email Subject:", font=('Helvetica', 12)).grid(row=6, column=0, sticky='w',
                                                                                          padx=10, pady=5)
        tk.Entry(self.settings_frame, textvariable=self.email_subject, font=('Helvetica', 12)).grid(row=6, column=1,
                                                                                                    sticky='ew', padx=10,
                                                                                                    pady=5)

        tk.Label(self.settings_frame, text="CPU Threshold (%):", font=('Helvetica', 12)).grid(row=7, column=0, sticky='w',
                                                                                              padx=10, pady=5)
        tk.Entry(self.settings_frame, textvariable=self.cpu_threshold, font=('Helvetica', 12)).grid(row=7, column=1,
                                                                                                    sticky='ew', padx=10,
                                                                                                    pady=5)

        tk.Label(self.settings_frame, text="RAM Threshold (%):", font=('Helvetica', 12)).grid(row=8, column=0, sticky='w',
                                                                                              padx=10, pady=5)
        tk.Entry(self.settings_frame, textvariable=self.ram_threshold, font=('Helvetica', 12)).grid(row=8, column=1,
                                                                                                    sticky='ew', padx=10,
                                                                                                    pady=5)

        # Column 2
        tk.Label(self.settings_frame, text="Email Frequency (minutes):", font=('Helvetica', 12)).grid(row=1, column=2,
                                                                                                      sticky='w', padx=10,
                                                                                                      pady=5)
        tk.Entry(self.settings_frame, textvariable=self.email_frequency, font=('Helvetica', 12)).grid(row=1, column=3,
                                                                                                      sticky='ew', padx=10,
                                                                                                      pady=5)

        tk.Checkbutton(self.settings_frame, text="Send Repeat Emails", variable=self.send_repeat_email,
                       font=('Helvetica', 12)).grid(row=2, column=2, columnspan=2, padx=10, pady=5)

        tk.Label(self.settings_frame, text="Max Restart Attempts:", font=('Helvetica', 12)).grid(row=3, column=2,
                                                                                                 sticky='w', padx=10,
                                                                                                 pady=5)
        tk.Entry(self.settings_frame, textvariable=self.max_restart_attempts, font=('Helvetica', 12)).grid(row=3, column=3,
                                                                                                           sticky='ew',
                                                                                                           padx=10, pady=5)

        tk.Checkbutton(self.settings_frame, text="Auto Restart Services", variable=self.auto_restart_service,
                       font=('Helvetica', 12)).grid(row=4, column=2, columnspan=2, padx=10, pady=5)

        tk.Checkbutton(self.settings_frame, text="Daily Report", variable=self.daily_report, font=('Helvetica', 12)).grid(
            row=5, column=2, columnspan=2, padx=10, pady=5)
        tk.Checkbutton(self.settings_frame, text="Weekly Report", variable=self.weekly_report, font=('Helvetica', 12)).grid(
            row=6, column=2, columnspan=2, padx=10, pady=5)
        tk.Checkbutton(self.settings_frame, text="Monthly Report", variable=self.monthly_report,
                       font=('Helvetica', 12)).grid(row=7, column=2, columnspan=2, padx=10, pady=5)

        tk.Checkbutton(self.settings_frame, text="Activate Reports", variable=self.report_active,
                       font=('Helvetica', 12)).grid(row=8, column=2, columnspan=2, padx=10, pady=5)

        tk.Label(self.settings_frame, text="Server IP:", font=('Helvetica', 12)).grid(row=9, column=0, sticky='w', padx=10,
                                                                                      pady=5)
        tk.Entry(self.settings_frame, textvariable=self.server_ip, font=('Helvetica', 12)).grid(row=9, column=1,
                                                                                                sticky='ew', padx=10,
                                                                                                pady=5)

        tk.Label(self.settings_frame, text="Server Port:", font=('Helvetica', 12)).grid(row=9, column=2, sticky='w',
                                                                                        padx=10, pady=5)
        tk.Entry(self.settings_frame, textvariable=self.server_port, font=('Helvetica', 12)).grid(row=9, column=3,
                                                                                                  sticky='ew', padx=10,
                                                                                                  pady=5)

        tk.Checkbutton(self.settings_frame, text="Enable Remote Monitoring", variable=self.enable_remote_monitoring,
                       font=('Helvetica', 12)).grid(row=10, column=0, columnspan=4, padx=10, pady=5)

        tk.Button(self.settings_frame, text="Set Disk Thresholds", command=self.set_disk_thresholds, font=('Helvetica', 12),
                  bg="#E0E0E0", fg="black").grid(row=11, column=0, columnspan=2, sticky='ew', padx=10, pady=5)
        tk.Button(self.settings_frame, text="Test Connection", command=self.test_email_connection, font=('Helvetica', 12),
                  bg="#E0E0E0", fg="black").grid(row=11, column=2, columnspan=2, sticky='ew', padx=10, pady=5)

        tk.Button(self.settings_frame, text="Save", command=self.save_settings, font=('Helvetica', 12), bg="#E0E0E0",
                  fg="black").grid(row=12, column=0, columnspan=4, sticky='ew', padx=10, pady=5)

        tk.Button(self.settings_frame, text="Send Instant Report", command=self.send_instant_report, font=('Helvetica', 12),
                  bg="#E0E0E0", fg="black").grid(row=13, column=0, columnspan=4, pady=10, sticky='ew', padx=10)

        tk.Button(self.settings_frame, text="Back", command=lambda: self.show_frame(self.system_status_frame),
                  font=('Helvetica', 12), bg="#E0E0E0", fg="black").grid(row=14, column=0, columnspan=4, pady=10,
                                                                         sticky='ew', padx=10)

    def show_frame(self, frame):
        self.notebook.select(frame)

    def create_tray_icon(self):
        # icon image for the tray
        image = Image.open("assets/icon.png")

        # tray icon menu
        menu = (
            item('Show', self.show_window),
            item('Exit', self.exit_app)
        )

        # tray icon
        self.tray_icon = pystray.Icon("CyberMoose Watch", image, "CyberMoose Watch", menu)

        # tray icon in a separate thread
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def hide_window(self):
        self.root.withdraw()

    def show_window(self, icon=None, item=None):
        self.root.deiconify()

    def exit_app(self, icon=None, item=None):
        self.tray_icon.stop()
        self.root.quit()

    def set_disk_thresholds(self):
        for partition in psutil.disk_partitions():
            if partition.fstype:  # Only consider partitions with a filesystem
                threshold = simpledialog.askinteger("Disk Threshold", f"Set threshold for {partition.device} (% used):",
                                                    minvalue=1, maxvalue=100)
                if threshold is not None:
                    self.disk_thresholds[partition.device] = threshold

    def test_email_connection(self):
        smtp_server = self.smtp_server.get()
        smtp_port = self.smtp_port.get()
        email_from = self.email_from.get()
        email_password = self.email_password.get()

        try:
            server = smtplib.SMTP(smtp_server, int(smtp_port))
            server.starttls()
            server.login(email_from, email_password)
            server.quit()
            messagebox.showinfo("Connection Test", "Connection successful!")
        except Exception as e:
            messagebox.showerror("Connection Test Failed", f"Failed to connect: {e}")

    def save_settings(self):
        self.config['EMAIL'] = {
            'SMTP_Server': self.smtp_server.get(),
            'SMTP_Port': self.smtp_port.get(),
            'From': self.email_from.get(),
            'Password': self.email_password.get(),
            'To': self.email_to.get(),
            'Subject': self.email_subject.get(),  # Ensure this is saving correctly
            'Frequency': str(self.email_frequency.get()),
            'SendRepeatEmail': str(self.send_repeat_email.get())
        }

        self.config['HARDWARE'] = {
            'CPU_Threshold': str(self.cpu_threshold.get()),
            'RAM_Threshold': str(self.ram_threshold.get()),
            'Max_Restart_Attempts': str(self.max_restart_attempts.get()),
            'Auto_Restart_Service': str(self.auto_restart_service.get())
        }

        self.config['REPORTS'] = {
            'DailyReport': str(self.daily_report.get()),
            'WeeklyReport': str(self.weekly_report.get()),
            'MonthlyReport': str(self.monthly_report.get()),
            'ReportActive': str(self.report_active.get())
        }

        self.config['SERVER'] = {
            'IP': self.server_ip.get(),
            'Port': str(self.server_port.get()),
            'EnableRemoteMonitoring': str(self.enable_remote_monitoring.get())
        }

        with open('config.ini', 'w') as configfile:
            self.config.write(configfile)

        messagebox.showinfo("Settings", "Settings saved successfully!")

        # Restart Flask server if the remote monitoring setting is changed
        if self.enable_remote_monitoring.get():
            self.setup_flask_routes()
            threading.Thread(target=self.run_flask_server, daemon=True).start()

    def save_monitoring_settings(self):
        self.config['MONITORING'] = {
            'Services': ','.join(self.selected_services),
            'Processes': ','.join(self.selected_processes)
        }

        with open('config.ini', 'w') as configfile:
            self.config.write(configfile)

        messagebox.showinfo("Settings", "Monitoring settings saved successfully!")

    def scan_services(self):
        self.services = [service.name() for service in psutil.win_service_iter()]
        self.update_filtered_items("", "services")

    def scan_processes(self):
        self.processes = [proc.info['name'] for proc in psutil.process_iter(['name'])]
        self.update_filtered_items("", "processes")

    def update_filtered_items(self, search_text, item_type):
        if item_type == "services":
            filtered_services = [service for service in self.services if search_text.lower() in service.lower()]
            self.lst_scanned_services.delete(0, tk.END)
            for service in filtered_services:
                self.lst_scanned_services.insert(tk.END, service)
        elif item_type == "processes":
            filtered_processes = [proc for proc in self.processes if search_text.lower() in proc.lower()]
            self.lst_scanned_processes.delete(0, tk.END)
            for proc in filtered_processes:
                self.lst_scanned_processes.insert(tk.END, proc)

    def add_to_monitor_list(self):
        selected_services = self.lst_scanned_services.curselection()
        selected_processes = self.lst_scanned_processes.curselection()

        for i in selected_services:
            service_name = self.lst_scanned_services.get(i)
            if service_name and service_name not in self.lst_monitored_services.get(0, tk.END):
                self.lst_monitored_services.insert(tk.END, service_name)
                self.selected_services.append(service_name)
                self.last_service_status[service_name] = "Running"  # Assuming the service is running initially

        for i in selected_processes:
            proc_name = self.lst_scanned_processes.get(i)
            if proc_name and proc_name not in self.lst_monitored_processes.get(0, tk.END):
                self.lst_monitored_processes.insert(tk.END, proc_name)
                self.selected_processes.append(proc_name)
                self.last_process_status[proc_name] = "Running"  # Assuming the process is running initially

        self.refresh_status()

    def remove_from_monitor_list(self):
        selected_services = self.lst_monitored_services.curselection()
        selected_processes = self.lst_monitored_processes.curselection()

        for i in reversed(selected_services):
            service_name = self.lst_monitored_services.get(i)
            self.lst_monitored_services.delete(i)
            if service_name in self.selected_services:
                self.selected_services.remove(service_name)
                del self.last_service_status[service_name]

        for i in reversed(selected_processes):
            proc_name = self.lst_monitored_processes.get(i)
            self.lst_monitored_processes.delete(i)
            if proc_name in self.selected_processes:
                self.selected_processes.remove(proc_name)
                del self.last_process_status[proc_name]

        self.refresh_status()

    def start_monitoring(self):
        self.monitoring_thread = threading.Thread(target=self.monitor_services_and_processes, daemon=True)
        self.monitoring_thread.start()

    def monitor_services_and_processes(self):
        while True:
            self.root.after(0, self.refresh_status)
            self.check_services()
            self.check_processes()
            self.check_cpu_ram_usage()
            self.check_disk_space()
            self.generate_reports_if_needed()
            time.sleep(5)  # Check every 5 seconds for live update

    def refresh_status(self):
        self.cpu_label.config(text=f"{psutil.cpu_percent(interval=1)}%")
        self.ram_label.config(text=f"{psutil.virtual_memory().percent}%")

        for i, (label, usage_label) in enumerate(self.disk_labels):
            partition = psutil.disk_partitions()[i]
            if partition.fstype:  # Check if the partition has a file system type
                usage = psutil.disk_usage(partition.mountpoint).percent
                usage_label.config(text=f"{usage}% used")

        # Refresh the status of services and processes here (not shown)

        for widget in self.services_frame.winfo_children():
            widget.destroy()

        for service in self.selected_services:
            service_status = self.get_service_status(service)
            tk.Label(self.services_frame, text=f"{service}: {service_status}", font=('Helvetica', 12)).grid(sticky="w")

        for widget in self.processes_frame.winfo_children():
            widget.destroy()

        for proc in self.selected_processes:
            proc_status = "Running" if proc in self.processes else "Running"
            tk.Label(self.processes_frame, text=f"{proc}: {proc_status}", font=('Helvetica', 12)).grid(sticky="w")

    def get_service_status(self, service_name):
        if not service_name:
            return "Monitored"
        try:
            service = psutil.win_service_get(service_name)
            return "Running" if service.status() == 'running' else "Stopped"
        except Exception as e:
            return "Not Found"

    def check_services(self):
        for service_name in self.selected_services:
            if service_name:
                current_status = "Running" if self.is_service_running(service_name) else "Stopped"
                if service_name not in self.last_service_status or current_status != self.last_service_status[
                    service_name]:
                    self.handle_service_status_change(service_name, current_status)
                    self.last_service_status[service_name] = current_status

                    if current_status == "Stopped" and self.auto_restart_service.get():
                        self.attempt_service_restart(service_name)

    def check_processes(self):
        running_processes = [proc.info['name'] for proc in psutil.process_iter(['name'])]
        for proc_name in self.selected_processes:
            if proc_name:
                current_status = "Running" if proc_name in running_processes else "Not Running"
                if proc_name not in self.last_process_status or current_status != self.last_process_status[proc_name]:
                    self.handle_process_status_change(proc_name, current_status)
                    self.last_process_status[proc_name] = current_status

    def handle_service_status_change(self, service_name, status):
        previous_status = self.last_service_status.get(service_name, "Unknown")
        subject = f"Service {status}"
        detailed_body = f"The service <strong>{service_name}</strong> is now in a <strong>{status}</strong> state."
        self.send_email(subject, detailed_body, service_or_process_name=service_name, status=status,
                        previous_status=previous_status)

    def handle_process_status_change(self, proc_name, status):
        previous_status = self.last_process_status.get(proc_name, "Unknown")
        subject = f"Process {status}"
        detailed_body = f"The process <strong>{proc_name}</strong> is now in a <strong>{status}</strong> state."
        self.send_email(subject, detailed_body, service_or_process_name=proc_name, status=status,
                        previous_status=previous_status)

    def attempt_service_restart(self, service_name):
        """Attempt to restart the service up to the maximum number of times specified."""
        success = False
        for attempt in range(1, self.max_restart_attempts.get() + 1):
            try:
                win32serviceutil.RestartService(service_name)
                success = True
                break
            except Exception as e:
                if attempt == self.max_restart_attempts.get():
                    success = False
                time.sleep(5)  #!!!!!!! Wait a bit before trying again !!!!!!#

        self.send_restart_report(service_name, success, attempt)

    def send_restart_report(self, service_name, success, attempt):
        """Send an email report after attempting to restart a service."""
        status = "Success" if success else "Failure"
        subject = f"Service Restart {status}: {service_name}"
        body = f"""
        <html>
        <body>
        <p>The service <strong>{service_name}</strong> was {status.lower()}fully restarted on attempt {attempt}.</p>
        <p><strong>Final Status:</strong> {status}</p>
        </body>
        </html>
        """
        self.send_email(subject, body)

    def check_cpu_ram_usage(self):
        cpu_usage = psutil.cpu_percent(interval=1)
        ram_usage = psutil.virtual_memory().percent

        if cpu_usage > self.cpu_threshold.get():
            self.handle_hardware_overload("CPU Usage", cpu_usage, self.cpu_threshold.get())
        if ram_usage > self.ram_threshold.get():
            self.handle_hardware_overload("RAM Usage", ram_usage, self.ram_threshold.get())

    def check_disk_space(self):
        for partition in psutil.disk_partitions():
            if partition.device in self.disk_thresholds:
                usage = psutil.disk_usage(partition.mountpoint).percent
                if usage > self.disk_thresholds[partition.device]:
                    self.handle_hardware_overload(f"Disk Space {partition.device}", usage,
                                                  self.disk_thresholds[partition.device])

    def handle_hardware_overload(self, name, current_usage, threshold):
        # current system uptime
        uptime_seconds = time.time() - psutil.boot_time()
        uptime = time.strftime("%H:%M:%S", time.gmtime(uptime_seconds))

        # current CPU and RAM usage details
        cpu_usage = psutil.cpu_percent(interval=1)
        ram_usage = psutil.virtual_memory().percent

        # disk usage details
        partitions = psutil.disk_partitions()
        disk_usage_details = []
        for partition in partitions:
            if partition.fstype:
                usage = psutil.disk_usage(partition.mountpoint).percent
                disk_usage_details.append(f"{partition.device}: {usage}% used")

        # top processes by CPU and memory
        top_processes_by_cpu = sorted(psutil.process_iter(['name', 'cpu_percent']),
                                      key=lambda p: p.info.get('cpu_percent', 0), reverse=True)[:5]
        top_processes_by_memory = sorted(psutil.process_iter(['name', 'memory_percent']),
                                         key=lambda p: p.info.get('memory_percent', 0), reverse=True)[:5]

        top_cpu_processes = "\n".join(
            [f"{proc.info['name']}: {proc.info.get('cpu_percent', 0)}% CPU" for proc in top_processes_by_cpu])

        top_memory_processes = "\n".join(
            [f"{proc.info['name']}: {proc.info['memory_percent']:.2f}% Memory" for proc in top_processes_by_memory])

        # detailed message
        detailed_body = f"""
        <html>
        <body>
        <h2><b>Hardware Overload Detected: {name}</b></h2>
        <p><b>Alert Timestamp:</b> {time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())}</p>
        <p><b>Current System Uptime:</b> {uptime}</p>

        <h3><b>Current Resource Usage:</b></h3>
        <ul>
            <li><b>CPU Usage:</b> {cpu_usage}%</li>
            <li><b>RAM Usage:</b> {ram_usage}%</li>
        </ul>

        <h3><b>Disk Usage:</b></h3>
        <ul>
            {"".join([f"<li>{detail}</li>" for detail in disk_usage_details])}
        </ul>

        <h3><b>Exceeded Threshold:</b></h3>
        <p>The {name} has exceeded the threshold:</p>
        <ul>
            <li><b>Current Usage:</b> {current_usage}%</li>
            <li><b>Threshold:</b> {threshold}%</li>
            <li><b>Exceeded by:</b> {current_usage - threshold}%</li>
        </ul>

        <h3><b>Top Processes by CPU Usage:</b></h3>
        <ul>
            {top_cpu_processes}
        </ul>

        <h3><b>Top Processes by Memory Usage:</b></h3>
        <ul>
            {top_memory_processes}
        </ul>

        </body>
        </html>
        """

        if name not in self.last_email_sent or (
                time.time() - self.last_email_sent[name]) > self.email_frequency.get() * 60:
            if self.send_repeat_email.get() or name not in self.last_email_sent:
                self.send_email(f"Hardware Overload: {name}", detailed_body, service_or_process_name=name,
                                status="Overloaded")
                self.last_email_sent[name] = time.time()

    def is_service_running(self, service_name):
        if not service_name:
            return False
        try:
            service = psutil.win_service_get(service_name)
            if service.status() == 'running':
                return True
        except Exception as e:
            print(f"Service {service_name} not found: {e}")
        return False

    def send_email(self, subject, body, service_or_process_name=None, status=None, previous_status=None,
                   custom_description=None):
        smtp_server = self.smtp_server.get()
        smtp_port = self.smtp_port.get()
        email_from = self.email_from.get()
        email_password = self.email_password.get()
        email_to_list = self.email_to.get().split(',')

        # Update the subject to ensure it's the latest from the settings
        subject = self.email_subject.get()

        # additional system details
        cpu_usage = psutil.cpu_percent(interval=1)
        ram_usage = psutil.virtual_memory().percent
        memory_info = psutil.virtual_memory()
        load_avg = psutil.getloadavg() if hasattr(psutil, 'getloadavg') else ('N/A', 'N/A', 'N/A')
        disk_usage_details = "\n".join([
                                           f"{partition.device}: {psutil.disk_usage(partition.mountpoint).percent}% used, Free: {psutil.disk_usage(partition.mountpoint).free // (1024 ** 2)} MB"
                                           for partition in psutil.disk_partitions() if partition.fstype])
        uptime_seconds = time.time() - psutil.boot_time()
        uptime_string = time.strftime("%H:%M:%S", time.gmtime(uptime_seconds))
        network_info = psutil.net_io_counters()
        network_usage = f"Sent: {network_info.bytes_sent} bytes, Received: {network_info.bytes_recv} bytes"
        active_processes = len(psutil.pids())
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())

        # Additional top processes by CPU and RAM usage
        top_processes_by_cpu = sorted(psutil.process_iter(['name', 'cpu_percent']),
                                      key=lambda p: p.info.get('cpu_percent', 0), reverse=True)[:5]
        top_processes_by_ram = sorted(psutil.process_iter(['name', 'memory_percent']),
                                      key=lambda p: p.info.get('memory_percent', 0), reverse=True)[:5]

        top_cpu_processes_details = "\n".join(
            [f"{proc.info['name']}: {proc.info.get('cpu_percent', 'N/A')}% CPU" for proc in top_processes_by_cpu])
        top_ram_processes_details = "\n".join(
            [f"{proc.info['name']}: {proc.info['memory_percent']:.2f}% RAM" for proc in top_processes_by_ram])

        # service or process name, make it bold in the body
        if service_or_process_name:
            body = f"""
            <html>
            <body>
            <p><strong>Alert:</strong> The following {('service' if 'Service' in subject else 'process')} has changed its status:</p>
            <p><strong>Name:</strong> <strong>{service_or_process_name}</strong></p>
            <p><strong>New Status:</strong> <strong>{status}</strong></p>
            <p><strong>Previous Status:</strong> {previous_status}</p>
            <p><strong>Description:</strong> {custom_description or 'N/A'}</p>
            <p><strong>Timestamp:</strong> {timestamp}</p>
            <hr>
            <p><strong>System Status at the Time of Alert:</strong></p>
            <ul>
                <li><strong>CPU Usage:</strong> {cpu_usage}%</li>
                <li><strong>RAM Usage:</strong> {ram_usage}% (Total: {memory_info.total // (1024 ** 2)} MB, Available: {memory_info.available // (1024 ** 2)} MB)</li>
                <li><strong>Disk Usage:</strong><br>{disk_usage_details}</li>
                <li><strong>Network Usage:</strong> {network_usage}</li>
                <li><strong>Active Processes:</strong> {active_processes}</li>
                <li><strong>System Uptime:</strong> {uptime_string}</li>
                <li><strong>Load Average (1, 5, 15 min):</strong> {load_avg[0]}, {load_avg[1]}, {load_avg[2]}</li>
            </ul>
            <hr>
            <p><strong>Top Processes by CPU Usage:</strong></p>
            <pre>{top_cpu_processes_details}</pre>
            <p><strong>Top Processes by RAM Usage:</strong></p>
            <pre>{top_ram_processes_details}</pre>
            <hr>
            <p>This is an automated message from CyberMoose Watch.</p>
            </body>
            </html>
            """
        else:
            body = f"""
            <html>
            <body>
            <p>{body}</p>
            <hr>
            <p>This is an automated message from CyberMoose Watch.</p>
            </body>
            </html>
            """

        msg = MIMEMultipart("alternative")
        msg['From'] = email_from
        msg['To'] = ', '.join(email_to_list)
        msg['Subject'] = subject  # Ensure the latest subject is used here

        # HTML version of the body
        msg.attach(MIMEText(body, 'html'))

        try:
            server = smtplib.SMTP(smtp_server, int(smtp_port))
            server.starttls()
            server.login(email_from, email_password)
            server.sendmail(email_from, email_to_list, msg.as_string())
            server.quit()
            print(f"Email sent with subject: {subject}")
        except Exception as e:
            print(f"Failed to send email: {e}")

    def load_monitored_items(self):
        """Load the monitored services and processes from config and populate the listboxes."""
        for service_name in self.selected_services:
            if service_name and service_name not in self.lst_monitored_services.get(0, tk.END):
                self.lst_monitored_services.insert(tk.END, service_name)

        for proc_name in self.selected_processes:
            if proc_name and proc_name not in self.lst_monitored_processes.get(0, tk.END):
                self.lst_monitored_processes.insert(tk.END, proc_name)

    def start_selected_service(self):
        selected_services = self.lst_monitored_services.curselection()
        for i in selected_services:
            service_name = self.lst_monitored_services.get(i)
            if service_name:
                self.control_service(service_name, 'start')

    def stop_selected_service(self):
        selected_services = self.lst_monitored_services.curselection()
        for i in selected_services:
            service_name = self.lst_monitored_services.get(i)
            if service_name:
                self.control_service(service_name, 'stop')

    def restart_selected_service(self):
        selected_services = self.lst_monitored_services.curselection()
        for i in selected_services:
            service_name = self.lst_monitored_services.get(i)
            if service_name:
                self.control_service(service_name, 'restart')

    def control_service(self, service_name, action):
        try:
            service = psutil.win_service_get(service_name)
            if action == 'start' and service.status() != 'running':
                win32serviceutil.StartService(service_name)
            elif action == 'stop' and service.status() == 'running':
                win32serviceutil.StopService(service_name)
            elif action == 'restart':
                win32serviceutil.RestartService(service_name)
            print(f"Service '{service_name}' {action}ed successfully.")
        except Exception as e:
            print(f"Failed to {action} service '{service_name}': {e}")

    def show_graphical_monitoring(self):
        #  new window for the graphs
        graph_window = tk.Toplevel(self.root)
        graph_window.title("Graphical Monitoring")
        graph_window.geometry("800x900")  # Set a default window size

        #  CPU usage graph
        cpu_fig, cpu_ax = plt.subplots()
        cpu_ax.set_title("CPU Usage Over Time")
        cpu_ax.set_xlabel("Time (s)")
        cpu_ax.set_ylabel("CPU Usage (%)")
        self.cpu_line, = cpu_ax.plot([], [], 'r-')
        cpu_canvas = FigureCanvasTkAgg(cpu_fig, master=graph_window)
        cpu_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        #  RAM usage graph
        ram_fig, ram_ax = plt.subplots()
        ram_ax.set_title("RAM Usage Over Time")
        ram_ax.set_xlabel("Time (s)")
        ram_ax.set_ylabel("RAM Usage (%)")
        self.ram_line, = ram_ax.plot([], [], 'b-')
        ram_canvas = FigureCanvasTkAgg(ram_fig, master=graph_window)
        ram_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # updating the graphs in a separate thread
        threading.Thread(target=self.update_graphs, args=(cpu_ax, ram_ax), daemon=True).start()

    def update_graphs(self, cpu_ax, ram_ax):
        cpu_usage_data = []
        ram_usage_data = []
        x_data = []
        start_time = time.time()

        while True:
            # Append the current CPU and RAM usage
            cpu_usage_data.append(psutil.cpu_percent(interval=1))
            ram_usage_data.append(psutil.virtual_memory().percent)
            x_data.append(time.time() - start_time)

            # Update the data of the lines
            self.cpu_line.set_data(x_data, cpu_usage_data)
            self.ram_line.set_data(x_data, ram_usage_data)

            # Adjust the axes limits
            cpu_ax.set_xlim(0, max(x_data) + 1)
            ram_ax.set_xlim(0, max(x_data) + 1)
            cpu_ax.set_ylim(0, 100)
            ram_ax.set_ylim(0, 100)

            # Redraw  canvas
            cpu_ax.figure.canvas.draw()
            ram_ax.figure.canvas.draw()

            time.sleep(1)  # Update every second

    def generate_reports_if_needed(self):
        if not self.report_active.get():
            return

        current_time = time.time()
        last_report_time = self.config.getfloat('REPORTS', 'LastReportTime', fallback=0)

        if self.daily_report.get() and (current_time - last_report_time) >= 86400:
            self.generate_report('Daily')
        elif self.weekly_report.get() and (current_time - last_report_time) >= 604800:
            self.generate_report('Weekly')
        elif self.monthly_report.get() and (current_time - last_report_time) >= 2592000:
            self.generate_report('Monthly')

    def generate_report(self, report_type):
        report_subject = f"{report_type} System Report"
        report_body = self.get_report_body(report_type)
        self.send_email(report_subject, report_body)
        self.config.set('REPORTS', 'LastReportTime', str(time.time()))
        with open('config.ini', 'w') as configfile:
            self.config.write(configfile)

    def send_instant_report(self):
        report_subject = "Instant System Report"
        report_body = self.get_report_body("Instant")
        self.send_email(report_subject, report_body)

    def get_report_body(self, report_type):
        """Generate the report content with system information."""
        cpu_usage = psutil.cpu_percent(interval=1)
        ram_usage = psutil.virtual_memory().percent
        uptime_seconds = time.time() - psutil.boot_time()
        uptime = time.strftime("%H:%M:%S", time.gmtime(uptime_seconds))
        disk_usage_details = "<br>".join([
            "{}: {}% used".format(partition.device, psutil.disk_usage(partition.mountpoint).percent)
            for partition in psutil.disk_partitions() if partition.fstype
        ])

        body = f"""
        <html>
        <body>
        <h2>{report_type} System Report</h2>
        <p><b>Timestamp:</b> {time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())}</p>
        <p><b>CPU Usage:</b> {cpu_usage}%</p>
        <p><b>RAM Usage:</b> {ram_usage}%</p>
        <p><b>System Uptime:</b> {uptime}</p>
        <p><b>Disk Usage:</b><br>{disk_usage_details.replace('n', '<br>')}</p>
        </body>
        </html>
        """
        return body

    def setup_flask_routes(self):
        self.app = Flask(__name__)

        @self.app.route('/')
        def index():
            return render_template('index.html')

        @self.app.route('/status')
        def status():
            status_data = {
                'cpu': psutil.cpu_percent(interval=1),
                'ram': psutil.virtual_memory().percent,
                'disks': {
                    partition.device: psutil.disk_usage(partition.mountpoint).percent
                    for partition in psutil.disk_partitions() if partition.fstype
                },
                'services': {service: self.get_service_status(service) for service in self.selected_services},
                'processes': {
                    proc: "Running" if proc in self.processes else "Not Running"
                    for proc in self.selected_processes
                }
            }
            return jsonify(status_data)

    def run_flask_server(self):
        self.app.run(host=self.server_ip.get(), port=self.server_port.get())


if __name__ == "__main__":
    root = tk.Tk()
    app = ServiceMonitorApp(root)
    root.mainloop()
