# Cyber_Moose_Watch

**Cyber_Moose_Watch** is a Python-based system monitoring application designed to keep track of CPU, RAM, disk usage, and system services. It features a Tkinter-based GUI and allows users to monitor, manage, and receive alerts on critical system statuses.

## Features

- **System Monitoring:** Monitor CPU, RAM, and disk usage in real-time.
- **Service Management:** Scan, add, remove, and monitor system services.
- **Process Management:** Monitor and manage system processes.
- **Automatic Restart:** Attempt to automatically restart failed services.
- **Email Alerts:** Receive email notifications for critical system events.
- **Tray Icon:** Minimal system tray icon with status updates (Windows only).

## Installation

### Prerequisites

- Python 3.11+
- Tkinter (usually included with Python on Windows)
- `psutil` for system resource monitoring
- `Pillow` for image handling (used in the tray icon)


## Usage

1. **Starting the application:**
    - Once the application is running, you will see the main GUI window.
    - The system status (CPU, RAM, and disks) is updated in real-time.

2. **Service Monitoring:**
    - Scan for available services and add them to the monitoring list.
    - The application will attempt to restart failed services automatically (if enabled in the settings).

3. **Process Monitoring:**
    - Manage and monitor system processes similarly to services.

4. **Settings:**
    - Configure email alerts and service restart behavior.
    - Save or load configuration settings from a file.


## Troubleshooting

- **TclError during Tkinter Initialization:**
  - Ensure that your Python installation includes Tkinter. Reinstall Python if necessary.
  
- **Missing Tray Icon Files:**
  - If you encounter file not found errors related to `assets/icon.png`, ensure that the image files are correctly placed in the `assets` directory.

## Contributing

Contributions are welcome! Please fork this repository and submit a pull request with your changes. Ensure that your code follows the existing style and includes relevant tests.


## Contact

For questions or suggestions, please contact:

**Alex Losev**  
Email: [cybermoose1989@gmail.com](mailto:cybermoose1989@gmail.com)  
GitHub: [Cyber-Moose89](https://github.com/Cyber-Moose89)


