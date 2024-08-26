import psutil
import threading
import time

def start_monitoring(app):
    monitoring_thread = threading.Thread(target=monitor_services_and_processes, args=(app,), daemon=True)
    monitoring_thread.start()

def monitor_services_and_processes(app):
    while True:
        app.root.after(0, app.refresh_status)
        check_services(app)
        check_processes(app)
        check_cpu_ram_usage(app)
        check_disk_space(app)
        generate_reports_if_needed(app)
        time.sleep(5)

def check_services(app):
    for service_name in app.selected_services:
        if service_name:
            current_status = "Running" if is_service_running(service_name) else "Stopped"
            if service_name not in app.last_service_status or current_status != app.last_service_status[service_name]:
                app.handle_service_status_change(service_name, current_status)
                app.last_service_status[service_name] = current_status

                if current_status == "Stopped" and app.auto_restart_service.get():
                    app.attempt_service_restart(service_name)

def check_processes(app):
    running_processes = [proc.info['name'] for proc in psutil.process_iter(['name'])]
    for proc_name in app.selected_processes:
        if proc_name:
            current_status = "Running" if proc_name in running_processes else "Not Running"
            if proc_name not in app.last_process_status or current_status != app.last_process_status[proc_name]:
                app.handle_process_status_change(proc_name, current_status)
                app.last_process_status[proc_name] = current_status

def check_cpu_ram_usage(app):
    cpu_usage = psutil.cpu_percent(interval=1)
    ram_usage = psutil.virtual_memory().percent

    if cpu_usage > app.cpu_threshold.get():
        app.handle_hardware_overload("CPU Usage", cpu_usage, app.cpu_threshold.get())
    if ram_usage > app.ram_threshold.get():
        app.handle_hardware_overload("RAM Usage", ram_usage, app.ram_threshold.get())

def check_disk_space(app):
    for partition in psutil.disk_partitions():
        if partition.device in app.disk_thresholds:
            usage = psutil.disk_usage(partition.mountpoint).percent
            if usage > app.disk_thresholds[partition.device]:
                app.handle_hardware_overload(f"Disk Space {partition.device}", usage, app.disk_thresholds[partition.device])


