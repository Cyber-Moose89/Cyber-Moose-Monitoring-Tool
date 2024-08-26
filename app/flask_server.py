from flask import Flask, jsonify, render_template
import psutil

app = Flask(__name__)

def setup_flask_routes(app_instance):
    @app.route('/')
    def index():
        return render_template('index.html')

    @app.route('/status')
    def status():
        status_data = {
            'cpu': psutil.cpu_percent(interval=1),
            'ram': psutil.virtual_memory().percent,
            'disks': {
                partition.device: psutil.disk_usage(partition.mountpoint).percent
                for partition in psutil.disk_partitions() if partition.fstype
            },
            'services': {service: app_instance.get_service_status(service) for service in app_instance.selected_services},
            'processes': {
                proc: "Running" if proc in app_instance.processes else "Not Running"
                for proc in app_instance.selected_processes
            }
        }
        return jsonify(status_data)

def run_flask_server(app_instance):
    setup_flask_routes(app_instance)
    app.run(host=app_instance.server_ip.get(), port=app_instance.server_port.get())
