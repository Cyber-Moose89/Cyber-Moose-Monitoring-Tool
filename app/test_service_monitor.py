import unittest
from unittest.mock import MagicMock, patch
import tkinter as tk
import psutil
from app.gui import ServiceMonitorApp
from PIL import Image


class TestServiceMonitorApp(unittest.TestCase):

    def setUp(self):
        """Set up the test environment."""
        self.root = tk.Tk()

        # Mock the iconphoto method to avoid errors related to missing images
        self.root.iconphoto = MagicMock()

        # Mock the ImageTk.PhotoImage and Image.open to avoid loading the actual image
        with patch('PIL.ImageTk.PhotoImage', return_value=MagicMock()):
            with patch('PIL.Image.open', return_value=MagicMock()):
                self.app = ServiceMonitorApp(self.root)

        # Mock initial variables and methods
        self.app.cpu_label = MagicMock()
        self.app.ram_label = MagicMock()
        self.app.disk_labels = [(MagicMock(), MagicMock()) for _ in range(2)]
        self.app.send_email = MagicMock()
        self.app.get_service_status = MagicMock()

    def test_add_service_to_monitor(self):
        """Test adding a service to the monitor list."""
        self.app.services = ['Service1', 'Service2']
        self.app.lst_scanned_services.insert(tk.END, 'Service1')
        self.app.lst_scanned_services.selection_set(0)
        self.app.add_to_monitor_list()
        monitored_services = self.app.lst_monitored_services.get(0, tk.END)
        self.assertIn('Service1', monitored_services)

    def test_remove_service_from_monitor(self):
        """Test removing a service from the monitor list."""
        self.app.selected_services = ['Service1', 'Service2']
        self.app.lst_monitored_services.insert(tk.END, 'Service1')
        self.app.lst_monitored_services.insert(tk.END, 'Service2')
        self.app.lst_monitored_services.selection_set(0)
        self.app.remove_from_monitor_list()
        monitored_services = self.app.lst_monitored_services.get(0, tk.END)
        self.assertNotIn('Service1', monitored_services)
        self.assertIn('Service2', monitored_services)

    def test_scan_services(self):
        """Test scanning services."""
        psutil.win_service_iter = MagicMock(return_value=[MagicMock(name='Service1'), MagicMock(name='Service2')])
        self.app.scan_services()
        scanned_services = self.app.lst_scanned_services.get(0, tk.END)
        self.assertIn('Service1', scanned_services)
        self.assertIn('Service2', scanned_services)

    def test_cpu_ram_threshold_exceeded(self):
        """Test CPU and RAM threshold exceeded."""
        psutil.cpu_percent = MagicMock(return_value=90)
        psutil.virtual_memory = MagicMock()
        psutil.virtual_memory.return_value.percent = 85
        self.app.cpu_threshold.set(80)
        self.app.ram_threshold.set(80)
        self.app.handle_hardware_overload = MagicMock()
        self.app.check_cpu_ram_usage()
        self.assertEqual(self.app.handle_hardware_overload.call_count, 2)
        self.app.handle_hardware_overload.assert_any_call("CPU Usage", 90, 80)
        self.app.handle_hardware_overload.assert_any_call("RAM Usage", 85, 80)

    def test_attempt_service_restart(self):
        """Test attempting to restart a service."""
        self.app.max_restart_attempts.set(3)
        self.app.send_restart_report = MagicMock()
        self.app.attempt_service_restart('DummyService')
        self.assertTrue(self.app.send_restart_report.called)

    def test_handle_service_status_change(self):
        """Test handling a service status change."""
        self.app.handle_service_status_change('DummyService', 'Stopped')
        self.app.send_email.assert_called_with(
            'Service Stopped',
            "<p>The service <strong>DummyService</strong> is now in a <strong>Stopped</strong> state.</p>",
            service_or_process_name='DummyService',
            status='Stopped',
            previous_status='Unknown'
        )

    def test_save_settings(self):
        """Test saving settings."""
        open_mock = MagicMock()
        self.app.cpu_threshold.set(75)
        self.app.email_subject.set('Test Alert')
        with unittest.mock.patch('builtins.open', open_mock):
            self.app.save_settings()
        open_mock.assert_called_with('config.ini', 'w')
        config_data = open_mock().write.call_args[0][0]
        self.assertIn('CPU_Threshold = 75', config_data)
        self.assertIn('Subject = Test Alert', config_data)

    def test_refresh_status(self):
        """Test refreshing the system status."""
        psutil.cpu_percent = MagicMock(return_value=20)
        psutil.virtual_memory = MagicMock()
        psutil.virtual_memory.return_value.percent = 50
        psutil.disk_partitions = MagicMock(return_value=[
            MagicMock(device='C:', mountpoint='/', fstype='NTFS', opts='rw'),
            MagicMock(device='D:', mountpoint='/data', fstype='NTFS', opts='rw')
        ])
        psutil.disk_usage = MagicMock(return_value=MagicMock(percent=30))
        self.app.refresh_status()
        self.app.cpu_label.config.assert_called_with(text="20%")
        self.app.ram_label.config.assert_called_with(text="50%")
        self.app.disk_labels[0][1].config.assert_called_with(text="30% used")
        self.app.disk_labels[1][1].config.assert_called_with(text="30% used")


if __name__ == '__main__':
    unittest.main()
