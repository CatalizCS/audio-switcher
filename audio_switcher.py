import json
import time
import sounddevice as sd
import keyboard
import pystray
from PIL import Image
import os
from threading import Thread, Timer
import logging
from datetime import datetime
import traceback
import sys
import subprocess
from overlay_notification import OverlayNotification
from enum import Enum
import ctypes
import win32api
import win32con
import win32security
import winreg
import win32com.client
import win32process
from app_mapping_gui import AppMappingGUI, run_mapping_gui_process
import queue
from pythoncom import CoInitialize, CoUninitialize
import tkinter as tk
from update_checker import UpdateChecker
import asyncio
from multiprocessing import Process, Queue, freeze_support
import os.path


class DeviceType(Enum):
    SPEAKER = "Speakers"
    HEADPHONE = "Headphones"


class AudioDeviceListener:
    """Monitors audio device changes"""

    def __init__(self, callback):
        self._callback = callback
        self._running = True
        self._known_devices = set()
        self._check_interval = 2.0  # seconds
        self._timer = None

    def start(self):
        """Start monitoring device changes"""
        self._known_devices = self._get_current_devices()
        self._schedule_check()

    def stop(self):
        """Stop monitoring device changes"""
        self._running = False
        if self._timer:
            self._timer.cancel()

    def _get_current_devices(self):
        """Get current set of device IDs"""
        devices = sd.query_devices()
        return {
            (d["name"], str(d["index"]))
            for d in devices
            if d["max_output_channels"] > 0
        }

    def _check_devices(self):
        """Check for device changes"""
        if not self._running:
            return

        current_devices = self._get_current_devices()

        # Check for new devices
        new_devices = current_devices - self._known_devices
        for name, dev_id in new_devices:
            self._callback("connected", name, dev_id)

        # Check for removed devices
        removed_devices = self._known_devices - current_devices
        for name, dev_id in removed_devices:
            self._callback("disconnected", name, dev_id)

        self._known_devices = current_devices
        self._schedule_check()

    def _schedule_check(self):
        """Schedule next device check"""
        if self._running:
            self._timer = Timer(self._check_interval, self._check_devices)
            self._timer.daemon = True
            self._timer.start()


class ProcessMonitor:
    def __init__(self, callback):
        self._callback = callback
        self._running = True
        self._current_process = None
        self._timer = None
        self._check_interval = 1.0  # seconds

    def start(self):
        self._schedule_check()

    def stop(self):
        self._running = False
        if self._timer:
            self._timer.cancel()

    def _get_foreground_process(self):
        try:
            import win32gui
            import win32process

            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return pid
        except:
            return None

    def _check_process(self):
        if not self._running:
            return

        try:
            current_pid = self._get_foreground_process()
            if current_pid and current_pid != self._current_process:
                self._current_process = current_pid
                self._callback(current_pid)
        except Exception as e:
            logging.error(f"Error checking process: {e}")
        finally:
            self._schedule_check()

    def _schedule_check(self):
        if self._running:
            self._timer = Timer(self._check_interval, self._check_process)
            self._timer.daemon = True
            self._timer.start()


class AudioSwitcher:
    VERSION = "1.0.2"

    def __init__(self):
        # Initialize single root window at the very beginning
        self.root = tk.Tk()
        self.root.withdraw()
        self.root.protocol("WM_DELETE_WINDOW", self._on_root_close)

        # Initialize queues first before anything else
        self.menu_event_queue = queue.Queue()
        self.gui_queue = queue.Queue()
        # Add thread-safe queue for GUI operations
        self.gui_action_queue = queue.Queue()

        # Initialize event loop
        self.loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self.loop)

        # Initialize async event handling
        self.gui_event_queue = asyncio.Queue()

        # Initialize COM in main thread
        CoInitialize()
        self._active = True
        self._error_count = 0
        self.MAX_ERRORS = 3
        self.config_file = "config.json"

        # Load debug mode setting first
        self.debug_mode = self._load_debug_setting()

        # Only setup logging if debug mode is enabled
        if self.debug_mode:
            self.setup_logging()
        else:
            # Suppress all logging when debug mode is disabled
            logging.getLogger().setLevel(logging.CRITICAL)
            logging.disable(logging.CRITICAL)

        # Initialize basic attributes first
        self.devices = {DeviceType.SPEAKER: [], DeviceType.HEADPHONE: []}
        self.current_type = DeviceType.SPEAKER
        self.current_device_index = {DeviceType.SPEAKER: 0, DeviceType.HEADPHONE: 0}
        self.hotkeys = {
            "switch_device": "ctrl+alt+s",
            "switch_type": "ctrl+alt+t",
        }
        self.kernel_mode_enabled = True
        self.force_start = False
        self.startup_enabled = self.is_startup_enabled()

        # Add new attribute for process tracking
        self._svcl_processes = set()

        self.app_device_map = {}
        self.auto_switch_enabled = False
        self.process_monitor = None
        self.mapping_gui = None
        self.gui_queue = queue.Queue()

        # Make sure root processes events
        self.root.update_idletasks()
        self.root.update()

        # Now load config which may override defaults
        self.load_config()

        # Check and request admin privileges
        if not self.is_elevated():
            self.request_elevation()
            return

        # Enable kernel mode if configured
        if self.kernel_mode_enabled:
            if not self.enable_kernel_mode():
                logging.error("Failed to enable kernel mode access")
                if not self.force_start:
                    sys.exit(1)
                logging.warning(
                    "Continuing without kernel mode due to force_start=true"
                )

        try:
            logging.debug("Audio Switcher initialization started")

            # Get application paths
            if getattr(sys, "frozen", False):
                # Running as compiled executable
                base_dir = os.path.dirname(sys.executable)
                self.resources_dir = os.path.join(base_dir, "resources")
                self.logs_dir = os.path.join(base_dir, "logs")
            else:
                # Running as script
                base_dir = os.path.dirname(os.path.abspath(__file__))
                self.resources_dir = os.path.join(base_dir, "resources")
                self.logs_dir = os.path.join(base_dir, "logs")

            # Create necessary directories
            os.makedirs(self.resources_dir, exist_ok=True)
            os.makedirs(self.logs_dir, exist_ok=True)

            # Find svcl.exe
            self.soundvolumeview_path = self._find_resource("svcl.exe")
            if not self.soundvolumeview_path:
                raise FileNotFoundError("svcl.exe not found in any expected location")

            # Find icon.png for tray
            self.icon_path = self._find_resource("icon.png")
            if not self.icon_path:
                raise FileNotFoundError("icon.png not found in any expected location")

            # Initialize notification system
            try:
                self.notifier = OverlayNotification()
                # Wait for notification system to be ready
                self.notifier._setup_done.wait(timeout=5.0)
                if not self.notifier._active:
                    raise RuntimeError("Notification system failed to initialize")
                logging.debug("Overlay notification system initialized")
            except Exception as e:
                logging.warning(f"Failed to initialize overlay notifications: {e}")
                self.notifier = None

            # Initialize basic attributes
            self.config_file = "config.json"
            self.devices = {DeviceType.SPEAKER: [], DeviceType.HEADPHONE: []}
            self.current_type = DeviceType.SPEAKER
            self.current_device_index = {DeviceType.SPEAKER: 0, DeviceType.HEADPHONE: 0}
            self.hotkeys = {
                "switch_device": "ctrl+alt+s",
                "switch_type": "ctrl+alt+t",
            }

            # Ensure working directory is script directory
            os.chdir(os.path.dirname(os.path.abspath(__file__)))

            # Initialize components
            self.load_config()
            self.init_devices()
            self.init_tray()

            # Initialize device listener
            self.device_listener = AudioDeviceListener(self._handle_device_change)
            self.device_listener.start()

            # Initialize process monitor if auto-switch is enabled
            if self.auto_switch_enabled:
                self.start_process_monitor()

            # Initialize Tkinter root window
            self.root.withdraw()

            # Initialize update checker
            self.update_checker = UpdateChecker(self.VERSION)
            self.check_for_updates()

            logging.info("Initialization completed successfully")

        except Exception as e:
            logging.critical(f"Initialization failed: {e}\n{traceback.format_exc()}")
            self.cleanup()
            sys.exit(1)

        # Start async event processing
        self.loop.create_task(self._process_gui_events())

        # Create GUI thread for handling Tkinter operations
        self.gui_thread = Thread(target=self._run_gui_loop, daemon=True)
        self.gui_thread.start()

        # Wait for GUI thread to initialize
        time.sleep(0.1)

        self.gui_process = None
        self.gui_queue = Queue()  # For communication with GUI process

        # Add freeze support for Windows
        if __name__ == "__main__":
            freeze_support()

        self._last_config_modified = (
            os.path.getmtime(self.config_file)
            if os.path.exists(self.config_file)
            else 0
        )

    async def _process_gui_events(self):
        """Process GUI events asynchronously"""
        while self._active:
            try:
                # Process Tkinter events safely
                if self.root and self.root.winfo_exists():
                    self.root.update()

                # Small delay to prevent CPU overuse
                await asyncio.sleep(0.01)

            except tk.TclError as e:
                if "application has been destroyed" not in str(e):
                    logging.error(f"Tkinter error: {e}")
            except Exception as e:
                logging.error(f"Error in event processing: {e}")

    def _run_gui_loop(self):
        """Run Tkinter event loop in a dedicated thread"""
        try:
            while self._active:
                if self.root and self.root.winfo_exists():
                    try:
                        self.root.update()
                        # Check config file every second
                        self._check_config_changes()
                        time.sleep(0.1)  # Prevent high CPU usage
                    except tk.TclError as e:
                        if "application has been destroyed" not in str(e):
                            logging.error(f"Tkinter error in GUI loop: {e}")
                            break
                else:
                    break
        except Exception as e:
            logging.error(f"Error in GUI loop: {e}")

    def _load_debug_setting(self):
        """Load debug mode setting from config"""
        try:
            with open(self.config_file, "r") as f:
                config = json.load(f)
                return config.get("debug_mode", False)
        except FileNotFoundError:
            return False
        except Exception as e:
            print(f"Error loading debug setting: {e}")
            return False

    def is_elevated(self):
        """Check if process has admin privileges"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except Exception as e:
            logging.error(f"Failed to check admin privileges: {e}")
            return False

    def request_elevation(self):
        """Request elevation through UAC"""
        logging.info("Requesting administrative privileges...")
        try:
            if sys.argv[0].endswith(".py"):
                # Running as Python script
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, f'"{sys.argv[0]}"', None, 1
                )
            else:
                # Running as executable
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.argv[0], None, None, 1
                )
            sys.exit(0)
        except Exception as e:
            logging.error(f"Failed to request elevation: {e}")
            sys.exit(1)

    def enable_kernel_mode(self):
        """Enable kernel mode access for audio operations"""
        try:
            # Get required privileges
            privileges = [
                win32security.SE_TCB_NAME,  # Operating system privileges
                win32security.SE_LOAD_DRIVER_NAME,  # Load device drivers
                win32security.SE_SYSTEM_PROFILE_NAME,  # System performance monitoring
            ]

            # Get process token
            token = win32security.OpenProcessToken(
                win32api.GetCurrentProcess(),
                win32con.TOKEN_ADJUST_PRIVILEGES | win32con.TOKEN_QUERY,
            )

            # Enable each privilege
            for privilege in privileges:
                try:
                    # Look up privilege ID
                    privilege_id = win32security.LookupPrivilegeValue(None, privilege)

                    # Enable privilege
                    win32security.AdjustTokenPrivileges(
                        token, False, [(privilege_id, win32con.SE_PRIVILEGE_ENABLED)]
                    )

                    logging.debug(f"Enabled privilege: {privilege}")
                except Exception as e:
                    logging.warning(f"Failed to enable privilege {privilege}: {e}")
                    return False

            return True

        except Exception as e:
            logging.error(f"Failed to enable kernel mode: {e}")
            return False

    def init_devices(self):
        """Initialize audio devices"""
        logging.debug("Checking audio devices...")
        devices = self.get_audio_devices()
        logging.info(f"Found {len(devices)} audio output devices:")
        for device in devices:
            logging.info(f"  - {device['name']} (index: {device['index']})")

    def init_tray(self):
        """Initialize tray icon separately"""
        logging.debug("Initializing tray icon...")
        self.setup_tray()
        # Create and start tray thread
        self.tray_thread = Thread(target=self._run_tray, daemon=True, name="TrayThread")
        self.tray_thread.start()
        # Wait briefly to ensure tray icon appears
        time.sleep(1)
        if not self.tray_thread.is_alive():
            raise RuntimeError("Tray thread failed to start")
        logging.info("Tray icon initialized successfully")

    def setup_logging(self):
        """Setup logging configuration"""
        try:
            # Use the logs directory from instance
            if hasattr(self, "logs_dir"):
                log_dir = self.logs_dir
            else:
                # Fallback to default
                log_dir = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)), "logs"
                )

            os.makedirs(log_dir, exist_ok=True)

            log_filename = os.path.join(
                log_dir,
                f'audio_switcher_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
            )

            logging.basicConfig(
                level=logging.DEBUG,
                format="%(asctime)s.%(msecs)03d - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
                handlers=[logging.FileHandler(log_filename), logging.StreamHandler()],
            )

            logging.info(f"Logging initialized. Log file: {log_filename}")

        except Exception as e:
            print(f"Failed to setup logging: {e}")
            # Set basic logging as fallback
            logging.basicConfig(level=logging.DEBUG)

    def is_admin(self):
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    def restart_as_admin(self):
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        sys.exit()

    def load_config(self):
        try:
            with open(self.config_file, "r") as f:
                config = json.load(f)
                # Load debug mode setting
                self.debug_mode = config.get("debug_mode", False)

                # Load kernel mode settings
                self.kernel_mode_enabled = config.get(
                    "kernel_mode_enabled", self.kernel_mode_enabled
                )
                self.force_start = config.get("force_start", self.force_start)
                self.hotkeys.update(config.get("hotkeys", {}))

                # Convert old config format if needed
                speakers = config.get("speakers", [])
                headphones = config.get("headphones", [])

                # Convert if old format (just indices)
                if speakers and isinstance(speakers[0], int):
                    speakers = [{"index": idx, "id": str(idx)} for idx in speakers]
                if headphones and isinstance(headphones[0], int):
                    headphones = [{"index": idx, "id": str(idx)} for idx in headphones]

                self.devices = {
                    DeviceType.SPEAKER: speakers,
                    DeviceType.HEADPHONE: headphones,
                }
                self.hotkeys = config.get("hotkeys", self.hotkeys)
                self.current_type = DeviceType(
                    config.get("current_type", DeviceType.SPEAKER.value)
                )

                # Load new settings with proper conversion
                self.app_device_map.clear()  # Clear existing mappings
                raw_mappings = config.get("app_device_map", {})
                for app, settings in raw_mappings.items():
                    if isinstance(settings, dict):
                        self.app_device_map[app] = {
                            "type": str(settings.get("type", "Speakers")),
                            "device_id": str(settings.get("device_id", "")),
                            "disabled": bool(settings.get("disabled", False)),
                        }
                    else:
                        # Handle legacy format
                        self.app_device_map[app] = {
                            "type": "Speakers",
                            "device_id": str(settings),
                            "disabled": False,
                        }

                logging.info(f"Loaded {len(self.app_device_map)} application mappings")
                logging.debug(f"Loaded mappings: {self.app_device_map}")

        except FileNotFoundError:
            # Set defaults for new settings
            self.kernel_mode_enabled = True
            self.force_start = False
            self.debug_mode = False
            self.app_device_map = {}
            self.auto_switch_enabled = False
            self.save_config()

    def save_config(self):
        """Save configuration with validation and backup"""
        try:
            # Load current config first
            current_config = {}
            if os.path.exists(self.config_file):
                with open(self.config_file, "r", encoding="utf-8") as f:
                    current_config = json.load(f)

            # Merge our changes with current config
            current_config.update(
                {
                    "speakers": self.devices[DeviceType.SPEAKER],
                    "headphones": self.devices[DeviceType.HEADPHONE],
                    "hotkeys": self.hotkeys,
                    "current_type": self.current_type.value,
                    "kernel_mode_enabled": self.kernel_mode_enabled,
                    "force_start": self.force_start,
                    "debug_mode": self.debug_mode,
                    "auto_switch_enabled": self.auto_switch_enabled,
                    "app_device_map": {
                        app: {
                            "type": str(settings["type"]),
                            "device_id": str(settings["device_id"]),
                            "disabled": bool(settings.get("disabled", False)),
                        }
                        for app, settings in current_config.get(
                            "app_device_map", {}
                        ).items()
                    },
                }
            )

            logging.debug(f"Current config state: {current_config}")

            # Create backup of existing config
            if os.path.exists(self.config_file):
                backup_file = f"{self.config_file}.bak"
                try:
                    os.replace(self.config_file, backup_file)
                except Exception as e:
                    logging.warning(f"Failed to create backup: {e}")

            # Save merged config
            with open(self.config_file, "w", encoding="utf-8") as f:
                json.dump(current_config, f, indent=4, ensure_ascii=False)

            # Verify the save
            with open(self.config_file, "r", encoding="utf-8") as f:
                saved_data = json.load(f)
                saved_mappings = len(saved_data.get("app_device_map", {}))
                logging.info(
                    f"Config saved successfully with {saved_mappings} mappings"
                )
                if os.path.exists(backup_file):
                    os.remove(backup_file)

            # Update last modified time
            self._last_config_modified = os.path.getmtime(self.config_file)
            return True

        except Exception as e:
            logging.error(f"Failed to save config: {e}", exc_info=True)
            return False

    def reload_config(self):
        """Reload configuration from file"""
        try:
            logging.info("Reloading configuration")
            old_map = self.app_device_map.copy()

            # Load fresh config
            with open(self.config_file, "r") as f:
                config = json.load(f)

            # Update app mappings
            self.app_device_map.clear()
            raw_mappings = config.get("app_device_map", {})
            for app, settings in raw_mappings.items():
                if isinstance(settings, dict):
                    self.app_device_map[app] = {
                        "type": str(settings.get("type", "Speakers")),
                        "device_id": str(settings.get("device_id", "")),
                        "disabled": bool(settings.get("disabled", False)),
                    }

            if self.app_device_map != old_map:
                logging.info(
                    f"Updated mappings from config: {len(self.app_device_map)} entries"
                )
                logging.debug(f"New mappings: {self.app_device_map}")
                self._last_config_modified = os.path.getmtime(self.config_file)
                self._refresh_interface()  # Update UI
                return True

            return False

        except Exception as e:
            logging.error(f"Failed to reload config: {e}", exc_info=True)
            return False

    def get_audio_devices(self):
        """Get audio devices using Windows Audio API directly"""
        output_devices = []
        try:
            # Ensure COM is initialized for this thread
            CoInitialize()

            # Get enumerator interface
            from pycaw.pycaw import AudioUtilities, EDataFlow
            from comtypes import CLSCTX_ALL, cast, POINTER
            from pycaw.pycaw import IMMDevice, IAudioEndpointVolume

            # Get all devices directly from MMDeviceEnumerator
            devices = AudioUtilities.GetAllDevices()

            # Filter and process devices
            for device in devices:
                try:

                    name = device.FriendlyName
                    sys_id = device.id

                    # Get index from sounddevice for compatibility
                    sd_devices = sd.query_devices()
                    index = next(
                        (
                            i
                            for i, d in enumerate(sd_devices)
                            if d["name"] == name and d["max_output_channels"] > 0
                        ),
                        0,
                    )

                    device_info = {"name": name, "index": index, "id": sys_id}
                    output_devices.append(device_info)
                    logging.debug(f"Found active output device: {device_info}")

                except Exception as e:
                    logging.warning(
                        f"Error processing device {getattr(device, 'FriendlyName', 'Unknown')}: {e}"
                    )
                    continue

        except Exception as e:
            logging.error(f"Error enumerating audio devices: {e}")
            logging.error(traceback.format_exc())
        finally:
            # Clean up COM
            CoUninitialize()

        if not output_devices:
            # Debug logging
            try:
                devices = AudioUtilities.GetAllDevices()
                logging.warning("All available devices:")
                for d in devices:
                    logging.warning(f"  - Name: {d.FriendlyName}")
                    logging.warning(f"    ID: {d.id}")
                    logging.warning(f"    State: {getattr(d, 'State', 'Unknown')}")
            except Exception as e:
                logging.error(f"Debug logging failed: {e}")

            logging.warning("No audio output devices found!")

        return output_devices

    def switch_device_type(self):
        logging.info(f"Switching device type from {self.current_type}")
        # Switch between speaker and headphone
        if self.current_type == DeviceType.SPEAKER:
            self.current_type = DeviceType.HEADPHONE
        else:
            self.current_type = DeviceType.SPEAKER

        # Switch to first device of the new type
        if self.devices[self.current_type]:
            device = self.devices[self.current_type][0]
            # Ensure both default and output device are changed
            self.set_default_audio_device(device)
            self._refresh_interface()  # Force immediate update
            device_name = sd.query_devices(device["index"])["name"]
            self.show_notification(
                "Switched Type", f"Changed to {self.current_type.value}: {device_name}"
            )
        else:
            self.show_notification(
                "Warning", f"No {self.current_type.value} devices configured"
            )

        self.save_config()
        logging.info(f"Switched to {self.current_type}")

    def switch_audio_device(self):
        if not self._active:
            logging.warning("Switch attempted while inactive")
            return
        self._safe_device_operation(self._switch_audio_device_impl)

    def _switch_audio_device_impl(self):
        current_devices = self.devices[self.current_type]
        if not current_devices:
            return

        # Update current device index for this type
        self.current_device_index[self.current_type] = (
            self.current_device_index[self.current_type] + 1
        ) % len(current_devices)

        device = current_devices[self.current_device_index[self.current_type]]
        device_idx = device["index"]
        self.set_default_audio_device(device)
        self._refresh_interface()  # Force immediate update
        device_name = sd.query_devices(device_idx)["name"]
        self.show_notification(
            f"Switched {self.current_type.value}", f"Now using: {device_name}"
        )

    def update_tray_title(self, device_info):
        """Update tray title with device name"""
        device_idx = (
            device_info["index"] if isinstance(device_info, dict) else device_info
        )
        device_name = sd.query_devices(device_idx)["name"]
        self.icon.title = f"Current {self.current_type.value}: {device_name}"

    def _process_notifications(self):
        """Process GUI events and notifications"""
        if self._active and self.root and self.root.winfo_exists():
            try:
                # Process Tkinter events
                self.root.update()

                # Process GUI actions
                self._process_gui_actions()

            except tk.TclError as e:
                if "application has been destroyed" not in str(e):
                    logging.error(f"Tkinter error: {e}")
            except Exception as e:
                logging.error(f"Error in event processing: {e}")

    def _process_gui_actions(self):
        """Process queued GUI actions"""
        try:
            while True:  # Process all pending actions
                action = self.gui_action_queue.get_nowait()
                if action == "show_mapping":
                    self._create_mapping_gui_safe()
        except queue.Empty:
            pass
        except Exception as e:
            logging.error(f"Error processing GUI action: {e}")

    def _process_menu_events(self):
        """Process queued menu events"""
        try:
            while True:
                try:
                    action = self.menu_event_queue.get_nowait()
                    if action == "show_mapping":
                        if self.root and self.root.winfo_exists():
                            self._create_mapping_gui_safe()
                except queue.Empty:
                    break
        except Exception as e:
            logging.error(f"Error processing menu events: {e}")

    def show_notification(self, title, message):
        """Show both tray and overlay notifications with error handling"""
        try:
            # Show overlay notification first
            if self._active and self.notifier:
                try:
                    # Queue notification and force immediate processing
                    self.notifier.show_notification(title, message, duration=2.5)
                    if hasattr(self.notifier, "root") and self.notifier.root:
                        self.notifier.root.event_generate("<<ProcessQueue>>")
                except Exception as e:
                    logging.warning(f"Overlay notification failed: {e}")

            # Show tray notification
            if self.icon and self._active:
                try:
                    tray_title = (title[:60] + "...") if len(title) > 63 else title
                    tray_message = (
                        (message[:60] + "...") if len(message) > 63 else message
                    )
                    self.icon.notify(tray_title, tray_message)
                except Exception as e:
                    logging.warning(f"Tray notification failed: {e}")

        except Exception as e:
            self._error_count += 1
            logging.error(f"Notification error: {e}", exc_info=True)
            if self._error_count >= self.MAX_ERRORS:
                self.cleanup()

    def set_default_audio_device(self, device_info):
        """Set default audio device using system device ID"""
        try:
            device_name = device_info.get("name", "Unknown Device")
            device_id = device_info.get("id")

            if not device_id or not device_id.startswith("{"):
                logging.warning(
                    f"Invalid device ID format for {device_name}, refreshing device info"
                )
                # Get fresh device info
                devices = self.get_audio_devices()
                matching_device = next(
                    (d for d in devices if d["name"] == device_name), None
                )
                if matching_device:
                    device_id = matching_device["id"]
                else:
                    raise ValueError(f"Could not find system ID for {device_name}")

            logging.info(f"Setting default device: {device_name} (ID: {device_id})")

            # Set up process info
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = win32con.SW_HIDE

            # Set as default for all roles in one command
            cmd_playback = [self.soundvolumeview_path, "/SetDefault", device_id, "all"]
            result = subprocess.run(
                cmd_playback,
                check=True,
                capture_output=True,
                text=True,
                startupinfo=startupinfo,
                creationflags=win32process.CREATE_NO_WINDOW,
            )

            # Quick verification using Windows API
            try:
                from pycaw.pycaw import AudioUtilities

                devices = AudioUtilities.GetAllDevices()
                default_device = next((d for d in devices if d.id == device_id), None)
                if not default_device:
                    logging.warning("Device not found in default devices after setting")
                    return False
            except Exception as e:
                logging.debug(f"Verification warning: {e}")

            logging.info(f"Successfully set {device_name} as default device")
            return True

        except subprocess.CalledProcessError as e:
            logging.error(f"SoundVolumeView failed: {e}")
            return False
        except Exception as e:
            logging.error(f"Error setting default device: {e}")
            logging.error(traceback.format_exc())
            return False

    def setup_tray(self):
        try:
            logging.debug("Setting up system tray icon...")

            # Get icon path from resources
            if getattr(sys, "frozen", False):
                icon_path = os.path.join(
                    os.path.dirname(sys.executable), "resources", "icon.png"
                )
            else:
                icon_path = os.path.join(
                    os.path.dirname(__file__), "resources", "icon.png"
                )

            if not os.path.exists(icon_path):
                logging.error("icon.png not found in: " + os.path.abspath(icon_path))
                raise FileNotFoundError("icon.png is missing")

            try:
                image = Image.open(icon_path)
                # Ensure icon is proper size
                if image.size != (32, 32):
                    image = image.resize((32, 32), Image.Resampling.LANCZOS)
                logging.debug(
                    f"Icon loaded and sized: {image.size[0]}x{image.size[1]}px"
                )
            except Exception as e:
                logging.error(f"Failed to load icon: {e}\n{traceback.format_exc()}")
                raise

            # Create menu with enhanced visuals
            menu = self.create_menu()

            # Create icon with tooltip showing current device
            current_device = "No device selected"
            if self.devices[self.current_type]:
                first_device = self.devices[self.current_type][0]
                device_idx = first_device["index"]
                current_device = sd.query_devices(device_idx)["name"]

            self.icon = pystray.Icon(
                "audio_switcher",
                image,
                f"Audio Switcher\n{self.current_type.value}: {current_device}",
                menu,
            )

            logging.debug("Initializing system tray icon...")
            self.icon = pystray.Icon("audio_switcher", image, "Audio Switcher", menu)
            logging.debug("System tray icon initialized")

            logging.debug("Setting up hotkeys...")
            keyboard.add_hotkey(self.hotkeys["switch_device"], self.switch_audio_device)
            keyboard.add_hotkey(self.hotkeys["switch_type"], self.switch_device_type)
            logging.debug("Hotkeys registered")

            self.tray_thread = Thread(target=self._run_tray, daemon=True)
            self.tray_thread.start()
            logging.info("Tray icon thread started")

        except Exception as e:
            logging.critical(f"Tray setup failed: {e}\n{traceback.format_exc()}")
            raise

    def _run_tray(self):
        """Run tray icon in a separate thread with COM initialization"""
        try:
            # Initialize COM in tray thread
            CoInitialize()
            logging.debug("Starting tray icon loop in thread")
            self.icon.run()
            logging.debug("Tray icon loop ended normally")
        except Exception as e:
            logging.error(f"Tray icon thread error: {e}\n{traceback.format_exc()}")
            self._error_count += 1
            if self._error_count >= self.MAX_ERRORS:
                self.cleanup()
        finally:
            # Clean up COM
            CoUninitialize()

    def create_fallback_menu(self):
        """Create a minimal fallback menu"""
        return pystray.Menu(
            pystray.MenuItem(text="❌ Exit", action=lambda: self.cleanup_and_exit())
        )

    def handle_device_click(self, device, device_type):
        """Handle device menu item click"""
        try:
            if isinstance(device, pystray.MenuItem):
                return

            logging.info(f"Device clicked: {device['name']} (Type: {device_type})")

            was_toggled = self.toggle_device(device, device_type)

            if was_toggled and device_type == self.current_type:
                logging.info(f"Setting {device['name']} as default")
                self.set_default_audio_device(device)

            # Force immediate menu update
            self._refresh_interface()

        except Exception as e:
            logging.error(f"Error handling device click: {e}", exc_info=True)

    def _refresh_interface(self):
        """Force refresh of all UI elements"""
        try:
            # Update menu structure
            menu = self.create_menu()
            self.icon.menu = menu

            # Force menu rebuild
            self.icon.update_menu()

            # Update current device in tray title
            if self.devices[self.current_type]:
                current_device = self.devices[self.current_type][
                    self.current_device_index[self.current_type]
                ]
                self.update_tray_title(current_device)

            # Force system tray refresh
            self.icon.remove_notification()  # Clear any existing notifications
            self.icon.visible = True  # Ensure icon is visible

        except Exception as e:
            logging.error(f"Error refreshing interface: {e}")

    def create_menu(self):
        try:

            def make_group_menu(devices, device_type, group_name):
                items = []
                # Create a dictionary to track name occurrences
                name_counter = {}

                for device in sorted(devices, key=lambda x: x["name"]):
                    device_id = device.get("id", str(device["index"]))
                    is_active = device_id in [
                        d.get("id", str(d["index"])) for d in self.devices[device_type]
                    ]
                    is_current = (
                        device_type == self.current_type
                        and self.devices[device_type]
                        and device_id
                        == self.devices[device_type][
                            self.current_device_index[device_type]
                        ].get("id")
                    )

                    name = device["name"].replace(group_name, "").strip()

                    if name in name_counter:
                        name_counter[name] += 1
                        name = f"{name} ({name_counter[name]})"
                    else:
                        name_counter[name] = 1

                    if is_current:
                        name = f"▶ {name}"
                    elif is_active:
                        name = f"✓ {name}"

                    def make_callback(dev=device, typ=device_type):
                        def callback(icon, item):
                            self.handle_device_click(dev, typ)

                        return callback

                    items.append(
                        pystray.MenuItem(
                            text=name,
                            action=make_callback(),
                            checked=lambda item, _is_active=is_active: _is_active,
                        )
                    )
                return items

            def make_device_menu(device_type):
                devices = self.get_audio_devices()
                if not devices:
                    return [
                        pystray.MenuItem(
                            text="No devices available", action=None, enabled=False
                        )
                    ]

                groups = {}
                for device in devices:
                    group = device["name"].split(" ", 1)[0]
                    if group not in groups:
                        groups[group] = []
                    groups[group].append(device)

                menu_items = []
                for group_name in sorted(groups.keys()):
                    group_devices = groups[group_name]
                    menu_items.append(
                        pystray.MenuItem(
                            text=group_name,
                            action=pystray.Menu(
                                *make_group_menu(group_devices, device_type, group_name)
                            ),
                        )
                    )
                return menu_items

            menu_items = [
                pystray.MenuItem(
                    text=f"● {self.current_type.value}", action=None, enabled=False
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(
                    text="🔊 Speakers",
                    action=pystray.Menu(*make_device_menu(DeviceType.SPEAKER)),
                ),
                pystray.MenuItem(
                    text="🎧 Headphones",
                    action=pystray.Menu(*make_device_menu(DeviceType.HEADPHONE)),
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(
                    text="⌨️ Controls",
                    action=pystray.Menu(
                        pystray.MenuItem(
                            text=f"Switch Device ({self.hotkeys['switch_device']})",
                            action=lambda _: self.switch_audio_device(),
                        ),
                        pystray.MenuItem(
                            text=f"Switch Type ({self.hotkeys['switch_type']})",
                            action=lambda _: self.switch_device_type(),
                        ),
                        pystray.Menu.SEPARATOR,
                        pystray.MenuItem(
                            text="🔒 Kernel Mode",
                            action=lambda _: self.toggle_kernel_mode(),
                            checked=lambda _: self.kernel_mode_enabled,
                        ),
                        pystray.MenuItem(
                            text="🚀 Start with Windows",
                            action=lambda _: self.toggle_startup(),
                            checked=lambda _: self.startup_enabled,
                        ),
                        pystray.MenuItem(
                            text="🔧 Debug Mode",
                            action=lambda _: self.toggle_debug_mode(),
                            checked=lambda _: self.debug_mode,
                        ),
                        pystray.Menu.SEPARATOR,
                        pystray.MenuItem(
                            text="🔄 Auto-Switch",
                            action=lambda _: self.toggle_auto_switch(),
                            checked=lambda _: self.auto_switch_enabled,
                        ),
                        pystray.MenuItem(
                            text="⚙️ Configure App Mappings",
                            action=lambda _: self.show_mapping_gui(),
                        ),
                    ),
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(
                    text="❌ Exit", action=lambda _: self.cleanup_and_exit()
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(
                    text=f"⏳ Version {self.VERSION}",
                    action=lambda _: self.update_checker.open_download_page(),
                ),
                pystray.MenuItem(
                    text="🔄 Check for Updates",
                    action=lambda _: self.check_for_updates(),
                ),
                pystray.MenuItem(text="ℹ️ Made by Tamaisme", action=None, enabled=False),
            ]

            return pystray.Menu(*menu_items)

        except Exception as e:
            logging.error(f"Error creating menu: {e}", exc_info=True)
            return self.create_fallback_menu()

    def toggle_device(self, device_info, device_type):
        """Toggle device in configuration with menu update"""
        try:
            device_id = device_info.get("id", str(device_info["index"]))
            existing = next(
                (d for d in self.devices[device_type] if d.get("id") == device_id), None
            )

            if existing:
                # Don't allow removing the last device of current type
                if (
                    device_type == self.current_type
                    and len(self.devices[device_type]) <= 1
                ):
                    self.show_notification(
                        "Warning", f"Cannot remove last {device_type.value} device"
                    )
                    return False

                self.devices[device_type].remove(existing)
                action = "removed from"

                # If this was the current device, switch to another one
                if device_type == self.current_type and self.current_device_index[
                    device_type
                ] >= len(self.devices[device_type]):
                    self.current_device_index[device_type] = 0
                    if self.devices[device_type]:
                        self.set_default_audio_device(self.devices[device_type][0])
            else:
                new_device = {
                    "index": device_info["index"],
                    "id": device_id,
                    "name": device_info["name"],
                }
                self.devices[device_type].append(new_device)
                action = "added to"

            self.save_config()
            self.show_notification(
                "Device Configuration",
                f"{device_info['name']} {action} {device_type.value}",
            )

            return True  # Indicates successful toggle

        except Exception as e:
            logging.error(f"Error toggling device: {e}", exc_info=True)
            self.show_notification("Error", f"Failed to update device: {e}")
            return False

    def cleanup_and_exit(self):
        """Clean exit handler for menu"""
        logging.info("Exit requested from menu")
        self.cleanup()
        self.icon.stop()

    def cleanup(self):
        if not self._active:
            return

        logging.info("Starting cleanup process")
        self._active = False

        try:
            # Kill any remaining svcl processes
            try:
                # Force kill any remaining svcl processes
                subprocess.run(
                    ["taskkill", "/F", "/IM", "svcl.exe"],
                    startupinfo=subprocess.STARTUPINFO(),
                    creationflags=win32process.CREATE_NO_WINDOW,
                    stderr=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                )
            except:
                pass

            # Clean up notifications first to ensure proper shutdown
            if hasattr(self, "notifier"):
                self.notifier.destroy()
                self.notifier = None

            if hasattr(self, "notification_thread"):
                self.notification_thread.join(timeout=1.0)

            # Stop device monitoring
            if hasattr(self, "device_listener"):
                self.device_listener.stop()

            # Stop process monitor
            if self.process_monitor:
                self.process_monitor.stop()
                self.process_monitor = None

            # Unhook keyboard
            keyboard.unhook_all()

            # Stop tray icon last
            if hasattr(self, "icon"):
                self.icon.stop()

            # Destroy root window
            if hasattr(self, "root"):
                self.root.quit()
                self.root.destroy()

            # Wait for GUI thread to finish
            if hasattr(self, "gui_thread") and self.gui_thread.is_alive():
                self.gui_thread.join(timeout=1.0)

            # Terminate GUI process if running
            if self.gui_process and self.gui_process.is_alive():
                self.gui_process.terminate()
                self.gui_process.join(timeout=1.0)

            # Clean up COM at the end
            CoUninitialize()

            logging.info("Cleanup completed successfully")
        except Exception as e:
            logging.error(f"Error during cleanup: {e}", exc_info=True)

        if self._error_count >= self.MAX_ERRORS:
            logging.critical("Maximum errors reached, forcing exit")
            os._exit(1)

    def _safe_device_operation(self, operation):
        """Wrapper for safe device operations"""
        try:
            return operation()
        except Exception as e:
            self._error_count += 1
            self.show_notification("Error", str(e))
            if self._error_count >= self.MAX_ERRORS:

                self.cleanup()
            return None

    def _handle_device_change(self, event_type, device_name, device_id):
        """Handle device connection/disconnection events"""
        if not self._active:
            return

        if event_type == "connected":
            message = f"🔌 Audio device connected: {device_name}"
            logging.info(f"Device connected: {device_name} (ID: {device_id})")
        else:
            message = f"❌ Audio device disconnected: {device_name}"
            logging.info(f"Device disconnected: {device_name} (ID: {device_id})")

            # Remove disconnected device from configurations
            self._remove_disconnected_device(device_id)

        self.show_notification("Device Change", message)
        self.icon.update_menu()

    def _remove_disconnected_device(self, device_id):
        """Remove disconnected device from configurations"""
        for device_type in DeviceType:
            self.devices[device_type] = [
                d
                for d in self.devices[device_type]
                if d.get("id", str(d["index"])) != device_id
            ]
        self.save_config()

    def toggle_kernel_mode(self):
        """Toggle kernel mode setting"""
        try:
            self.kernel_mode_enabled = not self.kernel_mode_enabled
            if self.kernel_mode_enabled:
                if self.enable_kernel_mode():
                    message = "Kernel mode enabled"
                else:
                    message = "Failed to enable kernel mode"
                    self.kernel_mode_enabled = False
            else:
                message = "Kernel mode disabled"

            self.save_config()
            self.show_notification("Kernel Mode", message)
            self._refresh_interface()
            self.icon.update_menu()

        except Exception as e:
            logging.error(f"Error toggling kernel mode: {e}")
            self.show_notification("Error", "Failed to toggle kernel mode")

    def setup_startup(self):
        """Setup application to run at startup with admin privileges"""
        try:
            # Get the path to the current executable or script
            if getattr(sys, "frozen", False):
                app_path = sys.executable
            else:
                app_path = os.path.abspath(sys.argv[0])

            # Create shortcut path in Startup folder
            startup_folder = os.path.join(
                os.getenv("APPDATA"),
                "Microsoft\\Windows\\Start Menu\\Programs\\Startup",
            )
            shortcut_path = os.path.join(startup_folder, "AudioSwitcher.lnk")

            # Create shortcut with admin privileges
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.TargetPath = app_path
            shortcut.WorkingDirectory = os.path.dirname(app_path)
            shortcut.Description = "Audio Switcher"
            shortcut.IconLocation = os.path.join(os.path.dirname(app_path), "icon.png")
            # Set to run as admin
            shortcut.Save()

            # Set registry key for admin privileges
            with winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers",
                0,
                winreg.KEY_SET_VALUE,
            ) as key:
                winreg.SetValueEx(key, shortcut_path, 0, winreg.REG_SZ, "RUNASADMIN")

            logging.info("Startup shortcut created successfully")
            return True
        except Exception as e:
            logging.error(f"Failed to setup startup: {e}")
            return False

    def remove_startup(self):
        """Remove application from startup"""
        try:
            startup_folder = os.path.join(
                os.getenv("APPDATA"),
                "Microsoft\\Windows\\Start Menu\\Programs\\Startup",
            )
            shortcut_path = os.path.join(startup_folder, "AudioSwitcher.lnk")

            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)

            # Remove registry key
            try:
                with winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers",
                    0,
                    winreg.KEY_SET_VALUE,
                ) as key:
                    winreg.DeleteValue(key, shortcut_path)
            except WindowsError:
                pass  # Key might not exist

            logging.info("Startup shortcut removed successfully")
            return True
        except Exception as e:
            logging.error(f"Failed to remove startup: {e}")
            return False

    def is_startup_enabled(self):
        """Check if application is set to run at startup"""
        startup_folder = os.path.join(
            os.getenv("APPDATA"), "Microsoft\\Windows\\Start Menu\\Programs\\Startup"
        )
        shortcut_path = os.path.join(startup_folder, "AudioSwitcher.lnk")
        return os.path.exists(shortcut_path)

    def toggle_startup(self):
        """Toggle startup status"""
        try:
            if self.is_startup_enabled():
                success = self.remove_startup()
                message = "Startup disabled" if success else "Failed to disable startup"
            else:
                success = self.setup_startup()
                message = "Startup enabled" if success else "Failed to enable startup"

            self.startup_enabled = self.is_startup_enabled()
            self.show_notification("Startup Settings", message)
            self._refresh_interface()
            self.icon.update_menu()
        except Exception as e:
            logging.error(f"Error toggling startup: {e}")
            self.show_notification("Error", "Failed to toggle startup setting")

    def toggle_debug_mode(self):
        """Toggle debug mode"""
        try:
            self.debug_mode = not self.debug_mode
            if self.debug_mode:
                self.setup_logging()
                message = "Debug mode enabled"
            else:
                logging.getLogger().setLevel(logging.CRITICAL)
                logging.disable(logging.CRITICAL)
                message = "Debug mode disabled"

            self.save_config()
            self.show_notification("Debug Mode", message)
            self._refresh_interface()
            self.icon.update_menu()
        except Exception as e:
            print(f"Error toggling debug mode: {e}")

    def _find_resource(self, filename):
        """Find a resource file in various locations"""
        search_paths = [
            os.path.join(self.resources_dir, filename),
            os.path.join(os.path.dirname(sys.executable), "resources", filename),
            os.path.join(
                os.path.dirname(os.path.abspath(__file__)), "resources", filename
            ),
            os.path.join(os.getcwd(), "resources", filename),
            os.path.join(os.getcwd(), filename),
            filename,
        ]

        for path in search_paths:
            if path:
                logging.info(f"Found {filename} at: {path}")
                return path

        paths_str = "\n  ".join(search_paths)
        logging.error(f"Could not find {filename} in any of:\n  {paths_str}")
        return None

    def start_process_monitor(self):
        if not self.process_monitor:
            self.process_monitor = ProcessMonitor(self._handle_process_change)
            self.process_monitor.start()

    def stop_process_monitor(self):
        if self.process_monitor:
            self.process_monitor.stop()
            self.process_monitor = None

    def _handle_process_change(self, pid):
        """Handle foreground process changes with window title matching"""
        try:
            # Check if config has changed
            if self._check_config_changes():
                logging.info("Config changed, reloading settings")
                self._force_reload_config()
                
            import psutil
            import win32gui
            import win32process

            # Get process name
            process = psutil.Process(pid)
            process_name = process.name().lower()
            process_base_name = os.path.splitext(process_name)[0]  # Remove extension

            # Get window title
            def get_window_title(hwnd):
                if win32process.GetWindowThreadProcessId(hwnd)[1] == pid:
                    return win32gui.GetWindowText(hwnd).lower()
                return None

            window_title = None
            win32gui.EnumWindows(
                lambda hwnd, _: (
                    setattr(process, "_window_title", get_window_title(hwnd))
                    if get_window_title(hwnd)
                    else None
                ),
                None,
            )
            window_title = getattr(process, "_window_title", "").lower()

            logging.debug(
                f"Active window - Process: {process_name} ({process_base_name}), Title: {window_title}"
            )

            for app_pattern, device_config in self.app_device_map.items():
                if device_config.get("disabled", False):
                    continue

                pattern = app_pattern.lower()
                pattern_parts = pattern.split()

                def all_parts_match(target):
                    return all(part in target for part in pattern_parts)

                is_match = (
                    pattern == process_name
                    or pattern == process_base_name
                    or (len(pattern) > 3 and pattern in process_base_name)
                    or (
                        window_title
                        and (
                            pattern == window_title
                            or (
                                len(pattern_parts) > 1 and all_parts_match(window_title)
                            )
                            or (len(pattern) > 3 and pattern in window_title)
                        )
                    )
                )

                if is_match:
                    device_type = DeviceType(device_config["type"])
                    device_id = device_config["device_id"]

                    device = next(
                        (
                            d
                            for d in self.devices[device_type]
                            if d.get("id") == device_id
                        ),
                        None,
                    )

                    if device:
                        # Switch to this device
                        self.current_type = device_type
                        self.set_default_audio_device(device)
                        match_type = (
                            "process name"
                            if pattern in process_name
                            else "window title"
                        )
                        self.show_notification(
                            "Auto-Switched Device",
                            f"Switched to {device.get('name')} for {match_type}: {app_pattern}",
                        )
                        self._refresh_interface()
                        break

        except Exception as e:
            logging.error(f"Error handling process change: {e}", exc_info=True)

    def toggle_auto_switch(self):
        """Toggle automatic device switching"""
        try:
            self.auto_switch_enabled = not self.auto_switch_enabled

            if self.auto_switch_enabled:
                self.start_process_monitor()
                message = "Automatic switching enabled"
            else:
                self.stop_process_monitor()
                message = "Automatic switching disabled"

            self.save_config()
            self.show_notification("Auto-Switch", message)
            self._refresh_interface()

        except Exception as e:
            logging.error(f"Error toggling auto-switch: {e}")
            self.show_notification("Error", "Failed to toggle automatic switching")

    def get_icon_path(self):
        """Get path to icon file"""
        if hasattr(self, "icon_path") and self.icon_path:
            # Return ICO version if it exists
            ico_path = self.icon_path.replace(".png", ".ico")
            if os.path.exists(ico_path):
                return ico_path
            return self.icon_path
        return None

    def show_mapping_gui(self):
        """Launch mapping GUI in a separate process with data passing"""
        try:
            if not self._active:
                return

            # Kill any existing GUI process
            if self.gui_process and self.gui_process.is_alive():
                self.gui_process.terminate()
                self.gui_process.join()

            # Create fresh queue and prepare data
            self.gui_queue = Queue()

            # Add icon path to GUI data
            icon_path = self.get_icon_path()
            
            gui_data = {
                "devices": {
                    "Speakers": [d.copy() for d in self.devices[DeviceType.SPEAKER]],
                    "Headphones": [d.copy() for d in self.devices[DeviceType.HEADPHONE]],
                },
                "app_device_map": {
                    app: {
                        "type": str(config.get("type", "Speakers")),
                        "device_id": str(config.get("device_id", "")),
                        "disabled": bool(config.get("disabled", False)),
                    }
                    for app, config in self.app_device_map.items()
                },
                "device_types": {"SPEAKER": "Speakers", "HEADPHONE": "Headphones"},
                "icon_path": icon_path  # Add icon path to data
            }

            # Launch GUI process
            self.gui_process = Process(
                target=run_mapping_gui_process, args=(gui_data, self.gui_queue)
            )
            self.gui_process.daemon = True
            self.gui_process.start()

            # Start monitoring the queue in main thread
            self.root.after(100, self._check_gui_queue)

        except Exception as e:
            logging.error(f"Error launching GUI process: {e}", exc_info=True)

    def _check_gui_queue(self):
        """Check GUI queue in main thread"""
        if not self._active:
            return

        try:
            while True:
                try:
                    action, data = self.gui_queue.get_nowait()
                    logging.debug(f"Received GUI message: {action} with data: {data}")

                    if action == "update_mapping" and isinstance(data, dict):
                        try:
                            if self._validate_mapping_data(data):
                                # Update mappings
                                self.app_device_map = {}
                                for app, config in data.items():
                                    self.app_device_map[app] = {
                                        "type": str(config["type"]),
                                        "device_id": str(config["device_id"]),
                                        "disabled": bool(config.get("disabled", False)),
                                    }

                                # Save and reload config
                                if self.save_config():
                                    if self.reload_config():
                                        logging.info(
                                            "Configuration updated and reloaded"
                                        )
                                    else:
                                        logging.error("Failed to reload configuration")
                                else:
                                    raise Exception("Failed to save configuration")

                            else:
                                raise ValueError("Invalid mapping data received")

                        except Exception as e:
                            logging.error(
                                f"Error updating mappings: {e}", exc_info=True
                            )
                            self.show_notification("Error", "Failed to update mappings")

                    elif action == "force_reload":
                        self._force_reload_config()

                    elif action == "force_save":
                        if self.save_config() and self.reload_config():
                            logging.info("Force save and reload successful")
                        else:
                            self.show_notification(
                                "Error", "Failed to save configuration"
                            )

                except queue.Empty:
                    break

            # Schedule next check if GUI is active
            if self.gui_process and self.gui_process.is_alive():
                self.root.after(100, self._check_gui_queue)

        except Exception as e:
            logging.error(f"Error in GUI queue handler: {e}", exc_info=True)

    def _validate_mapping_data(self, data):
        """Validate mapping data from GUI"""
        try:
            if not isinstance(data, dict):
                logging.error("Mapping data is not a dictionary")
                return False

            for app, config in data.items():
                if not isinstance(config, dict):
                    logging.error(f"Invalid config for app {app}")
                    return False

                # Check required fields
                if "type" not in config:
                    logging.error(f"Missing type in config for app {app}")
                    return False
                if "device_id" not in config:
                    logging.error(f"Missing device_id in config for app {app}")
                    return False

                # Validate field types
                if not isinstance(config["type"], str):
                    logging.error(f"Invalid type format for app {app}")
                    return False
                if not isinstance(config["device_id"], (str, int)):
                    logging.error(f"Invalid device_id format for app {app}")
                    return False
                if "disabled" in config and not isinstance(config["disabled"], bool):
                    logging.error(f"Invalid disabled format for app {app}")
                    return False

            logging.debug(f"Validated mapping data: {data}")
            return True

        except Exception as e:
            logging.error(f"Validation error: {e}", exc_info=True)
            return False

    def _create_mapping_gui_safe(self):
        """Create mapping GUI safely in main thread"""
        try:
            if not self.root or not self.root.winfo_exists():
                logging.error("Root window not available")
                return

            if (
                not hasattr(self, "mapping_gui")
                or not self.mapping_gui
                or not hasattr(self.mapping_gui, "window")
                or not self.mapping_gui.window.winfo_exists()
            ):

                logging.debug("Creating new mapping GUI")
                self.mapping_gui = AppMappingGUI(self, DeviceType)
                self.mapping_gui.window.lift()
                self.mapping_gui.window.focus_force()
                logging.debug("Mapping GUI created successfully")
            else:
                logging.debug("Reusing existing mapping GUI")
                self.mapping_gui.window.lift()
                self.mapping_gui.window.focus_force()

        except Exception as e:
            logging.error(f"Error creating mapping GUI: {e}")
            self.show_notification("Error", "Failed to open mapping configuration")

    def check_for_updates(self):
        """Check for available updates"""
        try:
            if self.update_checker.check_for_updates():
                self.show_notification(
                    "Update Available",
                    f"Version {self.update_checker.latest_version} is available",
                )
        except Exception as e:
            logging.error(f"Error checking for updates: {e}")

    def _on_root_close(self):
        """Handle root window close"""
        if self._active:
            self.cleanup()

    def _queue_menu_action(self, action):
        """Queue menu action for later processing"""
        try:
            self.menu_event_queue.put_nowait(action)
        except Exception as e:
            logging.error(f"Error queueing menu action: {e}")

    def _handle_menu_action(self, action):
        """Deprecated - use _process_menu_events instead"""
        pass

    def _check_config_changes(self):
        """Check if config file has been modified externally"""
        try:
            if not os.path.exists(self.config_file):
                return False

            current_mtime = os.path.getmtime(self.config_file)
            if current_mtime > self._last_config_modified:
                logging.info("Config file changed externally, reloading...")
                self._last_config_modified = current_mtime
                return self.reload_config()
            return False
        except Exception as e:
            logging.error(f"Error checking config changes: {e}")
            return False

    def _force_reload_config(self):
        """Force reload configuration and update monitoring"""
        try:
            logging.info("Force reloading configuration")
            
            # Store old state for comparison
            old_map = self.app_device_map.copy()
            old_auto_switch = self.auto_switch_enabled
            
            # Reload configuration
            if self.reload_config():
                # Check if auto-switch status changed
                if self.auto_switch_enabled != old_auto_switch:
                    if self.auto_switch_enabled:
                        self.start_process_monitor()
                    else:
                        self.stop_process_monitor()
                
                # If app mappings changed and monitoring is active, refresh
                if old_map != self.app_device_map and self.process_monitor:
                    self.stop_process_monitor()
                    self.start_process_monitor()
                
                self.show_notification(
                    "Configuration Reloaded",
                    "Settings updated successfully"
                )
                return True
            
            self.show_notification(
                "Reload Failed",
                "Failed to reload configuration"
            )
            return False
            
        except Exception as e:
            logging.error(f"Error force reloading config: {e}")
            self.show_notification("Error", "Failed to reload configuration")
            return False


if __name__ == "__main__":
    try:
        freeze_support()
        app = AudioSwitcher()
        logging.info("Application started")

        # Main event loop
        while app._active:
            try:
                app.root.update()
                time.sleep(0.01)

                # Check tray thread
                if not app.tray_thread.is_alive():
                    break

            except tk.TclError as e:
                if "application has been destroyed" not in str(e):
                    logging.error(f"Main loop error: {e}")
                break
            except Exception as e:
                logging.error(f"Main loop error: {e}")
                break

    except Exception as e:
        logging.critical(f"Application error: {e}", exc_info=True)
    finally:
        if "app" in locals():
            app.cleanup()
        logging.info("Application shutdown complete")
