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


class AudioSwitcher:
    def __init__(self):
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

            # Check if soundvolumeview exists
            self.soundvolumeview_path = os.path.join(
                os.path.dirname(__file__), "svcl.exe"
            )
            if not os.path.exists(self.soundvolumeview_path):
                logging.error(
                    "soundvolumeview.exe not found. Please download it from https://www.nirsoft.net/utils/sound_volume_view.html"
                )
                raise FileNotFoundError("soundvolumeview.exe is required but not found")

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

            logging.info("Initialization completed successfully")

        except Exception as e:
            logging.critical(f"Initialization failed: {e}\n{traceback.format_exc()}")
            self.cleanup()
            sys.exit(1)

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
        # Create logs directory if it doesn't exist
        if not os.path.exists("logs"):
            os.makedirs("logs")

        # Setup logging format and configuration
        log_filename = (
            f'logs/audio_switcher_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
        )
        logging.basicConfig(
            level=logging.DEBUG, 
            format="%(asctime)s.%(msecs)03d - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
            handlers=[logging.FileHandler(log_filename), logging.StreamHandler()],
        )

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
        except FileNotFoundError:
            # Set defaults for new settings
            self.kernel_mode_enabled = True
            self.force_start = False
            self.debug_mode = False
            self.save_config()

    def save_config(self):
        with open(self.config_file, "w") as f:
            json.dump(
                {
                    "speakers": self.devices[DeviceType.SPEAKER],
                    "headphones": self.devices[DeviceType.HEADPHONE],
                    "hotkeys": self.hotkeys,
                    "current_type": self.current_type.value,
                    "kernel_mode_enabled": self.kernel_mode_enabled,
                    "force_start": self.force_start,
                    "debug_mode": self.debug_mode,  # Save debug mode setting
                },
                f,
                indent=4,
            )

    def get_audio_devices(self):
        """Get audio devices using Windows Audio API directly"""
        output_devices = []
        try:
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
            self.update_tray_title(device)
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
        self.update_tray_title(device)
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
        """Process notifications in the main thread"""
        if self._active and hasattr(self, "notifier") and self.notifier:
            try:
                self.notifier.process_events()
            except Exception as e:
                if self._active:  # Only log if not shutting down
                    logging.error(f"Error processing notifications: {e}")
                    self._error_count += 1
                    if self._error_count >= self.MAX_ERRORS:
                        self.cleanup()

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

            # Set as default for all roles
            cmd_playback = [self.soundvolumeview_path, "/SetDefault", device_id, "all"]
            logging.debug(f"Running command: {' '.join(cmd_playback)}")

            result = subprocess.run(
                cmd_playback, check=True, capture_output=True, text=True
            )

            if result.stdout:
                logging.debug(f"Command output: {result.stdout}")
            if result.stderr:
                logging.warning(f"Command stderr: {result.stderr}")

            # Verify the change
            verify_cmd = [self.soundvolumeview_path, "/scomma", "verify.csv"]
            subprocess.run(verify_cmd, check=True, capture_output=True)

            try:
                with open("verify.csv", "r") as f:
                    current = f.read()
                    if device_id not in current:
                        logging.warning(
                            "Device ID not found in current devices after setting"
                        )
            except Exception as e:
                logging.warning(f"Could not verify device change: {e}")
            finally:
                try:
                    os.remove("verify.csv")
                except:
                    pass

            logging.info(f"Successfully set {device_name} as default device")
            return True

        except subprocess.CalledProcessError as e:
            logging.error(f"SoundVolumeView failed: {e}")
            logging.error(f"Command output: {e.stdout if e.stdout else 'No output'}")
            logging.error(
                f"Error output: {e.stderr if e.stderr else 'No error output'}"
            )
            return False
        except Exception as e:
            logging.error(f"Error setting default device: {e}")
            logging.error(f"Exception type: {type(e).__name__}")
            logging.error(traceback.format_exc())
            return False

    def setup_tray(self):
        try:
            logging.debug("Setting up system tray icon...")

            # Enhanced icon setup
            if not os.path.exists("icon.png"):
                logging.error("icon.png not found in: " + os.path.abspath("icon.png"))
                raise FileNotFoundError("icon.png is missing")

            try:
                image = Image.open("icon.png")
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

            logging.debug("Creating tray menu...")
            try:
                menu = self.create_menu()
                logging.debug("Menu structure created")
            except Exception as e:
                logging.error(f"Failed to create menu: {e}\n{traceback.format_exc()}")
                raise

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
            logging.debug("Starting tray icon loop in thread")
            self.icon.run()
            logging.debug("Tray icon loop ended normally")
        except Exception as e:
            logging.error(f"Tray icon thread error: {e}\n{traceback.format_exc()}")
            self._error_count += 1
            if self._error_count >= self.MAX_ERRORS:
                self.cleanup()

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

            self.icon.update_menu()
            if device_type == self.current_type:
                self.update_tray_title(device)

        except Exception as e:
            logging.error(f"Error handling device click: {e}", exc_info=True)

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

                    # Get base name without group prefix
                    name = device["name"].replace(group_name, "").strip()

                    # Count occurrences of this name
                    if name in name_counter:
                        name_counter[name] += 1
                        # Add a number suffix to duplicate names
                        name = f"{name} ({name_counter[name]})"
                    else:
                        name_counter[name] = 1

                    # Add status indicators
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

                # Group devices by manufacturer
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

            # Create main menu structure
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
                    ),
                ),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem(
                    text="❌ Exit", action=lambda _: self.cleanup_and_exit()
                ),
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
            self.show_notification("Error", f"Failed to update device: {str(e)}")
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
            # Clean up notifications first to ensure proper shutdown
            if hasattr(self, "notifier"):
                self.notifier.destroy()
                self.notifier = None

            if hasattr(self, "notification_thread"):
                self.notification_thread.join(timeout=1.0)

            # Stop device monitoring
            if hasattr(self, "device_listener"):
                self.device_listener.stop()

            # Unhook keyboard
            keyboard.unhook_all()

            # Stop tray icon last
            if hasattr(self, "icon"):
                self.icon.stop()

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
            self.icon.update_menu()
        except Exception as e:
            print(f"Error toggling debug mode: {e}")


if __name__ == "__main__":
    try:
        app = AudioSwitcher()
        logging.info("Application started, entering main loop")

        # Main event loop that processes both tray and notifications
        while app._active and app.tray_thread.is_alive():
            try:
                app._process_notifications()
                time.sleep(0.01)
            except KeyboardInterrupt:
                logging.info("Received keyboard interrupt")
                break
            except Exception as e:
                logging.error(f"Main loop error: {e}", exc_info=True)
                break

    except Exception as e:
        logging.critical(f"Application error: {e}", exc_info=True)
    finally:
        if "app" in locals():
            app.cleanup()
        logging.info("Application shutdown complete")
