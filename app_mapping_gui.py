import sys
import tkinter as tk
from tkinter import messagebox
import logging
import os
import customtkinter as ctk
from datetime import datetime
import json


def setup_logging():
    """Setup logging for GUI process"""
    try:
        # Create logs directory
        log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(log_dir, exist_ok=True)

        # Create log filename with timestamp
        log_file = os.path.join(
            log_dir, f'app_mapping_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log'
        )

        logging.basicConfig(
            level=logging.DEBUG,
            format="%(asctime)s.%(msecs)03d - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
            handlers=[logging.FileHandler(log_file), logging.StreamHandler()],
        )
        logging.info("AppMappingGUI logging initialized")
        return True
    except Exception as e:
        print(f"Failed to setup logging: {e}")
        return False


def run_mapping_gui_process(data, queue):
    """Run GUI in separate process"""
    try:
        setup_logging()
        logging.info(f"Starting GUI process with data: {data}")

        def send_message(msg_type, msg_data=None):
            """Send message to main process with validation"""
            try:
                queue.put_nowait((msg_type, msg_data))
                logging.debug(f"Sent message: {msg_type} with data: {msg_data}")
                return True
            except Exception as e:
                logging.error(f"Failed to send message {msg_type}: {e}")
                return False

        # Set theme and initialize GUI
        ctk.set_appearance_mode("dark")
        root = ctk.CTk()

        # Load and set window icon
        try:
            # Find icon in various locations
            icon_paths = [
                os.path.join(os.path.dirname(sys.executable), "resources", "icon.ico"),
                os.path.join(
                    os.path.dirname(os.path.abspath(__file__)), "resources", "icon.ico"
                ),
                os.path.join(os.getcwd(), "resources", "icon.ico"),
                "icon.ico",
            ]

            icon_path = None
            for path in icon_paths:
                if os.path.exists(path):
                    icon_path = path
                    break

            if icon_path:
                root.iconbitmap(icon_path)
                logging.info(f"Set window icon from: {icon_path}")
            else:
                # Try PNG if ICO not found
                png_paths = [p.replace(".ico", ".png") for p in icon_paths]
                for path in png_paths:
                    if os.path.exists(path):
                        # Convert PNG to ICO in memory and set
                        from PIL import Image

                        img = Image.open(path)
                        icon_path = os.path.join(os.path.dirname(path), "temp_icon.ico")
                        img.save(icon_path, format="ICO")
                        root.iconbitmap(icon_path)
                        os.remove(icon_path)  # Clean up temporary file
                        logging.info(f"Converted and set icon from PNG: {path}")
                        break
        except Exception as e:
            logging.warning(f"Failed to set window icon: {e}")

        gui = AppMappingGUI(root, data, send_message)
        root.title("Audio Mapper")
        root.geometry("500x600")

        # Handle window closing with state save
        def on_closing():
            try:
                logging.info("Saving final state before closing")
                success = gui.save_state()
                if success:
                    logging.info("Final state saved successfully")
                    send_message("force_save", None)  # Force config save
                else:
                    logging.error("Failed to save final state")
            except Exception as e:
                logging.error(f"Error during window closing: {e}")
            finally:
                root.destroy()

        root.protocol("WM_DELETE_WINDOW", on_closing)

        # Center window on screen
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f"+{x}+{y}")

        # Run mainloop
        root.mainloop()

    except Exception as e:
        logging.error(f"Error in GUI process: {e}", exc_info=True)


class AppMappingGUI:
    def __init__(self, root, data, send_message):
        self.root = root
        self.send_message = send_message
        logging.info(f"Initializing GUI with data: {data}")

        # Make deep copies to avoid reference issues
        self.devices = {k: list(v) for k, v in data.get("devices", {}).items()}
        # Convert app_device_map if needed
        raw_map = data.get("app_device_map", {})
        self.app_device_map = {}
        for app, settings in raw_map.items():
            if isinstance(settings, dict):
                self.app_device_map[app] = settings.copy()
            else:
                # Handle legacy format
                self.app_device_map[app] = {
                    "type": "Speakers",
                    "device_id": str(settings),
                    "disabled": False,
                }

        logging.debug(f"Initialized app_device_map: {self.app_device_map}")

        # Add state tracking
        self._last_saved_state = self.app_device_map.copy()
        self._changes_pending = False

        # Update color scheme with proper alpha values
        self.colors = {
            "bg": "#2b2b2b",
            "secondary_bg": "#323232",
            "accent": "#007acc",
            "accent_hover": "#1a8cd8",  # Lighter accent
            "text": "#ffffff",
            "text_secondary": "#cccccc",
            "success": "#4ecca3",
            "warning": "#ffbe0b",
            "error": "#ff006e",
            "error_hover": "#ff3385",  # Lighter error
            "border": "#404040",
            "hover": "#3a3a3a",  # General hover color
            "disabled": "#666666",  # Add disabled state color
            "disabled_bg": "#2a2a2a",  # Add disabled background color
        }

        # Initialize these before creating widgets
        self.type_var = ctk.StringVar(value="Speakers")
        self.device_var = ctk.StringVar()
        self.selected_mapping = None  # Track selected mapping

        self.config_file = "config.json"

        self._create_widgets()
        self._update_device_list()
        self._load_mappings()

        logging.info("AppMappingGUI initialized successfully")

    def _create_widgets(self):
        # Main container with padding
        self.main_container = ctk.CTkFrame(self.root)
        self.main_container.pack(fill="both", expand=True, padx=20, pady=20)

        # Title
        title = ctk.CTkLabel(
            self.main_container,
            text="Application Audio Mapping",
            font=("Segoe UI", 20, "bold"),
        )
        title.pack(pady=(0, 20))

        # Left panel - Mappings list
        list_frame = ctk.CTkFrame(self.main_container)
        list_frame.pack(fill="both", expand=True, pady=(0, 20))

        # Add refresh button next to search
        search_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
        search_frame.pack(fill="x", padx=10, pady=10)

        # Search box with callback binding
        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="🔍 Search mappings...",
            height=35,
            textvariable=self.search_var,
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        # Add refresh button
        refresh_btn = ctk.CTkButton(
            search_frame,
            text="🔄",
            width=35,
            height=35,
            command=self._force_reload_config
        )
        refresh_btn.pack(side="right")

        # Mappings list using CTkScrollableFrame
        self.list_container = ctk.CTkScrollableFrame(list_frame)
        self.list_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Bind click events for mappings
        self.mapping_widgets = []
        self.mapping_frames = {}  # Store frames with their app names as keys

        # Right panel - Add/Edit
        edit_frame = ctk.CTkFrame(self.main_container)
        edit_frame.pack(fill="x")

        # App name input with browse button container
        app_input_frame = ctk.CTkFrame(edit_frame, fg_color="transparent")
        app_input_frame.pack(fill="x", pady=(0, 10))

        self.app_entry = ctk.CTkEntry(
            app_input_frame, placeholder_text="Application name or path", height=35
        )
        self.app_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))

        browse_btn = ctk.CTkButton(
            app_input_frame,
            text="Browse",
            width=70,
            height=35,
            command=self._browse_application,
        )
        browse_btn.pack(side="right")

        # Device type selection
        self.type_combo = ctk.CTkSegmentedButton(
            edit_frame,
            values=["Speakers", "Headphones"],
            variable=self.type_var,
            command=self._on_type_change,  # Will handle any number of arguments
            height=35,
        )
        self.type_combo.pack(fill="x", pady=(0, 10))

        # Device selection
        self.device_combo = ctk.CTkOptionMenu(
            edit_frame,
            values=["Select device"],
            variable=self.device_var,
            height=35,
            dynamic_resizing=False,
        )
        self.device_combo.pack(fill="x", pady=(0, 20))

        # Buttons
        btn_frame = ctk.CTkFrame(edit_frame, fg_color="transparent")
        btn_frame.pack(fill="x")

        # Save button with updated styling
        save_btn = ctk.CTkButton(
            btn_frame,
            text="Save",
            command=self._add_mapping,
            height=35,
            fg_color=self.colors["accent"],
            hover_color=self.colors["accent_hover"],
        )
        save_btn.pack(side="left", expand=True, padx=(0, 5))

        # Delete button with updated styling
        delete_btn = ctk.CTkButton(
            btn_frame,
            text="Delete",
            command=self._delete_mapping,
            height=35,
            fg_color="transparent",
            hover_color=self.colors["error_hover"],
            border_color=self.colors["error"],
            border_width=2,
            text_color=self.colors["error"],
        )
        delete_btn.pack(side="left", expand=True, padx=(5, 0))

    def _create_mapping_widget(
        self, parent, app, device_type, device_name, search_text=""
    ):
        """Create a custom mapping list item with optional text highlighting"""
        frame = ctk.CTkFrame(
            self.list_container, fg_color="transparent", corner_radius=6
        )
        frame.pack(fill="x", pady=2, padx=5)

        # Left container for enable/disable checkbox and app name
        left_container = ctk.CTkFrame(frame, fg_color="transparent")
        left_container.pack(side="left", fill="x", expand=True)

        # Enable/Disable checkbox
        is_enabled = not self.app_device_map[app].get("disabled", False)
        checkbox = ctk.CTkCheckBox(
            left_container,
            text="",
            command=lambda: self._toggle_mapping_state(app, checkbox),
            width=20,
            height=20,
        )
        checkbox.pack(side="left", padx=(5, 10))
        checkbox.select() if is_enabled else checkbox.deselect()

        # Update app label with highlighted text if searching
        app_text = self._highlight_text(app, search_text)
        app_label = ctk.CTkLabel(
            left_container,
            text=app_text,
            font=("Segoe UI", 12),
            anchor="w",
            text_color=self.colors["text"] if is_enabled else self.colors["disabled"],
        )
        app_label.pack(side="left", fill="x", expand=True)

        # Update device info with highlighted text if searching
        info = f"{device_type} • {device_name}"
        if search_text:
            info = self._highlight_text(info, search_text)

        info_label = ctk.CTkLabel(
            frame,
            text=info,
            font=("Segoe UI", 11),
            text_color=(
                self.colors["text_secondary"] if is_enabled else self.colors["disabled"]
            ),
            anchor="e",
        )
        info_label.pack(side="right", padx=5)

        # Store frame and labels for updates
        self.mapping_frames[app] = {
            "frame": frame,
            "checkbox": checkbox,
            "app_label": app_label,
            "info_label": info_label,
        }

        # Bind click events
        for widget in (frame, app_label, info_label):
            widget.bind("<Button-1>", lambda e, a=app: self._on_mapping_click(a))

        # Show filepath if it exists
        filepath = self.app_device_map[app].get("filepath", "")
        if filepath:
            tooltip_text = f"Path: {filepath}"

            # Create tooltip label that shows on hover
            tooltip = ctk.CTkLabel(
                frame,
                text="📁",  # Folder icon
                font=("Segoe UI", 11),
                text_color=self.colors["text_secondary"],
                cursor="hand2",
            )
            tooltip.pack(side="right", padx=(0, 5))

            # Bind hover events for tooltip
            def show_tooltip(event):
                tooltip.configure(text=tooltip_text)

            def hide_tooltip(event):
                tooltip.configure(text="📁")

            tooltip.bind("<Enter>", show_tooltip)
            tooltip.bind("<Leave>", hide_tooltip)

        return frame

    def _toggle_mapping_state(self, app_name, checkbox):
        """Toggle mapping state with improved feedback"""
        try:
            is_enabled = checkbox.get()
            logging.info(
                f"Toggling {app_name} to {'enabled' if is_enabled else 'disabled'}"
            )

            # Store old state for rollback
            old_state = self.app_device_map[app_name].copy()

            # Update state
            self.app_device_map[app_name]["disabled"] = not is_enabled
            self._changes_pending = True

            # Try to save
            if not self.save_state():
                # Rollback on failure
                logging.warning("Failed to save, rolling back changes")
                self.app_device_map[app_name] = old_state
                checkbox.toggle()  # Revert checkbox
                self.show_error("Failed to update state")
                return

            # Update UI on success
            self._update_mapping_visuals(app_name, is_enabled)
            logging.info(f"Successfully toggled {app_name}")

        except Exception as e:
            logging.error(f"Error toggling state: {e}", exc_info=True)
            self.show_error("Failed to update state")

    def _update_mapping_visuals(self, app_name, is_enabled):
        """Update visual elements of a mapping"""
        try:
            widgets = self.mapping_frames[app_name]
            text_color = self.colors["text"] if is_enabled else self.colors["disabled"]
            text_secondary = (
                self.colors["text_secondary"] if is_enabled else self.colors["disabled"]
            )

            widgets["app_label"].configure(text_color=text_color)
            widgets["info_label"].configure(text_color=text_secondary)
        except Exception as e:
            logging.error(f"Error updating visuals: {e}")

    def _on_mapping_click(self, app_name):
        """Handle mapping selection"""
        try:
            logging.debug(f"Mapping clicked: {app_name}")

            # Reset previous selection
            if self.selected_mapping and self.selected_mapping in self.mapping_frames:
                self.mapping_frames[self.selected_mapping]["frame"].configure(
                    fg_color="transparent"
                )

            # Update selection
            self.selected_mapping = app_name
            if app_name in self.mapping_frames:
                self.mapping_frames[app_name]["frame"].configure(
                    fg_color=self.colors["accent"]
                )

            # Update input fields
            config = self.app_device_map.get(app_name, {})
            self.app_entry.delete(0, "end")
            self.app_entry.insert(0, app_name)

            device_type = config.get("type", "Speakers")
            self.type_var.set(device_type)
            self._update_device_list()

            # Find and set device name
            device_id = config.get("device_id")
            device_list = self.devices.get(device_type, [])
            device = next(
                (d for d in device_list if str(d["id"]) == str(device_id)), None
            )

            if device:
                self.device_combo.set(device["name"])

        except Exception as e:
            logging.error(f"Error handling mapping click: {e}", exc_info=True)

    def _load_mappings(self, search_text=""):
        """Refresh mapping list with optional search filter"""
        try:
            # Clear existing mappings
            for widget in self.mapping_widgets:
                widget.destroy()
            self.mapping_widgets.clear()
            self.mapping_frames.clear()

            # Filter and sort mappings
            filtered_mappings = []
            for app, config in self.app_device_map.items():
                if search_text:
                    # Search in app name and device type/name
                    device_type = config["type"]
                    device_id = config["device_id"]
                    device_list = self.devices.get(device_type, [])
                    device = next(
                        (d for d in device_list if str(d["id"]) == str(device_id)), None
                    )
                    device_name = device["name"] if device else "Unknown Device"
                    search_target = f"{app} {device_type} {device_name}".lower()

                    if search_text.lower() not in search_target:
                        continue

                filtered_mappings.append((app, config))

            # Sort filtered mappings
            sorted_mappings = sorted(filtered_mappings, key=lambda x: x[0].lower())

            # Create widgets for filtered mappings
            for app, config in sorted_mappings:
                try:
                    device_type = config["type"]
                    device_id = config["device_id"]
                    device_list = self.devices.get(device_type, [])
                    device = next(
                        (d for d in device_list if str(d["id"]) == str(device_id)), None
                    )
                    device_name = device["name"] if device else "Unknown Device"

                    # Create widget with highlighted text if searching
                    widget = self._create_mapping_widget(
                        self.list_container, app, device_type, device_name, search_text
                    )
                    self.mapping_widgets.append(widget)

                except Exception as e:
                    logging.error(f"Error loading mapping for {app}: {e}")

            if not filtered_mappings and search_text:
                # Show no results message
                no_results = ctk.CTkLabel(
                    self.list_container,
                    text="No matches found",
                    text_color=self.colors["text_secondary"],
                    font=("Segoe UI", 12),
                )
                no_results.pack(pady=20)
                self.mapping_widgets.append(no_results)

        except Exception as e:
            logging.error(f"Error loading mappings: {e}")
            messagebox.showerror("Error", "Failed to load device mappings")

    def _load_config(self):
        """Load configuration from file"""
        try:
            if not os.path.exists(self.config_file):
                return False

            with open(self.config_file, "r") as f:
                config = json.load(f)
                self.app_device_map = config.get("app_device_map", {})
                logging.info(f"Loaded {len(self.app_device_map)} mappings from config")
            return True
        except Exception as e:
            logging.error(f"Error loading config: {e}")
            return False

    def _save_config(self):
        """Save current mappings to config file"""
        try:
            # Read existing config first
            if os.path.exists(self.config_file):
                with open(self.config_file, "r") as f:
                    config = json.load(f)
            else:
                config = {}

            # Update only app_device_map section
            config["app_device_map"] = self.app_device_map

            # Create backup
            if os.path.exists(self.config_file):
                backup_file = f"{self.config_file}.bak"
                os.replace(self.config_file, backup_file)

            # Save updated config
            with open(self.config_file, "w") as f:
                json.dump(config, f, indent=4)

            logging.info(f"Saved {len(self.app_device_map)} mappings to config")
            return True

        except Exception as e:
            logging.error(f"Error saving config: {e}")
            return False

    def save_state(self):
        """Save current state and update config"""
        try:
            logging.info("Saving state changes...")
            current_state = self.app_device_map.copy()

            if current_state == self._last_saved_state:
                logging.debug("No changes to save")
                return True

            # Save to config file first
            if not self._save_config():
                logging.error("Failed to save to config file")
                return False

            # Now send to main process
            sanitized_state = {
                app: {
                    "type": str(config.get("type", "Speakers")),
                    "device_id": str(config.get("device_id", "")),
                    "disabled": bool(config.get("disabled", False)),
                }
                for app, config in current_state.items()
            }

            if self.send_message("update_mapping", sanitized_state):
                self._last_saved_state = current_state.copy()
                self._changes_pending = False
                logging.info(
                    f"State saved successfully with {len(sanitized_state)} mappings"
                )
                return True

            logging.error("Failed to send state update")
            return False

        except Exception as e:
            logging.error(f"Error saving state: {e}", exc_info=True)
            return False

    def _add_mapping(self):
        """Add or update mapping with config save"""
        try:
            app_name = self.app_entry.get().strip().lower()
            if not app_name:
                self.show_error(
                    "Please enter an application name or select using Browse"
                )
                return

            # Get filepath if it was set by browse
            filepath = getattr(self.app_entry, "_filepath", None)
            if filepath:
                # Store just the filename without extension as the app name
                app_name = os.path.splitext(os.path.basename(filepath))[0].lower()

            device_type = self.type_var.get()
            device_name = self.device_var.get()

            # Find the device
            device_list = self.devices.get(device_type, [])
            device = next((d for d in device_list if d["name"] == device_name), None)

            if not device:
                self.show_error("Selected device not found")
                return

            # Preserve disabled state if updating existing mapping
            disabled_state = False
            if app_name in self.app_device_map:
                disabled_state = self.app_device_map[app_name].get("disabled", False)

            # Update local state with filepath if available
            mapping_data = {
                "type": device_type,
                "device_id": device.get("id", str(device.get("index", ""))),
                "disabled": disabled_state,
            }

            # Add filepath if it exists
            if filepath:
                mapping_data["filepath"] = filepath

            self.app_device_map[app_name] = mapping_data
            self._changes_pending = True

            if self.save_state():
                logging.info(f"Added/updated mapping for {app_name}")
                self._load_mappings()
                self.app_entry.delete(0, "end")
                # Clear stored filepath
                if hasattr(self.app_entry, "_filepath"):
                    delattr(self.app_entry, "_filepath")
                self.show_success("Mapping saved successfully")
            else:
                self.app_device_map = self._last_saved_state.copy()
                self.show_error("Failed to save changes")

        except Exception as e:
            logging.error(f"Error in add_mapping: {e}", exc_info=True)
            self.show_error(f"Failed to add mapping: {str(e)}")

    def _show_save_success(self):
        """Show success message"""
        try:
            messagebox.showinfo("Success", "Settings saved successfully")
        except Exception as e:
            logging.error(f"Error showing save confirmation: {e}")

    def _confirm_save(self):
        """Show save confirmation"""
        try:
            messagebox.showinfo("Success", "Mapping saved successfully")
        except Exception as e:
            logging.error(f"Error showing save confirmation: {e}")

    def _delete_mapping(self):
        """Delete mapping and save"""
        try:
            if not self.selected_mapping:
                self.show_error("Please select a mapping to delete")
                return

            if not self.confirm_dialog("Delete Mapping", "Delete selected mapping?"):
                return

            app_name = self.selected_mapping
            if app_name in self.app_device_map:
                # Update local state
                del self.app_device_map[app_name]

                # Save to main process
                if self.save_state():
                    self._load_mappings()
                    self.selected_mapping = None
                    self.app_entry.delete(0, "end")
                    self.show_success("Mapping deleted successfully")
                else:
                    self.show_error("Failed to save changes")

        except Exception as e:
            logging.error(f"Error deleting mapping: {e}")
            self.show_error(f"Failed to delete mapping: {str(e)}")

    def show_error(self, message):
        """Show error dialog"""
        logging.error(message)
        messagebox.showerror("Error", message)

    def show_success(self, message):
        """Show success dialog"""
        logging.info(message)
        messagebox.showinfo("Success", message)

    def confirm_dialog(self, title, message):
        """Show confirmation dialog"""
        return messagebox.askyesno(title, message)

    def _set_window_properties(self):
        """Set window properties safely in main thread"""
        self.root.grab_set()
        self.root.focus_force()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        """Handle window close properly"""
        self.root.grab_release()
        self.root.destroy()

    def _on_type_change(self, value=None):
        """Handle device type change with optional value parameter"""
        try:
            logging.debug(f"Type changed to: {self.type_var.get()}")
            self._update_device_list()

            # Select first device by default
            device_values = self.device_combo.cget("values")
            if device_values and len(device_values) > 0:
                self.device_combo.set(device_values[0])
                logging.debug(f"Set default device: {device_values[0]}")
            else:
                logging.warning("No devices available after type change")

        except Exception as e:
            logging.error(f"Error in type change handler: {e}", exc_info=True)

    def _update_device_list(self, event=None):
        try:
            device_type = self.type_var.get()
            logging.debug(f"Updating device list for type: {device_type}")

            # Get device list using exact type name
            device_list = self.devices.get(device_type, [])
            logging.debug(f"Found devices: {device_list}")

            if not device_list:
                self.device_combo.configure(values=["No devices configured"])
                self.device_combo.set("No devices configured")
                return

            # Filter and validate devices
            valid_devices = []
            for device in device_list:
                name = device.get("name")
                if name:
                    valid_devices.append(name)
                    logging.debug(f"Added valid device: {name}")

            if valid_devices:
                self.device_combo.configure(values=valid_devices)
                self.device_combo.set(valid_devices[0])
                logging.info(f"Loaded {len(valid_devices)} devices")
            else:
                self.device_combo.configure(values=["No valid devices"])
                self.device_combo.set("No valid devices")
                logging.warning("No valid devices found")

        except Exception as e:
            logging.error(f"Error updating device list: {e}", exc_info=True)
            self.device_combo.configure(values=["Error loading devices"])
            self.device_combo.set("Error loading devices")

    def _on_select(self, event):
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            values = item["values"]

            self.app_entry.delete(0, tk.END)
            self.app_entry.insert(0, values[0])
            self.type_combo.set(values[1])
            self._update_device_list()
            self.device_combo.set(values[2])

    def _on_search_change(self, *args):
        """Handle search text changes"""
        try:
            search_text = self.search_var.get().lower().strip()
            self._load_mappings(search_text)
        except Exception as e:
            logging.error(f"Error in search: {e}")

    def _clear_search(self):
        """Clear search field and reset list"""
        self.search_var.set("")
        self.search_entry.focus_set()
        self._load_mappings()

    def _highlight_text(self, text, search_text):
        """Create highlighted text with search matches"""
        if not search_text:
            return text

        try:
            parts = []
            last_end = 0
            text_lower = text.lower()
            search_lower = search_text.lower()

            # Find all occurrences of search text
            start = text_lower.find(search_lower)
            while start != -1:
                # Add text before match
                if start > last_end:
                    parts.append(text[last_end:start])

                # Add highlighted match
                match_end = start + len(search_text)
                match_text = text[start:match_end]
                parts.append(f"**{match_text}**")  # Bold text for highlighting

                last_end = match_end
                start = text_lower.find(search_lower, last_end)

            # Add remaining text
            if last_end < len(text):
                parts.append(text[last_end:])

            return "".join(parts)
        except Exception as e:
            logging.error(f"Error highlighting text: {e}")
            return text

    def _browse_application(self):
        """Open file browser to select application"""
        try:
            from tkinter import filedialog

            filetypes = [("Applications", "*.exe"), ("All files", "*.*")]

            filepath = filedialog.askopenfilename(
                title="Select Application",
                filetypes=filetypes,
                initialdir=os.path.expandvars(r"%ProgramFiles%"),
            )

            if filepath:
                # Get the filename without extension for the app name
                app_name = os.path.splitext(os.path.basename(filepath))[0]

                # Clean up the name
                app_name = app_name.lower()

                # Update the entry field
                self.app_entry.delete(0, "end")
                self.app_entry.insert(0, app_name)

                # If this is a new application, set default device type
                if app_name not in self.app_device_map:
                    self.type_var.set("Speakers")
                    self._update_device_list()

                logging.info(f"Selected application: {app_name} from {filepath}")

                # Optional: Store the full path in a hidden variable
                self.app_entry._filepath = filepath

        except Exception as e:
            logging.error(f"Error browsing for application: {e}")
            self.show_error("Failed to select application")

    def _force_reload_config(self):
        """Force reload configuration from file"""
        try:
            logging.info("Force reloading configuration")
            if self._load_config():
                self._load_mappings()  # Refresh the mapping list
                self.show_success("Configuration reloaded successfully")
                # Send message to main process to reload
                self.send_message("force_reload", None)
            else:
                self.show_error("Failed to reload configuration")
        except Exception as e:
            logging.error(f"Error force reloading config: {e}")
            self.show_error("Failed to reload configuration")
