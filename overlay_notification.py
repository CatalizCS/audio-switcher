import logging
import tkinter as tk
from tkinter import ttk
import threading
import time
import queue


class OverlayNotification:
    def __init__(self):
        self.notification_queue = queue.Queue()
        self.is_showing = False
        self.lock = threading.Lock()
        self._active = True
        self._setup_done = threading.Event()

        # Initialize Tkinter in the main thread
        self.root = None
        self._init_root()

    def _init_root(self):
        """Initialize root window safely"""
        try:
            self.root = tk.Tk()
            self.root.withdraw()
            self.root.protocol("WM_DELETE_WINDOW", self.destroy)
            self.root.wm_attributes("-topmost", True)

            # Updated modern styles with new color scheme
            style = ttk.Style(self.root)
            style.configure(
                "Custom.TLabel",
                font=("Segoe UI", 10),
                background="#2B2B2B",
                foreground="#E8E8E8",
                padding=6,
            )
            style.configure(
                "Title.TLabel",
                font=("Segoe UI Semibold", 11),
                background="#2B2B2B",
                foreground="#FFFFFF",
                padding=(6, 6, 6, 2),
            )
            style.configure(
                "Custom.TFrame",
                background="#2B2B2B",
                borderwidth=0,
            )

            logging.debug("Notification system initialized")

            # Create message queue handler
            self.root.after(100, self._check_queue)
            self._setup_done.set()

        except Exception as e:
            logging.error(f"Failed to initialize notification window: {e}")
            self._active = False
            self._setup_done.set()
            raise

    def _check_queue(self):
        """Check message queue periodically"""
        if self._active and self.root and self.root.winfo_exists():
            if not self.is_showing:
                try:
                    title, message, duration = self.notification_queue.get_nowait()
                    self._show_notification(title, message, duration)
                except queue.Empty:
                    pass
            self.root.after(100, self._check_queue)

    def show_notification(self, title, message, duration=2.0):
        """Queue notification for display"""
        if not self._active:
            return

        try:
            self.notification_queue.put((title, message, duration))
        except Exception as e:
            logging.error(f"Error queueing notification: {e}")

    def _show_notification(self, title, message, duration):
        """Show actual notification window"""
        try:
            self.is_showing = True

            window = tk.Toplevel(self.root)
            window.overrideredirect(True)
            window.attributes("-topmost", True)

            # Enhanced shadow and border effect
            window.configure(bg="#1A1A1A")
            window.attributes("-transparentcolor", "#1A1A1A")

            # Create rounded frame with border
            content_frame = tk.Frame(
                window,
                bg="#2B2B2B",
                highlightbackground="#3B3B3B",
                highlightthickness=1,
            )
            content_frame.pack(padx=3, pady=3)

            # Create inner frame with additional styling
            inner_frame = tk.Frame(
                content_frame,
                bg="#2B2B2B",
                padx=2,
                pady=2,
            )
            inner_frame.pack(fill="both", expand=True)

            # Main content frame with rounded corners
            frame = ttk.Frame(
                inner_frame, style="Custom.TFrame", padding=(12, 8, 12, 8)
            )
            frame.pack(expand=True, fill="both")

            def create_rounded_frame(parent, bg_color="#2B2B2B", corner_radius=10):
                canvas = tk.Canvas(
                    parent,
                    bg=bg_color,
                    highlightthickness=0,
                    width=400,
                    height=200,
                )
                canvas.pack(expand=True, fill="both")

                # Create rounded rectangle
                canvas.create_rectangle(
                    corner_radius,
                    corner_radius,
                    canvas.winfo_reqwidth() - corner_radius,
                    canvas.winfo_reqheight() - corner_radius,
                    fill=bg_color,
                    outline=bg_color,
                )
                return canvas

            rounded_canvas = create_rounded_frame(frame)

            # Rest of the content (icon and text) now goes on the canvas
            icon_text = "🔊" if "speaker" in title.lower() else "🎧"
            icon_label = tk.Label(
                rounded_canvas,
                text=icon_text,
                font=("Segoe UI", 13),
                bg="#2B2B2B",
                fg="#FFFFFF",
            )
            icon_label.pack(side="left", padx=(8, 8))

            # Text container
            text_frame = tk.Frame(rounded_canvas, bg="#2B2B2B")
            text_frame.pack(side="left", fill="both", expand=True)

            # Title label
            title_label = tk.Label(
                text_frame,
                text=title,
                font=("Segoe UI Semibold", 11),
                bg="#2B2B2B",
                fg="#FFFFFF",
            )
            title_label.pack(anchor="w", pady=(0, 1))

            # Message label
            msg_label = tk.Label(
                text_frame,
                text=message,
                font=("Segoe UI", 10),
                bg="#2B2B2B",
                fg="#E8E8E8",
                wraplength=250,
            )
            msg_label.pack(anchor="w")

            window.update_idletasks()
            width = window.winfo_width() + 16
            height = window.winfo_height() + 8
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            x = (screen_width - width) // 2
            y = screen_height - height - 50

            window.attributes("-alpha", 0)
            original_y = y
            slide_distance = 15

            for i in range(6):
                current_y = int(original_y + (slide_distance * (5 - i) / 5))
                window.geometry(f"+{x}+{current_y}")
                window.attributes("-alpha", i / 5)
                window.update()
                time.sleep(0.01)

            # Force full opacity at end
            window.attributes("-alpha", 1)
            window.geometry(f"+{x}+{original_y}")

            # Schedule fade out
            def fade_out():
                try:
                    # Faster fade out
                    for i in range(5, -1, -1):
                        if window.winfo_exists():
                            window.attributes("-alpha", i / 5)
                            current_y = int(original_y + (slide_distance * (5 - i) / 5))
                            window.geometry(f"+{x}+{current_y}")
                            window.update()
                            time.sleep(0.02)
                    if window.winfo_exists():
                        window.destroy()
                except Exception as e:
                    logging.error(f"Error in fade out: {e}")
                finally:
                    self.is_showing = False

            window.after(int(duration * 1000), fade_out)

            def on_close():
                try:
                    if window.winfo_exists():
                        window.destroy()
                except:
                    pass
                finally:
                    self.is_showing = False

            # Ensure cleanup happens even if window is closed
            window.protocol("WM_DELETE_WINDOW", on_close)

        except Exception as e:
            logging.error(f"Error showing notification: {e}")
            self.is_showing = False

    def process_events(self):
        """Process Tkinter events in main thread"""
        if self._active and self.root and self.root.winfo_exists():
            try:
                self.root.update()
            except Exception as e:
                if self._active:  # Only log if not shutting down
                    logging.error(f"Error updating root window: {e}")

    def destroy(self):
        """Clean up resources"""
        self._active = False
        if self.root:
            try:
                self.root.quit()
                self.root.destroy()
            except:
                pass
            self.root = None
