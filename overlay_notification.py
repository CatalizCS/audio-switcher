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

            # Configure modern styles
            style = ttk.Style(self.root)
            style.configure(
                "Custom.TLabel",
                font=("Segoe UI", 11),
                background="#2D2D2D",
                foreground="#FFFFFF",
                padding=8,
            )
            style.configure(
                "Title.TLabel",
                font=("Segoe UI", 12, "bold"),
                background="#2D2D2D",
                foreground="#FFFFFF",
                padding=(8, 8, 8, 4),
            )
            style.configure(
                "Custom.TFrame", background="#2D2D2D", borderwidth=1, relief="solid"
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

            # Create notification window
            window = tk.Toplevel(self.root)
            window.overrideredirect(True)
            window.attributes("-topmost", True)
            window.configure(bg="#2D2D2D")

            # Add rounded corners effect
            window.attributes("-transparentcolor", "#2D2D2D")

            # Main frame with padding
            frame = ttk.Frame(window, style="Custom.TFrame", padding=(15, 10, 15, 12))
            frame.pack(expand=True, fill="both")

            # Icon frame for type indicator
            icon_text = "🔊" if "speaker" in title.lower() else "🎧"
            icon_label = ttk.Label(
                frame, text=icon_text, style="Title.TLabel", font=("Segoe UI", 14)
            )
            icon_label.pack(side="left", padx=(0, 10))

            # Text frame for title and message
            text_frame = ttk.Frame(frame, style="Custom.TFrame")
            text_frame.pack(side="left", fill="both", expand=True)

            # Title with bottom padding
            title_label = ttk.Label(text_frame, text=title, style="Title.TLabel")
            title_label.pack(anchor="w", pady=(0, 2))

            # Message with custom wrapping
            msg_label = ttk.Label(
                text_frame, text=message, style="Custom.TLabel", wraplength=300
            )
            msg_label.pack(anchor="w")

            # Position window - MODIFIED FOR CENTER BOTTOM
            window.update_idletasks()
            width = window.winfo_width() + 20  # Add padding
            height = window.winfo_height() + 10  # Add padding

            # Calculate center position
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()

            # Center horizontally, position at bottom with 60px margin
            x = (screen_width - width) // 2
            y = screen_height - height - 60  # 60px from bottom

            window.geometry(f"+{x}+{y}")

            # Smooth fade in with slide - FASTER ANIMATION
            window.attributes("-alpha", 0)
            original_y = y
            slide_distance = 15  # Reduced from 20

            # Faster fade in (reduced iterations and sleep time)
            for i in range(6):  # Reduced from 11
                current_y = int(original_y + (slide_distance * (5 - i) / 5))
                window.geometry(f"+{x}+{current_y}")
                window.attributes("-alpha", i / 5)
                window.update()
                time.sleep(0.01)  # Reduced from 0.02

            # Force full opacity at end
            window.attributes("-alpha", 1)
            window.geometry(f"+{x}+{original_y}")

            # Schedule fade out
            def fade_out():
                try:
                    # Faster fade out
                    for i in range(5, -1, -1):  # Reduced from 10
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
