from ctypes.wintypes import HWND, DWORD
from datetime import datetime
from threading import Lock
from threading import Thread
from tkinter import ttk
import csv
import ctypes
import os
import sys
import tempfile
import time
import tkinter as tk
import rethyxyz.rethyxyz

# Function prototype in Windows API
SetWindowDisplayAffinity = ctypes.windll.user32.SetWindowDisplayAffinity
SetWindowDisplayAffinity.argtypes = [HWND, DWORD]
SetWindowDisplayAffinity.restype = ctypes.c_bool

# Constants for display affinity
WDA_NONE = 0
WDA_MONITOR = 1

def set_window_monitor_only(hwnd):
    """
    Attempt to make the window content only visible on the monitor,
    making it more difficult to capture via screen capture software.
    """
    result = SetWindowDisplayAffinity(hwnd, WDA_MONITOR)
    if not result:
        print("Failed to set display affinity.")
    else:
        print("Display affinity set to monitor only.")

class TimeTracker(tk.Tk):
    def __init__(self):
        super().__init__()
        self.configure(bg='#EDEDED')  # SAP-like background color
        # Other initialization code...
        self.file_lock = Lock()  # Initialize the lock
        self.title("ConsultantTimeTracker")
        self.projects = self.load_projects("projects.txt")
        self.iconbitmap('consulting.ico')
        self.timers = {project: 0 for project in self.projects}  # Initialize timers for each project
        self.buttons = {}  # To store button widgets
        self.current_project = None
        self.running = False
        self.create_widgets()
        self.after(100, self.lock_window_size)  # Ensure widgets are drawn before locking size

        self.log_file_name = f"project_times_log_{datetime.now().strftime('%Y-%m-%d')}.txt"  # Dynamic log file name
        self.load_existing_times()

    def apply_display_affinity(self):
        hwnd = self.frame()  # This method might not directly give you the HWND. See note below.
        set_window_monitor_only(hwnd)

    def load_existing_times(self):
        """Load existing times from today's log file if it exists."""
        if not os.path.exists(self.log_file_name):
            return  # File doesn't exist, nothing to load

        with open(self.log_file_name, "r") as file:
            reader = csv.reader(file)
            next(reader, None)  # Skip header
            for row in reader:
                if len(row) < 2:
                    continue  # Skip malformed lines
                project, time_str = row
                project = project.lower()  # Match case-insensitive project names

                # Initialize total_seconds to accumulate time
                total_seconds = 0

                # Parse the time string
                time_parts = time_str.split()
                for part in time_parts:
                    if 'h' in part:
                        hours = int(part.replace('h', ''))
                        total_seconds += hours * 3600
                    elif 'm' in part:
                        minutes = int(part.replace('m', ''))
                        total_seconds += minutes * 60
                    elif 's' in part:
                        seconds = int(part.replace('s', ''))
                        total_seconds += seconds

                # Match and accumulate times case-insensitively
                for existing_project in self.timers.keys():
                    if existing_project.lower() == project:
                        self.timers[existing_project] += total_seconds
                        break

    def load_projects(self, filepath):
        if not os.path.exists(filepath):
            print(f"Failed to load project file \"{filepath}\"")
            sys.exit(0)
        """Load projects from a file."""
        with open(filepath, 'r') as file:
            projects = [line.strip() for line in file.readlines()]
        return projects

    def create_widgets(self):
        style = ttk.Style()
        style.configure('TButton', background='#D3D3D3', foreground='black', font=('Helvetica', 10))
        style.configure('Active.TButton', background='#A9A9A9', font=('Helvetica', 10, 'bold'))

        for project in self.projects:
            btn = ttk.Button(self, text=project, command=lambda proj=project: self.start_or_pause_timer(proj), style='TButton')
            btn.pack(pady=5, fill=tk.X, padx=10)
            self.buttons[project] = btn

        self.timer_label = ttk.Label(self, text="00:00:00", font=("Helvetica", 16), background='#EDEDED', foreground='black')
        self.timer_label.pack(pady=20)

    def start_or_pause_timer(self, project):
        """Start or pause the timer for the selected project, and log the time after each operation."""
        if project == self.current_project and self.running:
            self.running = False
            self.timer_thread.join()  # Wait for the timer to pause
            self.buttons[project].config(style='TButton')  # Reset the button style
            self.log_time()  # Log time when pausing
        elif project == self.current_project and not self.running:
            self.running = True
            self.buttons[project].config(style='Active.TButton')  # Highlight the button
            self.timer_thread = Thread(target=self.update_timer)
            self.timer_thread.start()
            self.log_time()  # Log time immediately after restarting the timer
        else:
            if self.current_project:
                self.running = False
                self.timer_thread.join()  # Ensure the current timer is stopped before switching
                self.buttons[self.current_project].config(style='TButton')  # Reset previous button style
                self.log_time()  # Log time when switching projects
            self.current_project = project
            self.running = True
            self.buttons[project].config(style='Active.TButton')  # Highlight the button
            self.timer_thread = Thread(target=self.update_timer)
            self.timer_thread.start()
            self.log_time()  # Log time immediately after starting a new timer

    def update_timer(self):
        """Update the timer for the current project."""
        while self.running:
            self.timers[self.current_project] += 1
            time_spent = self.timers[self.current_project]
            hours, remainder = divmod(time_spent, 3600)
            minutes, seconds = divmod(remainder, 60)
            time_str = f"{hours:02}:{minutes:02}:{seconds:02}"
            self.timer_label.config(text=time_str)
            self.update()
            time.sleep(1)

    def log_time(self):
        """Safely log the total time spent on each project to a CSV file."""
        temp_log_path = tempfile.NamedTemporaryFile(delete=False).name
        try:
            with open(temp_log_path, "w", newline='') as temp_file:
                writer = csv.writer(temp_file)
                writer.writerow(["PROJECT", "TOTAL_TIME"])  # Write header
                for project, seconds in self.timers.items():
                    hours, remainder = divmod(seconds, 3600)
                    minutes, seconds = divmod(remainder, 60)
                    time_str = f"{hours}h {minutes}m {seconds}s"
                    writer.writerow([project.upper(), time_str])
            # Replace the old log file with the new one
            os.replace(temp_log_path, self.log_file_name)
        except Exception as e:
            print(f"Error logging time: {e}")
            os.unlink(temp_log_path)  # Remove the temporary file in case of error

    def on_closing(self):
        """Handle window closing event."""
        if self.current_project and self.running:
            self.running = False
            self.timer_thread.join()  # Ensure the timer thread is stopped before closing
        self.log_time()  # Log time before closing
        self.destroy()

    def lock_window_size(self):
        """Lock the window size based on content."""
        self.update_idletasks()  # Ensure all widgets are drawn and sizes are updated
        window_width = max(btn.winfo_width() for btn in self.buttons.values()) + 200  # Adjust based on content and padding
        window_height = sum(btn.winfo_height() for btn in self.buttons.values()) + self.timer_label.winfo_height() + 100  # Adjust for spacing
        self.geometry(f"{window_width}x{window_height}")
        self.resizable(False, True)  # Prevent resizing

if __name__ == "__main__":
    rethyxyz.rethyxyz.show_intro("ConsultantTimeTracker")
    app = TimeTracker()

    # Style configuration for the active (current) project button
    style = ttk.Style()
    style.configure('TButton', background='lightgrey')  # Default button style
    style.configure('Active.TButton', background='lightgreen')  # Active button style

    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()
