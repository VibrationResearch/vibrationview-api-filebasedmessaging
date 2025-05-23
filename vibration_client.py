#!/usr/bin/env python3
"""
VibrationVIEW Client - Python refactor of VB6 application
A client application for controlling VibrationVIEW software through file-based communication.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import time
import subprocess
import threading
import winreg
from pathlib import Path
from datetime import datetime
import json

# Constants
VIBRATIONVIEW_VERSION = "2025.0"


class VibrationViewClient:
    def __init__(self, root):
        self.root = root
        self.root.title("VibrationVIEW Client")
        self.root.geometry("800x500")
        
        # Instance variables
        self.control_file = ""
        self.response_file = ""
        self.system_file_path = ""
        self.last_file_time = 0
        self.timeout_count = 0
        self.retry_count = 0
        self.current_command = ""
        self.refresh_timer = None
        self.timer_running = False
        
        # Initialize the application
        self.setup_file_paths()
        self.create_widgets()
        
    def setup_file_paths(self):
        """Setup file paths and registry settings"""

        self.profile_file_path = "C:\\VibrationVIEW\\Profiles\\"
        self.data_file_path = "C:\\VibrationVIEW\\Data\\"

        try:
            # Get application path
            app_path = os.path.dirname(os.path.abspath(__file__))
            self.control_file = os.path.join(app_path, "RemoteControl.txt")
            # Set response file path 
            # Note the RemoteControl.Status is in C:\Program Files by default, require a permissions modification,
            # or running vibrationview with elevated permissions (As Administrator)
            if self.system_file_path:
                self.response_file = os.path.join(self.system_file_path, "RemoteControl.Status")

            # when we add the registry setting to setup the response file - this is better            
            # self.response_file = os.path.join(app_path, "RemoteControl.Status")


            # Try to get system file path from registry
            try:
                keyname = rf"SOFTWARE\Vibration Research Corporation\{VIBRATIONVIEW_VERSION}"
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, keyname) as key:
                    system_path, _ = winreg.QueryValueEx(key, "System File Path")
                    self.system_file_path = system_path
                    if not self.system_file_path.endswith("\\"):
                        self.system_file_path += "\\"
                        
                    # Set remote control file in registry only if it's different
                    try:
                        # First check if the value needs to be updated
                        current_control_file = None
                        try:
                            keyparams = rf"SOFTWARE\Vibration Research Corporation\VibrationVIEW\{VIBRATIONVIEW_VERSION}\System Parameters"
                            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                              keyparams,
                                              0, winreg.KEY_READ) as param_key:
                                current_control_file, _ = winreg.QueryValueEx(param_key, "Remote Control File")
                        except (FileNotFoundError, OSError):
                            # Key or value doesn't exist
                            pass
                        
                        # Only write if the value is different or doesn't exist
                        if current_control_file != self.control_file:
                            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                              keyparams,
                                              0, winreg.KEY_SET_VALUE) as param_key:
                                winreg.SetValueEx(param_key, "Remote Control File", 0, 
                                                winreg.REG_SZ, self.control_file)
                                winreg.SetValueEx(param_key, "Remote Status File", 0, 
                                                winreg.REG_SZ, self.response_file)
                                
                                print(f"Updated registry: Remote Control File = {self.control_file}")
                                # Notify user that VibrationVIEW needs to be restarted
                                messagebox.showinfo("Registry Updated", 
                                                  "Registry has been updated with the remote control file path.\n\n"
                                                  "Please restart VibrationVIEW for the changes to take effect.")
                        else:
                            print(f"Registry already correct: Remote Control File = {self.control_file}")
                    except Exception as e:
                        print(f"Warning: Could not set registry value: {e}")
                        
            except FileNotFoundError:
                # Registry key not found, ask user for path
                messagebox.showwarning("Registry Not Found", 
                                     "VibrationVIEW registry entry not found. Please select installation directory.")
                self.system_file_path = filedialog.askdirectory(title="Select VibrationVIEW Installation Directory")
                if self.system_file_path and not self.system_file_path.endswith("\\"):
                    self.system_file_path += "\\"
                    
            
            # Initialize file time tracking
            self.last_file_time = self.get_file_time(self.response_file)
            
        except Exception as e:
            messagebox.showerror("Initialization Error", f"Error setting up file paths: {e}")
    
    def create_widgets(self):
        """Create the GUI widgets"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Left panel for buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=0, column=0, sticky=(tk.W, tk.N), padx=(0, 10))
        
        # Buttons
        ttk.Button(button_frame, text="Load", command=self.load_profile, width=12).grid(row=0, column=0, pady=5)
        ttk.Button(button_frame, text="Run", command=self.run_test, width=12).grid(row=1, column=0, pady=5)
        ttk.Button(button_frame, text="Stop", command=self.stop_test, width=12).grid(row=2, column=0, pady=5)
        ttk.Button(button_frame, text="Status", command=self.get_status, width=12).grid(row=3, column=0, pady=5)
        ttk.Button(button_frame, text="Convert", command=self.convert_data, width=12).grid(row=4, column=0, pady=5)
        ttk.Button(button_frame, text="Clear Log", command=self.clear_log, width=12).grid(row=5, column=0, pady=5)
        
        # Response text area
        text_frame = ttk.Frame(main_frame)
        text_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        # Text widget with scrollbar
        self.response_text = tk.Text(text_frame, wrap=tk.WORD, state=tk.DISABLED)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.response_text.yview)
        self.response_text.configure(yscrollcommand=scrollbar.set)
        
        self.response_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
    
    def get_file_time(self, file_path):
        """Get file modification time with high precision"""
        try:
            if os.path.exists(file_path):
                return os.path.getmtime(file_path)
            return 0
        except Exception:
            return 0
    
    def send_command(self, command):
        """Send command to the control file"""
        try:
            with open(self.control_file, 'w') as f:
                f.write(command)
            self.current_command = command
            self.log_command(command)
            self.start_refresh_timer()
            self.status_var.set(f"Sent command: {command}")
        except Exception as e:
            messagebox.showerror("Send Command Error", f"Error sending command: {e}")
    
    def log_command(self, command):
        """Log the sent command with timestamp"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        self.response_text.config(state=tk.NORMAL)
        
        # If this is not the first entry, add a separator
        if self.response_text.get(1.0, tk.END).strip():
            self.response_text.insert(tk.END, "\n")
        
        # Add timestamp and command
        self.response_text.insert(tk.END, f"[{timestamp}] Sent: {command}\n")
        
        self.response_text.config(state=tk.DISABLED)
        self.response_text.see(tk.END)
    
    def clear_log(self):
        """Clear the response log"""
        self.response_text.config(state=tk.NORMAL)
        self.response_text.delete(1.0, tk.END)
        self.response_text.config(state=tk.DISABLED)
        self.status_var.set("Log cleared")
    
    def start_refresh_timer(self):
        """Start the refresh timer to check for responses"""
        self.timeout_count = 0
        self.retry_count = 0
        self.timer_running = True
        self.check_response()
    
    def stop_refresh_timer(self):
        """Stop the refresh timer"""
        self.timer_running = False
        if self.refresh_timer:
            self.root.after_cancel(self.refresh_timer)
    
    def check_response(self):
        """Check for response file changes (equivalent to tmrRefresh_Timer)"""
        if not self.timer_running:
            return
            
        try:
            self.timeout_count += 1
            current_file_time = self.get_file_time(self.response_file)
            
            # Check if file has been modified
            if current_file_time != self.last_file_time and current_file_time > 0:
                self.last_file_time = current_file_time
                self.stop_refresh_timer()
                self.get_response()
                self.timeout_count = 0
                self.retry_count = 0
            else:
                # Handle timeout and retries
                if self.timeout_count > 12:  # 3 seconds at 250ms intervals
                    if self.retry_count > 3:
                        self.update_response_text("No Host Response")
                        self.stop_refresh_timer()
                        self.timeout_count = 0
                        self.retry_count = 0
                    else:
                        # Resend command
                        self.send_command_direct(self.current_command)
                        self.retry_count += 1
                        self.timeout_count = 0
            
            # Schedule next check
            if self.timer_running:
                self.refresh_timer = self.root.after(250, self.check_response)
                
        except Exception as e:
            self.update_response_text(f"Error checking response: {e}")
            self.stop_refresh_timer()
    
    def send_command_direct(self, command):
        """Send command directly without starting timer"""
        try:
            with open(self.control_file, 'w') as f:
                f.write(command)
        except Exception as e:
            print(f"Error sending command: {e}")
    
    def get_response(self):
        """Read the response file and update the text area"""
        try:
            if os.path.exists(self.response_file):
                with open(self.response_file, 'r') as f:
                    content = f.read()
                self.update_response_text(content)
            else:
                self.update_response_text("Response file not found")
        except Exception as e:
            self.update_response_text(f"Error reading response: {e}")
    
    def update_response_text(self, text):
        """Update the response text widget as a log with timestamps"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        self.response_text.config(state=tk.NORMAL)
        
        # If this is not the first entry, add a separator
        if self.response_text.get(1.0, tk.END).strip():
            self.response_text.insert(tk.END, "\n" + "-" * 50 + "\n")
        
        # Add timestamp and response
        self.response_text.insert(tk.END, f"[{timestamp}] Response:\n{text}\n")
        
        self.response_text.config(state=tk.DISABLED)
        self.response_text.see(tk.END)
    
    def load_profile(self):
        """Load a VibrationVIEW profile"""
        if not self.system_file_path:
            messagebox.showerror("Error", "System file path not set")
            return
            
        file_path = filedialog.askopenfilename(
            title="Select Profile",
            defaultextension=".vrp",
            filetypes=[("Random Profiles", "*.vrp"), ("All files", "*.*")],
            initialdir=self.profile_file_path
        )
        
        if file_path and os.path.exists(file_path):
            command = f"load {file_path}"
            self.send_command(command)
    
    def run_test(self):
        """Run the test"""
        self.send_command("run")
    
    def stop_test(self):
        """Stop the test"""
        self.send_command("stop")
    
    def get_status(self):
        """Get system status"""
        self.send_command("status")
    
    def convert_data(self):
        """Convert VRD file to CSV"""
        if not self.system_file_path:
            messagebox.showerror("Error", "System file path not set")
            return
            
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            defaultextension=".vrd",
            filetypes=[("Random Data", "*.vrd"), ("All files", "*.*")],
            initialdir=self.data_file_path
        )
        
        if file_path and os.path.exists(file_path):
            try:
                vibration_exe = os.path.join(self.system_file_path, "vibrationview.exe")
                if os.path.exists(vibration_exe):
                    # Run VibrationVIEW with CSV conversion parameter
                    subprocess.Popen([vibration_exe, "/csv", file_path])
                    self.status_var.set(f"Converted: {os.path.basename(file_path)}")
                else:
                    messagebox.showerror("Error", "VibrationVIEW executable not found")
            except Exception as e:
                messagebox.showerror("Conversion Error", f"Error converting file: {e}")
    
    def on_closing(self):
        """Handle application closing"""
        self.stop_refresh_timer()
        self.root.destroy()


def main():
    """Main application entry point"""
    root = tk.Tk()
    app = VibrationViewClient(root)
    
    # Handle window closing
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()


if __name__ == "__main__":
    main()