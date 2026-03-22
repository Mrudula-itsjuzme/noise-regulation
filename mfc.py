import serial
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import json
import csv
import os
from datetime import datetime
import pygame

# Set the correct COM port (Change as needed)
SERIAL_PORT = os.getenv("NOISE_MONITOR_SERIAL_PORT", "COM3")  # Override via env
# SERIAL_PORT = "/dev/ttyUSB0"  # Linux/Mac
BAUD_RATE = int(os.getenv("NOISE_MONITOR_BAUD", "115200"))

# Global variables
running = True
noise_history = []
max_history_length = 100
noise_min = 0
noise_max = 100  # Will be adjusted dynamically
threshold_crossed = False
threshold_time = 0
config_file = os.getenv("NOISE_MONITOR_CONFIG", "noise_config.json")

# Initialize volume control
try:
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    print("Volume control initialized successfully")
except Exception as e:
    print(f"Error initializing volume control: {e}")
    volume = None

# Initialize pygame for sounds
pygame.mixer.init()
try:
    alert_sound = pygame.mixer.Sound("alert.wav")
except:
    print("Alert sound file not found. Alerts will be silent.")
    alert_sound = None

# Create GUI window
root = tk.Tk()
root.title("Noise Level Monitor & Controller")
root.geometry("900x700")
root.configure(bg="#f0f0f0")

# Create style
style = ttk.Style()
style.theme_use('clam')
style.configure('TFrame', background='#f0f0f0')
style.configure('TLabelframe', background='#f0f0f0')
style.configure('TLabelframe.Label', background='#f0f0f0')
style.configure('TButton', background='#4a7abc', foreground='white')
style.map('TButton', background=[('active', '#5a8adc')])

# Top menu
menu_bar = tk.Menu(root)
root.config(menu=menu_bar)

# File menu
file_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Save Configuration", command=lambda: save_config())
file_menu.add_command(label="Load Configuration", command=lambda: load_config())
file_menu.add_separator()
file_menu.add_command(label="Export Data", command=lambda: export_data())
file_menu.add_separator()
file_menu.add_command(label="Exit", command=lambda: on_closing())

# Tools menu
tools_menu = tk.Menu(menu_bar, tearoff=0)
menu_bar.add_cascade(label="Tools", menu=tools_menu)
tools_menu.add_command(label="Calibrate", command=lambda: start_calibration())
tools_menu.add_command(label="Find COM Ports", command=lambda: find_com_ports())

# Create a better UI layout
main_frame = ttk.Frame(root, padding=10)
main_frame.pack(fill=tk.BOTH, expand=True)

# Create a notebook for tabs
notebook = ttk.Notebook(main_frame)
notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

# Main tab
main_tab = ttk.Frame(notebook)
notebook.add(main_tab, text="Monitor")

# Settings tab
settings_tab = ttk.Frame(notebook)
notebook.add(settings_tab, text="Settings")

# Log tab
log_tab = ttk.Frame(notebook)
notebook.add(log_tab, text="Event Log")

# Presets tab
presets_tab = ttk.Frame(notebook)
notebook.add(presets_tab, text="Presets")

# About tab
about_tab = ttk.Frame(notebook)
notebook.add(about_tab, text="About")

# Current values frame (on main tab)
current_frame = ttk.LabelFrame(main_tab, text="Current Values", padding=10)
current_frame.pack(fill=tk.X, padx=5, pady=5)

# Create grid for current values
ttk.Label(current_frame, text="Raw Noise Level:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
raw_value_var = tk.StringVar(value="0")
ttk.Label(current_frame, textvariable=raw_value_var, width=10).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

ttk.Label(current_frame, text="System Volume:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
volume_var = tk.StringVar(value="0%")
ttk.Label(current_frame, textvariable=volume_var, width=10).grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

ttk.Label(current_frame, text="Processed Noise:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
processed_var = tk.StringVar(value="0")
ttk.Label(current_frame, textvariable=processed_var, width=10).grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

ttk.Label(current_frame, text="Volume Level:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
volume_bar = ttk.Progressbar(current_frame, orient=tk.HORIZONTAL, length=150, mode='determinate')
volume_bar.grid(row=1, column=3, sticky=tk.W+tk.E, padx=5, pady=5)

# Status indicator
ttk.Label(current_frame, text="Status:").grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)
status_indicator = tk.Canvas(current_frame, width=20, height=20, bg="green")
status_indicator.grid(row=0, column=5, sticky=tk.W, padx=5, pady=5)

# Control frame
control_frame = ttk.LabelFrame(main_tab, text="Controls", padding=10)
control_frame.pack(fill=tk.X, padx=5, pady=5)

# Sensitivity control
sensitivity_var = tk.DoubleVar(value=3.0)
ttk.Label(control_frame, text="Sensitivity:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)

def update_sensitivity(val):
    sensitivity_label.config(text=f"{float(val):.1f}")

sensitivity_slider = ttk.Scale(control_frame, from_=1.0, to=5.0, orient=tk.HORIZONTAL, 
                             length=200, variable=sensitivity_var, command=update_sensitivity)
sensitivity_slider.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
sensitivity_label = ttk.Label(control_frame, text="3.0")
sensitivity_label.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

# Noise floor and ceiling adjustment
ttk.Label(control_frame, text="Min Threshold:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
min_var = tk.IntVar(value=0)
min_spin = ttk.Spinbox(control_frame, from_=0, to=1000, width=6, textvariable=min_var)
min_spin.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

ttk.Label(control_frame, text="Max Threshold:").grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)
max_var = tk.IntVar(value=3000)  # Higher value to accommodate the large readings
max_spin = ttk.Spinbox(control_frame, from_=100, to=10000, width=6, textvariable=max_var)
max_spin.grid(row=1, column=3, sticky=tk.W, padx=5, pady=5)

# Auto-calibration controls
auto_cal_var = tk.BooleanVar(value=True)
auto_cal_check = ttk.Checkbutton(control_frame, text="Auto-calibrate range", variable=auto_cal_var)
auto_cal_check.grid(row=0, column=3, padx=5, pady=5)

reset_button = ttk.Button(control_frame, text="Reset", 
                        command=lambda: [min_var.set(0), max_var.set(3000), noise_history.clear()])
reset_button.grid(row=0, column=4, padx=5, pady=5)

# Graphing frame
graph_frame = ttk.LabelFrame(main_tab, text="Real-time Monitoring", padding=10)
graph_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

# Set up matplotlib figure
fig, ax = plt.subplots(figsize=(8, 4))
canvas = FigureCanvasTkAgg(fig, graph_frame)
canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

# Initialize plots
noise_line, = ax.plot([], [], label='Noise Level', color='blue')
volume_line, = ax.plot([], [], label='System Volume', color='red')
threshold_line, = ax.plot([], [], label='Alert Threshold', color='green', linestyle='--')
ax.set_title('Noise and Volume Over Time')
ax.set_xlabel('Time (s)')
ax.set_ylabel('Level')
ax.legend()
ax.set_ylim(0, 100)
ax.grid(True)

# Status bar
status_var = tk.StringVar(value="Initializing...")
status_bar = ttk.Label(root, textvariable=status_var, relief=tk.SUNKEN, anchor=tk.W)
status_bar.pack(side=tk.BOTTOM, fill=tk.X)

# Settings Tab - Advanced Settings
settings_frame = ttk.LabelFrame(settings_tab, text="Connection Settings", padding=10)
settings_frame.pack(fill=tk.X, padx=5, pady=5)

# COM Port settings
ttk.Label(settings_frame, text="COM Port:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
com_port_var = tk.StringVar(value=SERIAL_PORT)
com_port_entry = ttk.Combobox(settings_frame, textvariable=com_port_var, width=15)
com_port_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

ttk.Label(settings_frame, text="Baud Rate:").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
baud_rate_var = tk.IntVar(value=BAUD_RATE)
baud_rate_combo = ttk.Combobox(settings_frame, textvariable=baud_rate_var, 
                              values=[9600, 19200, 38400, 57600, 115200], width=10)
baud_rate_combo.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

# Apply button
apply_button = ttk.Button(settings_frame, text="Apply Connection", 
                         command=lambda: restart_serial_connection())
apply_button.grid(row=0, column=4, padx=5, pady=5)

# Alert settings
alert_frame = ttk.LabelFrame(settings_tab, text="Alert Settings", padding=10)
alert_frame.pack(fill=tk.X, padx=5, pady=5)

# Alert threshold
ttk.Label(alert_frame, text="Alert Threshold:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
alert_threshold_var = tk.IntVar(value=80)
alert_threshold_spin = ttk.Spinbox(alert_frame, from_=0, to=100, width=5, textvariable=alert_threshold_var)
alert_threshold_spin.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

# Alert duration
ttk.Label(alert_frame, text="Required Duration (s):").grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
alert_duration_var = tk.DoubleVar(value=3.0)
alert_duration_spin = ttk.Spinbox(alert_frame, from_=0.5, to=10.0, increment=0.5, width=5, textvariable=alert_duration_var)
alert_duration_spin.grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)

# Alert enabled
alert_enabled_var = tk.BooleanVar(value=True)
alert_enabled_check = ttk.Checkbutton(alert_frame, text="Enable Alerts", variable=alert_enabled_var)
alert_enabled_check.grid(row=0, column=4, padx=5, pady=5)

# Sound alerts
sound_alert_var = tk.BooleanVar(value=True)
sound_alert_check = ttk.Checkbutton(alert_frame, text="Sound Alert", variable=sound_alert_var)
sound_alert_check.grid(row=1, column=0, padx=5, pady=5)

# Volume control settings
volume_frame = ttk.LabelFrame(settings_tab, text="Volume Control Settings", padding=10)
volume_frame.pack(fill=tk.X, padx=5, pady=5)

# Enable volume control
volume_control_var = tk.BooleanVar(value=True)
volume_control_check = ttk.Checkbutton(volume_frame, text="Enable Volume Control", variable=volume_control_var)
volume_control_check.grid(row=0, column=0, padx=5, pady=5)

# Default volume
ttk.Label(volume_frame, text="Default Volume (%):").grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
default_volume_var = tk.IntVar(value=50)
default_volume_spin = ttk.Spinbox(volume_frame, from_=0, to=100, width=5, textvariable=default_volume_var)
default_volume_spin.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

# Maximum volume
ttk.Label(volume_frame, text="Maximum Volume (%):").grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
max_volume_var = tk.IntVar(value=100)
max_volume_spin = ttk.Spinbox(volume_frame, from_=0, to=100, width=5, textvariable=max_volume_var)
max_volume_spin.grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)

# Data logging settings
logging_frame = ttk.LabelFrame(settings_tab, text="Data Logging Settings", padding=10)
logging_frame.pack(fill=tk.X, padx=5, pady=5)

# Enable logging
logging_var = tk.BooleanVar(value=False)
logging_check = ttk.Checkbutton(logging_frame, text="Enable Data Logging", variable=logging_var)
logging_check.grid(row=0, column=0, padx=5, pady=5)

# Logging interval
ttk.Label(logging_frame, text="Logging Interval (s):").grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
logging_interval_var = tk.DoubleVar(value=5.0)
logging_interval_spin = ttk.Spinbox(logging_frame, from_=1.0, to=60.0, increment=1.0, width=5, 
                                   textvariable=logging_interval_var)
logging_interval_spin.grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)

# Log file path
ttk.Label(logging_frame, text="Log File:").grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
log_file_var = tk.StringVar(value="noise_log.csv")
log_file_entry = ttk.Entry(logging_frame, textvariable=log_file_var, width=20)
log_file_entry.grid(row=0, column=4, sticky=tk.W, padx=5, pady=5)

# Browse button
browse_button = ttk.Button(logging_frame, text="...", width=3,
                          command=lambda: log_file_var.set(filedialog.asksaveasfilename(
                              defaultextension=".csv", filetypes=[("CSV files", "*.csv")])))
browse_button.grid(row=0, column=5, padx=5, pady=5)

# Log tab content
log_frame = ttk.Frame(log_tab, padding=10)
log_frame.pack(fill=tk.BOTH, expand=True)

# Log text box
log_text = tk.Text(log_frame, wrap=tk.WORD, width=80, height=20)
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Scrollbar for log
log_scrollbar = ttk.Scrollbar(log_frame, command=log_text.yview)
log_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_text.config(yscrollcommand=log_scrollbar.set)

# Clear log button
clear_log_button = ttk.Button(log_tab, text="Clear Log", 
                             command=lambda: log_text.delete(1.0, tk.END))
clear_log_button.pack(side=tk.BOTTOM, pady=5)

# Presets tab content
presets_frame = ttk.LabelFrame(presets_tab, text="Saved Presets", padding=10)
presets_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

# Presets listbox
presets_listbox = tk.Listbox(presets_frame, height=10)
presets_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

# Presets buttons frame
presets_buttons_frame = ttk.Frame(presets_frame, padding=5)
presets_buttons_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)

# Preset name entry
preset_name_var = tk.StringVar()
ttk.Label(presets_buttons_frame, text="Preset Name:").pack(anchor=tk.W, pady=2)
preset_name_entry = ttk.Entry(presets_buttons_frame, textvariable=preset_name_var, width=20)
preset_name_entry.pack(fill=tk.X, pady=2)

# Preset buttons
save_preset_button = ttk.Button(presets_buttons_frame, text="Save Current Settings", 
                               command=lambda: save_preset())
save_preset_button.pack(fill=tk.X, pady=2)

load_preset_button = ttk.Button(presets_buttons_frame, text="Load Selected", 
                               command=lambda: load_preset())
load_preset_button.pack(fill=tk.X, pady=2)

delete_preset_button = ttk.Button(presets_buttons_frame, text="Delete Selected", 
                                 command=lambda: delete_preset())
delete_preset_button.pack(fill=tk.X, pady=2)

# Preset description frame
preset_desc_frame = ttk.LabelFrame(presets_tab, text="Preset Description", padding=10)
preset_desc_frame.pack(fill=tk.X, padx=5, pady=5)

preset_desc_text = tk.Text(preset_desc_frame, wrap=tk.WORD, width=60, height=5)
preset_desc_text.pack(fill=tk.BOTH, expand=True)

# About tab content
about_frame = ttk.Frame(about_tab, padding=20)
about_frame.pack(fill=tk.BOTH, expand=True)

# App info
about_text = """
Noise Level Monitor & Controller v2.0

This application monitors ambient noise levels using an ESP32 microcontroller 
with a microphone module and adjusts system volume accordingly. 

Features:
- Real-time noise monitoring and visualization
- Automatic volume adjustment based on ambient noise
- Customizable sensitivity and thresholds
- Alert system for excessive noise
- Data logging and export
- Preset configurations for different environments

Created by: Your Name/Organization
License: MIT License
"""

about_label = ttk.Label(about_frame, text=about_text, wraplength=600, justify=tk.LEFT)
about_label.pack(padx=20, pady=20)

# Function to process noise value with high sensitivity to changes
def process_noise(raw_value):
    # Get current settings
    sensitivity = sensitivity_var.get()
    min_threshold = min_var.get()
    max_threshold = max_var.get()
    
    # Clamp value to thresholds
    clamped = max(min_threshold, min(raw_value, max_threshold))
    
    # Auto-adjust range if enabled
    if auto_cal_var.get() and raw_value > 0:
        global noise_min, noise_max
        noise_min = min(noise_min, raw_value) if noise_min > 0 else raw_value
        noise_max = max(noise_max, raw_value)
        min_var.set(noise_min)
        max_var.set(noise_max)
    
    # Normalize to 0-100 scale
    range_size = max_threshold - min_threshold
    if range_size <= 0:
        normalized = 0
    else:
        normalized = ((clamped - min_threshold) / range_size) * 100
    
    # Apply non-linear sensitivity curve (higher sensitivity = more dramatic response)
    # Using an exponential curve for sensitivity
    enhanced = (normalized / 100) ** (1 / sensitivity) * 100
    
    return enhanced

# Check if threshold is crossed
def check_threshold(processed_value):
    global threshold_crossed, threshold_time
    
    if not alert_enabled_var.get():
        return False
        
    threshold = alert_threshold_var.get()
    required_duration = alert_duration_var.get()
    
    if processed_value > threshold:
        if not threshold_crossed:
            threshold_crossed = True
            threshold_time = time.time()
            add_to_log(f"Threshold exceeded: {processed_value:.1f} > {threshold}")
        elif time.time() - threshold_time > required_duration:
            # Alert has been active long enough to trigger
            update_status_indicator("red")
            if sound_alert_var.get() and alert_sound:
                alert_sound.play()
            add_to_log(f"ALERT: Noise level {processed_value:.1f} exceeded threshold {threshold} for {required_duration}s")
            return True
    else:
        if threshold_crossed:
            threshold_crossed = False
            update_status_indicator("green")
            add_to_log(f"Noise level returned below threshold: {processed_value:.1f} < {threshold}")
    
    return False

# Read Serial Data
def read_serial():
    global running, noise_history
    
    last_log_time = time.time()
    
    try:
        ser = serial.Serial(com_port_var.get(), baud_rate_var.get(), timeout=1)
        status_var.set(f"Connected to {com_port_var.get()} at {baud_rate_var.get()} baud")
        add_to_log(f"Connected to {com_port_var.get()} at {baud_rate_var.get()} baud")
        time.sleep(2)  # Wait for connection to stabilize
        
        # Set initial volume
        if volume and volume_control_var.get():
            volume.SetMasterVolumeLevelScalar(default_volume_var.get() / 100, None)
        
        # For calculating a moving average
        value_buffer = []
        buffer_size = 5  # Increased buffer size for better smoothing
        
        while running:
            try:
                if ser.in_waiting > 0:
                    line = ser.readline().decode('utf-8', errors='replace').strip()
                    
                    # Try to extract noise value from various possible formats
                    try:
                        # Simple attempt to parse the raw value
                        noise_value = None
                        
                        # Try parsing numbers directly from the line
                        if line.isdigit():
                            # Just a number
                            noise_value = int(line)
                        elif "Noise Level:" in line:
                            # Format: "Raw Noise Level: X | Mapped Volume: Y"
                            parts = line.split('|')[0].split(':')
                            if len(parts) >= 2:
                                noise_part = parts[1].strip()
                                if noise_part.isdigit():
                                    noise_value = int(noise_part)
                        
                        # If we successfully extracted a value
                        if noise_value is not None:
                            # Add to buffer for smoothing
                            value_buffer.append(noise_value)
                            if len(value_buffer) > buffer_size:
                                value_buffer.pop(0)
                                
                            # Calculate smoothed value
                            smoothed_value = sum(value_buffer) / len(value_buffer)
                            
                            # Process with sensitivity adjustment
                            processed_value = process_noise(smoothed_value)
                            
                            # Check for threshold crossing
                            is_alert = check_threshold(processed_value)
                            
                            # Calculate volume level (0-100)
                            volume_level = min(int(processed_value), max_volume_var.get())
                            
                            # Update volume if not in alert state
                            if volume and volume_control_var.get() and not is_alert:
                                try:
                                    # Set volume (0.0 to 1.0)
                                    volume.SetMasterVolumeLevelScalar(volume_level / 100, None)
                                except Exception as e:
                                    print(f"Volume error: {e}")
                            
                            # Update history for graph
                            timestamp = time.time()
                            noise_history.append((timestamp, noise_value, processed_value, volume_level))
                            
                            # Trim history to maximum length
                            if len(noise_history) > max_history_length:
                                noise_history = noise_history[-max_history_length:]
                            
                            # Update UI (in a thread-safe way)
                            root.after(0, update_ui, noise_value, processed_value, volume_level)
                            
                            # Log data if enabled
                            if logging_var.get() and (timestamp - last_log_time) >= logging_interval_var.get():
                                log_data(timestamp, noise_value, processed_value, volume_level)
                                last_log_time = timestamp
                            
                        else:
                            print(f"Invalid data received: {line}")
                            
                    except (ValueError, IndexError) as e:
                        print(f"Invalid data received: {line}")
                        
            except Exception as e:
                print(f"Error processing data: {e}")
                time.sleep(0.5)  # Wait before trying again
                    
            # Brief pause to prevent CPU overload
            time.sleep(0.01)
            
    except Exception as e:
        status_var.set(f"Error: {e}")
        print(f"Serial error: {e}")
        add_to_log(f"Connection error: {e}")
        messagebox.showerror("Connection Error", f"Failed to connect to {com_port_var.get()}: {e}")
    finally:
        if 'ser' in locals() and ser.is_open:
            ser.close()
            print("Serial connection closed")
            add_to_log("Serial connection closed")

# Update status indicator color
def update_status_indicator(color):
    status_indicator.configure(bg=color)

# Add entry to log
def add_to_log(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = f"[{timestamp}] {message}\n"
    log_text.insert(tk.END, log_entry)
    log_text.see(tk.END)  # Scroll to see the latest entry

# Log data to CSV file
def log_data(timestamp, raw_value, processed_value, volume_level):
    try:
        file_exists = os.path.isfile(log_file_var.get())
        
        with open(log_file_var.get(), 'a', newline='') as file:
            writer = csv.writer(file)
            
            # Write header if file is new
            if not file_exists:
                writer.writerow(['Timestamp', 'ISO DateTime', 'Raw Noise', 'Processed Noise', 'Volume Level'])
            
            # Write data
            dt_string = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
            writer.writerow([timestamp, dt_string, raw_value, processed_value, volume_level])
            
    except Exception as e:
        print(f"Error logging data: {e}")
        add_to_log(f"Error logging data: {e}")

# Update UI with current values
def update_ui(raw_value, processed_value, volume_level):
    # Update text displays
    raw_value_var.set(str(int(raw_value)))
    processed_var.set(f"{processed_value:.1f}")
    volume_var.set(f"{volume_level}%")
    
    # Update progress bar
    volume_bar["value"] = volume_level
    
    # Update graph - but not on every update to avoid performance issues
    root.after(500, update_graph)  # Update graph every 500ms
    
    # Update status
    status_var.set(f"Running - Raw: {int(raw_value)}, Processed: {processed_value:.1f}, Volume: {volume_level}%")

# Update the graph
def update_graph():
    if not noise_history:
        return
        
    # Get the data for the last 30 seconds
    cutoff_time = time.time() - 30
    relevant_history = [x for x in noise_history if x[0] >= cutoff_time]
    
    if not relevant_history:
        return
        
    # Extract data
    times = [x[0] for x in relevant_history]
    base_time = times[0]  # Normalize times to start from 0
    times = [t - base_time for t in times]
    
    noise_values = [x[2] for x in relevant_history]  # Use processed noise values
    volume_values = [x[3] for x in relevant_history]  # Volume levels
    
    # Clear previous plots
    ax.clear()
    
    # Plot data
    ax.plot(times, noise_values, label='Noise Level', color='blue')
    ax.plot(times, volume_values, label='System Volume', color='red')
    
    # Plot threshold line
    threshold = alert_threshold_var.get()
    ax.axhline(y=threshold, color='green', linestyle='--', label='Alert Threshold')
    
    # Set labels and limits
    ax.set_title('Noise and Volume Over Time')
    ax.set_xlabel('Time (s)')
    ax.set_ylabel('Level')
    ax.set_ylim(0, 100)
    ax.legend()
    ax.grid(True)
    
    # Redraw canvas
    canvas.draw()

# Save current configuration
def save_config():
    config = {
        'com_port': com_port_var.get(),
        'baud_rate': baud_rate_var.get(),
        'sensitivity': sensitivity_var.get(),
        'min_threshold': min_var.get(),
        'max_threshold': max_var.get(),
        'auto_calibrate': auto_cal_var.get(),
        'alert_threshold': alert_threshold_var.get(),
        'alert_duration': alert_duration_var.get(),
        'alert_enabled': alert_enabled_var.get(),
        'sound_alert': sound_alert_var.get(),
        'volume_control': volume_control_var.get(),
        'default_volume': default_volume_var.get(),
        'max_volume': max_volume_var.get(),
        'logging_enabled': logging_var.get(),
        'logging_interval': logging_interval_var.get(),
        'log_file': log_file_var.get()
    }
    
    try:
        with open(config_file, 'w') as f:
            json.dump(config, f, indent=4)
        status_var.set("Configuration saved successfully")
        add_to_log("Configuration saved to " + config_file)
    except Exception as e:
        status_var.set(f"Error saving configuration: {e}")
        add_to_log(f"Error saving configuration: {e}")

# Load configuration
def load_config():
    try:
        if os.path.exists(config_file):
            with open(config_file, 'r') as f:
                config = json.load(f)
            
            # Apply loaded settings to variables
            com_port_var.set(config.get('com_port', SERIAL_PORT))
            baud_rate_var.set(config.get('baud_rate', BAUD_RATE))
            sensitivity_var.set(config.get('sensitivity', 3.0))
            min_var.set(config.get('min_threshold', 0))
            max_var.set(config.get('max_threshold', 3000))
            auto_cal_var.set(config.get('auto_calibrate', True))
            alert_threshold_var.set(config.get('alert_threshold', 80))
            alert_duration_var.set(config.get('alert_duration', 3.0))
            alert_enabled_var.set(config.get('alert_enabled', True))
            sound_alert_var.set(config.get('sound_alert', True))
            volume_control_var.set(config.get('volume_control', True))
            default_volume_var.set(config.get('default_volume', 50))
            max_volume_var.set(config.get('max_volume', 100))
            logging_var.set(config.get('logging_enabled', False))
            logging_interval_var.set(config.get('logging_interval', 5.0))
            log_file_var.set(config.get('log_file', 'noise_log.csv'))
            
            update_sensitivity(sensitivity_var.get())
            status_var.set("Configuration loaded successfully")
            add_to_log("Configuration loaded from " + config_file)
        else:
            status_var.set("No configuration file found")
            add_to_log("No configuration file found at " + config_file)
    except Exception as e:
        status_var.set(f"Error loading configuration: {e}")
        add_to_log(f"Error loading configuration: {e}")

# Export data to CSV
def export_data():
    if not noise_history:
        messagebox.showwarning("Export Data", "No data to export")
        return
        
    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="noise_export.csv"
        )
        
        if not file_path:
            return
            
        with open(file_path, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Timestamp', 'ISO DateTime', 'Raw Noise', 'Processed Noise', 'Volume Level'])
            
            for entry in noise_history:
                timestamp, raw, processed, volume_level = entry
                dt_string = datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
                writer.writerow([timestamp, dt_string, raw, processed, volume_level])
                
        status_var.set(f"Data exported to {file_path}")
        add_to_log(f"Data exported to {file_path}")
        messagebox.showinfo("Export Data", f"Data successfully exported to {file_path}")
    except Exception as e:
        status_var.set(f"Error exporting data: {e}")
        add_to_log(f"Error exporting data: {e}")
        messagebox.showerror("Export Error", f"Failed to export data: {e}")

# Restart serial connection
def restart_serial_connection():
    global running
    
    # Stop current thread if it's running
    running = False
    time.sleep(1)  # Give time for thread to finish
    
    # Start new thread
    running = True
    serial_thread = threading.Thread(target=read_serial, daemon=True)
    serial_thread.start()
    
    status_var.set(f"Connecting to {com_port_var.get()} at {baud_rate_var.get()} baud...")

# Find available COM ports
def find_com_ports():
    import serial.tools.list_ports
    
    ports = list(serial.tools.list_ports.comports())
    port_list = [port.device for port in ports]
    
    if not port_list:
        messagebox.showinfo("COM Ports", "No COM ports found")
        return
        
    # Update the combo box
    com_port_entry['values'] = port_list
    
    # Display the ports
    port_info = "Available COM ports:\n\n"
    for port in ports:
        port_info += f"{port.device}: {port.description}\n"
        
    messagebox.showinfo("COM Ports", port_info)

# Start calibration procedure
def start_calibration():
    # Reset min/max
    global noise_min, noise_max
    noise_min = 0
    noise_max = 100
    
    # Create calibration window
    cal_window = tk.Toplevel(root)
    cal_window.title("Calibration")
    cal_window.geometry("400x300")
    cal_window.transient(root)
    
    # Instructions
    ttk.Label(cal_window, text="Calibration Procedure", font=("Arial", 12, "bold")).pack(pady=10)
    ttk.Label(cal_window, text="1. Remain quiet for 5 seconds to measure ambient noise").pack(anchor=tk.W, padx=20)
    ttk.Label(cal_window, text="2. Make loud noises for 5 seconds to measure maximum levels").pack(anchor=tk.W, padx=20)
    ttk.Label(cal_window, text="3. System will automatically set appropriate thresholds").pack(anchor=tk.W, padx=20)
    
    # Progress bar
    progress_var = tk.DoubleVar()
    progress = ttk.Progressbar(cal_window, variable=progress_var, maximum=100)
    progress.pack(fill=tk.X, padx=20, pady=20)
    
    # Status display
    status_label = ttk.Label(cal_window, text="Ready to start calibration")
    status_label.pack(pady=10)
    
    # Current value display
    value_var = tk.StringVar(value="0")
    value_label = ttk.Label(cal_window, textvariable=value_var, font=("Arial", 14))
    value_label.pack(pady=10)
    
    # Cancel/Start/Done buttons
    button_frame = ttk.Frame(cal_window)
    button_frame.pack(side=tk.BOTTOM, pady=20)
    
    cancel_button = ttk.Button(button_frame, text="Cancel", command=cal_window.destroy)
    cancel_button.pack(side=tk.LEFT, padx=10)
    
    start_button = ttk.Button(button_frame, text="Start Calibration", 
                            command=lambda: start_cal_process())
    start_button.pack(side=tk.LEFT, padx=10)
    
    done_button = ttk.Button(button_frame, text="Done", command=cal_window.destroy)
    done_button.pack(side=tk.LEFT, padx=10)
    done_button.config(state=tk.DISABLED)
    
    # Calibration process
    def start_cal_process():
        # Reset values
        global noise_min, noise_max
        noise_min = 999999
        noise_max = 0
        
        # Disable start button
        start_button.config(state=tk.DISABLED)
        
        # Phase 1: Quiet
        status_label.config(text="Phase 1: Please remain quiet...")
        
        # Monitor progress and values for 5 seconds
        start_time = time.time()
        
        def update_cal_progress():
            current_time = time.time()
            elapsed = current_time - start_time
            
            # Get recent noise values (last 3 entries)
            recent_values = [x[1] for x in noise_history[-3:]] if noise_history else [0]
            current_value = sum(recent_values) / len(recent_values) if recent_values else 0
            
            # Update display
            value_var.set(f"Current value: {int(current_value)}")
            
            if elapsed < 5:  # Phase 1 (quiet)
                progress_var.set(elapsed * 10)  # 0-50%
                cal_window.after(100, update_cal_progress)
            elif elapsed < 10:  # Phase 2 (loud)
                if elapsed == 5:
                    status_label.config(text="Phase 2: Please make loud noises...")
                progress_var.set(50 + (elapsed - 5) * 10)  # 50-100%
                cal_window.after(100, update_cal_progress)
            else:
                # Calibration complete
                progress_var.set(100)
                status_label.config(text="Calibration complete!")
                
                # Set thresholds based on measurements
                if noise_min < 999999 and noise_max > 0:
                    # Add some margin
                    adjusted_min = max(0, noise_min - int(noise_min * 0.1))
                    adjusted_max = noise_max + int(noise_max * 0.1)
                    
                    min_var.set(adjusted_min)
                    max_var.set(adjusted_max)
                    
                    # Set alert threshold to 80% of the range
                    threshold_value = adjusted_min + int((adjusted_max - adjusted_min) * 0.8)
                    alert_threshold_var.set(threshold_value)
                    
                    value_var.set(f"Min: {adjusted_min}, Max: {adjusted_max}, Threshold: {threshold_value}")
                    add_to_log(f"Calibration completed: Min={adjusted_min}, Max={adjusted_max}, Threshold={threshold_value}")
                else:
                    status_label.config(text="Calibration failed, insufficient data")
                    add_to_log("Calibration failed: insufficient data")
                
                # Enable done button
                done_button.config(state=tk.NORMAL)
        
        # Start progress updates
        update_cal_progress()

# Save current settings as a preset
def save_preset():
    preset_name = preset_name_var.get().strip()
    if not preset_name:
        messagebox.showwarning("Save Preset", "Please enter a name for the preset")
        return
        
    # Get description
    preset_desc = preset_desc_text.get(1.0, tk.END).strip()
    
    # Create preset data
    preset = {
        'name': preset_name,
        'description': preset_desc,
        'date_created': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'settings': {
            'sensitivity': sensitivity_var.get(),
            'min_threshold': min_var.get(),
            'max_threshold': max_var.get(),
            'alert_threshold': alert_threshold_var.get(),
            'alert_duration': alert_duration_var.get(),
            'default_volume': default_volume_var.get(),
            'max_volume': max_volume_var.get()
        }
    }
    
    # Load existing presets
    presets = []
    presets_file = "presets.json"
    
    try:
        if os.path.exists(presets_file):
            with open(presets_file, 'r') as f:
                presets = json.load(f)
    except Exception:
        presets = []
    
    # Check if preset name already exists
    for i, p in enumerate(presets):
        if p.get('name') == preset_name:
            # Ask for confirmation to overwrite
            if messagebox.askyesno("Overwrite Preset", f"Preset '{preset_name}' already exists. Overwrite?"):
                presets[i] = preset
                break
            else:
                return
    else:
        # New preset
        presets.append(preset)
    
    # Save presets
    try:
        with open(presets_file, 'w') as f:
            json.dump(presets, f, indent=4)
        
        # Update listbox
        update_preset_list()
        
        status_var.set(f"Preset '{preset_name}' saved successfully")
        add_to_log(f"Preset '{preset_name}' saved")
    except Exception as e:
        status_var.set(f"Error saving preset: {e}")
        add_to_log(f"Error saving preset: {e}")

# Load selected preset
def load_preset():
    selection = presets_listbox.curselection()
    if not selection:
        messagebox.showwarning("Load Preset", "Please select a preset to load")
        return
    
    # Get the selected preset name
    preset_name = presets_listbox.get(selection[0])
    
    # Load presets
    presets_file = "presets.json"
    try:
        if os.path.exists(presets_file):
            with open(presets_file, 'r') as f:
                presets = json.load(f)
                
            # Find the matching preset
            preset = next((p for p in presets if p.get('name') == preset_name), None)
            
            if preset and 'settings' in preset:
                # Apply settings
                settings = preset['settings']
                sensitivity_var.set(settings.get('sensitivity', 3.0))
                min_var.set(settings.get('min_threshold', 0))
                max_var.set(settings.get('max_threshold', 3000))
                alert_threshold_var.set(settings.get('alert_threshold', 80))
                alert_duration_var.set(settings.get('alert_duration', 3.0))
                default_volume_var.set(settings.get('default_volume', 50))
                max_volume_var.set(settings.get('max_volume', 100))
                
                # Update sensitivity display
                update_sensitivity(sensitivity_var.get())
                
                # Show description
                preset_desc_text.delete(1.0, tk.END)
                preset_desc_text.insert(tk.END, preset.get('description', ''))
                
                status_var.set(f"Preset '{preset_name}' loaded successfully")
                add_to_log(f"Preset '{preset_name}' loaded")
            else:
                messagebox.showerror("Load Error", f"Invalid preset structure for '{preset_name}'")
    except Exception as e:
        status_var.set(f"Error loading preset: {e}")
        add_to_log(f"Error loading preset: {e}")

# Delete selected preset
def delete_preset():
    selection = presets_listbox.curselection()
    if not selection:
        messagebox.showwarning("Delete Preset", "Please select a preset to delete")
        return
    
    # Get the selected preset name
    preset_name = presets_listbox.get(selection[0])
    
    # Confirm deletion
    if not messagebox.askyesno("Delete Preset", f"Are you sure you want to delete preset '{preset_name}'?"):
        return
    
    # Load presets
    presets_file = "presets.json"
    try:
        if os.path.exists(presets_file):
            with open(presets_file, 'r') as f:
                presets = json.load(f)
                
            # Remove the matching preset
            presets = [p for p in presets if p.get('name') != preset_name]
            
            # Save updated presets
            with open(presets_file, 'w') as f:
                json.dump(presets, f, indent=4)
            
            # Update listbox
            update_preset_list()
            
            # Clear description
            preset_desc_text.delete(1.0, tk.END)
            
            status_var.set(f"Preset '{preset_name}' deleted successfully")
            add_to_log(f"Preset '{preset_name}' deleted")
    except Exception as e:
        status_var.set(f"Error deleting preset: {e}")
        add_to_log(f"Error deleting preset: {e}")

# Update the preset list
def update_preset_list():
    presets_listbox.delete(0, tk.END)
    
    presets_file = "presets.json"
    try:
        if os.path.exists(presets_file):
            with open(presets_file, 'r') as f:
                presets = json.load(f)
                
            # Add presets to listbox
            for preset in presets:
                presets_listbox.insert(tk.END, preset.get('name', 'Unnamed'))
    except Exception as e:
        print(f"Error loading presets: {e}")

# Show preset description when selected
def on_preset_select(event):
    selection = presets_listbox.curselection()
    if not selection:
        return
    
    # Get the selected preset name
    preset_name = presets_listbox.get(selection[0])
    
    # Load presets
    presets_file = "presets.json"
    try:
        if os.path.exists(presets_file):
            with open(presets_file, 'r') as f:
                presets = json.load(f)
                
            # Find the matching preset
            preset = next((p for p in presets if p.get('name') == preset_name), None)
            
            if preset:
                # Show description
                preset_desc_text.delete(1.0, tk.END)
                preset_desc_text.insert(tk.END, preset.get('description', ''))
                
                # Set name
                preset_name_var.set(preset_name)
    except Exception as e:
        print(f"Error loading preset description: {e}")

# Bind preset selection event
presets_listbox.bind('<<ListboxSelect>>', on_preset_select)

# Handle window closing
def on_closing():
    global running
    running = False
    time.sleep(0.5)  # Give threads time to cleanup
    root.destroy()

# Load config on startup
if os.path.exists(config_file):
    load_config()

# Update preset list on startup
update_preset_list()

# Start serial connection thread
serial_thread = threading.Thread(target=read_serial, daemon=True)
serial_thread.start()

# Set window close handler
root.protocol("WM_DELETE_WINDOW", on_closing)

# Start the main event loop
root.mainloop()