# Noise Level Monitor and Controller

A real-time noise-monitoring and acoustic-control application using Python, ESP32 sensor input, serial communication, and a desktop GUI.

The project is designed for environments such as classrooms, offices, libraries, and studios where noise levels need to be monitored or regulated.

---

## Problem statement

Ambient noise can affect concentration, comfort, communication, and productivity. Manual monitoring is unreliable, especially in shared spaces.

This project reads noise data from hardware sensors, visualizes it in real time, and supports threshold-based control or alerts.

---

## Features

- real-time noise-level monitoring
- ESP32 microcontroller integration
- microphone-sensor data input
- serial communication with Python
- Tkinter desktop GUI
- live Matplotlib plots
- threshold-based alerts
- CSV logging for historical analysis
- Windows system-volume control through Pycaw

---

## System overview

```text
Microphone Sensor
       ↓
ESP32
       ↓
Serial Communication
       ↓
Python Backend
       ↓
Tkinter Dashboard
       ↓
Visualization + Alerts + Control
```

---

## Quick start

```bash
git clone https://github.com/Mrudula-itsjuzme/noise-regulation.git
cd noise-regulation

pip install -r requirements.txt
python app.py
```

Before running, connect the ESP32 through USB and confirm the correct serial port is configured in the application.

---

## Calibration

Use the settings panel to configure:

- noise sensitivity
- custom threshold triggers
- room presets such as library, studio, or classroom
- recalibration based on current ambient noise

---

## Tech stack

- Python
- ESP32
- PySerial
- Tkinter
- Matplotlib
- Pycaw
- CSV logging

---

## Future improvements

- add screenshots of the GUI
- add ESP32 wiring diagram
- support Linux and macOS audio control
- add cloud-based long-term reporting
- classify noise types such as speech, traffic, and background hum
- add mobile dashboard support

---

## Author

Built by [Pedamallu Sai Mrudula](https://github.com/Mrudula-itsjuzme) as part of an IoT, automation, and environmental-monitoring portfolio.
