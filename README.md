# 🔊 Noise Level Monitor & Controller

A real-time, Python-powered application to tame the chaos of your acoustic jungle. Designed for classrooms, offices, studios, or anywhere silence is golden (or noise is a monster). This project lets you *monitor*, *visualize*, and *control* your system’s audio output based on environmental noise levels. Because your ears deserve better.

---

## 🧠 Project Overview

This project is a comprehensive **Noise Level Monitor & Controller** application, developed entirely in **Python** and designed for **real-time acoustic environment management**.

Leveraging an **ESP32 microcontroller** with a **microphone sensor**, the application continuously receives and analyzes ambient noise data via a **serial connection**.

It features a modern **graphical user interface (GUI)** built with **Tkinter**, offering live visualization of noise and system volume trends through **embedded Matplotlib plots**.

The system can automatically adjust the computer’s **audio output** in response to changing noise levels, with **customizable sensitivity**, **thresholds**, and **calibration routines** for different environments.

Users benefit from advanced features such as:

- Configurable alerts
- Data logging and export options
- Preset management

This makes the tool adaptable for **classrooms, offices, studios**, or any setting where **sound control is essential**.

The application integrates:

- Hardware interfacing
- Real-time data processing
- User-friendly controls
- System-level audio management

All within a single, robust Python program.

---

## 🎯 Features

- 📡 Real-time noise monitoring via ESP32 + microphone
- 📈 Live noise and system volume graphs (Tkinter + Matplotlib)
- 🔧 Auto-adjusts system volume based on detected noise levels
- ⚙️ Customizable sensitivity, thresholds, and environment presets
- 🚨 Alert system for excessive noise
- 📝 Data logging & CSV export support
- 🎛️ Preset management for fast environment switching
- 🔌 USB Serial communication with ESP32
- 🖥️ System-level audio control via Pycaw (Windows)

---

## 🛠️ Tech Stack

- **Python** (core application logic)
- **Tkinter** (GUI framework)
- **Matplotlib** (live data plots)
- **PySerial** (serial communication)
- **Pycaw** (audio control on Windows)
- **ESP32** with microphone sensor (hardware)

---

## 🧪 How It Works

1. **ESP32** reads noise levels from a mic sensor.
2. Sends values over serial to the Python app.
3. The GUI displays real-time graphs of noise and volume.
4. The system volume adjusts automatically based on noise input.
5. Alerts are triggered if noise breaches thresholds.
6. Data is logged for future analysis or compliance.

---

## 📦 Installation

### Clone the Repo
```bash
git clone https://github.com/your-username/noise-monitor-controller.git
cd noise-monitor-controller
```

### Install Requirements
```bash
pip install -r requirements.txt
```

### Run the App
Make sure your ESP32 is connected and run:
```bash
python app.py
```

---

## 🧩 Calibration & Presets

Navigate to the **Settings** panel within the GUI to:

- Adjust noise sensitivity
- Set custom thresholds
- Create room-specific presets (e.g., “Classroom,” “Studio,” “Library”)
- Recalibrate based on ambient levels

---

## 🌍 Ideal Use Cases

- 📚 **Classrooms** – Enforce noise discipline
- 🎧 **Studios** – Prevent recording disturbances
- 🧘 **Meditation Rooms** – Zen vibes only
- 🏢 **Open Offices** – Keep chatter in check

---

## 💡 Future Improvements

- 📲 Mobile App Companion (remote control!)
- ☁️ Cloud-based analytics and noise reporting
- 🧠 ML models to classify types of noise (speech, traffic, etc.)
- 🌐 Cross-platform system volume support (Linux & Mac)

---
