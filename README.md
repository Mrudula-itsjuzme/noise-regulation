# Noise Level Monitor & Controller

A real-time application designed for environmental acoustic management in settings such as classrooms, offices, and studios. This system monitors, visualizes, and controls audio output based on detected ambient noise levels.

## Project Overview

The Noise Level Monitor & Controller is a Python-based system integrated with an ESP32 microcontroller. It receives real-time noise data from a microphone sensor and analyzes it to manage the acoustic environment effectively.

### Key Features

- **Real-time Monitoring**: Continuous tracking of noise levels via serial communication with an ESP32.
- **Data Visualization**: Live plotting of noise trends and system volume using Matplotlib and Tkinter.
- **Automated Control**: Dynamic adjustment of system audio output in response to environmental noise.
- **Threshold Management**: Configurable sensitivity and alert systems for maintaining desired noise levels.
- **Data Logging**: Export capabilities for historical noise data analysis in CSV format.
- **System Integration**: Low-level audio management (primarily for Windows via Pycaw).

## Technical Architecture

- **Hardware**: ESP32 microcontroller with a compatible microphone sensor.
- **Software Core**: Python 3 code handling data ingestion and logic.
- **User Interface**: Tkinter-based GUI for real-time interaction.
- **Serial Communication**: PySerial for robust data transfer between hardware and software.

## Installation

### Repository Setup
1. Clone the repository to your local machine:
   ```bash
   git clone https://github.com/Mrudula-itsjuzme/noise-monitor-controller.git
   ```
2. Navigate to the project directory:
   ```bash
   cd noise-monitor-controller
   ```

### Dependency Installation
Install the required Python packages:
```bash
pip install -r requirements.txt
```

### Execution
Connect the ESP32 via USB and run the application:
```bash
python app.py
```

## Application Calibration

Access the Settings panel in the GUI to:
- Define noise sensitivity levels.
- Set custom threshold triggers for alerts.
- Manage room-specific presets (e.g., Library, Studio, Classroom).
- Perform recalibration based on current ambient noise.

## Future Roadmap

- Development of a mobile companion application for remote monitoring.
- Implementation of cloud-based analytics for long-term noise reporting.
- Integration of machine learning models for noise classification (e.g., speech vs. background traffic).
- Expansion of system volume control support to Linux and macOS environments.
