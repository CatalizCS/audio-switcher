# Audio Device Switcher

A kernel-mode system tray application that allows quick switching between audio devices using keyboard shortcuts.

## Features
- Separate management for Speakers and Headphones
- Two-level switching:
  - Switch between devices within same type (Ctrl+Alt+S)
  - Switch between Speakers and Headphones (Ctrl+Alt+T)
- System tray menu with device categories
- Persistent device configuration
- Kernel-mode audio switching
- Administrator privileges for low-level access

## Installation
1. Install requirements:
```bash
pip install -r requirements.txt
```

2. Add an icon.png file to the directory
3. Run the application with administrator privileges:
```bash
# Right-click and "Run as Administrator"
python audio_switcher.py
```

## Quick Start
1. Right-click `start.bat` and select "Run as Administrator"
2. Wait for installation and setup to complete
3. Look for the audio icon in your system tray
4. Right-click the icon to configure your devices

## First Time Setup
1. Right-click the tray icon
2. Under "Speakers", select your speaker devices
3. Under "Headphones", select your headphone devices
4. Use Ctrl+Alt+S to switch between devices of same type
5. Use Ctrl+Alt+T to switch between Speakers and Headphones

## Auto Start with Windows
1. Press Win+R
2. Type `shell:startup`
3. Create shortcut to start.bat in the opened folder
4. Right-click the shortcut and set "Run as Administrator"

## Usage
- Application must run with administrator privileges
- Right-click the tray icon to:
  - Configure Speaker devices
  - Configure Headphone devices
  - View current shortcuts
- Use Ctrl+Alt+S to switch between devices of same type
- Use Ctrl+Alt+T to switch between Speakers and Headphones
- Selected devices are marked with a checkmark in the menu

## Note
This application requires administrator privileges to access kernel-mode audio functions.
