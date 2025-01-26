# 🎧 Audio Device Switcher

A powerful, kernel-mode Windows application that enables seamless switching between audio devices using customizable keyboard shortcuts.

[![Windows Support](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

## ✨ Key Features

- **Quick Device Switching** - Switch audio outputs instantly with hotkeys
- **Dual Device Categories** - Separate management for Speakers and Headphones
- **Customizable Shortcuts**
  - `Ctrl+Alt+S` - Switch between devices of same type
  - `Ctrl+Alt+T` - Toggle between Speakers and Headphones
- **Smart Features**
  - System tray integration for easy access
  - Per-application audio device rules
  - Automatic device switching based on active applications
  - Visual notifications for device changes
- **System Integration**
  - Kernel-mode audio switching for reliability
  - Auto-start with Windows option
  - Persistent configuration

## 🚀 Getting Started

### Prerequisites

- Windows 10/11
- Python 3.8 or higher
- Administrator privileges

### Installation

1. **Clone the repository**

```bash
git clone https://github.com/catalizcs/audio-switcher.git
cd audio-switcher
```

2. **Install dependencies**

```bash
pip install -r requirements.txt
```

3. **First Run**

- Right-click `start.bat` and select "Run as Administrator"
- Look for the icon in your system tray

### Initial Configuration

1. Right-click the tray icon
2. Configure your devices:
   - Under "Speakers" - Select your speaker outputs
   - Under "Headphones" - Select your headphone outputs
3. Test the default shortcuts:
   - `Ctrl+Alt+S` - Switch between devices of same type
   - `Ctrl+Alt+T` - Switch between Speakers/Headphones

## 🛠️ Configuration

### Auto-Start Setup

1. Press `Win+R`
2. Type `shell:startup`
3. Create shortcut to `start.bat`
4. Right-click shortcut → Properties → "Run as Administrator"

### Application-Specific Rules

You can configure different audio devices for specific applications:

1. Right-click tray icon → "Settings"
2. Add application rules:
   - Select application
   - Choose preferred audio device
   - Enable/disable auto-switching

## ⚙️ Advanced Features

### Kernel Mode

- Enhanced system integration for reliable switching
- Lower latency device changes
- Enable/disable in settings

### Debug Mode

- Detailed logging for troubleshooting
- Access logs via settings menu
- Performance monitoring

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🔧 Troubleshooting

### Common Issues

1. **No tray icon visible**

   - Ensure running as Administrator
   - Check system tray overflow menu

2. **Hotkeys not working**

   - Verify no conflicting applications
   - Check keyboard settings

3. **Device not listed**
   - Update audio drivers
   - Reconnect device

### Support

- Create an issue for bugs/features
- Check existing issues first
- Provide system details when reporting

## 🌟 Acknowledgments

- Built with Python and Windows Audio APIs
- Uses pycaw for audio control
