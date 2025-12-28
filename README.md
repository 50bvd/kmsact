# KMS Activator

A modern, feature-rich Windows and Office activation tool with a beautiful WPF interface.

## Features

- **Windows Activation** - Activate all Windows editions (Pro, Education, Enterprise)
- **Office Activation** - Activate Microsoft Office suites
- **Edition Changer** - Upgrade from Home/Core to Pro/Enterprise editions
- **Office Installer** - Install Office LTSC 2024 with custom components
- **Auto-Renewal** - Schedule automatic activation renewal every 4 weeks
- **Uninstall Keys** - Remove all product keys and reset activation
- **Multi-language** - Support for English, French, Spanish, German, Italian
- **Dark/Light Theme** - Modern WPF interface with theme support
- **Configurable** - Custom KMS server settings

## Installation

### Option 1: Download Release
1. Download the latest `KMS_Activator.exe` from [Releases](https://github.com/50bvd/kmscact/releases)
2. Run as Administrator
3. Done!

### Option 2: Build from Source
```powershell
# Clone repository
git clone https://github.com/50bvd/kmscact.git
cd kmscact

# Compile
.\Compile.ps1

# Run
.\Run.ps1
```

## Usage

### Activate Windows
1. Click **"Activate Windows"**
2. Wait for activation to complete
3. Done!

### Change Windows Edition
1. Click **"Change Edition"** in Tools menu
2. If you're on Home/Core, you'll be prompted to upgrade to Pro first
3. Select target edition (Pro/Education/Enterprise)
4. Wait for activation

### Install Office
1. Click **"Install Office"** in Tools menu
2. Select components (Word, Excel, PowerPoint, etc.)
3. Wait for installation
4. Office will be automatically activated

### Schedule Auto-Renewal
1. Click **"Schedule Auto-Renewal"** in Tools menu
2. A scheduled task will run every 4 weeks to renew activation

## Configuration

Click the **Settings** button to configure:
- **Language**: English, French, Spanish, German, Italian
- **Theme**: Light or Dark mode
- **KMS Server**: Custom KMS server address (default: kms.50bvd.com)

## Technical Details

- **Framework**: PowerShell with WPF (Windows Presentation Foundation)
- **Architecture**: Modular design with separate activation, UI, and installer modules
- **Compilation**: Uses PS2EXE for creating standalone executable
- **Requirements**: Windows 10/11, PowerShell 5.1+, .NET Framework 4.7.2+

## Project Structure

```
kmscact/
â”œâ”€â”€ assets/              # Icons and images
â”œâ”€â”€ locales/             # Translation files (JSON)
â”œâ”€â”€ Modules/             # Core functionality modules
â”‚   â”œâ”€â”€ ActivationCore.ps1
â”‚   â”œâ”€â”€ EditionChanger.ps1
â”‚   â”œâ”€â”€ MessageBoxHelper.ps1
â”‚   â”œâ”€â”€ OfficeInstaller.ps1
â”‚   â””â”€â”€ UIHelper.ps1
â”œâ”€â”€ Resources/           # Configuration and themes
â”œâ”€â”€ UI/                  # XAML interface definitions
â”œâ”€â”€ src/                 # C# launcher source
â”œâ”€â”€ Compile.ps1          # Build script
â”œâ”€â”€ KMS_Activator_GUI.ps1 # Main application
â””â”€â”€ Run.ps1              # Quick run script
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License.

## Disclaimer

This tool is for educational purposes only. Use at your own risk. The authors are not responsible for any misuse or damage caused by this software.

## Credits

- Original concept and development by [50bvd](https://github.com/50bvd)
- KMS activation technology

## Support

- Report bugs: [Issues](https://github.com/50bvd/kmscact/issues)
- Discussions: [Discussions](https://github.com/50bvd/kmscact/discussions)

---

If you find this project useful, please give it a star!

