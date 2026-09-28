# 🔑 KMS Activator

[![Lint](https://github.com/50bvd/kmsact/actions/workflows/lint.yml/badge.svg?branch=main)](https://github.com/50bvd/kmsact/actions/workflows/lint.yml)
[![Release](https://img.shields.io/github/v/release/50bvd/kmsact?include_prereleases&sort=semver)](https://github.com/50bvd/kmsact/releases)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A modern Windows & Office volume-activation GUI for sysadmins — a clean WPF front end over the standard `slmgr` / `ospp.vbs` tooling, pointed at a KMS host you control.

![img](https://50bvd.com/assets/img/Capture%20d'%C3%A9cran%202025-12-28%20120017.png)

## ✨ Features

- 🪟 **Windows activation** — Pro, Education, Enterprise and more
- 📄 **Office activation** — Microsoft Office / LTSC suites
- ⬆️ **Edition changer** — upgrade Home/Core to Pro/Enterprise
- 📦 **Office installer** — install Office LTSC 2024 with the components you pick
- 🔄 **Auto-renewal** — scheduled task to re-activate every 4 weeks
- 🧹 **Uninstall keys** — remove product keys and reset activation
- 🌍 **Multi-language** — EN, FR, DE, ES, IT
- 🎨 **Light / Dark theme** — modern WPF interface
- 🔔 **Update check** — tells you when a newer version is available
- ⚙️ **Configurable KMS server** — target your own host

## 📥 Download

Get the latest version on the [releases page](https://github.com/50bvd/kmsact/releases/latest), or directly:

- **Windows**: [KMS_Activator.exe](https://github.com/50bvd/kmsact/releases/latest/download/KMS_Activator.exe)

Run it **as Administrator** (activation requires elevation). The app checks for updates and can open the download page when a newer version exists (**Tools › Check for Updates**).

## 🚀 Quick Start

1. Download `KMS_Activator.exe` from the [latest release](https://github.com/50bvd/kmsact/releases/latest)
2. Right-click → **Run as administrator**
3. Click **Activate Windows** (or **Activate Office**)
4. Done

> If Windows SmartScreen appears, click **More info → Run anyway** (the build is self-signed; see [Security](#-security)).

## 📖 Usage

### Activate Windows
1. Click **Activate Windows**
2. Wait for the sequence to finish

### Change Windows edition
1. **Tools → Change Edition**
2. From Home/Core you'll be prompted to upgrade to Pro first
3. Pick the target edition (Pro / Education / Enterprise)

### Install Office
1. **Tools → Install Office**
2. Select the components (Word, Excel, PowerPoint, …)
3. Office is activated automatically once installed

### Schedule auto-renewal
1. **Tools → Schedule Auto-Renewal**
2. A scheduled task renews activation every 4 weeks

## ⚙️ Configuration

Open **Settings** to configure:

- **Language** — EN, FR, DE, ES, IT
- **Theme** — Light or Dark
- **KMS server** — your KMS host address (default: `kms.50bvd.com`)

## 🛠️ Development

### Prerequisites
- Windows 10/11
- PowerShell 5.1+
- .NET Framework 4.7.2+ (provides `csc` for the launcher)

### Build
```powershell
git clone https://github.com/50bvd/kmsact.git
cd kmsact
./Compile.ps1     # builds Build/KMS_Activator.exe
./Run.ps1         # run from source
```

The CI release build (`Compile-CI.ps1`) additionally signs the executable — see [CODE_SIGNING.md](CODE_SIGNING.md).

## 🧱 Project structure

```
kmsact/
├── assets/                 # Icons and images
├── locales/                # Translations (JSON)
├── Modules/                # Core modules
│   ├── ActivationCore.ps1
│   ├── EditionChanger.ps1
│   ├── MessageBoxHelper.ps1
│   ├── OfficeInstaller.ps1
│   ├── UIHelper.ps1
│   └── Updater.ps1
├── Resources/              # Config, themes, languages
├── UI/                     # XAML interface
├── src/Launcher.cs         # C# launcher
├── KMS_Activator_GUI.ps1   # App entry point
├── Compile.ps1             # Build script
└── Run.ps1                 # Quick run
```

## 🤝 Contributing

Contributions are welcome — bug reports, translations and code.

- Read the [contribution guide](CONTRIBUTING.md). Pull requests target the **`dev`** branch; `main` only holds released versions.
- Test builds are produced by [Actions › Build & Release](https://github.com/50bvd/kmsact/actions/workflows/build-release.yml).
- Please follow the [Code of Conduct](CODE_OF_CONDUCT.md).
- See the [changelog](CHANGELOG.md) for what changed in each version.

## 🔒 Security

- Report vulnerabilities privately — see the [security policy](SECURITY.md), not a public issue.
- The app runs elevated; releases carry an Authenticode signature (self-signed until a trusted certificate is configured — see [CODE_SIGNING.md](CODE_SIGNING.md)). Verify the SHA-256 in `RELEASE_INFO.txt` against your download.

## 📄 License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

## 👤 Author

**Loup LIGNON KRASNIQI**
- GitHub: [@50bvd](https://github.com/50bvd)
- Email: loup.lk-pro@protonmail.ch

## ⚠️ Disclaimer

This tool provides a GUI over Microsoft's own volume-activation commands and is
intended for **legitimate activation against a KMS host you are authorized to
use**. You are responsible for holding the appropriate volume licenses. The
author is not responsible for misuse.

## 🙏 Acknowledgments

- Built with **PowerShell** and **WPF** (Windows Presentation Foundation)
- Packaged as a single executable via a small C# launcher

---

**⭐ Star this repo if you find it useful!**
