# CHANGELOG v3.5

## Version 3.5.0 - 2025-12-28

### 🎨 Complete UI Overhaul
- **NEW**: Modern WPF interface with beautiful design
- **NEW**: Dark/Light theme support with automatic detection
- **NEW**: Responsive layout with proper window sizing
- **NEW**: Custom styled buttons with hover effects and animations
- **NEW**: Professional message boxes (no more ugly Windows native dialogs)
- **NEW**: Settings window with icon and proper theming

### 🌍 Full Internationalization (i18n)
- **NEW**: 5 languages supported: English, French, Spanish, German, Italian
- **NEW**: JSON-based translation system
- **NEW**: Automatic language detection from Windows
- **NEW**: Easy to add more languages

### 🏗️ Complete Architecture Refactor
- **NEW**: Modular structure with separate modules:
  - ActivationCore.ps1 - Windows/Office activation logic
  - EditionChanger.ps1 - Edition changing with WPF interface
  - MessageBoxHelper.ps1 - Styled WPF message boxes
  - OfficeInstaller.ps1 - Office installation with component selection
  - UIHelper.ps1 - Logging and UI utilities
- **NEW**: Resources folder for configurations and themes
- **NEW**: Separate UI folder for XAML definitions
- **NEW**: Better error handling and real-time logging

### 🔧 Edition Changer Improvements
- **NEW**: Beautiful WPF interface with colored edition cards (Pro/Education/Enterprise)
- **NEW**: Automatic Core/Home detection with upgrade prompt
- **NEW**: Upgrade Core/Home to Pro using changepk.exe
- **NEW**: Service configuration (License Manager, Windows Update)
- **NEW**: Proper icons and theming
- **NEW**: Fully translated interface

### 📦 Office Installer Enhancements
- **NEW**: Modern WPF interface with checkboxes
- **NEW**: Component selection (Word, Excel, PowerPoint, Outlook, Access, Visio, Publisher, OneNote)
- **NEW**: Custom icons for each component (36px)
- **NEW**: Existing Office detection and uninstall prompt
- **NEW**: Auto-activation after installation
- **NEW**: Fully translated

### ⚙️ Settings Window
- **NEW**: Dedicated settings window with WPF styling
- **NEW**: Language selector (5 languages)
- **NEW**: Theme selector (Light/Dark)
- **NEW**: KMS Server configuration
- **NEW**: Save/Cancel buttons with proper order
- **NEW**: key.ico icon in window

### 🛠️ Technical Improvements
- **IMPROVED**: Real-time command output with UTF-8 encoding
- **IMPROVED**: Better process handling with exit codes
- **IMPROVED**: Responsive UI during long operations
- **IMPROVED**: Proper file paths and resource management
- **NEW**: GitHub Actions workflow for automatic builds
- **NEW**: Professional README with documentation
- **NEW**: Proper .gitignore
- **NEW**: Build directory with RELEASE_INFO.txt

### 🐛 Bug Fixes
- **FIXED**: MessageBox placeholders showing {0} instead of actual values
- **FIXED**: Settings button not working
- **FIXED**: Console not showing real-time output
- **FIXED**: Encoding issues with emojis
- **FIXED**: Double Get-String calls
- **FIXED**: All MessageBox now use WPF instead of native Windows dialogs
- **FIXED**: Edition changer format strings
- **FIXED**: Module loading issues

### 🗑️ Code Cleanup
- **REMOVED**: All temporary fix scripts
- **REMOVED**: Hardcoded text (100% translated)
- **REMOVED**: Unused functions
- **REMOVED**: Legacy code
- **CLEANED**: Proper UTF-8 encoding everywhere
- **CLEANED**: Consistent code style

### 📚 Documentation
- **NEW**: Professional README.md with features, usage, installation
- **NEW**: CHANGELOG.md (this file)
- **NEW**: Inline code comments
- **NEW**: Developer documentation

### 🚀 Build & Deployment
- **NEW**: GitHub Actions workflow for automatic releases
- **NEW**: Automatic .exe build on tag push
- **NEW**: Release artifacts (exe + info)
- **NEW**: Proper versioning

---

## Migration from v2.0

**Breaking Changes:**
- Complete rewrite - not compatible with v2.0 config
- New module structure
- New file organization

**What's Changed:**
- v2.0 was a single PowerShell script
- v3.5 is a full modular application with WPF UI
- v2.0 had basic CLI interface
- v3.5 has beautiful graphical interface with dark theme

**Upgrade Steps:**
1. Download v3.5 from Releases
2. Run as Administrator
3. All your settings will be fresh (no migration needed)

---

## Credits

- **Original v1.0-2.0**: Basic CLI activation script
- **v3.5 Complete Refactor**: Modern WPF application with full i18n
- **Developer**: 50bvd
- **Contributors**: Claude (Anthropic) for refactoring assistance

---

**Full Changelog**: https://github.com/50bvd/kmsact/compare/2.0...v3.5.0
