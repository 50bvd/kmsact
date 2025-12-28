# Developer Guide

Technical documentation for MS KMS Activator v3.5.0

## Architecture

### Tech Stack
- **C# Launcher** (`src/Launcher.cs`) - Admin elevation + resource extraction
- **PowerShell** - Main application logic (WPF)
- **XAML** - User interface (`UI/MainWindow.xaml`)
- **JSON** - Localization files (`locales/*.json`)
- **.NET Framework 4.0** - Runtime

### Project Structure

```
kmsact/
├── src/
│   └── Launcher.cs              # C# entry point
├── locales/                     # JSON translations
│   ├── en.json
│   ├── fr.json
│   ├── es.json
│   ├── de.json
│   └── it.json
├── Modules/                     # PowerShell modules
│   ├── ActivationCore.ps1      # KMS activation logic
│   ├── EditionChanger.ps1      # Edition upgrade
│   ├── OfficeInstaller.ps1     # Office install
│   ├── UIHelper.ps1            # Theme switching
│   └── MessageBoxHelper.ps1    # Custom dialogs
├── Resources/                   # Configuration
│   ├── GlobalInit.ps1
│   ├── Themes.ps1
│   └── OfficeConfig.ps1
├── UI/
│   └── MainWindow.xaml          # Main window layout
├── Compile.ps1                  # Build script
├── Run.ps1                      # Dev runner
└── icon.ico                     # Application icon
```

## Building

### Local Build

```powershell
# Simple build
.\Compile.ps1

# Output: Build/KMS_Activator.exe
```

The `Compile.ps1` script:
1. Finds C# compiler (`csc.exe`)
2. Compiles `src/Launcher.cs`
3. Embeds PowerShell scripts + XAML + icon
4. Generates signed EXE

### CI/CD Build (GitHub Actions)

Workflow (`.github/workflows/build.yml`) auto-builds on tag push:
- Generates self-signed certificate
- Compiles EXE
- Signs executable
- Creates GitHub Release
- Uploads signed EXE

**Required secret**: `CERT_PASSWORD` in GitHub repository settings

## Development

### Running in Dev Mode

```powershell
# Option 1: With launcher simulation
.\Run.ps1

# Option 2: Direct PowerShell
powershell -ExecutionPolicy Bypass -File .\KMS_Activator_GUI.ps1
```

### Adding a New Language

1. Create `locales/xx.json`:
```json
{
  "name": "LanguageName",
  "activateWindows": "Translation...",
  ...
}
```

2. Update `Resources/Languages.ps1` to load JSON:
```powershell
$langData = Get-Content "locales\xx.json" | ConvertFrom-Json
$Global:Languages.xx = $langData
```

3. Add to Settings dropdown in `KMS_Activator_GUI.ps1`

### Code Style

**PowerShell**:
- Functions: `Verb-Noun` (e.g., `Enable-Windows`)
- Variables: `$camelCase` for local, `$Global:PascalCase` for global
- Indentation: 4 spaces

**C#**:
- PascalCase for classes, methods
- camelCase for variables
- Braces on new line

**XAML**:
- Element names: PascalCase
- One attribute per line for complex elements

## Testing

**Manual checklist**:
- [ ] Windows activation
- [ ] Office activation
- [ ] Edition upgrade
- [ ] Office installation
- [ ] Language switching (test JSON loading)
- [ ] Theme switching
- [ ] Settings persistence

## Contributing

1. Fork repository
2. Create feature branch (`git checkout -b feature/name`)
3. Commit changes
4. Push to branch
5. Open Pull Request

## Release Process

1. Update version in `KMS_Activator_GUI.ps1`
2. Update `CHANGELOG.md`
3. Commit changes
4. Create and push tag:
   ```bash
   git tag -a v3.5.0 -m "Release v3.5.0"
   git push origin v3.5.0
   ```
5. GitHub Actions automatically builds and creates release

**Questions?** [Open an issue](https://github.com/50bvd/kmsact/issues)
