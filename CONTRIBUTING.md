# Contributing to KMS Activator

Thanks for taking the time to contribute! This project is a WPF (PowerShell)
GUI with a small C# launcher for activating Windows and Office against a
**KMS host you control**.

## Ground rules

- Be respectful — see [CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md).
- This tool is intended for **legitimate volume-activation** against a KMS host
  you are authorized to use. Please do not open issues/PRs asking for help with
  piracy or with evading antivirus/SmartScreen detection.
- Report security issues privately — see [SECURITY.md](SECURITY.md).

## Project layout

| Path | Purpose |
|------|---------|
| `KMS_Activator_GUI.ps1` | App entry point (loads modules, builds the window) |
| `Modules/` | Activation, Office install, edition change, UI helpers, updater |
| `Resources/` | Global config, themes, languages, Office config |
| `UI/MainWindow.xaml` | WPF layout |
| `locales/*.json` | Translations (en, fr, de, es, it) |
| `src/Launcher.cs` | C# launcher that extracts the payload and runs the GUI |
| `Compile.ps1` / `Compile-CI.ps1` | Build the single `KMS_Activator.exe` |

## Branch & PR flow

- `main` — released code. `dev` — integration branch.
- Branch from `dev`, keep changes focused, open a PR into `dev`.
- Describe **what** changed and **why**; link any related issue.
- Make sure the build works (`Compile-CI.ps1`) before requesting review.

## Coding style

- Follow the existing style and the rules in [.editorconfig](.editorconfig).
- Keep PowerShell functions small and add short comments for non-obvious logic.
- Any new user-facing string must be added to **all** `locales/*.json` (English
  is the fallback).

## Building locally (Windows)

```powershell
# Requires .NET Framework (csc) and PowerShell 5.1+
./Compile.ps1        # local build -> Build/KMS_Activator.exe
```

The release build (`Compile-CI.ps1`) additionally signs the executable; see
[CODE_SIGNING.md](CODE_SIGNING.md).
