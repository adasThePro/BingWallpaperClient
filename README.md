# Bing Wallpaper Client (BWC)

[![PowerShell](https://img.shields.io/badge/PowerShell-2.0+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Windows](https://img.shields.io/badge/Windows-7%2B-0078D6?style=flat&logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A lightweight CLI tool that automatically downloads and sets Bing's daily image as your desktop wallpaper. Built using **VBScript**, **PowerShell** and **C#**, BWC runs silently in the background to keep your desktop fresh with high-quality imagery from around the world.

>[!IMPORTANT]
> - Not affiliated with Microsoft or Bing.
> - Support for Linux is under development.
> - If you're using Windows 7, make sure to update to the latest version via [Legacy Update](https://legacyupdate.net/).

## Table of Contents
- [Features](#features)
- [Why BWC?](#why-bwc)
- [System Requirements](#system-requirements)
- [Installation](#quick-start)
- [Usage](#usage)
- [Configuration](#configuration)
- [Scheduled Task](#scheduled-task)
- [Project Structure](#project-structure)
- [Installation Structure](#installation-structure)
- [License](#license)

## Features

- **Automatic Daily Updates**: Downloads and sets Bing's image of the day as your wallpaper
- **Multi-Region Support**: Choose from 150+ Bing markets worldwide to get region-specific images
- **Smart Image Management**: Configurable image retention with automatic cleanup of old wallpapers
- **Scheduled Task**: Runs automatically via Windows Task Scheduler
- **Lightweight**: Pure VBScript implementation with minimal resource usage (~100 KB)
- **Zero Background Processes**: Only runs when scheduled, no constant RAM usage
- **No Dependencies**: Works on Windows 7 and later without additional software
- **Silent Operation**: Runs in the background without interrupting your workflow

## Why BWC?
<div align="center">

| Feature                 | Bing Wallpaper Client   | Official Bing Wallpaper App |
|-------------------------|-------------------------|-------------------------|
| **Size**                | ~100 KB                 | 200+ MB                 |
| **Telemetry**           | None                    | Yes                     |
| **Background Process**  | None                    | Always running          |
| **RAM Usage**           | 0 MB (when idle)        | 50-100 MB constantly    |
| **Customization**       | Full control via config | Limited                 |
| **Market Selection**    | 150+ markets, changeable| Locked to system locale |
| **Image Retention**     | Configurable (Unlimited / 1-1000 images) | Fixed  |

</div>

## System Requirements

- Windows 7 or later
- PowerShell 2.0 or later
- Internet connection

## Quick Start

1. Download the repository as a zip from [here](https://github.com/adasThePro/BingWallpaperClient/releases).
2. Press `Win + R`, type `powershell`, and press Enter to open PowerShell.
3. Navigate to the unzipped `bwc` directory:
```powershell
cd "$HOME\Desktop\bwc"
```
4. Run the installer script:
```powershell
# Temporarily bypass execution policy to run the installer
Set-ExecutionPolicy Bypass -Scope Process -Force
# Run the installer
.\installer.ps1
# To install without compiling SetWallpaper.exe (use existing one or if failed to compile)
./installer.ps1 -SkipCompile
```
5. Follow the interactive setup prompts to configure:
    - Market/region (e.g., en-US, ja-JP, fr-FR)
    - Download path for images (default: `%USERPROFILE%\Pictures\BingWallpapers`)
    - Maximum images to keep (0 for unlimited, default: 30)
    - Whether to automatically change desktop wallpaper

**The installer will:**
- Compile the `SetWallpaper.exe` component
- Copy application files to `%LOCALAPPDATA%\BingWallpaperClient`
- Create a scheduled task
- Add BWC to user's PATH environment variable
- Save your configuration

## Usage

1. Check status and configuration:
```powershell
bwc status
```

2. Manually apply wallpaper:
```powershell
bwc apply
```

3. Open configuration file:
```powershell
bwc config
```

4. View help or installation version:
```powershell
bwc help
bwc version
```

5. Uninstall BWC:
```powershell
bwc uninstall
```

## Configuration

Configuration is stored in `%LOCALAPPDATA%\BingWallpaperClient\config.json`. You can modify settings by editing this file or reinstalling the application.

Available configuration options:
- `Market`: Bing market code (e.g., "en-US", "de-DE", "ja-JP")
- `DownloadPath`: Directory where images are saved
- `MaxImages`: Maximum number of images to retain
- `KeepAllImages`: Whether to keep all downloaded images
- `ChangeDesktop`: Enable/disable automatic wallpaper changes
- `EnableLogging`: Enable logging (for troubleshooting)

## Scheduled Task

The installer creates a scheduled task (`BingWallpaperSync`) that runs:
- Daily at 9:00 AM (local time)
- At user logon
- When network connects (after 1 minute delay)

## Project Structure

```
BingWallpaperClient/
├── installer/
│   ├── Common.ps1              # Common installer utilities
│   ├── Configuration.ps1       # Interactive configuration setup
│   ├── FileOperations.ps1      # File copy and management
│   └── TaskScheduler.ps1       # Windows Task Scheduler 
├── src/
    ├── core/
    │   ├── ConfigManager.vbs   # Configuration management
    │   ├── ImageDownloader.vbs # Bing image download logic
    │   ├── SystemUtils.vbs     # System utilities and logging
    │   └── WallpaperCore.vbs   # Wallpaper setting functionality
    ├── lib/
    │   └── JSON.vbs            # JSON parser for VBScript
    ├── BingWallpaperClient.vbs # Main application entry point
├── installer.ps1               # Main installer script
├── markets.json                # Supported Bing markets configuration
├── README.md                   # Project documentation
├── SetWallpaper.cs             # C# wallpaper setter
└── SetWallpaper.exe            # Compiled wallpaper setter
```

## Installation Structure

After installation, BWC creates the following directory structure:

```
%LOCALAPPDATA%\BingWallpaperClient\
├── core\                       # Core application scripts
├── installer\                  # Installer scripts
├── lib\                        # Library scripts
├── BingWallpaperClient.vbs     # Main application entry point
├── bwc.cmd                     # Command-line wrapper
├── config.json                 # User configuration
├── installer.ps1               # Main installer script
├── markets.json                # Supported Bing markets
├── SetWallpaper.exe            # Compiled wallpaper setter

%USERPROFILE%\Pictures\BingWallpapers\
└── *.jpg                       # Downloaded wallpaper images
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.