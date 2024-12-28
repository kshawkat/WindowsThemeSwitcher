# Windows Theme Switcher

A PowerShell script that automatically switches between Light and Dark themes in Windows 11 based on sunrise and sunset times at your location. No Explorer restart is required, as it uses Windows API to broadcast theme changes.

## Features
- Automatically switches to Light theme at sunrise and Dark theme at sunset.
- Fetches sunrise/sunset times from the [Sunrise-Sunset API](https://sunrise-sunset.org/).
- Uses Windows API to apply theme changes without restarting Explorer.
- Schedules itself to run at the next theme transition time.
- Includes a boot-time task to ensure theme consistency on system startup.
- Logs all actions to a file for troubleshooting.
- Configurable latitude, longitude, and other settings.

## Requirements
- Windows 11
- PowerShell 5.1 or later
- Administrative privileges (required for theme changes and task scheduling)
- Internet connection (to fetch sunrise/sunset times)

## Installation
1. Clone or download this repository to your local machine.
2. Open the `ThemeSwitcher.ps1` script in a text editor.
3. Update the configuration section with your latitude and longitude:
   ```powershell
   $Config = @{
       Latitude = 40.063163025405856  # Replace with your latitude
       Longitude = -88.24559595778813 # Replace with your longitude
       TaskName = "ThemeSwitcher"
       TaskPath = "\CustomTasks\"
       ScriptPath = $PSCommandPath
       LogPath = "$env:TEMP\ThemeSwitcher.log"
       ForceExplorerRestart = $false
   }
   ```
   You can find your coordinates using services like [Google Maps](https://www.google.com/maps) or [latlong.net](https://www.latlong.net/).

4. Save the script to a permanent location (e.g., `%Appdata%\MyScripts\ThemeSwitcher.ps1`).

## Usage
1. Open PowerShell as Administrator.
2. Navigate to the script's directory:
   ```powershell
   cd %Appdata%\MyScripts
   ```
3. Run the script:
   ```powershell
   .\ThemeSwitcher.ps1
   ```
4. The script will:
   - Check the current time and sunrise/sunset data.
   - Set the appropriate theme (Light for daytime, Dark for nighttime).
   - Create a scheduled task to run at the next transition (sunrise or sunset).
   - Create a boot-time task to ensure the theme is set correctly on startup.

The script runs silently in the background and logs all actions to `$env:TEMP\ThemeSwitcher.log`.

## How It Works
- **Sunrise/Sunset Data**: Fetches times from the Sunrise-Sunset API based on your coordinates.
- **Theme Switching**: Modifies registry keys (`HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize`) to set Light or Dark theme.
  - **Light Theme**: Sets System to Dark (`SystemUsesLightTheme = 0`) and Apps to Light (`AppsUseLightTheme = 1`) by default. See **Customization** for modifying this behavior.
  - **Dark Theme**: Sets both System and Apps to Dark (`SystemUsesLightTheme = 0`, `AppsUseLightTheme = 0`).
- **Change Notification**: Uses `SendMessageTimeout` from `user32.dll` to broadcast theme changes, avoiding Explorer restarts.
- **Scheduling**: Creates a Windows Task Scheduler task to run at the next sunrise/sunset and at system boot.
- **Logging**: Writes detailed logs to `%TEMP%\ThemeSwitcher.log` for debugging.

## Configuration Options
- `Latitude` and `Longitude`: Set your geographic coordinates for accurate sunrise/sunset times.
- `TaskName`: Name of the scheduled task (default: `ThemeSwitcher`).
- `TaskPath`: Task Scheduler path (default: `\CustomTasks\`).
- `ScriptPath`: Path to the script (automatically set to `$PSCommandPath`).
- `LogPath`: Location for log files (default: `%TEMP%\ThemeSwitcher.log`).
- `ForceExplorerRestart`: Set to `$true` to force Explorer restart if the API broadcast fails (default: `$false`).
- **Note**: The Light theme currently sets the System to Dark and Apps to Light. To change this to a full Light theme (both System and Apps set to Light), see the **Customization** section.

## Customization
To modify the Light theme to apply Light mode to both System and Apps (instead of the default System: Dark, Apps: Light), edit the `Set-WindowsTheme` function in the script. Locate the following lines in the `Light` theme block (around line 138):

```powershell
Set-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -Value 0 -Type Dword -Force
Set-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -Value 1 -Type Dword -Force
```

Change the `SystemUsesLightTheme` value to `1`:

```powershell
Set-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -Value 1 -Type Dword -Force
Set-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -Value 1 -Type Dword -Force
```

Save the script and rerun it to apply the updated Light theme configuration.

## Troubleshooting
- Check the log file at `%TEMP%\ThemeSwitcher.log` for errors.
- Ensure the script runs with administrative privileges.
- Verify your internet connection for API calls.
- If theme changes don't apply, try setting `ForceExplorerRestart = $true` in the configuration.
- Confirm your coordinates are correct for accurate sunrise/sunset times.
- If the Light theme doesn't apply as expected, verify the `SystemUsesLightTheme` and `AppsUseLightTheme` settings in the script (see **Customization**).

## Uninstall
To remove the scheduled tasks:
1. Open PowerShell as Administrator.
2. Run:
   ```powershell
   Get-ScheduledTask -TaskName "ThemeSwitcher" -TaskPath "\CustomTasks\" | Unregister-ScheduledTask -Confirm:$false
   Get-ScheduledTask -TaskName "ThemeSwitcher_Boot" -TaskPath "\CustomTasks\" | Unregister-ScheduledTask -Confirm:$false
   ```
3. Delete the script and log files if desired.

## Notes
- **Author**: Khalid Shawkat
- **Version**: 2.0
- Tested on Windows 11 with PowerShell 5.1.
- The script requires an internet connection to fetch sunrise/sunset times.
- Some applications may not immediately reflect theme changes until restarted.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments
- Uses the [Sunrise-Sunset API](https://sunrise-sunset.org/) for time data.
- Inspired by the need for seamless theme switching without Explorer restarts.

---

⭐ Star the repo if it helps! Contributions welcome via pull requests. Questions? [Open an issue](https://github.com/kshawkat/WindowsThemeSwitcher).


