<#
.SYNOPSIS
    Automatically switches Windows 11 theme based on Sunrise/Sunset times
.DESCRIPTION
    This script fetches sunrise/sunset times, sets the appropriate theme,
    and schedules itself to run at the next transition time.
    NO EXPLORER RESTART NEEDED - uses Windows API to broadcast changes.
.NOTES
    Author: Khalid Shawkat
    Version: 2.0
#>

#Requires -RunAsAdministrator

# ============================================================================
# CONFIGURATION
# ============================================================================
$Config = @{
    Latitude = 40.063163025405856
    Longitude = -88.24559595778813
    TaskName = "ThemeSwitcher"
    TaskPath = "\CustomTasks\"
    ScriptPath = $PSCommandPath
    LogPath = "$env:TEMP\ThemeSwitcher.log"
    ForceExplorerRestart = $false  # Set to $true if API broadcast doesn't work
}

# ============================================================================
# LOGGING FUNCTION
# ============================================================================
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] $Message"
    Add-Content -Path $Config.LogPath -Value $logMessage
    Write-Host $logMessage
}

# ============================================================================
# WINDOWS API - BROADCAST THEME CHANGE (NO EXPLORER RESTART)
# ============================================================================
function Send-SettingChangeNotification {
    Write-Log "Broadcasting theme change notification to all windows..."
    
    $signature = @'
[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
public static extern IntPtr SendMessageTimeout(
    IntPtr hWnd,
    uint Msg,
    UIntPtr wParam,
    string lParam,
    uint fuFlags,
    uint uTimeout,
    out UIntPtr lpdwResult
);
'@
    
    try {
        $type = Add-Type -MemberDefinition $signature -Name Win32SendMessage -Namespace Win32Functions -PassThru
        
        $HWND_BROADCAST = [IntPtr]0xffff
        $WM_SETTINGCHANGE = 0x1a
        $SMTO_ABORTIFHUNG = 0x0002
        $result = [UIntPtr]::Zero
        
        # Broadcast theme change
        $type::SendMessageTimeout(
            $HWND_BROADCAST,
            $WM_SETTINGCHANGE,
            [UIntPtr]::Zero,
            "ImmersiveColorSet",
            $SMTO_ABORTIFHUNG,
            5000,
            [ref]$result
        ) | Out-Null
        
        Write-Log "Theme change notification sent successfully"
        return $true
    }
    catch {
        Write-Log "Error broadcasting theme change: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================================
# GET SUNRISE/SUNSET DATA
# ============================================================================
function Get-SunriseSunset {
    param(
        [double]$Latitude,
        [double]$Longitude,
        [datetime]$Date = (Get-Date)
    )
    
    $url = "https://api.sunrise-sunset.org/json?lat=$Latitude&lng=$Longitude&date=$($Date.ToString('yyyy-MM-dd'))&formatted=0"
    
    try {
        Write-Log "Fetching sunrise/sunset data from API..."
        $response = Invoke-RestMethod -Uri $url -UseBasicParsing -TimeoutSec 10
        
        if ($response.status -eq "OK") {
            $sunrise = [datetime]::Parse($response.results.sunrise).ToLocalTime()
            $sunset = [datetime]::Parse($response.results.sunset).ToLocalTime()
            
            Write-Log "Sunrise: $sunrise | Sunset: $sunset"
            
            return [pscustomobject]@{
                Sunrise = $sunrise
                Sunset = $sunset
            }
        } else {
            Write-Log "API returned error status: $($response.status)"
            return $null
        }
    }
    catch {
        Write-Log "Error fetching sunrise/sunset data: $($_.Exception.Message)"
        return $null
    }
}

# ============================================================================
# SET WINDOWS THEME
# ============================================================================
function Set-WindowsTheme {
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet("Light", "Dark")]
        [string]$Theme
    )
    
    $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    
    try {
        if ($Theme -eq "Light") {
            Write-Log "Setting Light Theme..."
            Set-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -Value 0 -Type Dword -Force
            Set-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -Value 1 -Type Dword -Force
        } else {
            Write-Log "Setting Dark Theme..."
            Set-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -Value 0 -Type Dword -Force
            Set-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -Value 0 -Type Dword -Force
        }
        
        # Broadcast the change (NO Explorer restart needed)
        $broadcasted = Send-SettingChangeNotification
        
        # Optional: Restart Explorer only if configured or if broadcast failed
        if ($Config.ForceExplorerRestart -or -not $broadcasted) {
            Write-Log "Restarting Explorer to apply theme changes..."
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
            Start-Process explorer
        }
        
        Write-Log "Theme changed to $Theme successfully"
        return $true
    }
    catch {
        Write-Log "Error setting theme: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================================
# GET CURRENT THEME
# ============================================================================
function Get-CurrentTheme {
    $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
    try {
        $value = Get-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -ErrorAction Stop
        return if ($value.AppsUseLightTheme -eq 1) { "Light" } else { "Dark" }
    }
    catch {
        return "Unknown"
    }
}

# ============================================================================
# CREATE SCHEDULED TASK
# ============================================================================
function New-ThemeSwitcherTask {
    param(
        [datetime]$TriggerTime,
        [string]$Description
    )
    
    try {
        # Remove existing task if it exists
        Get-ScheduledTask -TaskName $Config.TaskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false
        
        Write-Log "Creating scheduled task for $($TriggerTime.ToString('yyyy-MM-dd HH:mm:ss'))"
        
        # Create action
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($Config.ScriptPath)`""
        
        # Create trigger
        $trigger = New-ScheduledTaskTrigger -Once -At $TriggerTime
        
        # Create principal (run as current user)
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
        
        # Create settings
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        
        # Register task
        Register-ScheduledTask -TaskName $Config.TaskName -TaskPath $Config.TaskPath -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $Description -Force | Out-Null
        
        Write-Log "Scheduled task created successfully for $Description"
        return $true
    }
    catch {
        Write-Log "Error creating scheduled task: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================================
# CREATE BOOT TASK
# ============================================================================
function New-BootTask {
    try {
        $bootTaskName = "$($Config.TaskName)_Boot"
        
        # Remove existing boot task if it exists
        Get-ScheduledTask -TaskName $bootTaskName -ErrorAction SilentlyContinue | Unregister-ScheduledTask -Confirm:$false
        
        Write-Log "Creating boot-time task..."
        
        # Create action
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$($Config.ScriptPath)`""
        
        # Create trigger for boot
        $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
        
        # Create principal
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Highest
        
        # Create settings - delay by 30 seconds to allow network connectivity
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        $settings.ExecutionTimeLimit = "PT5M"  # 5 minute timeout
        
        # Register task
        Register-ScheduledTask -TaskName $bootTaskName -TaskPath $Config.TaskPath -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Auto Theme Switcher - Boot Task" -Force | Out-Null
        
        Write-Log "Boot task created successfully"
        return $true
    }
    catch {
        Write-Log "Error creating boot task: $($_.Exception.Message)"
        return $false
    }
}

# ============================================================================
# MAIN LOGIC
# ============================================================================
function Main {
    Write-Log "=========================================="
    Write-Log "Auto Theme Switcher Started"
    Write-Log "=========================================="
    
    # Get current time
    $currentTime = Get-Date
    Write-Log "Current time: $($currentTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    
    # Get current theme
    $currentTheme = Get-CurrentTheme
    Write-Log "Current theme: $currentTheme"
    
    # Fetch sunrise/sunset times
    $sunData = Get-SunriseSunset -Latitude $Config.Latitude -Longitude $Config.Longitude
    
    if (-not $sunData) {
        Write-Log "Failed to fetch sunrise/sunset data. Exiting."
        exit 1
    }
    
    # Determine if it's daytime or nighttime
    $isDaytime = ($currentTime -ge $sunData.Sunrise) -and ($currentTime -lt $sunData.Sunset)
    $desiredTheme = if ($isDaytime) { "Light" } else { "Dark" }
    
    Write-Log "Current period: $(if ($isDaytime) { 'Daytime' } else { 'Nighttime' })"
    Write-Log "Desired theme: $desiredTheme"
    
    # Set theme if it's different from current
    if ($currentTheme -ne $desiredTheme) {
        Write-Log "Theme mismatch detected. Changing theme from $currentTheme to $desiredTheme"
        Set-WindowsTheme -Theme $desiredTheme
    } else {
        Write-Log "Theme is already correct ($currentTheme). No change needed."
    }
    
    # Calculate next transition time
    if ($isDaytime) {
        # Currently daytime, schedule for sunset
        $nextTransition = $sunData.Sunset
        $nextDescription = "Switch to Dark theme at Sunset"
    } else {
        # Currently nighttime
        if ($currentTime -lt $sunData.Sunrise) {
            # Before sunrise today
            $nextTransition = $sunData.Sunrise
            $nextDescription = "Switch to Light theme at Sunrise"
        } else {
            # After sunset, schedule for tomorrow's sunrise
            $tomorrowSunData = Get-SunriseSunset -Latitude $Config.Latitude -Longitude $Config.Longitude -Date $currentTime.AddDays(1)
            if ($tomorrowSunData) {
                $nextTransition = $tomorrowSunData.Sunrise
                $nextDescription = "Switch to Light theme at Sunrise (tomorrow)"
            } else {
                Write-Log "Failed to get tomorrow's sunrise time"
                exit 1
            }
        }
    }
    
    Write-Log "Next transition scheduled: $($nextTransition.ToString('yyyy-MM-dd HH:mm:ss'))"
    
    # Create scheduled task for next transition
    New-ThemeSwitcherTask -TriggerTime $nextTransition -Description $nextDescription
    
    # Ensure boot task exists
    New-BootTask
    
    Write-Log "=========================================="
    Write-Log "Auto Theme Switcher Completed Successfully"
    Write-Log "=========================================="
}

# ============================================================================
# RUN MAIN FUNCTION
# ============================================================================
Main