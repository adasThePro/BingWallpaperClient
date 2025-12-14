# Bing Wallpaper Client Installer
# Windows 7 or later

#Requires -Version 2.0

param(
    [string]$Action = "install",
    [switch]$SkipCompile
)

$Script:AppName = "BingWallpaperClient"
$Script:Version = "1.0.1"
$Script:InstallPath = "$env:LOCALAPPDATA\$Script:AppName"
$Script:TaskName = "BingWallpaperSync"
$Script:ConfigPath = "$Script:InstallPath\config.json"
$Script:SourcePath = Split-Path -Parent $MyInvocation.MyCommand.Path

$validActions = @("install", "uninstall")

try {
    . "$Script:SourcePath\installer\Common.ps1"
    . "$Script:SourcePath\installer\Configuration.ps1"
    . "$Script:SourcePath\installer\FileOperations.ps1"
    . "$Script:SourcePath\installer\TaskScheduler.ps1"
} catch {
    Write-Host "[X] Failed to load installer modules: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Press Enter to exit..." -NoNewLine
    $null = Read-Host
    exit 1
}

if ($validActions -notcontains $Action.ToLower()) {
    Write-ColorOutput "[X] Invalid action specified: '$Action'. Valid actions are: $($validActions -join ', ')." Red
    exit 1
}

function Load-Markets {
    $marketsFile = "$Script:SourcePath\markets.json"

    if (-Not (Test-Path $marketsFile)) {
        Write-ColorOutput "[X] markets.json file not found at $marketsFile" "Red"
        Write-Host ""
        Write-Host "Press Enter to exit..." -NoNewLine
        $null = Read-Host
        exit 1
    }

    $jsonContent = [string]::Join("`n", (Get-Content $marketsFile))
    $marketsData = ConvertFrom-JsonPS2 -JsonString $jsonContent

    if ($marketsData -eq $null) {
        Write-ColorOutput "[X] Failed to parse markets.json file" "Red"
        Write-Host ""
        Write-Host "Press Enter to exit..." -NoNewLine
        $null = Read-Host
        exit 1
    }

    $Global:BingMarkets = @()
    $markets = $marketsData["markets"]

    foreach ($market in $markets) {
        $Global:BingMarkets += @{
            Code = $market["code"]
            Name = $market["name"]
        }
    }
}

Load-Markets

$Script:SourceFiles = @{
    Main = "src\BingWallpaperClient.vbs"
    Core = @(
        "src\core\WallpaperCore.vbs",
        "src\core\ConfigManager.vbs",
        "src\core\ImageDownloader.vbs",
        "src\core\SystemUtils.vbs"
    )
    Lib = @(
        "src\lib\JSON.vbs"
    )
    Data = @(
        "markets.json"
    )
    Installer = "installer.ps1"
    SetWallpaper = "SetWallpaper.cs"
}

function Compile-SetWallpaper {
    param([bool]$SkipCompile = $false)
    
    $csSource = "$Script:SourcePath\$($Script:SourceFiles.SetWallpaper)"
    $exeOutput = "$Script:SourcePath\SetWallpaper.exe"
    
    if ($SkipCompile) {
        Write-ColorOutput "[i] Skipping compilation (using existing SetWallpaper.exe)..." "Cyan"
        if (-Not (Test-Path $exeOutput)) {
            Write-ColorOutput "[X] SetWallpaper.exe not found at $exeOutput" "Red"
            Write-ColorOutput "[X] Please compile it first or remove -SkipCompile flag" "Red"
            return $false
        }
        Write-ColorOutput "[+] Found existing SetWallpaper.exe" "Green"
        return $true
    }
    
    Write-ColorOutput "[i] Compiling SetWallpaper.cs..." "Cyan"
    
    if (-Not (Test-Path $csSource)) {
        Write-ColorOutput "[X] SetWallpaper.cs not found at $csSource" "Red"
        return $false
    }
    
    $frameworkPaths = @(
        "$env:SystemRoot\Microsoft.NET\Framework64\v2.0.50727\csc.exe",
        "$env:SystemRoot\Microsoft.NET\Framework\v2.0.50727\csc.exe",
        "$env:SystemRoot\Microsoft.NET\Framework64\v3.5\csc.exe",
        "$env:SystemRoot\Microsoft.NET\Framework\v3.5\csc.exe",
        "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
        "$env:SystemRoot\Microsoft.NET\Framework\v4.0.30319\csc.exe"
    )
    
    $frameworkPath = $null
    foreach ($path in $frameworkPaths) {
        if (Test-Path $path) {
            $frameworkPath = $path
            break
        }
    }
    
    if (-Not $frameworkPath) {
        Write-ColorOutput "[X] C# compiler (csc.exe) not found" "Red"
        Write-ColorOutput "[X] .NET Framework 2.0 or later is required" "Red"
        return $false
    }
    
    try {
        $compileArgs = @(
            "/nologo",
            "/out:`"$exeOutput`"",
            "/target:exe",
            "/optimize+",
            "`"$csSource`""
        )
        
        $process = Start-Process -FilePath $frameworkPath -ArgumentList $compileArgs -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0 -and (Test-Path $exeOutput)) {
            Write-ColorOutput "[+] SetWallpaper.exe compiled successfully" "Green"
            return $true
        } else {
            Write-ColorOutput "[X] Compilation failed" "Red"
            return $false
        }
    } catch {
        Write-ColorOutput "[X] Compilation error: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Install-Client {
    Write-Host ""
    Write-ColorOutput "========================================" "Cyan"
    Write-ColorOutput "  Bing Wallpaper Client v$Script:Version" "Cyan"
    Write-ColorOutput "  Installation Wizard" "Cyan"
    Write-ColorOutput "========================================" "Cyan"
    Write-Host ""

    Write-ColorOutput "[i] Checking system requirements..." "Cyan"

    if (-Not (Test-WindowsVersionSupported)) {
        $winVersion = Get-WindowsVersion
        Write-Host ""
        Write-ColorOutput "========================================" "Red"
        Write-ColorOutput "  UNSUPPORTED WINDOWS VERSION" "Red"
        Write-ColorOutput "========================================" "Red"
        Write-Host ""

        Write-ColorOutput "This installer supports:" "Yellow"
        Write-Host "  - Windows 7"
        Write-Host "  - Windows 8 / 8.1"
        Write-Host "  - Windows 10"
        Write-Host "  - Windows 11"
        Write-Host ""

        Write-ColorOutput "Your Windows version is not supported: $winVersion" "Red"
        Write-Host ""
        Write-Host "Press Enter to exit..." -NoNewLine
        $null = Read-Host
        exit 1
    }

    Write-ColorOutput "[i] Windows version is supported." "Green"

    Write-Host ""
    Write-ColorOutput "[i] Starting installation..." "Cyan"
    Write-ColorOutput "[i] Installation directory: $Script:InstallPath" "Cyan"
    

    if (-Not (Test-Path $Script:InstallPath)) {
        try {
            $null = New-Item -ItemType Directory -Path $Script:InstallPath -Force
            Write-ColorOutput "[+] Created installation directory" "Green"
        } catch {
            Write-ColorOutput "[X] Failed to create installation directory: $($_.Exception.Message)" "Red"
            Write-Host ""
            Write-Host "Press Enter to exit..." -NoNewLine
            $null = Read-Host
            exit 1
        }
    }
    Write-Host ""
    
    if (-Not (Compile-SetWallpaper -SkipCompile $SkipCompile)) {
        Write-ColorOutput "[X] Failed to compile SetWallpaper. Use -SkipCompile to skip compilation. Installation aborted." "Red"
        Write-Host ""
        Write-Host "Press Enter to exit..." -NoNewLine
        $null = Read-Host
        exit 1
    }
    
    Write-Host ""
    
    if (-Not (Copy-SourceFiles)) {
        Write-ColorOutput "[X] Failed to copy source files. Installation aborted." "Red"
        Write-Host ""
        Write-Host "Press Enter to exit..." -NoNewLine
        $null = Read-Host
        exit 1
    }

    Write-Host ""

    if (-Not (Add-ToUserPath -Path $Script:InstallPath)) {
        Write-ColorOutput "[!] Warning: Could not add application to user PATH" "Yellow"
    }

    if (-Not (New-CommandAlias)) {
        Write-ColorOutput "[!] Warning: Could not create command alias" "Yellow"
    }

    Write-Host ""

    $config = Get-UserConfiguration

    if (-Not $config) {
        Write-ColorOutput "[X] Failed to get configuration. Installation aborted." "Red"
        Write-Host ""
        Write-Host "Press Enter to exit..." -NoNewLine
        $null = Read-Host
        exit 1
    }

    if (-Not (Save-Configuration -Config $config)) {
        Write-ColorOutput "[X] Failed to save configuration. Installation aborted." "Red"
        Write-Host ""
        Write-Host "Press Enter to exit..." -NoNewLine
        $null = Read-Host
        exit 1
    }

    Write-ColorOutput "[+] Configuration saved successfully." "Green"

    Write-Host ""
    Write-ColorOutput "[i] Setting up scheduled task..." "Cyan"
    Write-Host ""

    $vbsPath = "$Script:InstallPath\BingWallpaperClient.vbs"
    if (-Not (New-ScheduledTask -TaskName $Script:TaskName -ScriptPath $vbsPath)) {
        Write-ColorOutput "[!] Warning: Scheduled task setup failed." "Yellow"
        Write-Host ""
    }

    Write-ColorOutput "=== Installation Complete ===" "Green"

    Write-Host ""
    Write-ColorOutput "Commands:" "Cyan"
    Write-Host "  bwc apply      - Apply wallpaper now"
    Write-Host "  bwc status     - Check status"
    Write-Host "  bwc config     - Show configuration"
    Write-Host "  bwc uninstall  - Show uninstallation instructions"
    Write-Host "  bwc help       - Show help information"
    Write-Host ""
    Write-ColorOutput "[i] Note: Open a new terminal for 'bwc' command" "Yellow"
    Write-Host ""

    Write-Host "Press Enter to exit..." -NoNewLine
    $null = Read-Host
}

function Uninstall-Client {
    Write-Host ""
    Write-ColorOutput "========================================" "Cyan"
    Write-ColorOutput "  Bing Wallpaper Client $Script:Version" "Cyan"
    Write-ColorOutput "  Uninstallation Wizard" "Cyan"
    Write-ColorOutput "========================================" "Cyan"
    Write-Host ""

    $existingConfig = $null
    $imagesPath = $null

    if (Test-Path $Script:ConfigPath) {
        try {
            $json = [System.IO.File]::ReadAllText($Script:ConfigPath)
            $existingConfig = ConvertFrom-JsonPS2 -JsonString $json

            if ($existingConfig -and $existingConfig.ContainsKey("DownloadPath")) {
                $imagesPath = $existingConfig["DownloadPath"]
            }
        } catch {
            Write-ColorOutput "[!] Warning: Could not load configuration file: $($_.Exception.Message)" "Yellow"
        }
    }

    Write-Host "This will remove the Bing Wallpaper Client completely"
    Write-Host ""
    Write-Host "Continue? (Y/n): " -NoNewLine
    $confirmInput = Read-Host

    if ($confirmInput -ne "" -and $confirmInput.ToLower() -ne "y") {
        Write-ColorOutput "[i] Uninstallation cancelled by user" "Yellow"
        Write-Host ""
        Write-Host "Press Enter to exit..." -NoNewLine
        $null = Read-Host
        exit 0
    }

    Write-Host ""
    Write-ColorOutput "[i] Uninstalling..." "Cyan"
    Write-ColorOutput "[i] Installation directory: $Script:InstallPath" "Cyan"

    Write-Host ""
    Write-ColorOutput "[i] Stopping related processes..." "Cyan"
    try {
        $vbsPath = "$Script:InstallPath\BingWallpaperClient.vbs"
        $processes = Get-WmiObject Win32_Process | Where-Object {
            $_.CommandLine -like "*$vbsPath*" -or
            $_.CommandLine -like "*cscript.exe*BingWallpaperClient.vbs*" -or
            $_.CommandLine -like "*wscript.exe*BingWallpaperClient.vbs*"
        }

        if ($processes) {
            foreach ($process in $processes) {
                try {
                    $processId = $process.ProcessId
                    $null = $process.Terminate()
                    Write-ColorOutput "[+] Stopped process: $processId" "Green"
                } catch {
                    Write-ColorOutput "[!] Warning: Could not stop process $processId`: $($_.Exception.Message)" "Yellow"
                }
            }
            Start-Sleep -Seconds 2
        } else {
            Write-ColorOutput "[i] No running processes found" "Cyan"
        }
    } catch {
        Write-ColorOutput "[!] Warning: Could not stop processes: $($_.Exception.Message)" "Yellow"
    }

    Write-Host ""
    Write-ColorOutput "[i] Removing scheduled task..." "Cyan"
    try {
        $result = schtasks.exe /Query /TN $Script:TaskName 2>&1
        if ($LASTEXITCODE -eq 0) {
            $null = schtasks.exe /Delete /TN $Script:TaskName /F
            Write-ColorOutput "[+] Deleted scheduled task: $Script:TaskName" "Green"
        }
    } catch {
        Write-ColorOutput "[!] Warning: Could not delete scheduled task: $($_.Exception.Message)" "Yellow"
    }

    Write-Host ""
    Write-ColorOutput "[i] Removing bwc from user PATH..." "Cyan"
    $null = Remove-FromUserPath -Path $Script:InstallPath

    Write-Host ""
    Write-ColorOutput "[i] Removing installation directory..." "Cyan"
    if (Test-Path $Script:InstallPath) {
        try {
            Remove-Item -Path $Script:InstallPath -Recurse -Force -ErrorAction Stop
            Write-ColorOutput "[+] Installation directory removed" "Green"
        } catch {
            Write-ColorOutput "[!] Warning: Could not remove directory: $($_.Exception.Message)" "Yellow"
            Write-ColorOutput "[!] You may need to manually delete: $Script:InstallPath" "Yellow"
        }
    } else {
        Write-ColorOutput "[i] Installation directory not found" "Cyan"
    }

    Write-Host ""
    if ($imagesPath -and (Test-Path $imagesPath)) {
        Write-Host "Remove downloaded images? (y/N): " -NoNewLine
        $removeImagesInput = Read-Host
        if ($removeImagesInput.ToLower() -eq "y") {
            try {
                Remove-Item -Path $imagesPath -Recurse -Force -ErrorAction Stop
                Write-ColorOutput "[+] Downloaded images removed" "Green"                
            } catch {
                Write-ColorOutput "[!] Warning: Could not remove images directory: $($_.Exception.Message)" "Yellow"
                Write-ColorOutput "[!] You may need to manually delete: $imagesPath"
            }
        }
    } else {
        Write-ColorOutput "[i] No downloaded images directory found" "Cyan"
    }

    Write-Host ""
    Write-ColorOutput "=== Uninstallation Complete ===" "Green"
    Write-Host ""
    Write-ColorOutput "[i] Please restart your terminal to complete the PATH cleanup" "Yellow"
    Write-ColorOutput ""
    Write-Host "Press Enter to exit..." -NoNewLine
    $null = Read-Host
}

switch ($Action.ToLower()) {
    "install" {
        Install-Client
    }
    "uninstall" {
        Uninstall-Client
    }
    default {
        Install-Client
    }
}