# Configuration Module

function Get-UserConfiguration {
    Write-Host ""
    Write-ColorOutput "=== Bing Wallpaper Client Setup ===" "Cyan"
    Write-Host ""

    $config = @{}

    $marketCode = ""
    $validMarket = $false
    $marketCodes = $Global:BingMarkets | ForEach-Object { $_.Code }

    while (-Not $validMarket) {
        Write-Host "Enter market code (e.g., en-US, en-IN, fr-FR): " -NoNewLine
        $marketCode = Read-Host

        if ([string]::IsNullOrEmpty($marketCode.Trim())) {
            Write-ColorOutput "[!] Market code cannot be empty. Please try again." "Yellow"
            continue
        } elseif ($marketCodes -contains $marketCode) {
            $marketName = ($Global:BingMarkets | Where-Object { $_.Code -eq $marketCode }).Name
            Write-ColorOutput "[i] Selected market: $marketName ($marketCode)" "Cyan"
            Write-Host ""
            Write-Host "Confirm market? (Y/n): " -NoNewLine
            $confirmation = Read-Host
            if ($confirmation -eq "" -or $confirmation.ToLower() -eq "y") {
                $config.Market = $marketCode
                $validMarket = $true
            }
        } else {
            Write-ColorOutput "[!] Invalid market code '$marketCode'. Please try again." "Yellow"
        }
    }
    Write-Host ""

    $defaultImagesPath = "$env:USERPROFILE\Pictures\BingWallpapers"
    Write-Host "Images download path (default: $defaultImagesPath): " -NoNewLine  
    $imagesPath = Read-Host
    $config.DownloadPath = if ([string]::IsNullOrEmpty($imagesPath.Trim())) { $defaultImagesPath } else { $imagesPath }
    Write-ColorOutput "[+] Images download path set to: $($config.DownloadPath)" "Green"
    Write-Host ""

    if (-Not (Test-Path $config.DownloadPath)) {
        try {
            $null = New-Item -ItemType Directory -Path $config.DownloadPath -Force
            Write-ColorOutput "[+] Created images download directory: $($config.DownloadPath)" "Green"
        } catch {
            Write-ColorOutput "[X] Failed to create images download directory: $($config.DownloadPath)" "Red"
            return $null
        }
    }

    Write-Host "Maximum images to keep (0 for unlimited, default: 30 max: 1000): " -NoNewLine
    $config.KeepAllImages = $false
    $maxImagesInput = Read-Host
    if ([string]::IsNullOrEmpty($maxImagesInput.Trim())) {
        $config.MaxImages = 30
    } elseif ([int]::TryParse($maxImagesInput, [ref]$null) -and [int]$maxImagesInput -ge 0 -and [int]$maxImagesInput -le 1000) {
        $config.MaxImages = [int]$maxImagesInput
        $config.KeepAllImages = ($config.MaxImages -eq 0)
    } else {
        Write-ColorOutput "[!] Invalid input for maximum images to keep. Using default value of 30." "Yellow"
        $config.MaxImages = 30
    }
    Write-ColorOutput "[+] Maximum images to keep set to: $(if ($config.KeepAllImages) { 'Unlimited' } else { $config.MaxImages })" "Green"
    Write-Host ""

    Write-Host "Change desktop wallpaper? (Y/n): " -NoNewLine
    $changeDesktopWallpaperInput = Read-Host
    $config.ChangeDesktop = ($changeDesktopWallpaperInput -eq "" -or $changeDesktopWallpaperInput.ToLower() -eq "y")
    Write-ColorOutput "[+] Desktop wallpaper: $(if ($config.ChangeDesktop) { 'Enabled' } else { 'Disabled' })" "Green"
    Write-Host ""

    $config.EnableLogging = $false
    $config.LastUpdate = $null
    $config.Version = $Script:Version

    return $config
}

function Save-Configuration {
    param([hashtable]$Config)

    $configDir = Split-Path  $Script:ConfigPath -Parent
    if (-Not (Test-Path $configDir)) {
        try {
            $null = New-Item -ItemType Directory -Path $configDir -Force
            Write-ColorOutput "[+] Created configuration directory: $configDir" "Green"
        } catch {
            Write-ColorOutput "[X] Failed to create configuration directory: $configDir" "Red"
            return $false
        }
    }

    try {
        $json = ConvertTo-JsonPS2 -Data $Config
        [System.IO.File]::WriteAllText($Script:ConfigPath, $json)
        Write-ColorOutput "[+] Configuration saved to $Script:ConfigPath" "Green"
        return $true
    } catch {
        Write-ColorOutput "[X] Failed to save configuration: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Load-Configuration {
    param([string]$ConfigPath = $Script:ConfigPath)

    if (-Not (Test-Path $ConfigPath)) {
        return $null
    }

    try {
        $json = [System.IO.File]::ReadAllText($ConfigPath)
        $config = ConvertFrom-JsonPS2 -JsonString $json
        return $config
    } catch {
        Write-ColorOutput "[X] Failed to load configuration: $($_.Exception.Message)" "Red"
        return $null
    }
}