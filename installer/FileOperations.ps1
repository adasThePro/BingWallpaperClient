# File Operations Module

function Copy-SourceFiles {
    Write-ColorOutput "[i] Copying application files..." "Cyan"

    $destDirs = @{
        Main = "$Script:InstallPath"
        Core = "$Script:InstallPath\core"
        Lib = "$Script:InstallPath\lib"
    }

    foreach ($dir in $destDirs.Values) {
        if (-Not (Test-Path $dir)) {
           $null = New-Item -ItemType Directory -Path $dir -Force
        }
    }

    $mainSource = "$Script:SourcePath\$($Script:SourceFiles.Main)"
    if (-Not (Test-Path $mainSource)) {
        Write-ColorOutput "[X] Main source file not found: $mainSource" "Red"
        return $false
    }

    Copy-Item -Path $mainSource -Destination $destDirs.Main -Force
    Write-ColorOutput "[+] Copied: BingWallpaperClient.vbs" "Green"

    foreach ($file in $Script:SourceFiles.Core) {
        $source = "$Script:SourcePath\$file"
        if (-Not (Test-Path $source)) {
            Write-ColorOutput "[X] Core file not found: $source" "Red"
            return $false
        }
        Copy-Item -Path $source -Destination $destDirs.Core -Force
        Write-ColorOutput "[+] Copied: $(Split-Path $file -Leaf)" "Green"
    }

    foreach ($file in $Script:SourceFiles.Lib) {
        $source = "$Script:SourcePath\$file"
        if (-Not (Test-Path $source)) {
            Write-ColorOutput "[X] Library file not found: $source" "Red"
            return $false
        }
        Copy-Item -Path $source -Destination $destDirs.Lib -Force
        Write-ColorOutput "[+] Copied: $(Split-Path $file -Leaf)" "Green"
    }

    foreach ($file in $Script:SourceFiles.Data) {
        $source = "$Script:SourcePath\$file"
        if (-Not (Test-Path $source)) {
            Write-ColorOutput "[X] Data file not found: $source" "Red"
            return $false
        }
        Copy-Item -Path $source -Destination $Script:InstallPath -Force
        Write-ColorOutput "[+] Copied: $(Split-Path $file -Leaf)" "Green"
    }

    $installerSource = "$Script:SourcePath\$($Script:SourceFiles.Installer)"
    if (-Not (Test-Path $installerSource)) {
        Write-ColorOutput "[X] Installer file not found: $installerSource" "Red"
        return $false
    }
    Copy-Item -Path $installerSource -Destination $Script:InstallPath -Force
    Write-ColorOutput "[+] Copied: $(Split-Path $Script:SourceFiles.Installer -Leaf)" "Green"

    $installerModulesDir = "$Script:SourcePath\installer"
    if (-Not (Test-Path $installerModulesDir)) {
        Write-ColorOutput "[X] Installer modules directory not found: $installerModulesDir" "Red"
        return $false
    }

    $destInstallerModulesDir = "$Script:InstallPath\installer"
    if (-Not (Test-Path $destInstallerModulesDir)) {
        $null = New-Item -ItemType Directory -Path $destInstallerModulesDir -Force
    }
    Copy-Item -Path "$installerModulesDir\*" -Destination $destInstallerModulesDir -Recurse -Force
    Write-ColorOutput "[+] Copied installer modules" "Green"

    $setWallpaperSource = "$Script:SourcePath\SetWallpaper.exe"
    if (-Not (Test-Path $setWallpaperSource)) {
        Write-ColorOutput "[X] SetWallpaper.exe not found: $setWallpaperSource" "Red"
        return $false
    }
    Copy-Item -Path $setWallpaperSource -Destination $Script:InstallPath -Force
    Write-ColorOutput "[+] Copied: SetWallpaper.exe" "Green"

    return $true
}

function New-CommandAlias {
    $aliasPath = "$Script:InstallPath\bwc.cmd"
    $vbsPath = "$Script:InstallPath\BingWallpaperClient.vbs"
    $batchContent = @"
@echo off
setlocal
if "%1"=="" (
    cscript.exe //nologo "$vbsPath" help
) else (
    cscript.exe //nologo "$vbsPath" %*
)
endlocal
"@
    try {
        [System.IO.File]::WriteAllText($aliasPath, $batchContent)
        Write-ColorOutput "[+] Created command alias: bwc" "Green"
        return $true
    } catch {
        Write-ColorOutput "[X] Failed to create command alias: $($_.Exception.Message)" "Red"
        return $false
    }
}