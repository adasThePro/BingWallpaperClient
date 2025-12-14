# Common Functions Module

function Write-ColorOutput {
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    Write-Host $Message -ForegroundColor $Color
}

function Get-WindowsVersion {
    $os = Get-WmiObject -Class Win32_OperatingSystem
    $version = [System.Environment]::OSVersion.Version
    $productName = $os.Caption

    $result = @{
        Major = $version.Major
        Minor = $version.Minor
        Build = $version.Build
        ProductName = $productName
    }

    return $result
}

function Test-WindowsVersionSupported {
    $winVersion = Get-WindowsVersion
    $major = $winVersion.Major
    $minor = $winVersion.Minor

    if ($major -eq 6 -and $minor -ge 1) {
        return $true
    } elseif ($major -eq 10) {
        return $true
    }

    return $false
}

function ConvertTo-JsonPS2 {
    param([hashtable]$Data)

    $json = "{"
    $first = $true

    foreach ($key in $Data.Keys) {
        if (-Not $first) {
            $json += ","
        }
        $first = $false

        $json += "`n  `"$key`": "

        $value = $Data[$key]
        if ($value -eq $null) {
            $json += "null"
        } elseif ($value -is [bool]) {
            $json += if ($value) { "true" } else { "false" }
        } elseif ($value -is [int]) {
            $json += $value.ToString()
        } elseif ($value -is [string]) {
            $escapedValue = $value -replace '\\', '\\' -replace '"', '\"' -replace "`r`n", '\n' -replace "`r", "\r" -replace "`t", '\t'
            $json += "`"$escapedValue`""
        } else {
            $json += "`"$value`""
        }
    }

    $json += "`n}"
    return $json
}

function ConvertFrom-JsonPS2 {
    param([string]$JsonString)

    try {
        Add-Type -AssemblyName System.Web.Extensions
        $serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $obj = $serializer.DeserializeObject($JsonString)

        $result = @{}
        foreach ($key in $obj.Keys) {
            $result[$key] = $obj[$key]
        }
        return $result
    } catch {
        Write-ColorOutput "Error parsing JSON: $($_.Exception.Message)" "Red"
        return $null
    }
}

function Add-ToUserPath {
    param([string]$Path)

    try {
        $currentPath = [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::User)

        if ($currentPath -notLike "*$Path*") {
            if ($currentPath -and -not $currentPath.EndsWith(";")) {
                $newPath = "$currentPath;$Path"
            } elseif ($currentPath) {
                $newPath = "$currentPath$Path"
            } else {
                $newPath = $Path
            }

            [Environment]::SetEnvironmentVariable("Path", $newPath, [EnvironmentVariableTarget]::User)
            $env:Path = [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::Machine) + ";" + [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::User)

            Write-ColorOutput "[+] Added '$Path' to user PATH." "Green"
            return $true
        } else {
            Write-ColorOutput "[i] '$Path' is already in user PATH." "Cyan" 
            return $false
        }
    } catch {
        Write-ColorOutput "[X] Failed to add '$Path' to user PATH: $($_.Exception.Message)" "Red"
        return $false
    }
}

function Remove-FromUserPath {
    param([string]$Path)

    try {
        $currentPath = [Environment]::GetEnvironmentVariable("Path", [EnvironmentVariableTarget]::User)

        if ($currentPath) {
            $pathArray = $currentPath.Split(';') | Where-Object { $_ -and $_ -ne $Path }
            $newPath = $pathArray -join ';'

            [Environment]::SetEnvironmentVariable("Path", $newPath, [EnvironmentVariableTarget]::User)
            Write-ColorOutput "[+] Removed '$Path' from user PATH." "Green"
            return $true
        }
    } catch {
        Write-ColorOutput "[X] Failed to remove '$Path' from user PATH: $($_.Exception.Message)" "Red"
        return $false
    }
}