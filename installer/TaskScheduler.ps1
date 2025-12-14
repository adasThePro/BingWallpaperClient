# Task Scheduler Module

function New-ScheduledTask {
    param(
        [string]$TaskName,
        [string]$ScriptPath
    )

    try {
        $existingTask = schtasks.exe /Query /TN $TaskName 2>$null
        if ($LASTEXITCODE -eq 0) {
            $null = schtasks.exe /Delete /TN $TaskName /F
        }

        if ($env:USERDOMAIN -and ($env:USERNAME -ne $env:COMPUTERNAME)) {
            $userId = "$env:USERDOMAIN\$env:USERNAME"
        } else {
            $userId = "$env:COMPUTERNAME\$env:USERNAME"
        }
        $description = "Automatically changes Windows wallpaper for $env:USERNAME using Bing daily images."

        $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
    <RegistrationInfo>
        <Description>$description</Description>
    </RegistrationInfo>
    <Triggers>
        <CalendarTrigger>
            <StartBoundary>2025-01-01T09:00:00</StartBoundary>
            <Enabled>true</Enabled>
            <ScheduleByDay>
                <DaysInterval>1</DaysInterval>
            </ScheduleByDay>
        </CalendarTrigger>
        <LogonTrigger>
            <Enabled>true</Enabled>
            <UserId>$userId</UserId>
        </LogonTrigger>
        <EventTrigger>
            <Enabled>true</Enabled>
            <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="Microsoft-Windows-NetworkProfile/Operational"&gt;&lt;Select Path="Microsoft-Windows-NetworkProfile/Operational"&gt;*[System[(EventID=10000)]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
            <Delay>PT1M</Delay>
        </EventTrigger>
    </Triggers>
    <Principals>
        <Principal id="Author">
            <UserId>$userId</UserId>
            <LogonType>InteractiveToken</LogonType>
            <RunLevel>LeastPrivilege</RunLevel>
        </Principal>
    </Principals>
    <Settings>
        <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
        <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
        <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
        <AllowHardTerminate>true</AllowHardTerminate>
        <StartWhenAvailable>true</StartWhenAvailable>
        <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
        <IdleSettings>
            <StopOnIdleEnd>false</StopOnIdleEnd>
            <RestartOnIdle>false</RestartOnIdle>
        </IdleSettings>
        <AllowStartOnDemand>true</AllowStartOnDemand>
        <Enabled>true</Enabled>
        <Hidden>false</Hidden>
        <RunOnlyIfIdle>false</RunOnlyIfIdle>
        <WakeToRun>false</WakeToRun>
        <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
        <Priority>7</Priority>
    </Settings>
    <Actions Context="Author">
        <Exec>
            <Command>wscript.exe</Command>
            <Arguments>//nologo &quot;$ScriptPath&quot; apply /silent</Arguments>
        </Exec>
    </Actions>
</Task>
"@
        $tempXmlPath = [System.IO.Path]::GetTempFileName() + ".xml"
        [System.IO.File]::WriteAllText($tempXmlPath, $taskXml)

        $result = schtasks.exe /Create /TN $TaskName /XML $tempXmlPath /F
        Remove-Item $tempXmlPath -Force -ErrorAction SilentlyContinue

        if ($LASTEXITCODE -eq 0) {
            Write-ColorOutput "[+] Scheduled task created successfully" "Green"
            Write-ColorOutput "    - Daily at 9:00 AM" "Gray"
            Write-ColorOutput "    - At user logon" "Gray"
            Write-ColorOutput "    - 1 minute after network connects" "Gray"
            return $true
        } else {
            Write-ColorOutput "[X] Failed to create scheduled task" "Red"
            return $false
        }
    } catch {
        Write-ColorOutput "[X] Failed to create scheduled task: $($_.Exception.Message)" "Red"
        return $false
    }
}