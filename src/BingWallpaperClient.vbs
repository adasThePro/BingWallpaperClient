' Bing Wallpaper Client

Option Explicit

Const APP_NAME = "BingWallpaperClient"
Const APP_VERSION = "1.0.2"

Dim objFSO, objShell
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")

Dim strScriptDir, strCoreDir, strLibDir
strScriptDir = objFSO.GetParentFolderName(WScript.ScriptFullName)

strCoreDir = objFSO.BuildPath(strScriptDir, "core")
strLibDir = objFSO.BuildPath(strScriptDir, "lib")

ExecuteGlobal objFSO.OpenTextFile(objFSO.BuildPath(strLibDir, "JSON.vbs"), 1).ReadAll
ExecuteGlobal objFSO.OpenTextFile(objFSO.BuildPath(strCoreDir, "SystemUtils.vbs"), 1).ReadAll
ExecuteGlobal objFSO.OpenTextFile(objFSO.BuildPath(strCoreDir, "ConfigManager.vbs"), 1).ReadAll
ExecuteGlobal objFSO.OpenTextFile(objFSO.BuildPath(strCoreDir, "ImageDownloader.vbs"), 1).ReadAll
ExecuteGlobal objFSO.OpenTextFile(objFSO.BuildPath(strCoreDir, "WallpaperCore.vbs"), 1).ReadAll

Function IIf (bCondition, vTrue, vFalse)
    If bCondition Then
        IIf = vTrue
    Else
        IIf = vFalse
    End If
End Function

Class BingWallpaperApp
    Private m_objConfig
    Private m_objDownloader
    Private m_objWallpaper

    Private Sub Class_Initialize()
        Set m_objConfig = New ConfigManager
        Set m_objDownloader = New ImageDownloader
        Set m_objWallpaper = New WallpaperCore
    End Sub

    Private Sub Class_Terminate()
        Set m_objConfig = Nothing
        Set m_objDownloader = Nothing
        Set m_objWallpaper = Nothing
    End Sub

    Public Sub Run()
        Dim strCommand, bSilentMode, i

        If WScript.Arguments.Count = 0 Then
            strCommand = "help"
        Else
            strCommand = LCase(WScript.Arguments(0))
        End If

        ' Check if running in silent mode (automatic/scheduled task)
        bSilentMode = False
        For i = 0 To WScript.Arguments.Count - 1
            If LCase(WScript.Arguments(i)) = "/silent" Then
                bSilentMode = True
                Exit For
            End If
        Next

        Select Case strCommand
            Case "apply"
                ApplyWallpaper bSilentMode
            Case "status"
                ShowStatus
            Case "config"
                OpenConfig
            Case "help"
                ShowHelp
            Case "version"
                ShowVersion
            Case "uninstall"
                Uninstall
            Case Else
                WScript.Echo "Unknown command: " & strCommand
                WScript.Echo "Use 'help' command to see available commands"
        End Select
    End Sub

    Private Sub ApplyWallpaper(bSilentMode)
        WriteLog "=== Applying Bing Wallpaper ===", LOG_INFO

        If Not m_objConfig.ConfigExists() Then
            WriteLog "Configuration file not found: " & m_objConfig.GetConfigPath(), LOG_ERROR
            Exit Sub
        End If

        If Not m_objConfig.LoadConfig() Then
            WriteLog "Failed to load configuration from: " & m_objConfig.GetConfigPath(), LOG_ERROR
            Exit Sub
        End If

        If Not TestNetworkConnectivity() Then
            WriteLog "No network connection available", LOG_WARNING
            Exit Sub
        End If

        Dim arrImages,  i, objImageData, strImagePath
        arrImages = m_objDownloader.GetBingImages(m_objConfig.Market)

        If UBound(arrImages) < 0 Then
            WriteLog "Failed to fetch images from Bing API", LOG_ERROR
            Exit Sub
        End If

        If Not objFSO.FolderExists(m_objConfig.DownloadPath) Then
            objFSO.CreateFolder(m_objConfig.DownloadPath)
        End If

        Set objImageData = arrImages(0)
        strImagePath = m_objDownloader.DownloadImage(objImageData, m_objConfig.DownloadPath)

        If strImagePath = "" Then
            WriteLog "Failed to download desktop image", LOG_ERROR
            Exit Sub
        End If

        m_objWallpaper.Initialize m_objConfig

        Dim bDesktopSuccess, bShouldApplyWallpaper
        bDesktopSuccess = False
        bShouldApplyWallpaper = m_objConfig.ChangeDesktop Or (Not bSilentMode)

        If bShouldApplyWallpaper Then
            bDesktopSuccess = m_objWallpaper.SetDesktopWallpaper(strImagePath)
        End If

        m_objConfig.LastUpdate = FormatDateTime(Now)
        m_objConfig.SaveConfig

        If bShouldApplyWallpaper Then
            If bDesktopSuccess Then
                WriteLog "Desktop wallpaper updated successfully", LOG_SUCCESS
            Else
                WriteLog "Failed to set desktop wallpaper", LOG_ERROR
            End If
        End If

        WriteLog "=== Wallpaper Update Complete ===", LOG_SUCCESS

        If Not m_objConfig.KeepAllImages Then
            CleanupOldImages
        End If
    End Sub

    Private Sub CleanupOldImages()
        On Error Resume Next

        Dim objFolder, colFiles, arrFiles(), i, objFile
        Set objFolder = objFSO.GetFolder(m_objConfig.DownloadPath)
        Set colFiles = objFolder.Files

        i = 0
        ReDim arrFiles(colFiles.Count - 1)

        For Each objFile In colFiles
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "jpg" Then
                Set arrFiles(i) = objFile
                i = i + 1
            End If
        Next

        If i = 0 Then Exit Sub

        ReDim Preserve arrFiles(i - 1)

        Dim j, k, objTemp
        For j = 0 To UBound(arrFiles) - 1
            For k = j + 1 To UBound(arrFiles)
                If arrFiles(j).DateCreated < arrFiles(k).DateCreated Then
                    Set objTemp = arrFiles(j)
                    Set arrFiles(j) = arrFiles(k)
                    Set arrFiles(k) = objTemp
                End If
            Next
        Next

        For i = m_objConfig.MaxImages To UBound(arrFiles)
            WriteLog "Deleting old image: " & arrFiles(i).Name, LOG_INFO
            arrFiles(i).Delete
        Next

        On Error GoTo 0
    End Sub

    Private Sub ShowStatus()
        If Not m_objConfig.ConfigExists() Then
            WScript.Echo "Configuration not found"
            WScript.Echo "Expected location: " & m_objConfig.GetConfigPath()
            Exit Sub
        End If

        If Not m_objConfig.LoadConfig() Then
            WScript.Echo "Failed to load configuration"
            WScript.Echo "Configuration file may be corrupted"
            Exit Sub
        End If

        WScript.Echo ""
        WScript.Echo "=== " & APP_NAME & " Status ==="
        WScript.Echo ""
        WScript.Echo "Version: " & APP_VERSION
        WScript.Echo "Market: " & m_objConfig.GetMarketName(m_objConfig.Market) & " [" & m_objConfig.Market & "]"
        WScript.Echo "Download Path: " & m_objConfig.DownloadPath
        WScript.Echo "Max Images: " & IIf(m_objConfig.KeepAllImages, "Unlimited", CStr(m_objConfig.MaxImages))
        WScript.Echo "Desktop Wallpaper: " & IIf(m_objConfig.ChangeDesktop, "Enabled", "Disabled")
        WScript.Echo "Logging: " & IIf(m_objConfig.EnableLogging, "Enabled", "Disabled")
        WScript.Echo "Last Update: " & IIf(m_objConfig.LastUpdate = "", "Never", m_objConfig.LastUpdate)
        WScript.Echo ""

        WScript.Echo "Network: " & IIf(TestNetworkConnectivity(), "OK", "Unavailable")
        WScript.Echo "Bing API: " & IIf(TestBingAPIReachability(m_objConfig.Market), "OK", "Unavailable")

        Dim intImageCount
        intImageCount = CountImages()
        WScript.Echo "Downloaded Images: " & intImageCount
        WScript.Echo ""
    End Sub

    Private Sub Uninstall()
        Dim strInstallerPath, objShell
        Set objShell = CreateObject("WScript.Shell")

        strInstallerPath = objFSO.BuildPath(strScriptDir, "installer.ps1")

        If objFSO.FileExists(strInstallerPath) Then
            WScript.Echo "Launching uninstaller..."

            Dim strCommand
            strCommand = "powershell.exe -ExecutionPolicy Bypass -File """ & strInstallerPath & """ -Action uninstall"
            
            objShell.Run strCommand, 1, False
        Else
            WScript.Echo "Installer not found in program directory"
            WScript.Echo "To uninstall manually, delete the following:"
            WScript.Echo " - " & strScriptDir
            WScript.Echo " - Remove 'BingWallpaperSync' task from Task Scheduler"
            WScript.Echo " - Remove from PATH in User Environment Variables"
        End If

        Set objShell = Nothing
    End Sub

    Private Sub OpenConfig()
        Dim objShell
        Set objShell = CreateObject("WScript.Shell")

        If Not m_objConfig.ConfigExists() Then
            WScript.Echo "Configuration not found"
            WScript.Echo "Expected location: " & m_objConfig.GetConfigPath()
            Exit Sub
        End If

        objShell.Run """" & m_objConfig.GetConfigPath() & """", 1, False
        Set objShell = Nothing
    End Sub

    Private Sub ShowHelp()
        WScript.Echo ""
        WScript.Echo "=== " & APP_NAME & " v" & APP_VERSION & " ==="
        WScript.Echo ""
        WScript.Echo "Usage: bwc <command> [options]"
        WScript.Echo ""
        WScript.Echo "Commands:"
        WScript.Echo "  apply       - Download and apply new wallpaper"
        WScript.Echo "  status      - Show current status and configuration"
        WScript.Echo "  config      - Open configuration file in default text editor"
        WScript.Echo "  uninstall   - Uninstall the application"
        WScript.Echo "  help        - Show this help message"
        WScript.Echo "  version     - Show version information"
        WScript.Echo ""
        WScript.Echo "Options:"
        WScript.Echo "  /silent     - Run in silent mode (suppress console output)"
        WScript.Echo ""
        WScript.Echo "Examples:"
        WScript.Echo "  bwc apply          - Apply wallpaper now"
        WScript.Echo "  bwc apply /silent  - Apply wallpaper silently"
        WScript.Echo "  bwc status         - Show current status"
        WScript.Echo ""
    End Sub

    Private Sub ShowVersion()
        WScript.Echo APP_NAME & " v" & APP_VERSION
    End Sub

    Private Function CountImages()
        On Error Resume Next

        If Not objFSO.FolderExists(m_objConfig.DownloadPath) Then
            CountImages = 0
            Exit Function
        End If

        Dim objFolder, colFiles, objFile, intCount
        Set objFolder = objFSO.GetFolder(m_objConfig.DownloadPath)
        Set colFiles = objFolder.Files

        intCount = 0
        For Each objFile in colFiles
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "jpg" Then
                intCount = intCount + 1
            End If
        Next

        CountImages = intCount

        On Error GoTo 0
    End Function

End Class

Dim objApp
Set objApp = New BingWallpaperApp
objApp.Run
Set objApp = Nothing