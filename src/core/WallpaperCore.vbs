' Wallpaper Core Module

Class WallpaperCore
    Private m_objFSO
    Private m_objShell
    Private m_objConfig

    Private Sub Class_Initialize()
        Set m_objFSO = CreateObject("Scripting.FileSystemObject")
        Set m_objShell = CreateObject("WScript.Shell")
    End Sub

    Private Sub Class_Terminate()
        Set m_objFSO = Nothing
        Set m_objShell = Nothing
        Set m_objConfig = Nothing
    End Sub

    Public Sub Initialize(objConfig)
        Set m_objConfig = objConfig
    End Sub

    Public Function SetDesktopWallpaper(strImagePath)
        On Error Resume Next

        If Not m_objFSO.FileExists(strImagePath) Then
            WriteLog "Image not found: " & strImagePath, LOG_ERROR
            SetDesktopWallpaper = False
            Exit Function
        End If

        SetDesktopWallpaper = SetDesktopWallpaperUser(strImagePath)

        On Error GoTo 0
    End Function

    Private Function SetDesktopWallpaperUser(strImagePath)
        On Error Resume Next

        Dim strSetWallpaperExe, intExitCode
        strSetWallpaperExe = m_objFSO.GetParentFolderName(WScript.ScriptFullName) & "\SetWallpaper.exe"
        
        If Not m_objFSO.FileExists(strSetWallpaperExe) Then
            WriteLog "SetWallpaper.exe not found at: " & strSetWallpaperExe, LOG_ERROR
            SetDesktopWallpaperUser = False
            Exit Function
        End If
        
        intExitCode = m_objShell.Run("""" & strSetWallpaperExe & """ """ & strImagePath & """", 0, True)
        
        If intExitCode = 0 Then
            SetDesktopWallpaperUser = True
        Else
            WriteLog "SetWallpaper.exe failed with exit code: " & intExitCode, LOG_ERROR
            SetDesktopWallpaperUser = False
        End If

        On Error GoTo 0
    End Function

End Class