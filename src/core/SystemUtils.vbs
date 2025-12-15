' System Utilities Module

Const LOG_INFO = "Info"
Const LOG_SUCCESS = "Success"
Const LOG_WARNING = "Warning"
Const LOG_ERROR = "Error"

Sub WriteLog(strMessage, strLevel)
    Dim objConfig, strLogPath, strLogFile, objLogFile
    Dim strTimestamp, strLine

    If Not IsSilentMode() Then
        WScript.Echo "[" & Now & "] [" & strLevel & "] " & strMessage
    End If

    Set objConfig = New ConfigManager
    If Not objConfig.LoadConfig() Then Exit Sub

    If Not objConfig.EnableLogging Then
        Set objConfig = Nothing
        Exit Sub
    End If

    On Error Resume Next

    strLogPath = objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\BingWallpaperClient\logs"

    If Not objFSO.FolderExists(strLogPath) Then
        objFSO.CreateFolder(strLogPath)
    End If

    strLogFile = strLogPath & "\" & Year(Now) & "-" & _ 
                    Right("0" & Month(Now), 2) & "-" & _ 
                    Right("0" & Day(Now), 2) & ".log"

    strTimestamp = Year(Now) & "-" & Right("0" & Month(Now), 2) & "-" & _
                    Right("0" & Day(Now), 2) & " " & _ 
                    Right("0" & Hour(Now), 2) & ":" & _ 
                    Right("0" & Minute(Now), 2) & ":" & _ 
                    Right("0" & Second(Now), 2)
    
    strLine = "[" & strTimestamp & "] [" & strLevel & "] " & strMessage

    Set objLogFile = objFSO.OpenTextFile(strLogFile, 8, True)
    objLogFile.WriteLine strLine
    objLogFile.Close
    Set objLogFile = Nothing

    On Error GoTo 0
    Set objConfig = Nothing
End Sub

Function IsSilentMode()
    Dim i 
    IsSilentMode = False
    For i = 0 To WScript.Arguments.Count - 1
        If LCase(WScript.Arguments(i)) = "/silent" Then
            IsSilentMode = True
            Exit Function
        End If
    Next
End Function

Function TestNetworkConnectivity()
    On Error Resume Next
    Dim objPing, objStatus

    Set objPing = GetObject("winmgmts:\\.\root\cimv2")
    Set objStatus = objPing.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address = '1.1.1.1' AND Timeout = 3000")

    Dim objResult
    For Each objResult In objStatus
        If objResult.StatusCode = 0 Then
            TestNetworkConnectivity = True
            Exit Function
        End If
    Next

    TestNetworkConnectivity = False

    On Error GoTo 0
    Set objStatus = Nothing
    Set objPing = Nothing
End Function

Function TestBingApiReachability(strMarket)
    TestBingApiReachability = False
    On Error Resume Next

    Dim objHTTP, strUrl, intStatus
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    If Err.Number <> 0 Then
        Err.Clear
        Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
    End If
    
    If Err.Number <> 0 Or objHTTP Is Nothing Then
        On Error GoTo 0
        Exit Function
    End If
    
    strUrl = "https://services.bingapis.com/ge-apps/api/v2/bwc/hpimages?mkt=" & strMarket
    
    With objHTTP
        .Open "HEAD", strUrl, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        .Send
        
        If Err.Number <> 0 Then
            On Error GoTo 0
            Set objHTTP = Nothing
            Exit Function
        End If
        
        intStatus = .Status
        If Err.Number = 0 And intStatus = 200 Then
            TestBingApiReachability = True
        End If
    End With
    
    On Error GoTo 0
    Set objHTTP = Nothing
End Function

Function FormatDateTime(dtDate)
    FormatDateTime = Year(dtDate) & "-" & _
                     Right("0" & Month(dtDate), 2) & "-" & _
                     Right("0" & Day(dtDate), 2) & " " & _
                     Right("0" & Hour(dtDate), 2) & ":" & _
                     Right("0" & Minute(dtDate), 2) & ":" & _
                     Right("0" & Second(dtDate), 2)
End Function