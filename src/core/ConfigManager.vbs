' Configuration Manager Module

Class ConfigManager
    Public Market
    Public DownloadPath
    Public MaxImages
    Public KeepAllImages
    Public ChangeDesktop
    Public EnableLogging
    Public LastUpdate
    Public Version

    Private m_strConfigPath
    Private m_objFSO
    Private m_objShell

    Private m_arrValidMarkets
    Private m_dictMarkets

    Private Sub Class_Initialize()
        Set m_objFSO = CreateObject("Scripting.FileSystemObject")
        Set m_objShell = CreateObject("WScript.Shell")
        Set m_dictMarkets = CreateObject("Scripting.Dictionary")

        Dim strAppData
        strAppData = m_objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
        m_strConfigPath = m_objFSO.BuildPath(strAppData, "BingWallpaperClient\config.json")

        LoadMarkets

        SetDefaults
    End Sub

    Private Sub Class_Terminate()
        Set m_objFSO = Nothing
        Set m_objShell = Nothing
        Set m_dictMarkets = Nothing
    End Sub

    Private Sub SetDefaults()
        Market = "en-US"
        DownloadPath = m_objShell.ExpandEnvironmentStrings("%USERPROFILE%") & _
                        "\Pictures\BingWallpapers"
        MaxImages = 30
        KeepAllImages = False
        ChangeDesktop = True
        EnableLogging = False
        LastUpdate = ""
        Version = APP_VERSION
    End Sub

    Private Sub LoadMarkets()
        On Error Resume Next

        Dim strScriptDir, strMarketsPath, objFile, strJSON, objMarketsData
        strScriptDir = m_objFSO.GetParentFolderName(WScript.ScriptFullName)
        strMarketsPath = m_objFSO.BuildPath(strScriptDir, "markets.json")

        If m_objFSO.FileExists(strMarketsPath) Then
            Set objFile = m_objFSO.OpenTextFile(strMarketsPath, 1)
            strJSON = objFile.ReadAll
            objFile.Close
            Set objFile = Nothing

            Set objMarketsData = ParseJSON(strJSON)

            If Not objMarketsData Is Nothing And objMarketsData.Exists("markets") Then
                Dim objMarkets, i, objMarket
                Set objMarkets = objMarketsData("markets")

                ReDim m_arrValidMarkets(objMarkets("Count") - 1)

                For i = 0 To objMarkets("Count") - 1
                    Set objMarket = objMarkets(CStr(i))
                    m_arrValidMarkets(i) = objMarket("code")
                    m_dictMarkets.Add objMarket("code"), objMarket("name")
                Next

                Set objMarkets = Nothing
            End If

            Set objMarketsData = Nothing
        End If

        On Error GoTo 0
    End Sub

    Public Function LoadConfig()
        On Error Resume Next

        If Not m_objFSO.FileExists(m_strConfigPath) Then
            LoadConfig = True
            Exit Function
        End If

        Dim objFile, strJSON, objJSON
        Set objFile = m_objFSO.OpenTextFile(m_strConfigPath, 1)
        strJSON = objFile.ReadAll
        objFile.Close
        Set objFile = Nothing

        If Err.Number <> 0 Then
            LoadConfig = False
            Exit Function
        End If

        Set objJSON = ParseJSON(strJSON)

        If Not objJSON Is Nothing Then
            If objJSON.Exists("Market") Then Market = objJSON("Market")
            If objJSON.Exists("DownloadPath") Then DownloadPath = objJSON("DownloadPath")
            If objJSON.Exists("MaxImages") Then MaxImages = CInt(objJSON("MaxImages"))
            If objJSON.Exists("KeepAllImages") Then KeepAllImages = CBool(objJSON("KeepAllImages"))
            If objJSON.Exists("ChangeDesktop") Then ChangeDesktop = CBool(objJSON("ChangeDesktop"))
            If objJSON.Exists("LastUpdate") Then
                If Not IsNull(objJSON("LastUpdate")) Then
                    LastUpdate = objJSON("LastUpdate")
                Else
                    LastUpdate = ""
                End If
            End If
            If objJSON.Exists("EnableLogging") Then EnableLogging = CBool(objJSON("EnableLogging"))
            If objJSON.Exists("Version") Then Version = objJSON("Version")

            Set objJSON = Nothing
            LoadConfig = True
        Else
            LoadConfig = False
        End If

        On Error GoTo 0
    End Function

    Public Function SaveConfig()
        On Error Resume Next

        Dim strConfigDir, objFile, strJSON
        strConfigDir = m_objFSO.GetParentFolderName(m_strConfigPath)

        If Not m_objFSO.FolderExists(strConfigDir) Then
            m_objFSO.CreateFolder(strConfigDir)
        End If

        strJSON = "{" & vbCrLf & _
                "  ""Market"": """ & Market & """," & vbCrLf & _
                "  ""DownloadPath"": """ & Replace(DownloadPath, "\", "\\") & """," & vbCrLf & _
                "  ""MaxImages"": " & MaxImages & "," & vbCrLf & _
                "  ""KeepAllImages"": " & LCase(CStr(KeepAllImages)) & "," & vbCrLf & _
                "  ""ChangeDesktop"": " & LCase(CStr(ChangeDesktop)) & "," & vbCrLf & _
                "  ""EnableLogging"": " & LCase(CStr(EnableLogging)) & "," & vbCrLf & _
                "  ""LastUpdate"": """ & FormatDateTime(Now) & """," & vbCrLf & _
                "  ""Version"": """ & Version & """" & vbCrLf & _
                "}"
        
        Set objFile = m_objFSO.CreateTextFile(m_strConfigPath, True)
        objFile.Write strJSON
        objFile.Close
        Set objFile = Nothing

        SaveConfig = (Err.Number = 0)
        On Error GoTo 0
    End Function

    Public Function IsValidMarket(strMarket)
        Dim i
        IsValidMarket = False
        For i = LBound(m_arrValidMarkets) To UBound(m_arrValidMarkets)
            If m_arrValidMarkets(i) = strMarket Then
                IsValidMarket = True
                Exit Function
            End If
        Next
    End Function

    Public Function GetValidMarkets()
        GetValidMarkets = m_arrValidMarkets
    End Function

    Public Function GetMarketName(strCode)
        If m_dictMarkets.Exists(strCode) Then
            GetMarketName = m_dictMarkets(strCode)
        Else
            GetMarketName = strCode
        End If
    End Function

    Public Function GetConfigPath()
        GetConfigPath = m_strConfigPath
    End Function

    Public Function ConfigExists()
        ConfigExists = m_objFSO.FileExists(m_strConfigPath)
    End Function
    
End Class