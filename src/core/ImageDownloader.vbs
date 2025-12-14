' Image Downloader Module

Class ImageDownloader
    Private m_objFSO
    Private m_objShell

    Private Sub Class_Initialize()
        Set m_objFSO = CreateObject("Scripting.FileSystemObject")
        Set m_objShell = CreateObject("WScript.Shell")
    End Sub

    Private Sub Class_Terminate()
        Set m_objFSO = Nothing
        Set m_objShell = Nothing
    End Sub

    Public Function GetBingImages(strMarket)
        On Error Resume Next

        Dim objHTTP, strUrl, strResponse, objJSON, arrImages()
        Dim i, intCount

        strUrl = "https://services.bingapis.com/ge-apps/api/v2/bwc/hpimages?mkt=" & strMarket

        WriteLog "Fetching images from Bing API: " & strMarket, LOG_INFO

        Set objHTTP = CreateObject("WinHTTP.WinHttpRequest.5.1")
        If Err.Number <> 0 Then 
            Err.Clear
            Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        End If
        If Err.Number <> 0 Then 
            Err.Clear
            Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
        End If

        If Err.Number <> 0 Then
            WriteLog "Failed to create HTTP object: " & Err.Description, LOG_ERROR
            GetBingImages = Null
            Exit Function
        End If

        With objHTTP
            .Open "GET", strUrl, False
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"      
            .Send

            If Err.Number <> 0 Then
                WriteLog "Failed to fetch Bing images. Error: " & Err.Description, LOG_ERROR
                Err.Clear
                Set objHTTP = Nothing
                GetBingImages = Array()
                Exit Function
            End If

            If .Status <> 200 Then
                WriteLog "Failed to fetch Bing images. Status: " & .Status, LOG_ERROR
                Set objHTTP = Nothing
                GetBingImages = Array()
                Exit Function
            End If
            
            strResponse = .ResponseText
        End With
        Set objHTTP = Nothing

        Set objJSON = ParseJSON(strResponse)

        If objJSON Is Nothing Then
            WriteLog "Failed to parse API response JSON.", LOG_ERROR
            GetBingImages = Array()
            Exit Function
        End If

        If objJSON.Exists("images") Then
            Dim objImages, objImage, dictImage
            Set objImages = objJSON("images")

            intCount = 0
            ReDim arrImages(objImages.Count - 1)

            For i = 0 To objImages.Count - 1
                Set objImage = objImages(CStr(i))

                Set dictImage = CreateObject("Scripting.Dictionary")

                If objImage.Exists("urlbase") Then
                    dictImage.Add "url", objImage("urlbase")
                    Set arrImages(intCount) = dictImage
                    intCount = intCount + 1
                End If
            Next

            If intCount > 0 Then
                ReDim Preserve arrImages(intCount - 1)
            Else
                ReDim arrImages(-1)
            End If

            Set objImages = Nothing
        Else
            ReDim arrImages(-1)
        End If

        Set objJSON = Nothing

        GetBingImages = arrImages

        On Error GoTo 0
    End Function

    Public Function DownloadImage(objImageData, strDestPath)
        On Error Resume Next

        Dim strUrlBase, strHash, strFileName, strFilePath

        If objImageData Is Nothing Then
            DownloadImage = ""
            Exit Function
        End If

        strUrlBase = objImageData("url")
        strFileName = ExtractFilenameFromUrl(strUrlBase)
        strFilePath = m_objFSO.BuildPath(strDestPath, strFileName)

        If m_objFSO.FileExists(strFilePath) Then
            WriteLog "Image already exists: " & strFileName, LOG_INFO
            DownloadImage = strFilePath
            Exit Function
        End If

        WriteLog "Downloading image: " & strFileName, LOG_INFO

        Dim objHTTP, arrData
        Set objHTTP = CreateObject("WinHTTP.WinHttpRequest.5.1")
        If Err.Number <> 0 Then 
            Err.Clear
            Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        End If
        If Err.Number <> 0 Then 
            Err.Clear
            Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
        End If

        If Err.Number <> 0 Then
            WriteLog "Failed to create HTTP object for download", LOG_ERROR
            DownloadImage = ""
            Exit Function
        End If

        With objHTTP
            .Open "GET", strUrlBase, False
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
            .Send

            If Err.Number <> 0 Then
                WriteLog "Failed to download image. Error: " & Err.Description, LOG_ERROR
                Err.Clear
                Set objHTTP = Nothing
                DownloadImage = ""
                Exit Function
            End If

            If .Status <> 200 Then
                WriteLog "Failed to download image. Status: " & .Status, LOG_ERROR
                Set objHTTP = Nothing
                DownloadImage = ""
                Exit Function
            End If
            
            arrData = .ResponseBody
        End With
        Set objHTTP = Nothing

        If Not SaveBinaryFile(strFilePath, arrData) Then
            WriteLog "Failed to save image to disk: " & strFilePath, LOG_ERROR
            DownloadImage = ""
            Exit Function
        End If

        WriteLog "Downloaded successfully: " & strFileName, LOG_SUCCESS
        DownloadImage = strFilePath

        On Error GoTo 0
    End Function

    Private Function SaveBinaryFile(strPath, arrData)
        On Error Resume Next

        Dim objStream
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1
        objStream.Open
        objStream.Write arrData
        objStream.SaveToFile strPath, 2
        objStream.Close
        Set objStream = Nothing

        SaveBinaryFile = (Err.Number = 0)
        On Error GoTo 0
    End Function

    Private Function ExtractFilenameFromUrl(strUrl)
        Dim strFilename, intIdPos, intExtPos, strId
        
        intIdPos = InStr(strUrl, "id=")
        If intIdPos > 0 Then
            strId = Mid(strUrl, intIdPos + 3)
            intExtPos = InStr(strId, "&")
            If intExtPos > 0 Then
                strId = Left(strId, intExtPos - 1)
            End If
            strFilename = strId
        Else
            strFilename = m_objFSO.GetFileName(strUrl)
            intExtPos = InStr(strFilename, "?")
            If intExtPos > 0 Then
                strFilename = Left(strFilename, intExtPos - 1)
            End If
        End If

        If LCase(Right(strFilename, 4)) <> ".jpg" Then
            strFilename = strFilename & ".jpg"
        End If
        
        ExtractFilenameFromUrl = strFilename
    End Function

End Class
