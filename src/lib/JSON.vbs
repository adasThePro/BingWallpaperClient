' JSON parsing module for VBScript
' Simplified to work with basic JSON structures
' and 0 dependencies

Dim m_strJSON
Dim m_lngIndex

Function ParseJSON(strJSON)
    On Error Resume Next
    
    m_strJSON = strJSON
    m_lngIndex = 1

    SkipWhitespace
    Set ParseJSON = ParseValue()
    
    If Err.Number <> 0 Then
        Set ParseJSON = Nothing
    End If
    
    On Error GoTo 0
End Function

Private Function ParseValue()
    On Error Resume Next
    
    SkipWhitespace
    
    Dim ch
    ch = PeekChar()
    
    Select Case ch
        Case "{"
            Set ParseValue = ParseObject()
        Case "["
            Set ParseValue = ParseArray()
        Case """"
            ParseValue = ParseString()
        Case "t", "f"
            ParseValue = ParseBoolean()
        Case "n"
            ParseValue = ParseNull()
        Case Else
            If ch >= "0" And ch <= "9" Or ch = "-" Then
                ParseValue = ParseNumber()
            Else
                Err.Raise 5, , "Unexpected character: " & ch
            End If
    End Select
    
    On Error GoTo 0
End Function

Private Function ParseObject()
    Dim objDict
    Set objDict = CreateObject("Scripting.Dictionary")
    
    ConsumeChar
    SkipWhitespace
    
    If PeekChar() = "}" Then
        ConsumeChar
        Set ParseObject = objDict
        Exit Function
    End If
    
    Do
        SkipWhitespace

        Dim strKey
        strKey = ParseString()
        
        SkipWhitespace
        
        If ConsumeChar() <> ":" Then
            Err.Raise 5, , "Expected ':'"
        End If
        
        SkipWhitespace

        Dim varValue, ch
        ch = PeekChar()
        
        If ch = "{" Or ch = "[" Then
            Set varValue = ParseValue()
            Set objDict(strKey) = varValue
        Else
            varValue = ParseValue()
            objDict(strKey) = varValue
        End If
        
        SkipWhitespace

        ch = PeekChar()
        If ch = "," Then
            ConsumeChar
        ElseIf ch = "}" Then
            ConsumeChar
            Exit Do
        Else
            Err.Raise 5, , "Expected ',' or '}'"
        End If
    Loop
    
    Set ParseObject = objDict
End Function

Private Function ParseArray()
    Dim objDict
    Set objDict = CreateObject("Scripting.Dictionary")
    
    ConsumeChar
    SkipWhitespace
    
    Dim lngIndex
    lngIndex = 0

    If PeekChar() = "]" Then
        ConsumeChar
        objDict("Count") = 0
        Set ParseArray = objDict
        Exit Function
    End If
    
    Do
        SkipWhitespace

        Dim varValue, chNext, ch
        chNext = PeekChar()
        
        If chNext = "{" Or chNext = "[" Then
            Set varValue = ParseValue()
            Set objDict(CStr(lngIndex)) = varValue
        Else
            varValue = ParseValue()
            objDict(CStr(lngIndex)) = varValue
        End If
        
        lngIndex = lngIndex + 1
        
        SkipWhitespace

        ch = PeekChar()
        If ch = "," Then
            ConsumeChar
        ElseIf ch = "]" Then
            ConsumeChar
            Exit Do
        Else
            Err.Raise 5, , "Expected ',' or ']'"
        End If
    Loop
    
    objDict("Count") = lngIndex
    Set ParseArray = objDict
End Function

Private Function ParseString()
    ConsumeChar
    
    Dim strResult
    strResult = ""
    
    Do While m_lngIndex <= Len(m_strJSON)
        Dim ch
        ch = Mid(m_strJSON, m_lngIndex, 1)
        m_lngIndex = m_lngIndex + 1
        
        If ch = """" Then
            Exit Do
        ElseIf ch = "\" Then
            If m_lngIndex <= Len(m_strJSON) Then
                ch = Mid(m_strJSON, m_lngIndex, 1)
                m_lngIndex = m_lngIndex + 1
                
                Select Case ch
                    Case """"
                        strResult = strResult & """"
                    Case "\"
                        strResult = strResult & "\"
                    Case "/"
                        strResult = strResult & "/"
                    Case "b"
                        strResult = strResult & Chr(8)
                    Case "f"
                        strResult = strResult & Chr(12)
                    Case "n"
                        strResult = strResult & vbLf
                    Case "r"
                        strResult = strResult & vbCr
                    Case "t"
                        strResult = strResult & vbTab
                    Case "u"
                        Dim strHex
                        strHex = Mid(m_strJSON, m_lngIndex, 4)
                        m_lngIndex = m_lngIndex + 4
                        strResult = strResult & ChrW("&H" & strHex)
                    Case Else
                        strResult = strResult & ch
                End Select
            End If
        Else
            strResult = strResult & ch
        End If
    Loop
    
    ParseString = strResult
End Function

Private Function ParseNumber()
    Dim strNumber
    strNumber = ""
    
    Do While m_lngIndex <= Len(m_strJSON)
        Dim ch
        ch = Mid(m_strJSON, m_lngIndex, 1)
        
        If (ch >= "0" And ch <= "9") Or ch = "-" Or ch = "+" Or ch = "." Or ch = "e" Or ch = "E" Then
            strNumber = strNumber & ch
            m_lngIndex = m_lngIndex + 1
        Else
            Exit Do
        End If
    Loop
    
    If InStr(strNumber, ".") > 0 Or InStr(LCase(strNumber), "e") > 0 Then
        ParseNumber = CDbl(strNumber)
    Else
        ParseNumber = CLng(strNumber)
    End If
End Function

Private Function ParseBoolean()
    If Mid(m_strJSON, m_lngIndex, 4) = "true" Then
        m_lngIndex = m_lngIndex + 4
        ParseBoolean = True
    ElseIf Mid(m_strJSON, m_lngIndex, 5) = "false" Then
        m_lngIndex = m_lngIndex + 5
        ParseBoolean = False
    Else
        Err.Raise 5, , "Invalid boolean"
    End If
End Function

Private Function ParseNull()
    If Mid(m_strJSON, m_lngIndex, 4) = "null" Then
        m_lngIndex = m_lngIndex + 4
        ParseNull = Null
    Else
        Err.Raise 5, , "Invalid null"
    End If
End Function

Private Sub SkipWhitespace()
    Do While m_lngIndex <= Len(m_strJSON)
        Dim ch
        ch = Mid(m_strJSON, m_lngIndex, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            m_lngIndex = m_lngIndex + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function PeekChar()
    If m_lngIndex <= Len(m_strJSON) Then
        PeekChar = Mid(m_strJSON, m_lngIndex, 1)
    Else
        PeekChar = ""
    End If
End Function

Private Function ConsumeChar()
    If m_lngIndex <= Len(m_strJSON) Then
        ConsumeChar = Mid(m_strJSON, m_lngIndex, 1)
        m_lngIndex = m_lngIndex + 1
    Else
        ConsumeChar = ""
    End If
End Function

Function ToJSON(objValue)
    Dim strJSON

    If IsNull(objValue) Then
        ToJSON = "null"
    ElseIf IsEmpty(objValue) Then
        ToJSON = "null"
    ElseIf IsObject(objValue) Then
        If TypeName(objValue) = "Dictionary" Then
            strJSON = "{"
            Dim arrKeys, i, key, bFirst
            arrKeys = objValue.Keys
            bFirst = True

            For i = 0 To UBound(arrKeys)
                key = arrKeys(i)

                If Not bFirst Then
                    strJSON = strJSON & ","
                End If
                bFirst = False

                strJSON = strJSON & """" & key & """:" & ToJSON(objValue(key))
            Next

            strJSON = strJSON & "}"
            ToJSON = strJSON
        Else
            ToJSON = "null"
        End If
    ElseIf VarType(objValue) = vbBoolean Then
        If objValue Then
            ToJSON = "true"
        Else
            ToJSON = "false"
        End If
    ElseIf VarType(objValue) = vbString Then
        ToJSON = """" & JSONEscape(objValue) & """"
    ElseIf IsNumeric(objValue) Then
        ToJSON = CStr(objValue)
    Else
        ToJSON = """" & JSONEscape(CStr(objValue)) & """"
    End If
End Function

Function JSONEscape(strValue)
    Dim strResult, i, ch 
    strResult = strValue

    strResult = Replace(strResult, "\", "\\")
    strResult = Replace(strResult, """", "\""")
    strResult = Replace(strResult, vbCrLf, "\n")
    strResult = Replace(strResult, vbCr, "\r")
    strResult = Replace(strResult, vbLf, "\n")
    strResult = Replace(strResult, vbTab, "\t")

    JSONEscape = strResult
End Function