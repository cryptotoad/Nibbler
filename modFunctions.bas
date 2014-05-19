Attribute VB_Name = "modFunctions"


Public Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

Public Function HexToString(ByVal HexToStri As String) As String
    Dim strTemp   As String
    Dim strReturn As String
    Dim i         As Long
    
    For i = 1 To Len(HexToStri) Step 2
        strTemp = Chr$(Val("&H" & Mid$(HexToStri, i, 2)))
        strReturn = strReturn & strTemp
    Next i
    
    HexToString = strReturn
End Function

    Public Function HexToStr(ByVal Data As String) As String
    Dim Buffer As String
     
        If Len(Data) Mod 2 <> 0 Then
            HexToStr = vbNullString
        Else
            For i = 1 To Len(Data) - 1 Step 2
                Buffer = Buffer & Chr("&H" & Mid(Data, i, 2))
            Next i
            HexToStr = Buffer
        End If
    End Function

Public Function StringToHex(ByVal StrToHex As String) As String
    Dim strTemp   As String
    Dim strReturn As String
    Dim i         As Long
    
    For i = 1 To Len(StrToHex)
        strTemp = Hex$(Asc(Mid$(StrToHex, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & Space$(1) & strTemp
    Next i
    
    If Left(strReturn, 1) = " " Then
        strReturn = Right(strReturn, Len(strReturn) - 1)
    End If
    
    StringToHex = strReturn
End Function


