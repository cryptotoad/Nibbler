Attribute VB_Name = "modFunctions"
Public Function RC4(ByVal Expression As String, ByVal Password As String) As String
On Error Resume Next
Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
If Len(Password) = 0 Then
    Exit Function
End If
If Len(Expression) = 0 Then
    Exit Function
End If
If Len(Password) > 256 Then
    Key() = StrConv(Left$(Password, 256), vbFromUnicode)
Else
    Key() = StrConv(Password, vbFromUnicode)
End If
For X = 0 To 255
    RB(X) = X
Next X
X = 0
Y = 0
Z = 0
For X = 0 To 255
    Y = (Y + RB(X) + Key(X Mod Len(Password))) Mod 256
    Temp = RB(X)
    RB(X) = RB(Y)
    RB(Y) = Temp
Next X
X = 0
Y = 0
Z = 0
ByteArray() = StrConv(Expression, vbFromUnicode)
For X = 0 To Len(Expression)
    Y = (Y + 1) Mod 256
    Z = (Z + RB(Y)) Mod 256
    Temp = RB(Y)
    RB(Y) = RB(Z)
    RB(Z) = Temp
    ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
Next X
RC4 = StrConv(ByteArray, vbUnicode)
End Function

Public Function HexToString(ByVal HexToStr As String) As String
    Dim strTemp   As String
    Dim strReturn As String
    Dim I         As Long
    
    For I = 1 To Len(HexToStr) Step 3
        strTemp = Chr$(Val("&H" & Mid$(HexToStr, I, 2)))
        strReturn = strReturn & strTemp
    Next I
    
    HexToString = strReturn
End Function
    Public Function HexToStr(ByVal Data As String) As String
    Dim Buffer As String
     
        If Len(Data) Mod 2 <> 0 Then
            HexToStr = vbNullString
        Else
            For I = 1 To Len(Data) - 1 Step 2
                Buffer = Buffer & Chr("&H" & Mid(Data, I, 2))
            Next I
            HexToStr = Buffer
        End If
    End Function

Public Function StringToHex(ByVal StrToHex As String) As String
    Dim strTemp   As String
    Dim strReturn As String
    Dim I         As Long
    
    For I = 1 To Len(StrToHex)
        strTemp = Hex$(Asc(Mid$(StrToHex, I, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & Space$(1) & strTemp
    Next I
    
    If Left(strReturn, 1) = " " Then
        strReturn = Right(strReturn, Len(strReturn) - 1)
    End If
    
    StringToHex = strReturn
End Function


