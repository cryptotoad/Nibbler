Attribute VB_Name = "modCryptoAlgs"
Public Function RC4(ByVal Expression As String, ByVal Password As String) As String
On Error Resume Next
Dim RB(0 To 255) As Integer, x As Long, y As Long, Z As Long, KEY() As Byte, ByteArray() As Byte, Temp As Byte
If Len(Password) = 0 Then
    Exit Function
End If
If Len(Expression) = 0 Then
    Exit Function
End If
If Len(Password) > 256 Then
    KEY() = StrConv(Left$(Password, 256), vbFromUnicode)
Else
    KEY() = StrConv(Password, vbFromUnicode)
End If
For x = 0 To 255
    RB(x) = x
Next x
x = 0
y = 0
Z = 0
For x = 0 To 255
    y = (y + RB(x) + KEY(x Mod Len(Password))) Mod 256
    Temp = RB(x)
    RB(x) = RB(y)
    RB(y) = Temp
Next x
x = 0
y = 0
Z = 0
ByteArray() = StrConv(Expression, vbFromUnicode)
For x = 0 To Len(Expression)
    y = (y + 1) Mod 256
    Z = (Z + RB(y)) Mod 256
    Temp = RB(y)
    RB(y) = RB(Z)
    RB(Z) = Temp
    ByteArray(x) = ByteArray(x) Xor (RB((RB(y) + RB(Z)) Mod 256))
Next x
RC4 = StrConv(ByteArray, vbUnicode)
End Function


Public Function XOREnc(DataIn As String, CodeKey) As String
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer, intXOrValue2 As Integer


    For lonDataPtr = 1 To Len(DataIn)
        'The first value to be XOr-ed comes from
        '     the data to be encrypted
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        'The second value comes from the code ke
        '     y
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
   XOREnc = strDataOut
End Function






