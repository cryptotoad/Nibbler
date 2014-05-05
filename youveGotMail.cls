VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "youveGotMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Cryptotoad - 2014

'Implement hold the phone functionality
'wait for the recieving socket to connect before sending traffic
'All functions return true if successful, false otherwise

Public packetsCount As Integer ' the number of packets we're holding
Dim holdingMessages() As String ' A string array containing our messages
Public myData As String ' this is where we will fetch our output

Public Function addBuffer() As Boolean
    On Error GoTo retFalse
    ReDim Preserve holdingMessages(UBound(holdingMessages) + 1)
    addBuffer = True
    Exit Function
    
retFalse:
    addBuffer = False
End Function

Public Function remBuffer() As Boolean
    On Error GoTo retFalse
    ReDim Preserve holdingMessages(UBound(holdingMessages) - 1)
    remBuffer = True
    Exit Function
    
retFalse:
    remBuffer = False
    
End Function

Public Function pushMessage() As Boolean
    If addBuffer Then
        holdingMessages(UBound(holdingMessages)) = myData
    Else
        Debug.Print "Error increasing buffer length!"
    End If
End Function

Public Function popMessage() As Boolean
    Dim strMessage As String
    
    strMessage = holdingMessages(UBound(holdingMessages))
    
    remBuffer
    
    myData = strMessage
End Function

Public Sub clearAll()
    ReDim holdingMessages(0)
    myData = vbNullString
End Sub


