VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSploit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public retOffset As String      ' return offset in hex
Public shellcode As String      ' shellcode in hex
Public evilPacket As String     ' our final packet is stored here
Public bufferLen As Integer     ' the length we have to work with
Public prepend As Boolean       ' do we have any data to prepend?
Public prependBytes As String   ' What data are we prepending?
'Anything below here is just for flavor.
Public expName As String
Public CVEID As String
Public OSINFO As String




Private Sub Class_Initialize()

    'Debug.Print "New exploit initialized!"
    
End Sub

Public Function craftSploitCode()

    'MsgBox makeNops(10)
    
End Function

Private Function willShellcodeFit() As Boolean

    'returns true if we have enough space for our shellcode
    If countPadding + Len(shellcode) <= bufferLen Then
        willShellcodeFit = True
    Else
        willShellcodeFit = False
    End If
    
End Function

Private Function countPadding() As Integer

    'returns the total number of nops we're going to need.
    countPadding = bufferLen - Len(shellcode) - Len(prependBytes)
End Function

Private Function makeNops(count As Integer) As String

    'creates some nops for us
    Dim myNops As String
    
    For I = 1 To count
        myNops = myNops & HexToString(90)
    Next I
    
    makeNops = myNops
    
End Function
