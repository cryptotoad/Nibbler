VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacketFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public filterType As Integer ' 0 = string, 1 = byte
Public encoding As Integer ' 0 = standard, 1 = unicode, 2 = hexadecimal

Public toFilter As String ' text to find in the packet
Public newData As String ' text to replace in the packet

Public isActive As Boolean



Public Function applyFilters(packetStream As String) As String

    Dim hexPlaceholder As String
    
    If Not isActive Then
        applyFilters = packetStream
        Exit Function
    End If

    'Applies our filters to the packet stream of our choice
    Dim packets As String
    
    packets = packetStream
    
    Debug.Print "[p]" & packets
    
    If encoding = 1 Then ' unicode
        packets = Replace(packets, toFilter, newData)
    ElseIf encoding = 0 Then
        hexPlaceholder = StringToHex(packets)
        Debug.Print hexPlaceholder
        hexPlaceholder = Replace(hexPlaceholder, StringToHex(toFilter), StringToHex(newData))
        Debug.Print hexPlaceholder
        packets = HexToString(hexPlaceholder)
    ElseIf encoding = 2 Then
        Debug.Print "Attempting to filter HEX data in packet"
        hexPlaceholder = StringToHex(packets)
        Debug.Print hexPlaceholder
        hexPlaceholder = Replace(hexPlaceholder, toFilter, newData)
        Debug.Print hexPlaceholder
        packets = HexToString(hexPlaceholder)
    Else
        Debug.Print "Dafuq? select your encoding first."
    End If
    
    Debug.Print "[p2]" & packets
    
    applyFilters = packets
    
End Function
