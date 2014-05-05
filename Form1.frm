VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   Caption         =   "Nibbler"
   ClientHeight    =   4170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   3450
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Key entered in hexadecimal"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   3375
   End
   Begin VB.OptionButton Option4 
      Caption         =   "N/A"
      Height          =   195
      Left            =   2640
      TabIndex        =   14
      Top             =   3000
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      Caption         =   "XOR"
      Height          =   195
      Left            =   1800
      TabIndex        =   13
      Top             =   3000
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "AES"
      Height          =   195
      Left            =   960
      TabIndex        =   12
      Top             =   3000
      Width           =   735
   End
   Begin VB.OptionButton optRC4 
      Caption         =   "RC4"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox tEncKey 
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Listen"
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtLocal 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Text            =   "999"
      Top             =   360
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock iSpeak 
      Left            =   3480
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Connect Client"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Text            =   "998"
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock iListen 
      Left            =   3480
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      Top             =   3840
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Encryption Key"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "Port (L)"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Port (R)"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "SYN IP"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   " Connection analysis and fuzzing utility"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   3840
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public clientStream As String
Dim clientStream2() As Byte
Public botStream As String
Dim botStream2() As Byte
Public isfirst As Boolean
Public inboundFilter As clsPacketFilter
Public outboundFilter As clsPacketFilter


Private Sub Command1_Click()
iSpeak.RemoteHost = txtIP.Text
iSpeak.RemotePort = txtPort.Text
iSpeak.Connect
End Sub


Private Sub Command2_Click()
    inboundFilter.isActive = False
    outboundFilter.isActive = False
End Sub

Private Sub Command5_Click()

If iListen.State <> sckListening And iListen.State <> sckConnected Then
iListen.LocalPort = txtLocal.Text
iListen.Listen
End If

End Sub

Private Sub Form_Load()
    'Set up packet filtering
    Set outboundFilter = New clsPacketFilter
    
    outboundFilter.encoding = 2 ' turn on hex
    outboundFilter.newData = StringToHex(StrConv("Skidszz", vbUnicode))
    outboundFilter.toFilter = StringToHex(StrConv("AntiVir", vbUnicode))
    outboundFilter.isActive = True
    'WPE style packet filters :P
    
    
    Set inboundFilter = New clsPacketFilter
    
    inboundFilter.encoding = 2 ' turn on hex
    inboundFilter.newData = "90" ' let's be crazy and filter nulls to nops for shits n gigs
    inboundFilter.toFilter = "00"
    inboundFilter.isActive = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub iListen_Close()
'MsgBox "Server seems to have crashed. Please restart Nibbler and the server."
    Shape1.BackColor = &HFF&

End Sub

Private Sub iListen_ConnectionRequest(ByVal requestID As Long)
    
    If iListen.State <> sckClosed Then
          iListen.Close
    End If
    
    iListen.Accept (requestID)
    Shape1.BackColor = &HFF00&
    'MsgBox "Connection recieved!"
    'iListen.SendData "Hello?"
    
    Form2.Show
    Form2.Caption = "DataWatcher - " & iListen.RemoteHostIP
    

End Sub

Private Sub iListen_DataArrival(ByVal bytesTotal As Long)
    iListen.GetData botStream
    

    botStream = applyCrypto(botStream, 0)
    Debug.Print "After apply 1 - " & botStream
    
    If strData = Chr(0) Then                    'server says c'mon in
        Exit Sub                                'no need to process further
    End If
    
    botStream = outboundFilter.applyFilters(botStream)
    
    Debug.Print "Filtering data"
    
    Form2.Show
    Form2.text1.Text = Form2.text1.Text & vbNewLine & "[BOT] " & botStream

    Form2.text2.Text = Form2.text2.Text & vbNewLine & "[BOT] " & StringToHex(botStream)
    
    Form2.text3.Text = Form2.text3.Text & vbNewLine & "[BOT] " & StrConv(botStream, vbFromUnicode)
    
    botStream = applyCrypto(botStream, 1)
    
    iSpeak.SendData (botStream)
    
End Sub

Public Function applyCrypto(packetStream As String, ident As Integer) As String
    Debug.Print "Attempting to apply cryptographic stripping. " & ident
    
    If tEncKey.Text = "" Or Option4.Value = True Then
        Debug.Print "No cryptography enabled. Exiting crypto function."
        applyCrypto = packetStream
        Exit Function
    End If

    Dim tempStream As String
    
    If ident = 0 Then ' if we're decrypting
        If optRC4.Value = True Then
            'we decrypt the stream as it comes in
            If Check1.Value = True Then
                'if we're using hexadecimal for the key
                tempStream = RC4(packetStream, HexToString(tEncKey.Text))
            Else
                tempStream = RC4(packetStream, tEncKey.Text)
            End If
        End If
    ElseIf ident = 1 Then ' if we're encrypting
        If optRC4.Value = True Then
            'we decrypt the stream as it comes in
            If Check1.Value = True Then
                'if we're using hexadecimal for the key
                tempStream = RC4(packetStream, HexToString(tEncKey.Text))
            Else
                tempStream = RC4(packetStream, tEncKey.Text)
            End If
        End If
    End If
    
    applyCrypto = tempStream
    
End Function

Private Sub iSpeak_DataArrival(ByVal bytesTotal As Long)

    If iListen.State <> sckConnected Then
        'this is where we "hold the phone"
        
    End If
    
    iSpeak.GetData clientStream
    'iSpeak.GetData clientStream2
    
    
    If clientStream = Chr(0) Then               'server says c'mon in
        Exit Sub
    End If
    
    Debug.Print "****" & clientStream
    
    clientStream = applyCrypto(clientStream, 0) ' run decryption
    
    Debug.Print "****" & clientStream
    
    clientStream = inboundFilter.applyFilters(clientStream)
    
    Form2.Show
    Form2.text1.Text = Form2.text1.Text & vbNewLine & "[SERV] " & clientStream
    
    Form2.text2.Text = Form2.text2.Text & vbNewLine & "[SERV] " & StringToHex(clientStream)
    
    Form2.text3.Text = Form2.text3.Text & vbNewLine & "[SERV] " & StrConv(clientStream, vbFromUnicode)
    
    Debug.Print clientStream
    
    clientStream = applyCrypto(clientStream, 1) ' run encryption so we can pass the stream along as if it was never messed with :P
    
    iListen.SendData (clientStream)
    
    
End Sub

