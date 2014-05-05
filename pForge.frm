VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form pForge 
   Caption         =   "Packet Forge"
   ClientHeight    =   6045
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form3"
   ScaleHeight     =   6045
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Hex to Text"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Text to Hex"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send to Server"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   5640
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send to bot"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5640
      Width           =   3495
   End
   Begin RichTextLib.RichTextBox asciibox 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4048
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"pForge.frx":0000
   End
   Begin RichTextLib.RichTextBox hexbox 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"pForge.frx":0082
   End
   Begin VB.Label Label2 
      Caption         =   "Hex Rendering"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "ASCII Rendering"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "pForge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.iListen.SendData HexToString(hexbox.Text)
    Form2.text1.Text = Form2.text1.Text & vbNewLine & "[SPOOF][SERV] " & HexToString(hexbox.Text)
    Form2.text2.Text = Form2.text2.Text & vbNewLine & "[SPOOF][SERV] " & hexbox.Text
End Sub

Private Sub Command2_Click()
    Form1.iSpeak.SendData HexToString(hexbox.Text)
    Form2.text1.Text = Form2.text1.Text & vbNewLine & "[SPOOF][BOT] " & HexToString(hexbox.Text)
    Form2.text2.Text = Form2.text2.Text & vbNewLine & "[SPOOF][BOT] " & hexbox.Text
End Sub

Private Sub Command3_Click()
    hexbox.Text = StringToHex(asciibox.Text)
End Sub

Private Sub Command4_Click()
    asciibox.Text = HexToString(hexbox.Text)
End Sub

Private Sub Form_Load()
'frmFilters.Show

End Sub
