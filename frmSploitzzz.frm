VERSION 5.00
Begin VB.Form frmSploitzzz 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Craft Exploit"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   Icon            =   "frmSploitzzz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAppend 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Text            =   "4141414141414141"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CheckBox chkAppend 
      Caption         =   "Append After Shellcode"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.FileListBox fileShellcode 
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtPrepend 
      Height          =   285
      Left            =   2520
      TabIndex        =   13
      Text            =   "0F2E445723"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CheckBox chkPrepend 
      Caption         =   "Prepend Before Shellcode"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtNops 
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Text            =   "25"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CheckBox chkNops 
      Caption         =   "Nop Sled (Wheee!)"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send To Packet Forge"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5160
      Width           =   4095
   End
   Begin VB.TextBox txtPacket 
      Height          =   2055
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CheckBox chkEndian 
      Caption         =   "Little Endian"
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      ToolTipText     =   $"frmSploitzzz.frx":15371
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtRet 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Text            =   "41424142"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   4095
   End
   Begin VB.ComboBox cmbShellcode 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "Shellcodes/calc.nib"
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox txtBuf 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "1024"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblSize 
      Height          =   255
      Left            =   1440
      TabIndex        =   18
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Total Length"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Return Adress"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "byte buffer"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Generate Exploit For"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
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
    formSploit.append = chkAppend.Value
    
    If chkPrepend.Value = 1 Then
        formSploit.prependBytes = txtPrepend.Text
    End If
    
    If chkAppend.Value = 1 Then
        formSploit.appendBytes = txtAppend.Text
    End If
    
    If chkNops.Value = 1 Then
        formSploit.useNops = True
        formSploit.nopCount = Val(txtNops.Text)
        'MsgBox formSploit.nopCount
    Else
        formSploit.useNops = False
    End If
    
    If chkEndian.Value = 0 Then
        formSploit.retOffset = txtRet.Text
        'MsgBox formSploit.retOffset
       ' MsgBox HexToStr(formSploit.retOffset)
    Else
        formSploit.retOffset = Right(txtRet.Text, 2)
        formSploit.retOffset = formSploit.retOffset & Left(Right(txtRet.Text, 4), 2)
        formSploit.retOffset = formSploit.retOffset & Left(Right(txtRet.Text, 6), 2)
        formSploit.retOffset = formSploit.retOffset & Left(Right(txtRet.Text, 8), 2)
        'MsgBox formSploit.retOffset
    End If
    
    
    formSploit.craftSploitCode
    
    txtPacket.Text = StringToHex(formSploit.evilPacket)
    
    lblSize.Caption = Len(formSploit.evilPacket)
    
    
End Sub

Private Sub Command2_Click()
    frmMain.hexbox.Text = txtPacket.Text
    frmMain.asciibox.Text = HexToString(txtPacket.Text)
End Sub

Private Sub Form_Load()
    'Now we configure exploit creation menu
    fileShellcode.Path = App.Path & "\shellcodes"
    For i = 0 To fileShellcode.ListCount - 1
        fileShellcode.ListIndex = i
        cmbShellcode.AddItem ("Shellcodes/" & fileShellcode.FileName)
    Next
End Sub

Private Sub Label5_Click()

End Sub

