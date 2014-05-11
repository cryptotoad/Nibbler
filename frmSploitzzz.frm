VERSION 5.00
Begin VB.Form frmSploitzzz 
   Caption         =   "Craft Exploit"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Options"
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send To Packet Forge"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   3855
   End
   Begin VB.TextBox Text3 
      Height          =   1935
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Little Endian"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      ToolTipText     =   $"frmSploitzzz.frx":0000
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "008043f2"
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generate"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Text            =   "Shellcodes/dlExec"
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Text            =   "1024"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Return Adress"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "byte buffer"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1095
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
