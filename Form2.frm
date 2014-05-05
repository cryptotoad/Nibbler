VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form2 
   Caption         =   "DataWatcher - "
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form2"
   ScaleHeight     =   8070
   ScaleWidth      =   8985
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox text1 
      Height          =   3615
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form2.frx":0000
   End
   Begin RichTextLib.RichTextBox text2 
      Height          =   3615
      Left            =   0
      TabIndex        =   3
      Top             =   4320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form2.frx":0082
   End
   Begin RichTextLib.RichTextBox text3 
      Height          =   3615
      Left            =   4560
      TabIndex        =   4
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   6376
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form2.frx":0104
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Hexadecimal"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Unicode"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ASCII"
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
pForge.Show

End Sub
