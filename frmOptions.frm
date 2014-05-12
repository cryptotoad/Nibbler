VERSION 5.00
Begin VB.Form frmOptions 
   Caption         =   "Options"
   ClientHeight    =   3180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3180
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Settings"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CheckBox chkArrival 
         Caption         =   "Bring to top on new data"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton btnDefaults 
         Caption         =   "Defaults"
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton btnApply 
         Caption         =   "Apply"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2640
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
