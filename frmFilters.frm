VERSION 5.00
Begin VB.Form frmFilters 
   Caption         =   "Filter Management"
   ClientHeight    =   5925
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form3"
   ScaleHeight     =   5925
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    ListView1.ListItems(1).ListSubItems.Add "Test", 0
    
    
End Sub
