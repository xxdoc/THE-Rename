VERSION 5.00
Begin VB.Form fView 
   AutoRedraw      =   -1  'True
   Caption         =   "View Picture"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   Icon            =   "fView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   4845
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "fView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
 Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 Unload Me
End Sub

Private Sub Form_Load()
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
