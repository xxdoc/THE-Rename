VERSION 5.00
Begin VB.Form fwnew 
   AutoRedraw      =   -1  'True
   Caption         =   "What's new ?"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   Icon            =   "fwnew.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4395
   End
End
Attribute VB_Name = "fwnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim WhatsNew As String
Me.Move (Screen.width - Me.width) / 2, (Screen.height - Me.height) / 2
WhatsNew = "Beta 1.7a,   No new important features but some corrections" & vbCrLf
Text1.Text = WhatsNew
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 Text1.left = 0
 Text1.top = 0
 Text1.width = Me.ScaleWidth
 Text1.height = Me.ScaleHeight
End Sub
