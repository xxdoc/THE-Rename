VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10950
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   10950
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2790
      ItemData        =   "Form1.frx":0000
      Left            =   5640
      List            =   "Form1.frx":006A
      TabIndex        =   2
      Top             =   300
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   4260
      TabIndex        =   1
      Top             =   3540
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   2790
      Left            =   128
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   300
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Dim debut As Integer
    Dim i As Integer
    debut = 220
    Text1.Text = ""

    For i = 0 To List1.ListCount - 1
        Text1.Text = Text1.Text & "LngCmd(" & i + debut & ", 1)=" & Len(Trim(List1.List(i))) - 1 & vbTab & "' " & Trim(List1.List(i)) & vbCrLf
        Text1.Text = Text1.Text & "LngCmd(" & i + debut & ", 2)=-2" & vbCrLf
    Next
End Sub
