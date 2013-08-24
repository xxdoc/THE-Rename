VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movie Credits"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   1920
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   600
      ScaleHeight     =   1155
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   480
      Width           =   5655
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000007&
         Height          =   6495
         Left            =   -50
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "scroll.frx":0000
         Top             =   -240
         Width           =   5600
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6315
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()

       Timer1.Enabled = True
End Sub


Private Sub Form_Load()

       Timer1.Interval = 100
       VScroll1.Max = Picture1.Height
       VScroll1.Min = 0 - Text1.Height
       VScroll1.Value = VScroll1.Max
End Sub


Private Sub Timer1_Timer()


              If VScroll1.Value >= VScroll1.Min + 30 Then
                      VScroll1.Value = VScroll1.Value - 20
              Else
                      VScroll1.Value = VScroll1.Max
                      DoEvents
              End If

       Text1.Top = VScroll1.Value
       Text1.Visible = True

              DoEvents
              End Sub
