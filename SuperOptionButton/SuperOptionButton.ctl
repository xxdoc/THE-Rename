VERSION 5.00
Begin VB.UserControl SuperOptionButton 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3390
   ControlContainer=   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   3390
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   2250
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   1170
      TabIndex        =   1
      Top             =   120
      Width           =   915
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   915
   End
End
Attribute VB_Name = "SuperOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub UserControl_Initialize()

End Sub

Private Sub UserControl_Resize()
    TextWidth (Option1(0).Caption)
End Sub
