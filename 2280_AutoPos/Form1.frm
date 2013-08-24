VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Auto resizing horizontally + vertically"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Relative to form's bottom border"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Auto Resizing vertically"
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Auto resizing horizontally"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Relative to form's right border"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oAutoPos As New clsAutoPositioner


Private Sub Form_Load()
' Always relative to container's right border
m_oAutoPos.AddAssignment Me.Command1, Me, tCONTAINER_RELATIVE_POS_RIGHT

' Auto resizing horizontally
m_oAutoPos.AddAssignment Me.Command2, Me, tCONTAINER_WIDTH_DELTA_RIGHT

' Auto resizing vertically
m_oAutoPos.AddAssignment Me.Command3, Me, tCONTAINER_HEIGHT_DELTA_BOTTOM

' Always relative to container's bottom border
m_oAutoPos.AddAssignment Me.Command4, Me, tCONTAINER_RELATIVE_POS_BOTTOM

' Auto resizing horizontally + Auto resizing vertically
m_oAutoPos.AddAssignment Me.Command5, Me, tCONTAINER_WIDTH_DELTA_RIGHT
m_oAutoPos.AddAssignment Me.Command5, Me, tCONTAINER_HEIGHT_DELTA_BOTTOM
End Sub

Private Sub Form_Resize()
m_oAutoPos.RefreshPositions
End Sub
