VERSION 5.00
Begin VB.Form FCopyName 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copy Names"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2130
   ControlBox      =   0   'False
   HelpContextID   =   286
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   2130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   210
      TabIndex        =   8
      Top             =   2400
      Width           =   810
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      Left            =   1110
      TabIndex        =   7
      Top             =   2400
      Width           =   810
   End
   Begin VB.Frame Frame2 
      Caption         =   "Which Files ? "
      Height          =   615
      Left            =   38
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
      Begin VB.OptionButton Option2 
         Caption         =   "Selected"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Copy What ? "
      Height          =   1575
      Left            =   38
      TabIndex        =   9
      Top             =   0
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "Path + Full filename"
         Height          =   255
         Index           =   4
         Left            =   140
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Path only"
         Height          =   255
         Index           =   3
         Left            =   140
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Prefix + Extension"
         Height          =   255
         Index           =   2
         Left            =   140
         TabIndex        =   2
         Top             =   720
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Extension only"
         Height          =   255
         Index           =   1
         Left            =   140
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Prefix only"
         Height          =   255
         Index           =   0
         Left            =   140
         TabIndex        =   0
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FCopyName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    LOk = True
    LesOptions.Misc1 = LOption1
    LesOptions.Misc2 = LOption2
    Unload Me
End Sub

Private Sub Command2_Click()
    LOk = False
    Unload Me
End Sub

Private Sub Form_Load()
    LOption1 = LesOptions.Misc1
    LOption2 = LesOptions.Misc2
    Option1(LOption1).Value = True
    Option2(LOption2).Value = True
End Sub

Private Sub Option1_Click(Index As Integer)
    LOption1 = Index
End Sub

Private Sub Option2_Click(Index As Integer)
    LOption2 = Index
End Sub
