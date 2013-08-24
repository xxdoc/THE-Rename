VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form progress 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Progress ..."
   ClientHeight    =   2070
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnok 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   285
      Left            =   518
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
   Begin VB.CommandButton btncancel 
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   1913
      TabIndex        =   3
      Top             =   1440
      Width           =   1140
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   150
      Left            =   45
      TabIndex        =   0
      Top             =   1935
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   265
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label foliotage 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   833
      TabIndex        =   2
      Top             =   900
      Width           =   1995
   End
   Begin VB.Label avancement 
      Alignment       =   2  'Center
      Height          =   555
      Left            =   113
      TabIndex        =   1
      Top             =   45
      Width           =   3345
   End
End
Attribute VB_Name = "progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncancel_Click()
 annuler = True
 Unload Me
End Sub


Private Sub btnok_Click()
 annuler = False
 Unload Me
End Sub

Private Sub Form_Load()
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub
