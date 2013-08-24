VERSION 5.00
Begin VB.Form fgoto 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Go to"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4200
   ControlBox      =   0   'False
   Icon            =   "fgoto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Text1 
      Height          =   315
      Left            =   540
      TabIndex        =   0
      ToolTipText     =   "Enter a directory where to go"
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Browse ..."
      Height          =   300
      Left            =   349
      TabIndex        =   1
      ToolTipText     =   "Browse for a folder"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2757
      TabIndex        =   3
      ToolTipText     =   "Goto this directory"
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1553
      TabIndex        =   2
      ToolTipText     =   "Close this window"
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Go"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   165
      Width           =   210
   End
End
Attribute VB_Name = "fgoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cHist9 As New cHistory
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
 On Error GoTo erreur
 If Text1.Text <> "" Then
  cHist9.AddNewItem Text1.Text
  RENAME.MousePointer = 11
  ChDir Text1.Text
  TemMove = False
  RENAME.FolderTreeview1(0).Visible = False
  RENAME.FolderTreeview1(0).SelectedFolder = Text1.Text
  Dir1Path = Text1.Text
  RENAME.FolderTreeview1(0).Visible = True
  RENAME.mundo.Enabled = False
  RENAME.List2.Clear
  RENAME.List3.Clear
  RENAME.MousePointer = 0
End If
Unload Me

erreur:
If Err.Number = 76 Then
    MsgBox "Invalid path !", vbOKOnly, "Error"
    Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
 Dim szFilename As String
 szFilename = BrowseFolder(Me, "Browse for folder:")
 If Len(Trim$(szFilename)) = 0 Then
  Exit Sub
 End If
 Text1.Text = szFilename
 Text1.SetFocus
End Sub

Private Sub Form_Load()
cHist9.sKey = "GotoString"
cHist9.Items Text1
End Sub
Private Sub Text1_GotFocus()
    SelAll Text1
End Sub
