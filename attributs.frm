VERSION 5.00
Begin VB.Form attributs 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change attributes"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2490
   ControlBox      =   0   'False
   HelpContextID   =   38
   Icon            =   "attributs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Archive 
      Caption         =   "Archive"
      Height          =   195
      HelpContextID   =   38
      Left            =   1185
      TabIndex        =   1
      Top             =   75
      Width           =   1290
   End
   Begin VB.CheckBox System 
      Caption         =   "System"
      Height          =   195
      HelpContextID   =   38
      Left            =   1185
      TabIndex        =   3
      Top             =   385
      Width           =   1290
   End
   Begin VB.CheckBox Hidden 
      Caption         =   "Hidden"
      Height          =   195
      HelpContextID   =   38
      Left            =   1185
      TabIndex        =   7
      Top             =   1005
      Width           =   1290
   End
   Begin VB.CheckBox ReadOnly 
      Caption         =   "Read-Only"
      Height          =   195
      HelpContextID   =   38
      Left            =   1185
      TabIndex        =   5
      Top             =   695
      Width           =   1290
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Set"
      Height          =   195
      HelpContextID   =   38
      Left            =   15
      TabIndex        =   4
      Top             =   690
      Width           =   1080
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Set"
      Height          =   195
      HelpContextID   =   38
      Left            =   15
      TabIndex        =   6
      Top             =   1005
      Width           =   1080
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Set"
      Height          =   195
      HelpContextID   =   38
      Left            =   15
      TabIndex        =   2
      Top             =   385
      Width           =   1080
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Set"
      Height          =   195
      HelpContextID   =   38
      Left            =   15
      TabIndex        =   0
      Top             =   75
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   38
      Left            =   1275
      TabIndex        =   9
      ToolTipText     =   "Files attributes will be changed"
      Top             =   2265
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   38
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Files attributes will not be changed"
      Top             =   2265
      Width           =   1095
   End
   Begin VB.Label Warning 
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   60
      TabIndex        =   10
      Top             =   1380
      Width           =   2415
   End
End
Attribute VB_Name = "attributs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LesAttrib As New CAttrib
Private Sub Check1_Click()
 If Check1.Value = 1 Then
  ReadOnly.Value = 0
  ReadOnly.Enabled = False
 Else
  ReadOnly.Enabled = True
 End If
End Sub

Private Sub Check2_Click()
 If Check2.Value = 1 Then
  Hidden.Value = 0
  Hidden.Enabled = False
 Else
  Hidden.Enabled = True
 End If
End Sub

Private Sub Check3_Click()
 If Check3.Value = 1 Then
  System.Value = 0
  System.Enabled = False
 Else
  System.Enabled = True
 End If
End Sub

Private Sub Check4_Click()
 If Check4.Value = 1 Then
  Archive.Value = 0
  Archive.Enabled = False
 Else
  Archive.Enabled = True
 End If
End Sub

Private Sub cmdCancel_Click()
 LesAttrib.AtrOk = False
 Unload Me
End Sub

Private Sub cmdOK_Click()
 LesAttrib.ResetAttrib
 LesAttrib.AtrOk = True
 If Archive.Value = 1 Then
   LesAttrib.Archive = True
 End If
 If System.Value = 1 Then
  LesAttrib.System = True
 End If
 If ReadOnly.Value = 1 Then
  LesAttrib.ReadOnly = True
 End If
 If Hidden.Value = 1 Then
  LesAttrib.Hidden = True
 End If
 If Check4.Value = 1 Then
  LesAttrib.SETArchive = True
 End If
 If Check3.Value = 1 Then
  LesAttrib.SETSystem = True
 End If
 If Check1.Value = 1 Then
  LesAttrib.SETReadOnly = True
 End If
 If Check2.Value = 1 Then
  LesAttrib.SETHidden = True
 End If
 Unload Me
End Sub

Private Sub Form_Load()
  Select Case AttrEncours
   Case 1
    Set LesAttrib = Attr1
    Warning.Caption = "Warning files will have their attributes changed only when you will rename files"
   Case 2
    Set LesAttrib = Attr2
    Warning.Caption = "Warning files will have their attributes changed only when you will copy or paste files"
   Case 3
    Set LesAttrib = Attr3
    Warning.Caption = "Warning files will have their attributes changed only when you will drop files on the multiple copy button files"
   Case 4
    Set LesAttrib = Attr4
    Warning.Caption = ""
  End Select
  If LesAttrib.SETArchive = True Then
   Check4.Value = 1
  End If
  If LesAttrib.Archive = True Then
   Archive.Value = 1
  End If
  If LesAttrib.SETSystem = True Then
   Check3.Value = 1
  End If
  If LesAttrib.System = True Then
   System.Value = 1
  End If
  If LesAttrib.SETReadOnly = True Then
   Check1.Value = 1
  End If
  If LesAttrib.ReadOnly = True Then
   ReadOnly.Value = 1
  End If
  If LesAttrib.SETHidden = True Then
   Check2.Value = 1
  End If
  If LesAttrib.Hidden = True Then
   Hidden.Value = 1
  End If
' Traduction **********************************
'LoadAllName Me, AddBackSlash(App.Path) & "Francais.lng"
  
End Sub
