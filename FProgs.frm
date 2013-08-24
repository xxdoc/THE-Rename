VERSION 5.00
Begin VB.Form FProgs 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registred programs"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7965
   ControlBox      =   0   'False
   HelpContextID   =   537
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   3000
      Width           =   7815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   14
      Left            =   4035
      TabIndex        =   3
      Top             =   3420
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   14
      Left            =   2835
      TabIndex        =   2
      Top             =   3420
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2880
      IntegralHeight  =   0   'False
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   60
      Width           =   7815
   End
End
Attribute VB_Name = "FProgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cR As New cRegistry
Private mok As Boolean
Private fonct As String
Private Sub cmdCancel_Click()
    mok = False
    Unload Me
    Exit Sub
End Sub
Private Sub cmdOK_Click()
    mok = True
    Unload Me
    Exit Sub
End Sub
Private Sub List1_Click()
    If List1.ListIndex = -1 Then Exit Sub
    With cR
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" & List1.List(List1.ListIndex)
        .ValueKey = ""
        .ValueType = REG_SZ
        Text1.Text = .Value
        Text1.Text = Replace(Text1.Text, Chr$(34), "")
        fonct = Text1.Text
    End With
End Sub

Public Function GetProgram(fonction As String) As Boolean
    Dim sKeys() As String, iKeyCount As Long, iKey As Long
    mok = False
    With cR
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
        .EnumerateSections sKeys(), iKeyCount
        For iKey = 1 To iKeyCount
            List1.AddItem sKeys(iKey)
        Next
    End With
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
    Me.Show vbModal
    If mok = True Then
        fonction = fonct
    Else
        fonction = ""
    End If
    GetProgram = mok
End Function

Private Sub List1_DblClick()
    cmdOK_Click
End Sub
