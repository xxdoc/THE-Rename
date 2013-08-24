VERSION 5.00
Begin VB.Form FFoldMan 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create a list of folders manually"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   HelpContextID   =   76
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   83
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   4515
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      ToolTipText     =   "Don't create folders"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      HelpContextID   =   14
      Left            =   2385
      TabIndex        =   2
      ToolTipText     =   "Create folders"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Type the list of folders you want to create. Enter one name by line. Folders will be created in the current directory."
      Height          =   435
      Left            =   143
      TabIndex        =   3
      Top             =   120
      Width           =   4395
   End
End
Attribute VB_Name = "FFoldMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer, vnb As Integer
    Dim vtmp As String, vtmp2 As String
    vtmp = AddBackSlash(Dir1Path)
    i = 1
    vtmp2 = GetToken(Text1.Text, vbCrLf, i)
    While Trim$(vtmp2) <> ""
        vtmp2 = Replace(vtmp2, Chr$(13), "")
        vtmp2 = Replace(vtmp2, Chr$(10), "")
        vtmp2 = Trim$(vtmp2)
        If vtmp2 <> "" Then
            vnb = CreateNestedFoldersByPath(vtmp & vtmp2)
        End If
        i = i + 1
        vtmp2 = GetToken(Text1.Text, vbCrLf, i)
    Wend
    RENAME.Refresh
    Unload Me
End Sub
