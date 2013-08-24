VERSION 5.00
Begin VB.Form FRegRename 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RegExp rename"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2985
   ControlBox      =   0   'False
   HelpContextID   =   303
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   2985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "1"
      Top             =   1740
      Width           =   495
   End
   Begin VB.CheckBox Check2 
      Caption         =   "&Replace all"
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Match Case"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   352
      TabIndex        =   7
      Top             =   2850
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      Left            =   1537
      TabIndex        =   8
      Top             =   2850
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Which Files ? "
      Height          =   585
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   2775
      Begin VB.OptionButton Option2 
         Caption         =   "&Selected"
         Height          =   255
         Index           =   1
         Left            =   1215
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&All"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.ComboBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1000
      Width           =   2805
   End
   Begin VB.ComboBox Text4 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2805
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "# of substitutions"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rename to"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   760
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter regular expression"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "FRegRename"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cHist15 As New cHistory
Private Sub Check2_Click()
    If Check2.Value = 0 Then
        Label3.Enabled = True
        Text2.Enabled = True
    Else
        Label3.Enabled = False
        Text2.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
 If Trim$(Text4.Text) = "" Then
    MsgBox "Please, enter a regular expression"
    Text4.SetFocus
    Exit Sub
 End If
 If Trim$(Text1.Text) = "" Then
    MsgBox "Please enter a string to rename to"
    Text1.SetFocus
    Exit Sub
 End If
 cHist14.AddNewItem Text4.Text
 cHist15.AddNewItem Text1.Text
 LOk = True
 If Option2(0).Value = True Then
    LOption1 = 1
 Else
    LOption1 = 2
 End If
 If Check1.Value = 1 Then
    LOption2 = 1
 Else
    LOption2 = 0
 End If
 If Check2.Value = 1 Then
    LOption3 = 9999
 Else
    LOption3 = Val(Text2.Text)
 End If
 LChaine1 = Text4.Text
 LChaine2 = Text1.Text
 Unload Me
End Sub

Private Sub Command2_Click()
    LOk = False
    Unload Me
End Sub

Private Sub Form_Load()
  cHist14.sKey = "RegPattern"
  cHist14.Items Text4
  cHist15.sKey = "RegRename"
  cHist15.Items Text1
  Check2_Click
End Sub
Private Sub Text1_GotFocus()
    SelAll Text1
End Sub

Private Sub Text2_GotFocus()
    SelAll Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
    If Check2.Value = 0 Then
        If Val(Text2.Text) < 1 Then
            Cancel = True
        End If
    End If
End Sub
Private Sub Text4_GotFocus()
    SelAll Text4
End Sub
