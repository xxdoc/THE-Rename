VERSION 5.00
Begin VB.Form fmdgroupe 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create a group of directories"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   HelpContextID   =   159
   Icon            =   "fmdgroupe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Text1 
      Height          =   315
      Left            =   705
      TabIndex        =   0
      Text            =   "Folder"
      Top             =   0
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      HelpContextID   =   159
      ItemData        =   "fmdgroupe.frx":000C
      Left            =   1080
      List            =   "fmdgroupe.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Choose file's counter representation (Dec/Hex/Oct)"
      Top             =   1035
      Width           =   1300
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   285
      HelpContextID   =   159
      Left            =   2085
      TabIndex        =   4
      Text            =   "1"
      Top             =   690
      Width           =   465
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   285
      HelpContextID   =   159
      Left            =   1080
      TabIndex        =   3
      Text            =   "1"
      Top             =   690
      Width           =   465
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   159
      Left            =   555
      TabIndex        =   8
      ToolTipText     =   "Close this window"
      Top             =   2000
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   159
      Left            =   1560
      TabIndex        =   9
      ToolTipText     =   "Create directories"
      Top             =   2000
      Width           =   915
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Numbers to the right"
      Height          =   240
      HelpContextID   =   159
      Left            =   570
      TabIndex        =   6
      Top             =   1410
      Value           =   -1  'True
      Width           =   1890
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Numbers to the left"
      Height          =   240
      HelpContextID   =   159
      Left            =   570
      TabIndex        =   7
      Top             =   1650
      Width           =   1740
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   285
      HelpContextID   =   159
      Left            =   2085
      TabIndex        =   2
      Text            =   "9"
      Top             =   390
      Width           =   465
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      HelpContextID   =   159
      Left            =   1080
      TabIndex        =   1
      Text            =   "1"
      Top             =   390
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Format"
      Height          =   195
      Left            =   480
      TabIndex        =   15
      Top             =   1110
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Digits"
      Height          =   195
      Left            =   1605
      TabIndex        =   14
      Top             =   720
      Width           =   390
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Step"
      Height          =   195
      Left            =   480
      TabIndex        =   13
      Top             =   720
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "To"
      Height          =   195
      Left            =   1605
      TabIndex        =   12
      Top             =   435
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "From"
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   435
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label"
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   60
      Width           =   390
   End
End
Attribute VB_Name = "fmdgroupe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cHist10 As New cHistory
Private Sub Command1_Click()
 Dim i As Long
 Dim vnom As String
 Dim vtmp As String
 On Error Resume Next
 cHist10.AddNewItem Text1.Text
 vtmp = AddBackSlash(Dir1Path)
 If Val(Text3.Text) < Val(Text2.Text) Then
  MsgBox "Error, the 'To' value must be greater than the 'from' value"
  Exit Sub
 End If
 For i = Val(Text2.Text) To Val(Text3.Text) Step Val(Text4.Text)
  If Option2.Value = True Then
   vnom = Text1 & Compteur(i, Val(Text5.Text), Combo3.ListIndex)
  Else
   vnom = Compteur(i, Val(Text5.Text), Combo3.ListIndex) & Text1
  End If
  MkDir vtmp & vnom
 Next
 Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cHist10.sKey = "CreateGroupDir"
    cHist10.Items Text1
   Combo3.ListIndex = 0
End Sub

Private Sub Text1_Change()
    CharInterdits Text1.Text
End Sub

Private Sub Text1_GotFocus()
    SelAll Text1
End Sub
Private Sub Text2_GotFocus()
    SelAll Text2
End Sub

Private Sub Text3_Change()
    Text5.Text = Len(Text3.Text)
End Sub

Private Sub Text3_GotFocus()
    SelAll Text3
End Sub

Private Sub Text4_GotFocus()
    SelAll Text4
End Sub
Private Sub Text5_GotFocus()
    SelAll Text5
End Sub
