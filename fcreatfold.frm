VERSION 5.00
Begin VB.Form fcreatfold 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create folders with names"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3105
   ControlBox      =   0   'False
   HelpContextID   =   287
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "How to use it ? "
      Height          =   1455
      Left            =   38
      TabIndex        =   17
      Top             =   640
      Width           =   3015
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         ScaleHeight     =   615
         ScaleWidth      =   2655
         TabIndex        =   18
         Top             =   720
         Width           =   2655
         Begin VB.OptionButton Option5 
            Caption         =   "From character"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   23
            Top             =   40
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Text            =   "1"
            Top             =   0
            Width           =   375
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2160
            TabIndex        =   21
            Text            =   "1"
            Top             =   0
            Width           =   375
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Last"
            Height          =   195
            Index           =   1
            Left            =   0
            TabIndex        =   20
            Top             =   320
            Width           =   735
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   19
            Text            =   "1"
            Top             =   280
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "to"
            Height          =   195
            Left            =   1920
            TabIndex        =   25
            Top             =   40
            Width           =   135
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "characters"
            Height          =   195
            Left            =   1200
            TabIndex        =   24
            Top             =   320
            Width           =   750
         End
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Use a part of the name"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Use the whole name"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Capitalize all words"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   4680
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Replace _ with space"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   4410
      Value           =   1  'Checked
      Width           =   1995
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Stop to the first numeric character"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   4140
      Value           =   1  'Checked
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      Left            =   1616
      TabIndex        =   12
      Top             =   5025
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   573
      TabIndex        =   13
      Top             =   5025
      Width           =   915
   End
   Begin VB.Frame Frame3 
      Caption         =   "What to do ? "
      Height          =   1095
      Left            =   38
      TabIndex        =   16
      Top             =   2180
      Width           =   3015
      Begin VB.OptionButton Option3 
         Caption         =   "Move file to new folder"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Copy file to folder"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Just create folder"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "What to use ? "
      Height          =   555
      Left            =   38
      TabIndex        =   15
      ToolTipText     =   "To create folder's name"
      Top             =   0
      Width           =   3015
      Begin VB.OptionButton Option1 
         Caption         =   "Extension"
         Height          =   255
         Index           =   1
         Left            =   1500
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Prefix"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Which Files ? "
      Height          =   615
      Left            =   38
      TabIndex        =   14
      Top             =   3360
      Width           =   3015
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Selected"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
End
Attribute VB_Name = "fcreatfold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vtmp1 As Integer
Dim vtmp2 As Integer
Dim vtmp3 As Integer

Private Sub Command1_Click()
    LOk = True
    LesOptions.Misc3 = vtmp1
    LesOptions.Misc4 = vtmp2
    LesOptions.Misc5 = vtmp3
    LesOptions.Misc6 = Check1.Value
    LesOptions.Misc7 = Check2.Value
    LesOptions.Misc8 = Check3.Value
    LesOptions.Misc11 = Text1.Text
    LesOptions.Misc12 = Text2.Text
    LesOptions.Misc13 = Text3.Text
    Unload Me
End Sub

Private Sub Command2_Click()
    LOk = False
    Unload Me
End Sub
Private Sub Form_Load()
    Option1(LesOptions.Misc3).Value = True
    Option3(LesOptions.Misc4).Value = True
    Option2(LesOptions.Misc5).Value = True
    Check1.Value = LesOptions.Misc6
    Check2.Value = LesOptions.Misc7
    Check3.Value = LesOptions.Misc8
    Option4(LesOptions.Misc9).Value = True
    Option5(LesOptions.Misc10).Value = True
    Text1.Text = LesOptions.Misc11
    Text2.Text = LesOptions.Misc12
    Text3.Text = LesOptions.Misc13
    etat1
End Sub

Private Sub Option1_Click(Index As Integer)
    vtmp1 = Index
End Sub

Private Sub Option2_Click(Index As Integer)
    vtmp3 = Index
End Sub

Private Sub Option3_Click(Index As Integer)
    vtmp2 = Index
End Sub

Private Sub etat1()
    If Option4(0).Value = True Then
        Option5(0).Enabled = False
        Option5(1).Enabled = False
        Text1.Enabled = False
        Text2.Enabled = False
        Text3.Enabled = False
        Label1.Enabled = False
        Label2.Enabled = False
    Else
        Option5(0).Enabled = True
        Option5(1).Enabled = True
        If Option5(0).Value = True Then
            Text1.Enabled = True
            Label1.Enabled = True
            Text2.Enabled = True
            Text3.Enabled = False
            Label2.Enabled = False
        Else
            Text1.Enabled = False
            Label1.Enabled = False
            Text2.Enabled = False
            Text3.Enabled = True
            Label2.Enabled = True
        End If
    End If
End Sub

Private Sub Option4_Click(Index As Integer)
    etat1
    LesOptions.Misc9 = Index
End Sub

Private Sub Option5_Click(Index As Integer)
    etat1
    LesOptions.Misc10 = Index
End Sub
