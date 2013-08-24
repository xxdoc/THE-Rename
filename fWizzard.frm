VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form fWizzard 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Free Form Wizzard"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2213
      TabIndex        =   4
      ToolTipText     =   "Don't use this command and close this window"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   3653
      TabIndex        =   3
      ToolTipText     =   "Use this command line in the free form"
      Top             =   6360
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Prefix"
      TabPicture(0)   =   "fWizzard.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Combo1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ListView1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdDown"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdUp"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Combo2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Command3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Command4"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Extension"
      TabPicture(1)   =   "fWizzard.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Both/Other"
      TabPicture(2)   =   "fWizzard.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.TextBox Text5 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   22
         Text            =   "fWizzard.frx":0054
         ToolTipText     =   "This is the help about the current command"
         Top             =   1800
         Width           =   6615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Modify"
         Height          =   375
         Left            =   5640
         TabIndex        =   21
         ToolTipText     =   "Modify this rule"
         Top             =   2880
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         ToolTipText     =   "Add this rule"
         Top             =   2880
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "fWizzard.frx":007D
         Left            =   120
         List            =   "fWizzard.frx":00A8
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1400
         Width           =   6615
      End
      Begin VB.CommandButton Command5 
         Height          =   330
         HelpContextID   =   168
         Left            =   6430
         Picture         =   "fWizzard.frx":01C3
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Remove selected rule"
         Top             =   3720
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3960
         TabIndex        =   16
         Top             =   2925
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2280
         TabIndex        =   14
         Top             =   2925
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   2925
         Width           =   855
      End
      Begin VB.CommandButton cmdUp 
         Enabled         =   0   'False
         Height          =   330
         HelpContextID   =   20
         Left            =   6430
         Picture         =   "fWizzard.frx":02AD
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Move Up"
         Top             =   4305
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdDown 
         Enabled         =   0   'False
         Height          =   330
         HelpContextID   =   20
         Left            =   6430
         Picture         =   "fWizzard.frx":03AF
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Move Down"
         Top             =   4680
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   3600
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2990
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "fWizzard.frx":04B1
         Left            =   120
         List            =   "fWizzard.frx":04DF
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "What sub option would you like to use ?"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1155
         Width           =   2835
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Option3"
         Height          =   195
         Left            =   3360
         TabIndex        =   15
         Top             =   2955
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Option2"
         Height          =   195
         Left            =   1680
         TabIndex        =   13
         Top             =   2955
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Option1"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2955
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Current rules for prefix"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "What do you want to do on prefix ?"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2490
      End
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   113
      TabIndex        =   0
      ToolTipText     =   "That's the command line THE Rename wil use for the free form option"
      Top             =   6000
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Resulting command line"
      Height          =   195
      Left            =   2700
      TabIndex        =   1
      Top             =   5760
      Width           =   1680
   End
End
Attribute VB_Name = "fWizzard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0
    Text5.Text = "Ligne 1" & vbCrLf & "Ligne 2" & vbCrLf & "Ligne 3" & vbCrLf & "Ligne 4"
End Sub
