VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form FTokOption 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tokens Options"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   159
      Left            =   1447
      TabIndex        =   4
      ToolTipText     =   "Create directories"
      Top             =   4920
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   159
      Left            =   2422
      TabIndex        =   5
      ToolTipText     =   "Close this window"
      Top             =   4920
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   45
      TabIndex        =   10
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Prefix Tokens"
      TabPicture(0)   =   "FTokOption.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ListView1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Combo1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "&Extension Tokens"
      TabPicture(1)   =   "FTokOption.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "ListView2"
      Tab(1).Control(2)=   "Combo2"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "Label3"
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame2 
         Caption         =   "Range "
         Height          =   855
         Left            =   -74880
         TabIndex        =   14
         Top             =   350
         Width           =   4455
         Begin VB.OptionButton Option1 
            Caption         =   "&Apply options to all tokens"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Apply options &individually for each token"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   3135
         End
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   9
         Top             =   2280
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   -74880
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1560
         Width           =   4455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   1560
         Width           =   4455
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4048
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Frame Frame1 
         Caption         =   "Range "
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   350
         Width           =   4455
         Begin VB.OptionButton Option1 
            Caption         =   "Apply options &individually for each token"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   3135
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&Apply options to all tokens"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tokens list"
         Height          =   195
         Left            =   -74880
         TabIndex        =   16
         Top             =   2040
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Option to use"
         Height          =   195
         Left            =   -74880
         TabIndex        =   15
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Option to use"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tokens list"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   765
      End
   End
End
Attribute VB_Name = "FTokOption"
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

