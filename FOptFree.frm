VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FOptFree 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options for free form"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   14
      Left            =   3285
      TabIndex        =   10
      ToolTipText     =   "Save settings for ALL sessions"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   14
      Left            =   2100
      TabIndex        =   9
      ToolTipText     =   "Cancel your selection"
      Top             =   1560
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   2566
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Counters"
      TabPicture(0)   =   "FOptFree.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lang3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lang2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdtxt1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdtxt2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Cyclic texts"
      TabPicture(1)   =   "FOptFree.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Tokens"
      TabPicture(2)   =   "FOptFree.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Files"
      TabPicture(3)   =   "FOptFree.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lang4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdtxt3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.TextBox cmdtxt3 
         Height          =   285
         HelpContextID   =   41
         Left            =   -72120
         TabIndex        =   8
         Text            =   "25"
         ToolTipText     =   "Enter begin's value"
         Top             =   480
         WhatsThisHelpID =   211
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   300
         HelpContextID   =   168
         Left            =   2340
         Picture         =   "FOptFree.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Set options for Cyclic selection"
         Top             =   480
         WhatsThisHelpID =   198
         Width           =   300
      End
      Begin VB.TextBox cmdtxt2 
         Height          =   285
         HelpContextID   =   41
         Left            =   -72360
         TabIndex        =   2
         Text            =   "1"
         ToolTipText     =   "Enter increment value (1 for example)"
         Top             =   480
         WhatsThisHelpID =   210
         Width           =   375
      End
      Begin VB.TextBox cmdtxt1 
         Height          =   285
         HelpContextID   =   41
         Left            =   -73230
         TabIndex        =   1
         Text            =   "1"
         ToolTipText     =   "Enter begin's value"
         Top             =   480
         WhatsThisHelpID =   209
         Width           =   375
      End
      Begin VB.Label lang4 
         AutoSize        =   -1  'True
         Caption         =   "Number of characters to take from file"
         Height          =   195
         Left            =   -74880
         TabIndex        =   7
         Top             =   510
         WhatsThisHelpID =   211
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Click to configure cyclic texts"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   510
         Width           =   2055
      End
      Begin VB.Label lang2 
         AutoSize        =   -1  'True
         Caption         =   "Counters begin value"
         Height          =   195
         Left            =   -74880
         TabIndex        =   4
         Top             =   510
         WhatsThisHelpID =   209
         Width           =   1500
      End
      Begin VB.Label lang3 
         AutoSize        =   -1  'True
         Caption         =   "Step"
         Height          =   195
         Left            =   -72810
         TabIndex        =   3
         Top             =   510
         WhatsThisHelpID =   210
         Width           =   330
      End
   End
End
Attribute VB_Name = "FOptFree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command6_Click()
 OptionsCyclic = False
 Fcyclic.Show 1
End Sub
