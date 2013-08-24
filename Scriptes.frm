VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Scriptes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Script"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7800
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   14
      Left            =   3945
      TabIndex        =   2
      ToolTipText     =   "Save settings for ALL sessions"
      Top             =   5775
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   14
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Cancel your selection"
      Top             =   5775
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5565
      Left            =   90
      TabIndex        =   0
      Top             =   75
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   9816
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Scriptes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
End
Attribute VB_Name = "Scriptes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
