VERSION 5.00
Begin VB.Form FParamCmd 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Free Form parameters"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Cyclic Selections "
      Height          =   675
      Left            =   113
      TabIndex        =   6
      Top             =   780
      Width           =   4455
      Begin VB.CommandButton Command6 
         Height          =   300
         HelpContextID   =   168
         Left            =   1980
         Picture         =   "FParamCmd.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Set options for Cyclic selection"
         Top             =   240
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   198
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Set their values"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Counter "
      Height          =   675
      Left            =   113
      TabIndex        =   5
      Top             =   0
      Width           =   4455
      Begin THERename.LabelText cmdtxt1 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Enter begin's value"
         Top             =   250
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   503
         Caption         =   "Counter's intial value"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   1600
         MousePointer    =   0
         Text            =   "1"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignement  =   1
      End
      Begin THERename.LabelText cmdtxt2 
         Height          =   285
         Left            =   2190
         TabIndex        =   1
         ToolTipText     =   "Enter increment value (1 for example)"
         Top             =   250
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Caption         =   "Step"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   500
         MousePointer    =   0
         Text            =   "1"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignement  =   1
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   14
      Left            =   2393
      TabIndex        =   4
      ToolTipText     =   "Save parameters"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   14
      Left            =   1193
      TabIndex        =   3
      ToolTipText     =   "Cancel your selection"
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "FParamCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    RENAME.cmdtxt1.Text = cmdtxt1.Text
    RENAME.cmdtxt2.Text = cmdtxt2.Text
    Unload Me
End Sub

Private Sub Command6_Click()
 OptionsCyclic = False
 Fcyclic.Show 1
End Sub

Private Sub Form_Load()
    cmdtxt1.Text = RENAME.cmdtxt1.Text
    cmdtxt2.Text = RENAME.cmdtxt2.Text
End Sub
