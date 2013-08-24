VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FArrays 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrays"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7470
   ControlBox      =   0   'False
   Icon            =   "FArrays.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Values "
      Height          =   1815
      Left            =   80
      TabIndex        =   28
      Top             =   3480
      Width           =   7275
      Begin VB.CommandButton Command9 
         Caption         =   "Save"
         Height          =   1035
         Left            =   6480
         TabIndex        =   21
         ToolTipText     =   "Save all the values"
         Top             =   420
         Width           =   615
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Modify"
         Height          =   315
         Left            =   5580
         TabIndex        =   20
         ToolTipText     =   "Modify the value"
         Top             =   1140
         Width           =   800
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Delete"
         Height          =   315
         Left            =   5580
         TabIndex        =   19
         ToolTipText     =   "Delete selected value"
         Top             =   780
         Width           =   800
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Add"
         Height          =   315
         Left            =   5580
         TabIndex        =   18
         ToolTipText     =   "Add a new value"
         Top             =   420
         Width           =   800
      End
      Begin VB.TextBox Text8 
         Height          =   285
         HelpContextID   =   110
         Left            =   3360
         TabIndex        =   17
         Top             =   1200
         WhatsThisHelpID =   220
         Width           =   1890
      End
      Begin VB.TextBox Text7 
         Height          =   285
         HelpContextID   =   110
         Left            =   3360
         TabIndex        =   16
         Top             =   540
         WhatsThisHelpID =   220
         Width           =   1890
      End
      Begin MSComctlLib.ListView LstValues 
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   300
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2355
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Initial Value"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Return value"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Replace it with"
         Height          =   195
         Left            =   3360
         TabIndex        =   30
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "When value is equal to"
         Height          =   195
         Left            =   3360
         TabIndex        =   29
         Top             =   300
         Width           =   1635
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Deactivate all arrays"
      Height          =   315
      Left            =   4080
      TabIndex        =   4
      Top             =   1480
      Width           =   1700
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Activate all arrays"
      Height          =   315
      Left            =   5835
      TabIndex        =   5
      Top             =   1480
      Width           =   1550
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add..."
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      ToolTipText     =   "Add a new array"
      Top             =   1480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Duplicate"
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      ToolTipText     =   "Duplicate selected array"
      Top             =   1480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Delete selected array"
      Top             =   1480
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extract information "
      Height          =   1395
      Left            =   80
      TabIndex        =   24
      Top             =   1980
      Width           =   7275
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   285
         HelpContextID   =   110
         Left            =   6540
         TabIndex        =   14
         Top             =   960
         WhatsThisHelpID =   220
         Width           =   510
      End
      Begin VB.TextBox Text5 
         Height          =   285
         HelpContextID   =   110
         Left            =   2040
         TabIndex        =   13
         Top             =   960
         WhatsThisHelpID =   220
         Width           =   2850
      End
      Begin VB.OptionButton Option1 
         Caption         =   "A regular expression"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   1020
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   285
         HelpContextID   =   110
         Left            =   3555
         TabIndex        =   11
         Top             =   600
         WhatsThisHelpID =   220
         Width           =   1290
      End
      Begin VB.TextBox Text3 
         Height          =   285
         HelpContextID   =   110
         Left            =   2040
         TabIndex        =   10
         Top             =   600
         WhatsThisHelpID =   220
         Width           =   1050
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Between expression"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         HelpContextID   =   110
         Left            =   3000
         TabIndex        =   8
         Top             =   240
         WhatsThisHelpID =   220
         Width           =   510
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         HelpContextID   =   110
         Left            =   2040
         TabIndex        =   7
         Top             =   240
         WhatsThisHelpID =   220
         Width           =   510
      End
      Begin VB.OptionButton Option1 
         Caption         =   "From position"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "use backreference #"
         Height          =   195
         Left            =   4980
         TabIndex        =   27
         Top             =   1020
         WhatsThisHelpID =   220
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "and"
         Height          =   195
         Left            =   3180
         TabIndex        =   26
         Top             =   645
         WhatsThisHelpID =   220
         Width           =   270
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   2685
         TabIndex        =   25
         Top             =   285
         WhatsThisHelpID =   220
         Width           =   195
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1395
      Left            =   60
      TabIndex        =   0
      ToolTipText     =   "This is the list of all your arrays"
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2461
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Rule Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Rule Number"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      HelpContextID   =   168
      Left            =   6255
      TabIndex        =   23
      ToolTipText     =   "Save arrays"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   168
      Left            =   5100
      TabIndex        =   22
      ToolTipText     =   "Don't save arrays"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      Begin VB.Menu mload 
         Caption         =   "&Load arrays from a file..."
      End
      Begin VB.Menu msave 
         Caption         =   "&Save arrays to a file..."
      End
      Begin VB.Menu msep0 
         Caption         =   "-"
      End
      Begin VB.Menu MHelp 
         Caption         =   "&Help..."
      End
   End
End
Attribute VB_Name = "FArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim itmX As ListItem
    Set itmX = ListView1.ListItems.Add(, , "Coucou")
End Sub
