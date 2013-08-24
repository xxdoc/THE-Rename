VERSION 5.00
Begin VB.Form FFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3075
   ControlBox      =   0   'False
   HelpContextID   =   44
   Icon            =   "FFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.ComboBox Text1 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Text            =   "*.*"
      ToolTipText     =   "Enter a filter or type a regular expression"
      Top             =   135
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Regular expression"
      Height          =   285
      HelpContextID   =   44
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Check this option to use regular expressions"
      Top             =   585
      Width           =   1725
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   44
      Left            =   1575
      TabIndex        =   3
      ToolTipText     =   "Search"
      Top             =   1530
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   44
      Left            =   382
      TabIndex        =   2
      ToolTipText     =   "Cancel search"
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Note, for regular expressions, you can use * ? # [] !"
      Height          =   420
      Left            =   150
      TabIndex        =   5
      Top             =   990
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter filter"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   180
      Width           =   705
   End
End
Attribute VB_Name = "FFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cHist8 As New cHistory
Private Sub cmdCancel_Click()
    FilterOk = False
    If Check1.Value = 1 Then
        FilterRegular = True
    Else
        FilterRegular = False
    End If
    FilterExpr = Text1.Text
    Unload Me
End Sub

Private Sub cmdOK_Click()
If Trim$(Text1.Text) = "" Then Text1.SetFocus
FilterOk = True
If Check1.Value = 1 Then
    FilterRegular = True
Else
    FilterRegular = False
End If
FilterExpr = Text1.Text
cHist8.AddNewItem Text1.Text
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = FilterExpr
cHist8.sKey = "FilterString"
cHist8.Items Text1
End Sub

Private Sub Text1_GotFocus()
    SelAll Text1
End Sub
