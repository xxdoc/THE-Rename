VERSION 5.00
Begin VB.Form flheure 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select time"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2190
   ControlBox      =   0   'False
   HelpContextID   =   42
   Icon            =   "time.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   2190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   285
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "time.frx":000C
      Left            =   1485
      List            =   "time.frx":00C4
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select seconds"
      Top             =   270
      Width           =   645
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "time.frx":01B8
      Left            =   765
      List            =   "time.frx":0270
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Select minutes"
      Top             =   270
      Width           =   645
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "time.frx":0364
      Left            =   45
      List            =   "time.frx":03B0
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select hours"
      Top             =   270
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      Left            =   1140
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Sec."
      Height          =   195
      Left            =   1485
      TabIndex        =   7
      Top             =   45
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Min."
      Height          =   195
      Left            =   810
      TabIndex        =   6
      Top             =   45
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hours"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   45
      Width           =   420
   End
End
Attribute VB_Name = "flheure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim zzheure As String
Dim vtemoin As Boolean
Private Sub Command1_Click()
vtemoin = True
zzheure = Combo1.Text + ":" + Combo2.Text + ":" + Combo3.Text
Unload Me
End Sub

Private Sub Command2_Click()
 vtemoin = False
 Unload Me
End Sub
Public Function GetTime(heure As String) As Boolean
 Combo1.ListIndex = Val(Left$(heure, 2))
 Combo2.ListIndex = Val(Mid$(heure, 4, 2))
 Combo3.ListIndex = Val(Right$(heure, 2))
 Me.Show vbModal
 If vtemoin = True Then
  heure = zzheure
  GetTime = True
 Else
  GetTime = False
 End If
End Function
