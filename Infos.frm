VERSION 5.00
Begin VB.Form Infos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Information"
   ClientHeight    =   2115
   ClientLeft      =   3630
   ClientTop       =   2955
   ClientWidth     =   7005
   ControlBox      =   0   'False
   HelpContextID   =   15
   Icon            =   "Infos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   15
      Left            =   2347
      TabIndex        =   1
      Top             =   1785
      Width           =   2310
   End
   Begin VB.TextBox Text1 
      Height          =   1680
      HelpContextID   =   15
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   45
      Width           =   6765
   End
End
Attribute VB_Name = "Infos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oAutoPos As New clsAutoPositioner

Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Form_Load()
Dim ret As String
ret = Chr$(13) + Chr$(10)
Text1.Text = "Program installation's date is " + LesOptions.FirstDateUse + ret
Text1.Text = Text1.Text + "Last time when you used THE Rename was " + LesOptions.LastUseDate + " at " + LesOptions.LastUseTime + ret
Text1.Text = Text1.Text + "You ran the program " + Str$(LesOptions.NumberOfRuns) + " times" + ret
Text1.Text = Text1.Text + "Last time, you renamed " + Str$(LesOptions.NumberOFiles) + " files in " + LesOptions.LastDirectory + ret
Text1.Text = Text1.Text + "Startup directory : " + LesOptions.StartupDir + ret
Text1.Text = Text1.Text + "Don' forget to see my home page (http://www.herve-thouzard.com/therename.phtml) to have the latest version of THE Rename or mail me at herve@herve-thouzard.com to send me your comments suggestions or bug reports." + ret

m_oAutoPos.AddAssignment Me.Text1, Me, tCONTAINER_WIDTH_DELTA_RIGHT
m_oAutoPos.AddAssignment Me.Text1, Me, tCONTAINER_HEIGHT_DELTA_BOTTOM
m_oAutoPos.AddAssignment Me.Command1, Me, tCONTAINER_RELATIVE_POS_BOTTOM
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    m_oAutoPos.RefreshPositions
    Command1.Left = (Me.ScaleWidth / 2) - (Command1.width / 2)
End Sub
