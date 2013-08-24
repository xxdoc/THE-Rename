VERSION 5.00
Begin VB.Form datetime 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Date and Time"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   ControlBox      =   0   'False
   Icon            =   "datetime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Height          =   300
      Left            =   2160
      TabIndex        =   1
      ToolTipText     =   "Click to select date"
      Top             =   90
      Width           =   1050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select"
      Height          =   300
      Left            =   2160
      TabIndex        =   7
      ToolTipText     =   "Click to select date"
      Top             =   900
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Select"
      Height          =   300
      Left            =   2160
      TabIndex        =   4
      ToolTipText     =   "Click to select date"
      Top             =   495
      Width           =   1050
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Time"
      Height          =   300
      Left            =   4410
      TabIndex        =   2
      ToolTipText     =   "Click to select time"
      Top             =   90
      Width           =   1050
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Time"
      Height          =   300
      Left            =   4410
      TabIndex        =   8
      ToolTipText     =   "Click to select time"
      Top             =   900
      Width           =   1050
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Time"
      Height          =   300
      Left            =   4410
      TabIndex        =   5
      ToolTipText     =   "Click to select time"
      Top             =   495
      Width           =   1050
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Keep"
      Height          =   195
      Left            =   1305
      TabIndex        =   0
      ToolTipText     =   "Don't modify date and time"
      Top             =   180
      Width           =   735
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Keep"
      Height          =   195
      Left            =   1305
      TabIndex        =   6
      ToolTipText     =   "Don't modify date and time"
      Top             =   945
      Width           =   735
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Keep"
      Height          =   195
      Left            =   1305
      TabIndex        =   3
      ToolTipText     =   "Don't modify date and time"
      Top             =   540
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   2121
      TabIndex        =   9
      ToolTipText     =   "Cancel your selection"
      Top             =   1410
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   3325
      TabIndex        =   10
      ToolTipText     =   "Save settings"
      Top             =   1410
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Creation Date"
      Height          =   195
      Left            =   135
      TabIndex        =   19
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Access"
      Height          =   195
      Left            =   135
      TabIndex        =   18
      Top             =   945
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Modified"
      Height          =   195
      Left            =   135
      TabIndex        =   17
      Top             =   540
      Width           =   945
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "??/??/????"
      Height          =   195
      Left            =   3330
      TabIndex        =   16
      ToolTipText     =   "It will be the date"
      Top             =   540
      Width           =   870
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "??/??/????"
      Height          =   195
      Left            =   3330
      TabIndex        =   15
      ToolTipText     =   "It will be the date"
      Top             =   945
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "??/??/????"
      Height          =   195
      Left            =   3330
      TabIndex        =   14
      ToolTipText     =   "It will be the date"
      Top             =   180
      Width           =   870
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "??:??:??"
      Height          =   195
      Left            =   5625
      TabIndex        =   13
      ToolTipText     =   "It will be the time"
      Top             =   540
      Width           =   630
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "??:??:??"
      Height          =   195
      Left            =   5625
      TabIndex        =   12
      ToolTipText     =   "It will be the time"
      Top             =   945
      Width           =   630
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "??:??:??"
      Height          =   195
      Left            =   5625
      TabIndex        =   11
      ToolTipText     =   "It will be the time"
      Top             =   180
      Width           =   630
   End
End
Attribute VB_Name = "datetime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check2_Click()
 If Check2.Value = 1 Then
  Command1.Enabled = False
  Command4.Enabled = False
 Else
  Command1.Enabled = True
  Command4.Enabled = True
 End If
End Sub

Private Sub Check3_Click()
 If Check3.Value = 1 Then
  Command2.Enabled = False
  Command5.Enabled = False
 Else
  Command2.Enabled = True
  Command5.Enabled = True
 End If
End Sub

Private Sub Check4_Click()
 If Check4.Value = 1 Then
  Command3.Enabled = False
  Command6.Enabled = False
 Else
  Command3.Enabled = True
  Command6.Enabled = True
 End If

End Sub

Private Sub cmdCancel_Click()
 modifdate = False
 Unload Me
End Sub

Private Sub cmdOK_Click()
 modifdate = False
 If Label6.Caption <> "??/??/????" Then
  modifdate = True
 End If
 If Label5.Caption <> "??/??/????" Then
  modifdate = True
 End If
 If Label4.Caption <> "??/??/????" Then
  modifdate = True
 End If
 If Label9.Caption <> "??:??:??" Then
  modifdate = True
 End If
 If Label8.Caption <> "??:??:??" Then
  modifdate = True
 End If
 If Label7.Caption <> "??:??:??" Then
  modifdate = True
 End If

 If Check2.Value = 1 Then
  temdate1 = False
 Else
  temdate1 = True
 End If
 
 If Check3.Value = 1 Then
  temdate2 = False
 Else
  temdate2 = True
 End If
 If Check4.Value = 1 Then
  temdate3 = False
 Else
  temdate3 = True
 End If
  ccheck2 = Check2.Value
  ccheck3 = Check3.Value
  ccheck4 = Check4.Value
  dDate1 = Label6.Caption
  dDate2 = Label5.Caption
  dDate3 = Label4.Caption
  heure1 = Label9.Caption
  heure2 = Label8.Caption
  heure3 = Label7.Caption
  Unload Me
End Sub

Private Sub Command1_Click()
Dim UserDate As Date
Dim verif As Date
UserDate = Date
If frmCalendar.GetDate(UserDate) Then
    Label6.Caption = UserDate
    If Label5.Caption <> "??/??/????" And Check3.Value = 0 Then
     verif = Label5.Caption
     If UserDate > verif Then
      MsgBox "Warning, Creation date is greater than the last access date !"
     End If
    End If
    If Label4.Caption <> "??/??/????" And Check4.Value = 0 Then
     verif = Label4.Caption
     If UserDate > verif Then
      MsgBox "Warning, Creation date is greater than the last modified date !"
     End If
    End If
 End If
End Sub

Private Sub Command2_Click()
Dim UserDate As Date
Dim verif As Date
UserDate = Date
If frmCalendar.GetDate(UserDate) Then
    Label5.Caption = UserDate
    If Label6.Caption <> "??/??/????" And Check2.Value = 0 Then
     verif = Label6.Caption
     If UserDate < verif Then
      MsgBox "Warning, this date is lower than the creation date !"
     End If
    End If
 End If
End Sub
Private Sub Command3_Click()
Dim UserDate As Date
Dim verif As Date
UserDate = Date
If frmCalendar.GetDate(UserDate) Then
    Label4.Caption = UserDate
    If Label6.Caption <> "??/??/????" And Check2.Value = 0 Then
     verif = Label6.Caption
     If UserDate < verif Then
      MsgBox "Warning, this date is lower than the creation date !"
     End If
    End If
 End If
End Sub

Private Sub Command4_Click()
 flheure.Show 1
 Label9.Caption = lheure
End Sub

Private Sub Command5_Click()
 flheure.Show 1
 Label8.Caption = lheure
End Sub

Private Sub Command6_Click()
 flheure.Show 1
 Label7.Caption = lheure
End Sub

Private Sub Form_Load()
  Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
  modifdate = False
  Check2.Value = ccheck2
  Check3.Value = ccheck3
  Check4.Value = ccheck4
  
  Label6.Caption = dDate1
  Label5.Caption = dDate2
  Label4.Caption = dDate3
  Label9.Caption = heure1
  Label8.Caption = heure2
  Label7.Caption = heure3
End Sub
