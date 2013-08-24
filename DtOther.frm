VERSION 5.00
Begin VB.Form DtOther 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Other"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "DtOther.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Command1 
      Caption         =   "&Default"
      Height          =   300
      HelpContextID   =   14
      Left            =   1755
      TabIndex        =   4
      ToolTipText     =   "Set date's format to it''s default value"
      Top             =   3600
      WhatsThisHelpID =   231
      Width           =   1185
   End
   Begin VB.TextBox Text3 
      Height          =   330
      HelpContextID   =   14
      Left            =   1770
      TabIndex        =   3
      ToolTipText     =   "Example for the current function only"
      Top             =   2880
      WhatsThisHelpID =   230
      Width           =   2790
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   14
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Use personal date format"
      Top             =   4140
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   14
      Left            =   1425
      TabIndex        =   5
      ToolTipText     =   "Cancel any modification"
      Top             =   4140
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   1170
      HelpContextID   =   14
      Left            =   1770
      MultiLine       =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Inline help for the current function"
      Top             =   1320
      WhatsThisHelpID =   228
      Width           =   2790
   End
   Begin VB.ListBox List1 
      Height          =   1890
      HelpContextID   =   14
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Double click to copy it in the text box"
      Top             =   1320
      WhatsThisHelpID =   227
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   315
      HelpContextID   =   14
      Left            =   75
      TabIndex        =   0
      ToolTipText     =   "Enter here some functions to create a date's format"
      Top             =   300
      WhatsThisHelpID =   226
      Width           =   4515
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Label6"
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   75
      TabIndex        =   12
      Top             =   3375
      Width           =   4515
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Height          =   240
      Left            =   90
      TabIndex        =   11
      ToolTipText     =   "This is the result of your selection"
      Top             =   675
      WhatsThisHelpID =   229
      Width           =   4515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sample for this function"
      Height          =   195
      Left            =   1800
      TabIndex        =   10
      Top             =   2610
      WhatsThisHelpID =   230
      Width           =   1650
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   1800
      TabIndex        =   9
      Top             =   1020
      WhatsThisHelpID =   228
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Select a function"
      Height          =   195
      Left            =   150
      TabIndex        =   8
      Top             =   1020
      WhatsThisHelpID =   227
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter date's format"
      Height          =   195
      Left            =   75
      TabIndex        =   7
      Top             =   75
      WhatsThisHelpID =   226
      Width           =   1320
   End
End
Attribute VB_Name = "DtOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OtDate(23, 2) As String  ' Contient les descriptions et commandes des dates personnalisés
Private Sub Command1_Click()
 Text1.Text = "short Date"
End Sub

Private Sub Command3_Click()
 If QDateTravail = 1 Then
    LesOptions.PersonnalDate = Text1.Text
 Else
    LesOptions.DisplayDate = Text1.Text
 End If
 Unload Me
End Sub

Private Sub Command4_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 Dim i As Integer
OtDate(1, 1) = "d"
OtDate(2, 1) = "dd"
OtDate(3, 1) = "ddd"
OtDate(4, 1) = "dddd"
OtDate(5, 1) = "ddddd"
OtDate(6, 1) = "w"
OtDate(7, 1) = "ww"
OtDate(8, 1) = "m"
OtDate(9, 1) = "mm"
OtDate(10, 1) = "mmm"
OtDate(11, 1) = "mmmm"
OtDate(12, 1) = "y"
OtDate(13, 1) = "yy"
OtDate(14, 1) = "yyyy"
OtDate(15, 1) = "h"
OtDate(16, 1) = "hh"
OtDate(17, 1) = "n"
OtDate(18, 1) = "nn"
OtDate(19, 1) = "s"
OtDate(20, 1) = "ss"
OtDate(21, 1) = "ttttt"
OtDate(22, 1) = "AM/PM"
OtDate(23, 1) = "am/pm"
OtDate(1, 2) = "Return the day as a number without any leading zero (1 - 31)."
OtDate(2, 2) = "Return the day as a number with a leading zero (01 for example)"
OtDate(3, 2) = "Return the day as an abbreviation (Sun - Sat)"
OtDate(4, 2) = "Return the day as a full name (Sunday - Saturday)"
OtDate(5, 2) = "Return the date as a complete date (including day, month, and year), formatted according to your  system's short date format setting"
OtDate(6, 2) = "Return the day of the week as a number (1=Sunday ... 7=Saturday)"
OtDate(7, 2) = "Return the week of the year as a number (1 - 54)"
OtDate(8, 2) = "Return the month as a number without any leading zero (1 - 12)"
OtDate(9, 2) = "Return the month as a number with a leading zero (01 - 12)"
OtDate(10, 2) = "Return the month as an abbreviation (Jan - Dec)"
OtDate(11, 2) = "Return the month as a full month name (January - December)"
OtDate(12, 2) = "Return the day of the year as a number (1 - 366)"
OtDate(13, 2) = "Return the year as a 2-digit number (00 - 99)"
OtDate(14, 2) = "Return the year as a 4-digit number (100 - 9999)"
OtDate(15, 2) = "Return the hour as a number without leading zeros (0 - 23)"
OtDate(16, 2) = "Return the hour as a number with leading zeros (00 - 23)"
OtDate(17, 2) = "Return the minute as a number without leading zeros (0 - 59)"
OtDate(18, 2) = "Return the minute as a number with leading zeros (00 - 59)"
OtDate(19, 2) = "Return the second as a number without leading zeros (0 - 59)"
OtDate(20, 2) = "Return the second as a number with leading zeros (00 - 59)"
OtDate(21, 2) = "Return a time as a complete time (including hour, minute, and second)"
OtDate(22, 2) = "Use the 12-hour clock and return an uppercase AM with any hour before noon; display an uppercase PM with any hour between noon and 11:59 P.M"
OtDate(23, 2) = "Use the 12-hour clock and return a lowercase AM with any hour before noon; display a lowercase PM with any hour between noon and 11:59 P.M"
 
 If QDateTravail = 1 Then
    Me.Caption = "Other - " + Str$(Now)
    Text1.Text = LesOptions.PersonnalDate
    Command1.Visible = False
    Label6.Caption = "Warning, the following characters are not legal in a filename :" + vbCrLf + "\ / : * ? < > | " + Chr$(34)
 Else
    Me.Caption = "Personalize display date - " + Str$(Now)
    Text1.Text = LesOptions.DisplayDate
    Label6.Caption = ""
    Command1.Visible = True
 End If
 For i = 1 To 23
    List1.AddItem OtDate(i, 1)
 Next
End Sub

Private Sub List1_Click()
Dim leformat As String
Dim i As Integer
leformat = Trim$(List1.List(List1.ListIndex))
 For i = 1 To 23
    If List1.List(List1.ListIndex) = OtDate(i, 1) Then
        Text2.Text = OtDate(i, 2)
        Text3.Text = Format$(Now, leformat)
        i = 23
    End If
 Next
End Sub

Private Sub List1_DblClick()
 InsertTextInTextBox Text1, List1
End Sub

Private Sub Text1_Change()
 Dim leformat As String
 If QDateTravail = 1 Then
    CharInterdits Text1.Text
 End If
 leformat = Text1.Text
 leformat = Replace(leformat, "%", "")
 On Error GoTo erreur
 Label5.Caption = Format$(Now, leformat)
 Exit Sub
 
erreur:
  Label5.Caption = "<Invalid command in your expression>"
End Sub
