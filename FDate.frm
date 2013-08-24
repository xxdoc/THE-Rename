VERSION 5.00
Begin VB.Form FDate 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Select a date"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   38
      Left            =   360
      TabIndex        =   3
      Top             =   540
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      HelpContextID   =   38
      Left            =   1560
      TabIndex        =   4
      Top             =   540
      Width           =   1095
   End
   Begin VB.TextBox syear 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "2001"
      Top             =   60
      Width           =   495
   End
   Begin VB.TextBox smonth 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "1"
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox sday 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   420
      TabIndex        =   0
      Text            =   "1"
      Top             =   60
      Width           =   315
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "year"
      Height          =   195
      Left            =   1860
      TabIndex        =   7
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "month"
      Height          =   195
      Left            =   900
      TabIndex        =   6
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Day"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   285
   End
End
Attribute VB_Name = "FDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private s_Ok As Boolean
Private TheDate As Date
Public Function GetDateManually(zdate As Date) As Boolean
    ' Initialisation des données avec une date
    s_Ok = False
    If IsDate(zdate) Then
        sday = Day(zdate)
        smonth = Month(zdate)
        syear = Year(zdate)
    Else
        syear = Year(Date)
        sday = Day(Date)
        smonth = Month(Date)
    End If
    Me.Show vbModal
    If s_Ok Then
        zdate = TheDate
    End If
    GetDateManually = s_Ok
End Function
Private Sub cmdCancel_Click()
    s_Ok = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim zdate As String
    zdate = sday + "/" + smonth + "/" + syear
    If Not IsDate(zdate) Then
        MsgBox "Error, your date is not valid", vbOKOnly, "Invalid date"
        sday.SetFocus
    Else
        s_Ok = True
        TheDate = CDate(sday & "/" & smonth & "/" & syear)
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub sday_GotFocus()
    SelAll sday
End Sub

Private Sub sday_Validate(Cancel As Boolean)
    If Val(sday) <= 0 Then
        Cancel = True
    End If
    
    If Val(sday) > 31 Then
        Cancel = True
    End If
End Sub

Private Sub smonth_GotFocus()
    SelAll smonth
End Sub

Private Sub smonth_Validate(Cancel As Boolean)
    If Val(smonth) <= 0 Then
        Cancel = True
    End If
    If Val(smonth) > 12 Then
        Cancel = True
    End If
End Sub

Private Sub syear_GotFocus()
    SelAll syear
End Sub

Private Sub syear_Validate(Cancel As Boolean)
    If Val(syear) = 0 Then
        Cancel = True
    End If
End Sub
