VERSION 5.00
Begin VB.Form advsearch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search and replace"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   HelpContextID   =   110
   Icon            =   "advsearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4260
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Use Regular Expression"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   4035
      Width           =   2100
   End
   Begin VB.Frame Frame2 
      Caption         =   "Strings "
      Height          =   2490
      Left            =   53
      TabIndex        =   17
      Top             =   75
      Width           =   4155
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   285
         HelpContextID   =   110
         Left            =   2310
         TabIndex        =   4
         Text            =   "1"
         Top             =   2130
         WhatsThisHelpID =   220
         Width           =   510
      End
      Begin VB.ComboBox Text3 
         Height          =   315
         Left            =   1155
         TabIndex        =   1
         Top             =   1440
         Width           =   2805
      End
      Begin VB.ComboBox Text4 
         Height          =   315
         Left            =   1155
         TabIndex        =   0
         Top             =   960
         Width           =   2805
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         HelpContextID   =   110
         Left            =   1365
         TabIndex        =   12
         Top             =   585
         WhatsThisHelpID =   220
         Width           =   510
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         HelpContextID   =   110
         Left            =   2325
         TabIndex        =   13
         Top             =   585
         WhatsThisHelpID =   220
         Width           =   510
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Match case"
         Height          =   285
         HelpContextID   =   110
         Left            =   120
         TabIndex        =   2
         Top             =   1815
         WhatsThisHelpID =   223
         Width           =   1230
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   30
         ScaleHeight     =   285
         ScaleWidth      =   4035
         TabIndex        =   18
         Top             =   300
         Width           =   4035
         Begin VB.OptionButton Option1 
            Caption         =   "Search from &right"
            Height          =   195
            HelpContextID   =   110
            Index           =   1
            Left            =   2295
            TabIndex        =   11
            Top             =   0
            WhatsThisHelpID =   219
            Width           =   1500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Search from &left"
            Height          =   195
            HelpContextID   =   110
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   0
            WhatsThisHelpID =   219
            Width           =   1500
         End
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Repl&ace all"
         Height          =   285
         HelpContextID   =   110
         Left            =   1365
         TabIndex        =   3
         Top             =   1815
         WhatsThisHelpID =   224
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Number of substitutions "
         Height          =   195
         Left            =   540
         TabIndex        =   24
         Top             =   2175
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From position"
         Height          =   195
         Left            =   345
         TabIndex        =   22
         Top             =   630
         WhatsThisHelpID =   220
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "To"
         Height          =   195
         Left            =   2010
         TabIndex        =   21
         Top             =   630
         WhatsThisHelpID =   220
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Replace with"
         Height          =   195
         Left            =   75
         TabIndex        =   20
         Top             =   1455
         WhatsThisHelpID =   222
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Find"
         Height          =   195
         Left            =   75
         TabIndex        =   19
         Top             =   1005
         WhatsThisHelpID =   221
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Characters "
      Height          =   1215
      Left            =   53
      TabIndex        =   14
      Top             =   2715
      Width           =   4155
      Begin VB.ComboBox Text6 
         Height          =   315
         Left            =   1155
         TabIndex        =   6
         Top             =   720
         Width           =   2805
      End
      Begin VB.ComboBox Text5 
         Height          =   315
         Left            =   1155
         TabIndex        =   5
         Top             =   300
         Width           =   2805
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Replace with"
         Height          =   195
         Left            =   75
         TabIndex        =   16
         Top             =   795
         WhatsThisHelpID =   225
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Find"
         Height          =   195
         Left            =   75
         TabIndex        =   15
         Top             =   360
         WhatsThisHelpID =   225
         Width           =   300
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   110
      Left            =   764
      TabIndex        =   9
      ToolTipText     =   "Don't use search and replace"
      Top             =   4920
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   110
      Left            =   2176
      TabIndex        =   8
      ToolTipText     =   "Use search and replace"
      Top             =   4920
      Width           =   1320
   End
   Begin VB.Label Label7 
      Caption         =   "Remember, you can configure when to launch search and replace in the options."
      Height          =   420
      Left            =   60
      TabIndex        =   23
      Top             =   4425
      Width           =   4155
   End
End
Attribute VB_Name = "advsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim larech As New CSearch
Dim Supp As String
Dim cHist1 As New cHistory
Dim cHist2 As New cHistory
Dim cHist3 As New cHistory
Dim cHist4 As New cHistory

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Frame1.Enabled = False
        Option1(0).Enabled = False
        Option1(1).Enabled = False
        Label1.Enabled = False
        Label6.Enabled = False
        Text5.Enabled = False
        Text6.Enabled = False
    Else
        Frame1.Enabled = True
        Option1(0).Enabled = True
        Option1(1).Enabled = True
        Label1.Enabled = True
        Label6.Enabled = True
        Text5.Enabled = True
        Text6.Enabled = True
    End If
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then
        Label8.Enabled = False
        Text7.Enabled = False
    Else
        Label8.Enabled = True
        Text7.Enabled = True
    End If
End Sub

Private Sub Command1_Click()
  If Val(Text2.Text) <> 0 And Val(Text1.Text) = 0 Then
    MsgBox "Error, If you want to search from a position to an other position, the 'From position' must be a positive value or set both values to 0"
    Text1.SetFocus
    Exit Sub
  End If
  If RechEnCours = 1 Then
    RechPref = True
  Else
    If RechEnCours = 2 Then
        RechSuff = True
    Else
        RechGlob = True
    End If
  End If
  With larech
    If Option1(0).Value = True Then
        .SearchFromleft = True
    Else
        .SearchFromleft = False
    End If
    .SearchFrom = Val(Text1.Text)
    .SearchTo = Val(Text2.Text)
    .NumberOfSubst = Val(Text7.Text)
    If Check1.Value = 1 Then
        .MatchCase = True
    Else
        .MatchCase = False
    End If
    If Check4.Value = 1 Then
        .ReplaceAll = True
    Else
        .ReplaceAll = False
    End If
    .UseRegExp = Check2.Value
    .SearchString = Text4.Text
    .ReplaceString = Text3.Text
    .SearchCharacters = Text5.Text
    .ReplaceCharacters = Text6.Text
 End With
 
 cHist1.AddNewItem Text4.Text
 cHist2.AddNewItem Text3.Text
 cHist3.AddNewItem Text5.Text
 cHist4.AddNewItem Text6.Text
 Unload Me
End Sub

Private Sub Command2_Click()
  If RechEnCours = 1 Then
   RechPref = False
  Else
    If RechEnCours = 2 Then
        RechSuff = False
    Else
        RechGlob = False
    End If
  End If
 Unload Me
End Sub
Private Sub Form_Load()
  cHist1.sKey = "FindString"
  cHist1.Items Text4
  cHist2.sKey = "ReplaceString"
  cHist2.Items Text3
  cHist3.sKey = "FindChar"
  cHist3.Items Text5
  cHist4.sKey = "ReplaceChar"
  cHist4.Items Text6
  If RechEnCours = 1 Then ' Ouverture pour le préfixe
    Set larech = rech1
    advsearch.Caption = "Search and replace - Prefix"
    Supp = "Prefix"
  Else ' Ouverture pour le suffixe
    If RechEnCours = 2 Then
        Set larech = rech2
        advsearch.Caption = "Search and replace - Extension"
        Supp = "Extension"
    Else
        Set larech = rech3
        advsearch.Caption = "Global Search and replace"
    End If
  End If
  If larech.SearchFromleft = True Then
   Option1(0).Value = True
   Option1(1).Value = False
  Else
   Option1(0).Value = False
   Option1(1).Value = True
  End If
  Check2.Value = larech.UseRegExp
  Text1.Text = larech.SearchFrom
  Text2.Text = larech.SearchTo
  If larech.MatchCase = True Then
   Check1.Value = 1
  Else
   Check1.Value = 0
  End If
  If larech.ReplaceAll = True Then
   Check4.Value = 1
  Else
   Check4.Value = 0
  End If
  Text5.Text = larech.SearchCharacters
  Text6.Text = larech.ReplaceCharacters
  Text4.Text = larech.SearchString
  Text3.Text = larech.ReplaceString
  Text7.Text = larech.NumberOfSubst
  Check2_Click
    If Check4.Value = 1 Then
        Label8.Enabled = False
        Text7.Enabled = False
    Else
        Label8.Enabled = True
        Text7.Enabled = True
    End If

End Sub

Private Sub Text1_GotFocus()
    SelAll Text1
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
    If Val(Text1.Text) < 0 Then Cancel = True
End Sub

Private Sub Text2_GotFocus()
    SelAll Text2
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
 If Val(Text2.Text) < 0 Then Cancel = True
End Sub

Private Sub Text3_Change()
    If Check2.Value = 0 Then
        CharInterdits Text3.Text
    End If
End Sub

Private Sub Text3_GotFocus()
    SelAll Text3
End Sub

Private Sub Text4_GotFocus()
    SelAll Text4
End Sub

Private Sub Text5_GotFocus()
    SelAll Text5
End Sub

Private Sub Text6_Change()
 CharInterdits Text6.Text
End Sub

Private Sub Text6_GotFocus()
    SelAll Text6
End Sub

Private Sub Text7_GotFocus()
    SelAll Text7
End Sub

Private Sub Text7_Validate(Cancel As Boolean)
    If Val(Text7.Text) < 1 And Text7.Enabled = True Then Cancel = True
End Sub

