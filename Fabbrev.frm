VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Fabbrev 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abbreviations"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   HelpContextID   =   266
   Icon            =   "Fabbrev.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Text2 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   420
      Width           =   1815
   End
   Begin VB.ComboBox Text1 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Top             =   80
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   "Use Regular Expression"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   105
      WhatsThisHelpID =   273
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modify"
      Height          =   315
      Left            =   6300
      TabIndex        =   6
      ToolTipText     =   "Modify an item"
      Top             =   400
      WhatsThisHelpID =   272
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Height          =   330
      HelpContextID   =   168
      Left            =   1200
      Picture         =   "Fabbrev.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Load from a file"
      Top             =   4680
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   282
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Height          =   330
      HelpContextID   =   168
      Left            =   780
      Picture         =   "Fabbrev.frx":060C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Save to a File"
      Top             =   4680
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   281
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   315
      Left            =   5520
      TabIndex        =   5
      ToolTipText     =   "Add current Abbreviation"
      Top             =   400
      WhatsThisHelpID =   271
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Height          =   330
      HelpContextID   =   168
      Left            =   60
      Picture         =   "Fabbrev.frx":07D6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Remove selected abbreviation"
      Top             =   4680
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   280
      Width           =   330
   End
   Begin VB.Frame Frame2 
      Caption         =   "Where ? "
      Height          =   615
      Left            =   2610
      TabIndex        =   23
      Top             =   780
      Width           =   4755
      Begin VB.OptionButton Option1 
         Caption         =   "Both"
         Height          =   195
         Index           =   2
         Left            =   3360
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         WhatsThisHelpID =   278
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "In Extension"
         Height          =   195
         Index           =   1
         Left            =   1710
         TabIndex        =   10
         Top             =   240
         WhatsThisHelpID =   277
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "In Prefix"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         WhatsThisHelpID =   276
         Width           =   915
      End
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Match case"
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   105
      WhatsThisHelpID =   270
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Replace "
      Height          =   615
      Left            =   150
      TabIndex        =   21
      Top             =   780
      Width           =   2295
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Top             =   210
         WhatsThisHelpID =   275
         Width           =   675
      End
      Begin VB.CheckBox Check1 
         Caption         =   "All"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         WhatsThisHelpID =   274
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "times"
         Enabled         =   0   'False
         Height          =   195
         Left            =   1500
         TabIndex        =   22
         Top             =   240
         Width           =   360
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4500
      TabIndex        =   3
      Text            =   "1"
      Top             =   420
      WhatsThisHelpID =   269
      Width           =   555
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   168
      Left            =   5145
      TabIndex        =   16
      ToolTipText     =   "Don't use abbreviations"
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      HelpContextID   =   168
      Left            =   6300
      TabIndex        =   17
      ToolTipText     =   "Use abbreviations and save"
      Top             =   4680
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3075
      Left            =   0
      TabIndex        =   12
      Top             =   1500
      WhatsThisHelpID =   279
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5424
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Abbreviation"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Substitution"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Starting"
         Object.Width           =   1270
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Replace All"
         Object.Width           =   1826
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Match Case"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Everywhere"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "RegExp"
         Object.Width           =   1852
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Starting position"
      Height          =   195
      Left            =   3000
      TabIndex        =   20
      Top             =   465
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Substitution"
      Height          =   195
      Left            =   60
      TabIndex        =   19
      Top             =   465
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Abbreviation"
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "Fabbrev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cHist6 As New cHistory
Dim cHist7 As New cHistory
Dim lindex As Integer
Dim Signature As String
Dim SIni As New cInifile
Dim itemclick As Integer
Private Sub Check1_Click()
    If Check1.Value = 1 Then
        Text4.Enabled = False
        Label4.Enabled = False
    Else
        Text4.Enabled = True
        Label4.Enabled = True
    End If
End Sub

Private Sub Check3_Click()
 If Check3.Value = 1 Then
    Text3.Enabled = False
    Label3.Enabled = False
 Else
    Text3.Enabled = True
    Label3.Enabled = True
 End If
End Sub

Private Sub cmdCancel_Click()
    OkUseAbbrev = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    Dim vnb As Long
    Dim cH As String
    OkUseAbbrev = True
    ' Suppression du contenu initial de la collection
    Do While CollAbrev.Count > 0
        CollAbrev.Remove 1
    Loop
    ' Sauvegarde dans la collection des éléments entrés
    vnb = ListView1.ListItems.Count
    For i = 1 To vnb
        cH = ListView1.ListItems(i).Text + Chr$(254) ' Ajout de l'abréviation
        cH = cH + ListView1.ListItems(i).SubItems(1) + Chr$(254)
        cH = cH + ListView1.ListItems(i).SubItems(2) + Chr$(254)
        cH = cH + ListView1.ListItems(i).SubItems(3) + Chr$(254)
        cH = cH + ListView1.ListItems(i).SubItems(4) + Chr$(254)
        cH = cH + ListView1.ListItems(i).SubItems(5) + Chr$(254)
        cH = cH + ListView1.ListItems(i).SubItems(6)
        CollAbrev.Add cH, Str$(i)
        cH = ""
    Next
    If vnb = 0 Then
        OkUseAbbrev = False
    End If
    Unload Me
End Sub

Private Sub Command1_Click()
Dim itmX As ListItem
If Trim$(Text1.Text) = "" Then
    MsgBox "Type an abbreviation to add"
    Exit Sub
End If
 cHist6.AddNewItem Text1.Text
 cHist7.AddNewItem Text2.Text

Set itmX = ListView1.ListItems.Add(, , Text1.Text)
itmX.Text = Text1.Text
itmX.SubItems(1) = Text2.Text
itmX.SubItems(2) = Text3.Text
If Check1.Value = 1 Then
    itmX.SubItems(3) = "yes"
Else
    itmX.SubItems(3) = Text4.Text
End If
If Check2.Value = 1 Then
    itmX.SubItems(4) = "yes"
Else
    itmX.SubItems(4) = "no"
End If

If Check3.Value = 1 Then
    itmX.SubItems(6) = "yes"
Else
    itmX.SubItems(6) = "no"
End If

Select Case lindex
    Case 0
        itmX.SubItems(5) = "Prefix"
    Case 1
        itmX.SubItems(5) = "Extension"
    Case 2
        itmX.SubItems(5) = "yes"
End Select
Text1.Text = ""
Text2.Text = ""
Text3.Text = "1"
Check1.Value = 1
Text4.Text = ""
Option1(2).Value = True
Text1.SetFocus
End Sub

Private Sub Command2_Click()
Dim szFilename As String
Dim i As Long
Dim vnb As Long
Dim cH As String

If ListView1.ListItems.Count = 0 Then
    MsgBox "Sorry there are no abbreviations to save !", vbOKOnly, "Warning"
    Exit Sub
End If

szFilename = DialogFile(Me.hWnd, 2, "Save abbreviations as", "abbrev.abr", "Abbreviation" & Chr$(0) & "*.abr" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "Abbreviation")
If szFilename = "" Then
    Exit Sub
End If
With SIni
     .Path = szFilename
     .Section = "General"
     .Key = "NumberOfAbbreviations"
     .Value = Trim$(Str$(ListView1.ListItems.Count))
End With
With SIni
     .Path = szFilename
     .Section = "General"
     .Key = "Signature"
     .Value = Signature
End With
cH = ""
vnb = ListView1.ListItems.Count
For i = 1 To vnb
    cH = ListView1.ListItems(i).Text + Chr$(254) ' Ajout de l'abréviation
    cH = cH + ListView1.ListItems(i).SubItems(1) + Chr$(254)
    cH = cH + ListView1.ListItems(i).SubItems(2) + Chr$(254)
    cH = cH + ListView1.ListItems(i).SubItems(3) + Chr$(254)
    cH = cH + ListView1.ListItems(i).SubItems(4) + Chr$(254)
    cH = cH + ListView1.ListItems(i).SubItems(5) + Chr$(254)
    cH = cH + ListView1.ListItems(i).SubItems(6)
    With SIni
        .Path = szFilename
        .Section = "Abbreviations"
        .Key = "Abbrev" + Trim$(Str$(i))
        .Value = cH
    End With
    cH = ""
Next
End Sub

Private Sub Command3_Click()
Dim szFilename As String
Dim vretour As Integer
Dim sValue As String
Dim i As Integer, j As Integer, vnb As Integer, cH As String
Dim itmX As ListItem
szFilename = DialogFile(Me.hWnd, 1, "Open Abbreviations", "abbrev.abr", "Abbreviation" & Chr$(0) & "*.abr" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "Abbreviation")
If Trim$(szFilename) = "" Then
    Exit Sub
End If

If ListView1.ListItems.Count > 0 Then
    vretour = MsgBox("Delete current abbreviations ?", vbYesNoCancel, "Warning")
    Select Case vretour
        Case vbCancel
            Exit Sub
        Case vbYes
            ListView1.ListItems.Clear
    End Select
End If
' Vérification de la signature
With SIni
    .Path = szFilename
     .Section = "General"
     .Key = "Signature"
    sValue = .Value
End With
If sValue <> Signature Then
    MsgBox "Sorry but it's not an abbreviation's file coming from THE Rename !"
    Exit Sub
End If
' Lecture du nombre d'abréviations
With SIni
     .Path = szFilename
     .Section = "General"
     .Key = "NumberOfAbbreviations"
     sValue = .Value
End With
If Val(sValue) = 0 Then
    MsgBox "Sorry, there's no abbreviations in this file !"
    Exit Sub
End If
vnb = Val(sValue)

' Chargement dans le listview
For i = 1 To vnb
    With SIni
        .Path = szFilename
        .Section = "Abbreviations"
        .Key = "Abbrev" + Trim$(Str$(i))
        sValue = .Value
    End With
    cH = GetToken(sValue, Chr$(254), 1)
    Set itmX = ListView1.ListItems.Add(, , cH)
    itmX.Text = cH
    For j = 2 To 7
        itmX.SubItems(j - 1) = GetToken(sValue, Chr$(254), j)
    Next
Next
End Sub

Private Sub Command4_Click()
If itemclick = -1 Then
 Exit Sub
End If
If Trim$(Text1.Text) = "" Then
    MsgBox "Type an abbreviation to add"
    Exit Sub
End If
If Trim$(Text2.Text) = "" Then
    MsgBox "Type a substitution text"
    Exit Sub
End If

ListView1.ListItems(itemclick).Text = Text1.Text
ListView1.ListItems(itemclick).SubItems(1) = Text2.Text
ListView1.ListItems(itemclick).SubItems(2) = Text3.Text

If Check1.Value = 1 Then
    ListView1.ListItems(itemclick).SubItems(3) = "yes"
Else
    ListView1.ListItems(itemclick).SubItems(3) = Text4.Text
End If
If Check2.Value = 1 Then
    ListView1.ListItems(itemclick).SubItems(4) = "yes"
Else
    ListView1.ListItems(itemclick).SubItems(4) = "no"
End If
Select Case lindex
    Case 0
        ListView1.ListItems(itemclick).SubItems(5) = "Prefix"
    Case 1
        ListView1.ListItems(itemclick).SubItems(5) = "Extension"
    Case 2
        ListView1.ListItems(itemclick).SubItems(5) = "yes"
End Select
If Check3.Value = 1 Then
    ListView1.ListItems(itemclick).SubItems(6) = "yes"
Else
    ListView1.ListItems(itemclick).SubItems(6) = "no"
End If

End Sub

Private Sub Command5_Click()
If itemclick = -1 Then
 Exit Sub
End If
ListView1.ListItems.Remove itemclick
itemclick = -1
Text1.Text = ""
Text2.Text = ""
Text3.Text = "1"
Check1.Value = 1
Text4.Text = ""
Option1(2).Value = True
'Text1.SetFocus
ListView1.SetFocus
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    Dim cH As String
    Dim itmX As ListItem
    cHist6.sKey = "FindAbbrev"
    cHist6.Items Text1
    cHist7.sKey = "ReplaceAbbrev"
    cHist7.Items Text2
    lindex = 2
    Signature = "THE Rename's abbreviation file by Hervé Thouzard (hthouzard@bigfoot.com) - version 1.00"
    itemclick = -1
    ' Chargement de la collection dans le listview
    For i = 1 To CollAbrev.Count
        cH = GetToken(CollAbrev.Item(i), Chr$(254), 1)
        Set itmX = ListView1.ListItems.Add(, , cH)
        itmX.Text = cH
        For j = 2 To 7
            itmX.SubItems(j - 1) = GetToken(CollAbrev.Item(i), Chr$(254), j)
        Next
    Next
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub ListView1_DblClick()
    Text1.SetFocus
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
 itemclick = Item.Index
 Text1.Text = Item.Text
 Text2.Text = Item.ListSubItems(1).Text
 Text3.Text = Item.ListSubItems(2).Text
 If Item.ListSubItems(3).Text = "yes" Then
    Check1.Value = 1
    Text4.Text = ""
 Else
    Check1.Value = 0
    Text4.Text = Item.ListSubItems(3).Text
 End If
 If Item.ListSubItems(4).Text = "yes" Then
    Check2.Value = 1
 Else
    Check2.Value = 0
 End If
 If Item.ListSubItems(6).Text = "yes" Then
    Check3.Value = 1
 Else
    Check3.Value = 0
 End If
 Select Case UCase$(Item.ListSubItems(5).Text)
    Case "PREFIX"
        Option1(0).Value = True
    Case "EXTENSION"
        Option1(1).Value = True
    Case "YES"
        Option1(2).Value = True
 End Select
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        Command5_Click
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
 lindex = Index
End Sub

Private Sub Text1_GotFocus()
    SelAll Text1
End Sub

Private Sub Text2_GotFocus()
    SelAll Text2
End Sub

Private Sub Text3_GotFocus()
    SelAll Text3
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
    If Val(Text3.Text) < 1 And Trim$(Text3.Text) <> "" Then Cancel = True
End Sub

Private Sub Text4_GotFocus()
    SelAll Text4
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
    If Val(Text4.Text) < 1 And Trim$(Text4.Text) <> "" Then Cancel = True
End Sub
