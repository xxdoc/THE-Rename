VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FPresets 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presets"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "FPresets.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   3345
      TabIndex        =   21
      ToolTipText     =   "This is the replace expression"
      Top             =   2535
      WhatsThisHelpID =   253
      Width           =   3300
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   6225
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   2760
      TabIndex        =   18
      ToolTipText     =   "This is the current replace used by THE Rename"
      Top             =   3600
      WhatsThisHelpID =   261
      Width           =   2647
   End
   Begin VB.ListBox List3 
      Height          =   1815
      HelpContextID   =   251
      Left            =   3360
      MultiSelect     =   2  'Extended
      TabIndex        =   16
      Top             =   300
      WhatsThisHelpID =   251
      Width           =   3300
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   20
      Left            =   6690
      Picture         =   "FPresets.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Move Down"
      Top             =   1215
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   264
      Width           =   330
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   20
      Left            =   6690
      Picture         =   "FPresets.frx":024C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Move Up"
      Top             =   720
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   263
      Width           =   330
   End
   Begin VB.ListBox List1 
      Height          =   1815
      HelpContextID   =   251
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   11
      Top             =   300
      WhatsThisHelpID =   251
      Width           =   3300
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   10
      ToolTipText     =   "This is the search description"
      Top             =   2175
      WhatsThisHelpID =   253
      Width           =   5595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Use &Now"
      Height          =   315
      Left            =   60
      TabIndex        =   9
      ToolTipText     =   "Use the selected search now and close this window"
      Top             =   2940
      WhatsThisHelpID =   255
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   975
      TabIndex        =   8
      ToolTipText     =   "Remove the selected search"
      Top             =   2940
      WhatsThisHelpID =   256
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Set as default search"
      Height          =   315
      Left            =   1725
      TabIndex        =   7
      ToolTipText     =   "The selected search will be used when you use search and replace"
      Top             =   2940
      WhatsThisHelpID =   257
      Width           =   1995
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Modify"
      Height          =   315
      Left            =   5700
      TabIndex        =   6
      ToolTipText     =   "Modify selected command"
      Top             =   2175
      WhatsThisHelpID =   254
      Width           =   915
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   75
      TabIndex        =   5
      ToolTipText     =   "This is the current search used by THE Rename"
      Top             =   3600
      WhatsThisHelpID =   261
      Width           =   2647
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Add to list"
      Height          =   315
      Left            =   5460
      TabIndex        =   4
      ToolTipText     =   "Add current search and replace to saved list"
      Top             =   3600
      WhatsThisHelpID =   262
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Close"
      Height          =   315
      Left            =   6105
      TabIndex        =   3
      ToolTipText     =   "Close and save"
      Top             =   2940
      WhatsThisHelpID =   260
      Width           =   795
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   7560
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clear default search"
      Height          =   315
      Left            =   3780
      TabIndex        =   1
      ToolTipText     =   "Don't use any default search"
      Top             =   2940
      WhatsThisHelpID =   258
      Width           =   1875
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "This is the search expression"
      Top             =   2535
      WhatsThisHelpID =   253
      Width           =   3300
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3413
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
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Descr"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Search forr"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Replace with"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Position"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "To"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Match case"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "# subst."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Searched Char."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Replace Char."
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Current replace"
      Height          =   195
      Left            =   2760
      TabIndex        =   19
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Replace with"
      Height          =   195
      Left            =   4335
      TabIndex        =   17
      Top             =   45
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Search for"
      Height          =   195
      Left            =   990
      TabIndex        =   15
      Top             =   30
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current search"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Width           =   1035
   End
End
Attribute VB_Name = "FPresets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SIni As New cInifile
Dim Synchro As Boolean
Dim Supp As String
Dim modifs As Boolean
Dim EnCours As Boolean
Private Sub cmdDown_Click()
  On Error Resume Next
  Dim nItem As Integer
  List2.ListIndex = List1.ListIndex
  List3.ListIndex = List1.ListIndex
  EnCours = True
  With List1
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
    .AddItem .Text, nItem + 2
    .RemoveItem nItem
    .Selected(nItem + 1) = True
  End With
  With List2
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
    .AddItem .Text, nItem + 2
    .RemoveItem nItem
    .Selected(nItem + 1) = True
  End With
  With List3
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
    .AddItem .Text, nItem + 2
    .RemoveItem nItem
    .Selected(nItem + 1) = True
  End With
  modifs = True
'  List1_Click
End Sub

Private Sub cmdUp_Click()
  On Error Resume Next
  Dim nItem As Integer
  EnCours = True
  List2.ListIndex = List1.ListIndex
  List3.ListIndex = List1.ListIndex
  With List1
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'can't move 1st item up
    .AddItem .Text, nItem - 1
    .RemoveItem nItem + 1
    .Selected(nItem - 1) = True
  End With
  With List2
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'can't move 1st item up
    .AddItem .Text, nItem - 1
    .RemoveItem nItem + 1
    .Selected(nItem - 1) = True
  End With
  With List3
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'can't move 1st item up
    .AddItem .Text, nItem - 1
    .RemoveItem nItem + 1
    .Selected(nItem - 1) = True
  End With
  modifs = True
'  List1_Click
End Sub

Private Sub Command1_Click()
    If List1.ListIndex = -1 Then
        Exit Sub
    End If
    If Trim(advsearch.Text4.Text) <> "" Then
        If MsgBox("Warning, your actual search line contains some text, do you want to replace it with this search ?", vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    advsearch.Text4.Text = List1.List(List1.ListIndex)
    advsearch.Text3.Text = List3.List(List3.ListIndex)
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim Index As Integer
    Dim i As Integer
    Dim vnb As Integer
    Dim chemin As String
    
    chemin = AppPath + "presets.ini"
    
    If List1.ListIndex = -1 Then
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete this search ?" + vbCrLf + List1.List(List1.ListIndex), vbYesNo, "Warning") = vbNo Then
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    Index = List1.ListIndex
    StatusBar1.SimpleText = "Removing command"
    List1.RemoveItem (Index)
    List2.RemoveItem (Index)
    List3.RemoveItem (Index)
    StatusBar1.SimpleText = "Updating presets file"
    vnb = List1.ListCount
    ' Sauvegarde du nombre de commandes
    With SIni
         .Path = chemin
         .Section = "General"
         .Key = "NumberOfPresets" & Supp
         .Value = Trim(Str(vnb))
    End With
    ' Suppression des commandes actuelles
    With SIni
         .Path = chemin
         .Section = "Find" & Supp
         .DeleteSection
    End With
    With SIni
         .Path = chemin
         .Section = "Replace" & Supp
         .DeleteSection
    End With
     ' Suppression des descriptions associées
    With SIni
         .Path = chemin
         .Section = "Descriptions" & Supp
         .DeleteSection
    End With
    
    For i = 0 To vnb - 1
        With SIni
            .Path = chemin
            .Section = "Find" & Supp
            .Key = "Find" & Trim(Str(i + 1))
            .Value = List1.List(i)
        End With
        With SIni
            .Path = chemin
            .Section = "Replace" & Supp
            .Key = "Replace" & Trim(Str(i + 1))
            .Value = List3.List(i)
        End With
        
        With SIni
            .Path = chemin
            .Section = "Descriptions" & Supp
            .Key = "Desc" & Trim(Str(i + 1))
            .Value = List2.List(i)
        End With
    Next i
    StatusBar1.SimpleText = "Command deleted !"
    Text2.Text = ""
    Text4.Text = ""
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
        List3.ListIndex = 0
    End If
    Me.MousePointer = vbNormal
End Sub

Private Sub Command3_Click()
    If List1.ListIndex = -1 Then
        Exit Sub
    End If
    If RechEnCours = 1 Then
        DefaultSearchPr = List1.List(List1.ListIndex)
        DefaultReplacePr = List3.List(List3.ListIndex)
    Else
        DefaultSearchExt = List1.List(List1.ListIndex)
        DefaultReplaceExt = List3.List(List3.ListIndex)
    End If
    
End Sub

Private Sub Command4_Click()
    Dim Index As Integer
    Dim chemin As String
    
    chemin = AppPath + "presets.ini"
    
    If List1.ListIndex = -1 Then
        Exit Sub
    End If
    Index = List1.ListIndex
    List2.List(Index) = Text1.Text
    List1.List(Index) = Text3.Text
    List3.List(Index) = Text5.Text
    With SIni
        .Path = chemin
        .Section = "Descriptions" & Supp
        .Key = "Desc" & Trim(Str(Index + 1))
        .Value = Text1.Text
    End With
    With SIni
        .Path = chemin
        .Section = "Find" & Supp
        .Key = "Find" & Trim(Str(Index + 1))
        .Value = Text3.Text
    End With
    With SIni
        .Path = chemin
        .Section = "Replace" & Supp
        .Key = "Replace" & Trim(Str(Index + 1))
        .Value = Text5.Text
    End With
End Sub

Private Sub Command5_Click()
    Dim i As Integer
    Dim descr As String
    Dim chemin As String
    Dim Index As Integer
    Dim vnb As Integer
    chemin = AppPath + "presets.ini"
    
    If Trim(Text2.Text) = "" Then
        Exit Sub
    End If
    ' D'abord on vérifie que la commande n'est pas déjà présente dans la liste
    For i = 0 To List1.ListCount - 1
        If Trim(List1.List(i)) = Trim(Text2.Text) Then
            MsgBox "Warning, I can't add this search to the list because it is already in the list !", vbOKOnly, "Warning"
            Exit Sub
        End If
    Next
    ' Elle n'est pas présente donc on peut l'ajouter. On demande sa description
    descr = InputBox("Enter a description for this search", "Description")
    If descr = "" Then
        descr = " "
    End If
    List1.AddItem Text2.Text
    List3.AddItem Text5.Text
    List2.AddItem descr
    Me.MousePointer = vbHourglass
    StatusBar1.SimpleText = "Updating presets file"
    vnb = List1.ListCount
    ' Sauvegarde du nombre de commandes
    With SIni
         .Path = chemin
         .Section = "General"
         .Key = "NumberOfPresets" & Supp
         .Value = Trim(Str(vnb))
    End With
    ' Suppression des commandes actelles
    With SIni
         .Path = chemin
         .Section = "Find" & Supp
         .DeleteSection
    End With
    With SIni
         .Path = chemin
         .Section = "Replace" & Supp
         .DeleteSection
    End With
     ' Suppression des descriptions associées
    With SIni
         .Path = chemin
         .Section = "Descriptions" & Supp
         .DeleteSection
    End With
    For i = 0 To vnb - 1
        With SIni
            .Path = chemin
            .Section = "Find" & Supp
            .Key = "Find" & Trim(Str(i + 1))
            .Value = List1.List(i)
        End With
        With SIni
            .Path = chemin
            .Section = "Replace" & Supp
            .Key = "Replace" & Trim(Str(i + 1))
            .Value = List3.List(i)
        End With
        
        With SIni
            .Path = chemin
            .Section = "Descriptions" & Supp
            .Key = "Desc" & Trim(Str(i + 1))
            .Value = List2.List(i)
        End With
    Next i
    StatusBar1.SimpleText = "Search expression added !"
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
        List3.ListIndex = 0
    End If
    Text2.Text = ""
    Text4.Text = ""
    Me.MousePointer = vbNormal
End Sub

Private Sub Command6_Click()
    Dim i As Integer
    Dim descr As String
    Dim chemin As String
    Dim Index As Integer
    Dim vnb As Integer
    chemin = AppPath + "presets.ini"
    If modifs = False Then
        Unload Me
    End If
    
    Me.MousePointer = vbHourglass
    StatusBar1.SimpleText = "Updating presets file"
    vnb = List1.ListCount
    ' Sauvegarde du nombre de commandes
    With SIni
         .Path = chemin
         .Section = "General" & Supp
         .Key = "NumberOfPresets" & Supp
         .Value = Trim(Str(vnb))
    End With
    ' Suppression des commandes actelles
    With SIni
         .Path = chemin
         .Section = "Find" & Supp
         .DeleteSection
    End With
    With SIni
         .Path = chemin
         .Section = "Replace" & Supp
         .DeleteSection
    End With
     ' Suppression des descriptions associées
    With SIni
         .Path = chemin
         .Section = "Descriptions" & Supp
         .DeleteSection
    End With
    For i = 0 To vnb - 1
        With SIni
            .Path = chemin
            .Section = "Find" & Supp
            .Key = "Find" & Trim(Str(i + 1))
            .Value = List1.List(i)
        End With
        With SIni
            .Path = chemin
            .Section = "Replace" & Supp
            .Key = "Replace" & Trim(Str(i + 1))
            .Value = List3.List(i)
        End With
        With SIni
            .Path = chemin
            .Section = "Descriptions" & Supp
            .Key = "Desc" & Trim(Str(i + 1))
            .Value = List2.List(i)
        End With
    Next i
    Me.MousePointer = vbNormal
    Unload Me
End Sub
Private Sub Command8_Click()
    If RechEnCours = 1 Then
        DefaultSearchPr = ""
        DefaultReplacePr = ""
    Else
        DefaultSearchExt = ""
        DefaultReplaceExt = ""
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim sValue As String
    Dim vnbcmd As Integer
    Dim chemin As String
    Synchro = False
    If RechEnCours = 1 Then ' Ouverture pour le préfixe
        Supp = "Prefix"
        Me.Caption = Me.Caption & " for prefix"
    Else
        Supp = "Extension"
        Me.Caption = Me.Caption & " for extension"
    End If
    EnCours = False
    modifs = False
    chemin = AppPath + "presets.ini"
    sValue = ""
    With SIni
     .Path = chemin
     .Section = "General"
     .Key = "NumberOfPresets" & Supp
     sValue = .Value
    End With
    vnbcmd = Val(sValue)
    For i = 1 To vnbcmd
        With SIni
            .Path = chemin
            .Section = "Find" & Supp
            .Key = "Find" & Trim(Str(i))
            sValue = .Value
        End With
        List1.AddItem sValue
        With SIni
            .Path = chemin
            .Section = "Replace" & Supp
            .Key = "Replace" & Trim(Str(i))
            sValue = .Value
        End With
        List3.AddItem sValue
        With SIni
            .Path = chemin
            .Section = "Descriptions" & Supp
            .Key = "Desc" & Trim(Str(i))
            sValue = .Value
        End With
        List2.AddItem sValue
    Next
    Text2.Text = advsearch.Text4.Text
    Text4.Text = advsearch.Text3.Text
End Sub
Private Sub List1_Click()
If Not Synchro Then
    If EnCours = False Then
        Text1.Text = List2.List(List1.ListIndex)
        Text3.Text = List1.List(List1.ListIndex)
        Text5.Text = List3.List(List1.ListIndex)
        List2.ListIndex = List1.ListIndex
        Synchro = True
        List3.Selected(List3.ListIndex) = False
        List3.ListIndex = List1.ListIndex
        List3.Selected(List1.ListIndex) = True
        Synchro = False
        SetListButtons List1, cmdUp, cmdDown
    Else
        EnCours = False
    End If
End If
End Sub
Private Sub List1_DblClick()
    Text3.SetFocus
End Sub
Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then    ' Touche Suppr
        Command2_Click
    End If
End Sub

Private Sub List3_Click()
If Not Synchro Then
    If EnCours = False Then
        Text1.Text = List2.List(List1.ListIndex)
        Text3.Text = List1.List(List3.ListIndex)
        Text5.Text = List3.List(List3.ListIndex)
        List2.ListIndex = List3.ListIndex
        Synchro = True
        List1.Selected(List1.ListIndex) = False
        List1.ListIndex = List3.ListIndex
        List1.Selected(List3.ListIndex) = True
        SetListButtons List3, cmdUp, cmdDown
        Synchro = False
    Else
        EnCours = False
    End If
End If
End Sub

Private Sub List3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then    ' Touche Suppr
        Command2_Click
    End If
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
Private Sub Text4_GotFocus()
    SelAll Text4
End Sub

Private Sub Text5_GotFocus()
    SelAll Text5
End Sub
