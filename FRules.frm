VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form FRules 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rules"
   ClientHeight    =   4905
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6165
   ControlBox      =   0   'False
   HelpContextID   =   472
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "&Activate all rules"
      Height          =   315
      Left            =   4500
      TabIndex        =   5
      Top             =   2040
      Width           =   1550
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Use the folowing rules to rename files"
         Object.Width           =   9411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   3855
      TabIndex        =   13
      Top             =   3780
      Width           =   3855
      Begin VB.OptionButton Option1 
         Caption         =   "Rename files when they don't satisfy conditions"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   7
         Top             =   300
         Width           =   3675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Rename files when they satisfy conditions"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Value           =   -1  'True
         Width           =   3315
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   2460
      Width           =   5535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   38
      Left            =   1957
      TabIndex        =   8
      ToolTipText     =   "Don't save rules"
      Top             =   4500
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   38
      Left            =   3112
      TabIndex        =   9
      ToolTipText     =   "Save rules"
      Top             =   4500
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Deactivate all rules"
      Height          =   315
      Left            =   2745
      TabIndex        =   4
      Top             =   2040
      Width           =   1700
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   20
      Left            =   5760
      Picture         =   "FRules.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Move Up"
      Top             =   660
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   20
      Left            =   5760
      Picture         =   "FRules.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Move Down"
      Top             =   1140
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   1890
      TabIndex        =   3
      ToolTipText     =   "Delete selected rule"
      Top             =   2040
      Width           =   800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Modify..."
      Height          =   315
      Left            =   1035
      TabIndex        =   2
      ToolTipText     =   "Modify selected rule"
      Top             =   2040
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&New..."
      Height          =   315
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "Add a new rule"
      Top             =   2040
      Width           =   800
   End
   Begin VB.Menu mfiles 
      Caption         =   "&Files"
      Begin VB.Menu mload 
         Caption         =   "&Load rules from file..."
         Shortcut        =   ^O
      End
      Begin VB.Menu msave 
         Caption         =   "&Save rules to a file..."
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "FRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RglTmp As New Rules
Dim UneRegle As New Rule
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdDown_Click()
    If LV1.ListItems.Count > 0 Then
        If LV1.SelectedItem.Index <> LV1.ListItems.Count Then
            MoveFilesDown LV1, FRules
            EtatHautBas LV1, cmdUp, cmdDown
        End If
    End If
End Sub
Private Sub cmdOK_Click()
    Dim i As Integer
    Dim vnb As Integer
    Dim OnRglTmp As New Rule
    vnb = LV1.ListItems.Count
    ' On supprime les "anciennes" règles
    Set LesRegles = Nothing
    If Option1(0).Value = True Then
        LesRegles.RulesType = 0
    Else
        LesRegles.RulesType = 1
    End If
    
    For i = 1 To vnb
        Set OnRglTmp = RglTmp.GetRule(LV1.ListItems(i).SubItems(1))
        If LV1.ListItems(i).Checked = True Then
            OnRglTmp.RuleActive = True
        Else
            OnRglTmp.RuleActive = False
        End If
        LesRegles.AddRule OnRglTmp
    Next
    LesRegles.SaveRulesToFile AppPath & "Rules.ini"
    Unload Me
End Sub
Private Sub cmdUp_Click()
    If LV1.ListItems.Count > 0 Then
        If LV1.SelectedItem.Index <> 1 Then
            MoveFilesUp LV1, FRules
            EtatHautBas LV1, cmdUp, cmdDown
        End If
    End If
End Sub
Private Sub Command1_Click()
    If FDetRule.GetRule(1, RglTmp, 0, "Add a rule") Then
        LoadRules
        LV1.ListItems(1).Selected = True
        LV1_Click
    End If
End Sub
Private Sub Command2_Click()
    Dim indice As Integer
    If LV1.ListItems.Count > 0 Then
        LV1_Click
    End If
    indice = LV1.SelectedItem.Index
    indice = LV1.ListItems(indice).SubItems(1)
    If indice > 0 Then
        If FDetRule.GetRule(2, RglTmp, indice, "Modify a rule") Then
            LoadRules
            LV1.ListItems(indice).Selected = True
            LV1_Click
        End If
    End If
End Sub
Private Sub Command3_Click()
    If LV1.SelectedItem.Index > 0 Then
        If MsgBox("Do you really want to delete the rule '" + LV1.SelectedItem.Text + "' ?", vbYesNo, "Warning") = vbYes Then
            RglTmp.RemoveRule LV1.SelectedItem.Index
            LoadRules
            If LV1.ListItems.Count > 0 Then
                LV1.ListItems(1).Selected = True
                LV1_Click
            Else
                Text1.Text = ""
            End If
        End If
    End If
End Sub
Private Sub SelUnsel(Etat As Boolean)
    Dim i As Integer
    Dim vnb As Integer
    vnb = LV1.ListItems.Count
    For i = 1 To vnb
        LV1.ListItems(i).Checked = Etat
    Next
End Sub
Private Sub Command4_Click()
    SelUnsel False
End Sub
Private Sub Command5_Click()
    SelUnsel True
End Sub
Private Sub Form_Load()
    Set RglTmp = Nothing
    Set UneRegle = Nothing
    Set RglTmp = LesRegles
    Option1(LesRegles.RulesType).Value = True
    LoadRules
    LV1_Click
End Sub
Private Sub LV1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub
Private Sub LV1_Click()
    Dim indice As Integer
    If LV1.ListItems.Count > 0 Then
        indice = LV1.SelectedItem.Index
        indice = LV1.ListItems(indice).SubItems(1)
        Set UneRegle = RglTmp.GetRule(indice)
        Text1.Text = UneRegle.RuleDescription
        EtatHautBas LV1, cmdUp, cmdDown
    End If
End Sub
Private Sub LoadRules()
    Dim i As Integer
    Dim vnb As Integer
    Dim itmX As ListItem
    Set RglTmp = Nothing
    Set UneRegle = Nothing
    Set RglTmp = LesRegles
    
    LV1.ListItems.Clear
    
    vnb = RglTmp.RulesCount
    For i = 1 To vnb
        Set UneRegle = RglTmp.GetRule(i)
        Set itmX = LV1.ListItems.Add(, , UneRegle.RuleName)
        If UneRegle.RuleActive Then
            itmX.Checked = True
        End If
        itmX.SubItems(1) = UneRegle.RuleIndex
    Next
End Sub
Private Sub LV1_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 46 ' Suppr
            Command3_Click
        Case 45 ' Ins
            Command1_Click
    End Select
End Sub

Private Sub mload_Click()
Dim szFilename As String
szFilename = DialogFile(Me.hWnd, 1, "Open Rules", "Rules.ini", "Rules" & Chr$(0) & "*.ini" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "Rules")
If Trim$(szFilename) = "" Then
    Exit Sub
End If

If Not (LesRegles.LoadRulesFromFile(szFilename)) Then
    MsgBox "There was a problem while loading rules !"
End If
LoadRules
LV1_Click
End Sub

Private Sub msave_Click()
Dim szFilename As String

If LV1.ListItems.Count = 0 Then
    Exit Sub
End If

szFilename = DialogFile(Me.hWnd, 2, "Save rules as", "Rules.ini", "Rules" & Chr$(0) & "*.ini" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "Rules")
If szFilename = "" Then
    Exit Sub
End If
If Not (LesRegles.SaveRulesToFile(szFilename)) Then
    MsgBox "There was a problem while saving rules to file !"
End If
End Sub
