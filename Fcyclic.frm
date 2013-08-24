VERSION 5.00
Begin VB.Form Fcyclic 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options for cyclic selection"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3225
   ControlBox      =   0   'False
   HelpContextID   =   168
   Icon            =   "Fcyclic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Text7 
      Height          =   315
      Left            =   75
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Height          =   330
      HelpContextID   =   168
      Left            =   2835
      Picture         =   "Fcyclic.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Save to a File"
      Top             =   3240
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   281
      Width           =   330
   End
   Begin VB.CommandButton Command3 
      Height          =   330
      HelpContextID   =   168
      Left            =   2835
      Picture         =   "Fcyclic.frx":01D6
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Load from a file"
      Top             =   2880
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   282
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      Height          =   330
      HelpContextID   =   168
      Left            =   2835
      Picture         =   "Fcyclic.frx":03A0
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Copy items to clipboard"
      Top             =   2400
      Width           =   330
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      HelpContextID   =   168
      Left            =   720
      ScaleHeight     =   735
      ScaleWidth      =   1575
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   1575
      Begin VB.OptionButton Option1 
         Caption         =   "Replace prefix"
         Height          =   255
         HelpContextID   =   168
         Index           =   2
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Texts will replace prefix"
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Add to the left"
         Height          =   195
         HelpContextID   =   168
         Index           =   1
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Texts will be added to the left"
         Top             =   270
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Add to the right"
         Height          =   255
         HelpContextID   =   168
         Index           =   0
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Texts will be added to the right"
         Top             =   0
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      HelpContextID   =   168
      Left            =   2835
      Picture         =   "Fcyclic.frx":048A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Paste from clipboard"
      Top             =   2040
      Width           =   330
   End
   Begin VB.ListBox List2 
      Height          =   2205
      HelpContextID   =   168
      ItemData        =   "Fcyclic.frx":058C
      Left            =   60
      List            =   "Fcyclic.frx":058E
      TabIndex        =   2
      Top             =   780
      Width           =   2640
   End
   Begin VB.CommandButton Command4 
      Default         =   -1  'True
      Height          =   330
      HelpContextID   =   168
      Left            =   2835
      Picture         =   "Fcyclic.frx":0590
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Add item to the list"
      Top             =   360
      Width           =   330
   End
   Begin VB.CommandButton Command5 
      Height          =   330
      HelpContextID   =   168
      Left            =   2835
      Picture         =   "Fcyclic.frx":062A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Remove item from the list"
      Top             =   765
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   168
      Left            =   2835
      Picture         =   "Fcyclic.frx":0714
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Move Down"
      Top             =   1560
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   168
      Left            =   2835
      Picture         =   "Fcyclic.frx":0816
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Move Up"
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   300
      HelpContextID   =   168
      Left            =   1650
      TabIndex        =   12
      ToolTipText     =   "Use cyclic selection wit these parameters"
      Top             =   4035
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   168
      Left            =   480
      TabIndex        =   11
      ToolTipText     =   "Don't use cyclic selection"
      Top             =   4035
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Enter text"
      Height          =   195
      Left            =   75
      TabIndex        =   15
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "Fcyclic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cHist5 As New cHistory
Private Sub cmdCancel_Click()
    UseCylcic = False
    Unload Me
End Sub
Private Sub cmdDown_Click()
    ButtonDown List2
End Sub

Private Sub cmdOK_Click()
    UseCylcic = True
    SaveCyclic List2
    Unload Me
End Sub

Private Sub cmdUp_Click()
    ButtonUp List2
End Sub

Private Sub Command1_Click()
 Dim lachaine As String
 Dim i As Integer
 Dim l As Long
 Dim LePressePapier As String
 If Clipboard.GetFormat(vbCFText) Then
  List2.Visible = False
  LePressePapier = Clipboard.GetText
  l = Len(LePressePapier)
  lachaine = ""
  For i = 1 To l
   If Mid$(LePressePapier, i, 1) <> Chr$(13) And Mid$(LePressePapier, i, 1) <> Chr$(10) Then
    lachaine = lachaine + Mid$(LePressePapier, i, 1)
   Else
    If Mid$(LePressePapier, i, 1) <> Chr$(10) Then
        List2.AddItem lachaine
        lachaine = ""
    End If
   End If
  Next
  List2.Visible = True
 End If
 If List2.ListCount > 0 Then
  List2.ListIndex = List2.ListCount - 1
 End If
End Sub

Private Sub Command2_Click()
 Dim i As Integer
 Dim chaine As String
 Dim vnb As Integer
 vnb = List2.ListCount - 1
 chaine = ""
 For i = 0 To vnb
  chaine = chaine + List2.List(i) + vbCrLf
 Next
 Clipboard.Clear
 Clipboard.SetText chaine
 chaine = ""
End Sub

Private Sub Command3_Click()
Dim szFilename As String, vretour As Integer, vligne As String
Dim ff As Integer
szFilename = DialogFile(Me.hWnd, 1, "Open Cyclic selections", "cyclic.cyc", "Cyclic File" & Chr$(0) & "*.cyc" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "Cyclic File")
If Trim$(szFilename) = "" Then Exit Sub

If List2.ListCount > 0 Then
    vretour = MsgBox("Delete current content ?", vbYesNoCancel, "Warning")
    Select Case vretour
        Case vbCancel
            Exit Sub
        Case vbYes
            List2.Clear
    End Select
End If

ff = FreeFile
Open szFilename For Input As #ff
Line Input #ff, vligne
While Not EOF(ff)
    List2.AddItem vligne
    Line Input #ff, vligne
Wend
If vligne <> "" Then
    List2.AddItem vligne
End If
Close #ff
End Sub

Private Sub Command4_Click()
 List2.AddItem Text7.Text
 cHist5.AddNewItem Text7.Text
 Text7.Text = ""
 List2.ListIndex = List2.ListCount - 1
End Sub
Private Sub Command5_Click()
 If List2.ListIndex = -1 Then
  MsgBox "You must first select an item before to delete it"
  Exit Sub
 End If
 List2.RemoveItem (List2.ListIndex)
 If List2.ListCount > 0 Then
  List2.ListIndex = List2.ListCount - 1
 End If
End Sub

Private Sub Command6_Click()
Dim szFilename As String, i As Integer
Dim ff As Integer

If List2.ListCount = 0 Then
    MsgBox "Sorry there are no cyclic texts to save...", vbOKOnly, "Warning"
    Exit Sub
End If

szFilename = DialogFile(Me.hWnd, 2, "Save cyclic selections as", "cyclic.cyc", "Cyclic File" & Chr$(0) & "*.cyc" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "Cyclic File")
If szFilename = "" Then Exit Sub

ff = FreeFile
Open szFilename For Output As #ff
For i = 0 To List2.ListCount - 1
    Print #ff, List2.List(i)
Next
Close #ff
End Sub

Private Sub Form_Load()
 cHist5.sKey = "CyclicString"
 cHist5.Items Text7
 If OptionsCyclic = True Then
  Picture1.Visible = True
 End If
 LoadCyclic List2
 Option1(PlacementCyclic).Value = True
End Sub
Private Sub List2_Click()
 SetListButtons List2, cmdUp, cmdDown
End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 46
   Command5_Click
  End Select
End Sub

Private Sub Option1_Click(Index As Integer)
 PlacementCyclic = Index
End Sub

Private Sub Text7_GotFocus()
    SelAll Text7
End Sub
