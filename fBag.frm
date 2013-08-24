VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form fBag 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bin"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9405
   ControlBox      =   0   'False
   HelpContextID   =   66
   Icon            =   "fBag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "A&ttributes"
      Height          =   300
      Left            =   6447
      TabIndex        =   6
      ToolTipText     =   "Click to change files attributes while they are copy"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Dat&e && Time"
      Height          =   300
      Left            =   7376
      TabIndex        =   7
      ToolTipText     =   "Click to change files attributes while they are copy"
      Top             =   3960
      Width           =   1100
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Copy names"
      Height          =   300
      Left            =   5303
      TabIndex        =   5
      ToolTipText     =   "Copy selected names to clipboard"
      Top             =   3960
      Width           =   1070
   End
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   300
      Left            =   8552
      TabIndex        =   8
      ToolTipText     =   "Close this window"
      Top             =   3960
      Width           =   800
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Paste to and keep..."
      Height          =   300
      Left            =   1101
      TabIndex        =   2
      ToolTipText     =   "Past selected files to a folder and keep them in bin"
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear from bin"
      Height          =   300
      Left            =   4014
      TabIndex        =   4
      ToolTipText     =   "Remove selected files from bin"
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete files"
      Height          =   300
      Left            =   2870
      TabIndex        =   3
      ToolTipText     =   "Send selected files to recycle bin"
      Top             =   3960
      Width           =   1070
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Paste to..."
      Height          =   300
      Left            =   52
      TabIndex        =   1
      ToolTipText     =   "Paste selected files to a folder and remove them from bin"
      Top             =   3960
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3480
      Left            =   75
      TabIndex        =   0
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   6138
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Action"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   8114
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Created"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Attrib."
         Object.Width           =   706
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Information"
      Height          =   255
      Left            =   247
      TabIndex        =   9
      ToolTipText     =   "Information"
      Top             =   3645
      Width           =   8910
   End
End
Attribute VB_Name = "fBag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
 Dim szFilename As String, qfaire As String
 Dim i As Long, vnb As Long
 Dim fileop As New CSHFileOp
 Dim chemin3 As String
 With fileop
    .ParentWnd = hWnd
    .ConfirmOperation = LesOptions.ConfirmOperation
    .RenameOnCollision = LesOptions.RenameOnCollision
    .SilentMode = LesOptions.SilentMode
    .AllowUndo = LesOptions.AllowUndo
    .ConfirmMakeDir = LesOptions.ConfirmMakeDir
 End With
 
 If LVGetCountSelected(ListView1) = 0 Then
  Label1.Caption = "Select items please"
  Exit Sub
 End If
 
 szFilename = ""
 szFilename = BrowseFolder(Me, "Browse for folder:")
 If Len(Trim$(szFilename)) = 0 Then
  Exit Sub
 End If
 fBag.MousePointer = 11
 i = LVGetItemSelected(ListView1, -1)
 While i <> -1
  qfaire = LVGetName(ListView1, i) ' Action
  fileop.ClearSourceFiles
  fileop.ClearDestFiles
  fileop.AddSourceFile LVGetItemName(ListView1, i, 1)
  fileop.AddDestFile szFilename
  If qfaire = "Cut to bin" Or qfaire = "Cut additive" Then
   Label1.Caption = "Move " + LVGetItemName(ListView1, i, 1) + " to " + szFilename
   DoEvents
   If fileop.MoveFiles Then
      chemin3 = AddBackSlash(szFilename)
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView1, i, 1)) + "." + Suffixe(LVGetItemName(ListView1, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
   End If
  Else
   Label1.Caption = "Copy " + LVGetItemName(ListView1, i, 1) + " to " + szFilename
   DoEvents
   If fileop.CopyFiles Then
      chemin3 = AddBackSlash(szFilename)
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView1, i, 1)) + "." + Suffixe(LVGetItemName(ListView1, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
   End If
  End If
  i = LVGetItemSelected(ListView1, i)
 Wend
 
 Dim vretour As Boolean
 vnb = ListView1.ListItems.Count
 For i = vnb To 1 Step -1
  If ListView1.ListItems(i).Selected = True Then
   vretour = LVRemoveItem(ListView1, i - 1)
  End If
 Next
 Label1.Caption = ListView1.ListItems.Count & " Items"
 fBag.MousePointer = 0
End Sub

Private Sub Command3_Click()
 Dim vretour As Boolean
 Dim i As Long
 Dim fileop As New CSHFileOp
 Dim vnb As Long
 If LVGetCountSelected(ListView1) = 0 Then
  Label1.Caption = "Select items please"
  Exit Sub
 End If
 fBag.MousePointer = 11
 With fileop
    .ParentWnd = hWnd
    .ConfirmOperation = LesOptions.ConfirmOperation
    .RenameOnCollision = LesOptions.RenameOnCollision
    .SilentMode = LesOptions.SilentMode
    .AllowUndo = LesOptions.AllowUndo
    .ConfirmMakeDir = LesOptions.ConfirmMakeDir
 End With
 fileop.ClearSourceFiles
 fileop.ClearDestFiles
 vnb = ListView1.ListItems.Count
 For i = vnb To 1 Step -1
  If ListView1.ListItems(i).Selected = True Then
   Label1.Caption = ListView1.ListItems(i).SubItems(1)
   DoEvents
   fileop.AddSourceFile ListView1.ListItems(i).SubItems(1)
  End If
 Next
 fileop.DeleteFiles
 vnb = ListView1.ListItems.Count
 For i = vnb To 1 Step -1
  If ListView1.ListItems(i).Selected = True Then
   vretour = LVRemoveItem(ListView1, i - 1)
  End If
 Next
 Label1.Caption = ListView1.ListItems.Count & " Items"
 fBag.MousePointer = 0
End Sub

Private Sub Command4_Click()
 Dim vretour As Boolean
 Dim i As Long
 Dim vnb As Long
 If LVGetCountSelected(ListView1) = 0 Then
  Label1.Caption = "Select items please"
  Exit Sub
 End If
 fBag.MousePointer = 11
 Label1.Caption = "Removing..."
 vnb = ListView1.ListItems.Count
 For i = vnb To 1 Step -1
  If ListView1.ListItems(i).Selected = True Then
   vretour = LVRemoveItem(ListView1, i - 1)
  End If
 Next
 Label1.Caption = ListView1.ListItems.Count & " Items"
 fBag.MousePointer = 0
End Sub

Private Sub Command5_Click()
 Dim szFilename As String, qfaire As String
 Dim i As Long
 Dim fileop As New CSHFileOp
 With fileop
    .ParentWnd = hWnd
    .ConfirmOperation = LesOptions.ConfirmOperation
    .RenameOnCollision = LesOptions.RenameOnCollision
    .SilentMode = LesOptions.SilentMode
    .AllowUndo = LesOptions.AllowUndo
    .ConfirmMakeDir = LesOptions.ConfirmMakeDir
 End With
 
 If LVGetCountSelected(ListView1) = 0 Then
  Label1.Caption = "Select items please"
  Exit Sub
 End If
 
 szFilename = ""
 szFilename = BrowseFolder(Me, "Browse for folder:")
 If Len(Trim$(szFilename)) = 0 Then
  Exit Sub
 End If
 fBag.MousePointer = 11
 i = LVGetItemSelected(ListView1, -1)
 While i <> -1
  qfaire = LVGetName(ListView1, i) ' Action
  fileop.ClearSourceFiles
  fileop.ClearDestFiles
  fileop.AddSourceFile LVGetItemName(ListView1, i, 1)
  fileop.AddDestFile szFilename
  If qfaire = "Cut to bin" Or qfaire = "Cut additive" Then
   Label1.Caption = "Move " + LVGetItemName(ListView1, i, 1) + " to " + szFilename
   DoEvents
   fileop.MoveFiles
  Else
   Label1.Caption = "Copy " + LVGetItemName(ListView1, i, 1) + " to " + szFilename
   DoEvents
   fileop.CopyFiles
  End If
  i = LVGetItemSelected(ListView1, i)
 Wend
 
 Dim vretour As Boolean
 For i = ListView1.ListItems.Count To 1 Step -1
  If ListView1.ListItems(i).Selected = True Then
   qfaire = LVGetName(ListView1, i - 1) ' Action
   If qfaire = "Cut to bin" Or qfaire = "Cut additive" Then
    vretour = LVRemoveItem(ListView1, i - 1)
   End If
  End If
 Next
 Label1.Caption = ListView1.ListItems.Count & " Items"
 fBag.MousePointer = 0
End Sub

Private Sub Command6_Click()
 Dim i As Long
 Dim sItem As String
 Dim itmX As ListItem
 Dim vnb As Long
 RENAME.ListView3.ListItems.Clear
 fBag.MousePointer = 11
 Label1.Caption = "Updating"
 DoEvents
 vnb = ListView1.ListItems.Count - 1
 For i = 0 To vnb
  sItem = LVGetName(ListView1, i)
  Set itmX = RENAME.ListView3.ListItems.Add(, , sItem)
  itmX.Text = sItem
  itmX.SubItems(1) = LVGetItemName(ListView1, i, 1)
  itmX.SubItems(2) = LVGetItemName(ListView1, i, 2)
  itmX.SubItems(3) = LVGetItemName(ListView1, i, 3)
  itmX.SubItems(4) = LVGetItemName(ListView1, i, 4)
 Next
 ListView1.ListItems.Clear
 fBag.MousePointer = 0
 Unload Me
End Sub

Private Sub Command7_Click()
 Dim i As Long
 Dim vnb As Long
 Dim chaine As String
 If LVGetCountSelected(ListView1) = 0 Then
  Label1.Caption = "Select items please"
  Exit Sub
 End If
 
 chaine = ""
 fBag.MousePointer = 11
 vnb = ListView1.ListItems.Count
 For i = 1 To vnb
  If ListView1.ListItems(i).Selected = True Then
   Label1.Caption = "Copying " + ListView1.ListItems(i).SubItems(1) + " to clipboard"
   chaine = chaine + ListView1.ListItems(i).SubItems(1) + vbCrLf
  End If
 Next
Clipboard.Clear
Clipboard.SetText chaine
chaine = ""
Label1.Caption = ListView1.ListItems.Count & " Items"
fBag.MousePointer = 0
Beep
End Sub

Private Sub Command8_Click()
DTEnCours = 2
fDT.Show 1
End Sub

Private Sub Command9_Click()
 AttrEncours = 2
 attributs.Show 1
End Sub

Private Sub Form_Load()
 Dim i As Long
 Dim vnb As Long
 Dim sItem As String
 Dim itmX As ListItem
 fBag.MousePointer = 11
 ListView1.Visible = False
 ListView1.ListItems.Clear
  vnb = RENAME.ListView3.ListItems.Count - 1
  For i = 0 To vnb
  sItem = LVGetName(RENAME.ListView3, i)
  Set itmX = ListView1.ListItems.Add(, , sItem)
  itmX.Text = sItem
  itmX.SubItems(1) = LVGetItemName(RENAME.ListView3, i, 1)
  itmX.SubItems(2) = LVGetItemName(RENAME.ListView3, i, 2)
  itmX.SubItems(3) = LVGetItemName(RENAME.ListView3, i, 3)
  itmX.SubItems(4) = LVGetItemName(RENAME.ListView3, i, 4)
 Next
 ListView1.Visible = True
 Label1.Caption = ListView1.ListItems.Count & " Items"
 fBag.MousePointer = 0
End Sub
