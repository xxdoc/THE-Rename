VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fHistory 
   AutoRedraw      =   -1  'True
   Caption         =   "History"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   HelpContextID   =   59
   Icon            =   "fHistory.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   3090
      HelpContextID   =   59
      Left            =   45
      TabIndex        =   0
      ToolTipText     =   "Right click for a contextual menu"
      Top             =   60
      Visible         =   0   'False
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   5450
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Time"
         Object.Width           =   1217
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Original Name"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "New Name"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Menu mcontextuel 
      Caption         =   "Registeredl"
      Visible         =   0   'False
      Begin VB.Menu mclear 
         Caption         =   "Clear history"
      End
      Begin VB.Menu mclose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "fHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
 Dim i As Long, vnb As Long
 Dim chaine As String
 Dim itmX As ListItem
 fHistory.MousePointer = 11
 ListView1.ListItems.Clear
 vnb = RENAME.lhistory.ListCount - 1
 For i = 0 To vnb
  chaine = RENAME.lhistory.List(i)
  Set itmX = ListView1.ListItems.Add(, , GetToken(chaine, "|", 1))
  itmX.Text = GetToken(chaine, "|", 1)
  itmX.SubItems(1) = GetToken(chaine, "|", 2)
  itmX.SubItems(2) = GetToken(chaine, "|", 3)
  itmX.SubItems(3) = GetToken(chaine, "|", 4)
 Next
 ListView1.Visible = True
 SendKeys "^{+}"
 fHistory.MousePointer = 0
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 ListView1.Left = 0
 ListView1.Top = 0
 ListView1.width = Me.ScaleWidth
 ListView1.height = Me.ScaleHeight
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 2 Then
  PopupMenu mcontextuel
 End If
End Sub
Private Sub mclear_Click()
 RENAME.lhistory.Clear
 ListView1.ListItems.Clear
 Beep
End Sub

Private Sub mclose_Click()
  Unload Me
End Sub
