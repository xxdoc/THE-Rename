VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form RenList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rename from a list"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PanelList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   1080
      ScaleHeight     =   3495
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton Command4 
         Caption         =   "Sa&ve list as ..."
         Height          =   300
         Left            =   3360
         TabIndex        =   6
         ToolTipText     =   "Start renaming files"
         Top             =   2760
         Width           =   1155
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Open an existing list"
         Height          =   300
         Left            =   1680
         TabIndex        =   5
         ToolTipText     =   "Start renaming files"
         Top             =   2760
         Width           =   1605
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Remove"
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Copy selected to list"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   3120
         Width           =   1770
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Copy all"
         Height          =   300
         Left            =   1965
         TabIndex        =   2
         Top             =   3120
         Width           =   1050
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   2655
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Press F2 to edit filename"
         Top             =   0
         Width           =   6200
         _ExtentX        =   10927
         _ExtentY        =   4683
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "New Name"
            Object.Width           =   5080
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Original name"
            Object.Width           =   5080
         EndProperty
      End
   End
End
Attribute VB_Name = "RenList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vfaufaire As Boolean
Dim vancien As Integer

Private Sub Command1_Click()
 MsgBox "Not yet done"
End Sub
Private Sub Command15_Click()
Dim itmX
Dim chemin As String
chemin = Trim(Dir1Path)
If Right(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If
Dim i As Integer
 For i = 0 To File1.ListCount - 1
  If File1.Selected(i) = True Then
   Set itmX = ListView1.ListItems.Add(, , File1.List(i))
   itmX.Text = File1.List(i)
   itmX.SubItems(1) = File1.List(i)
  End If
 Next i
End Sub

Private Sub Command16_Click()
Dim chemin As String
Dim itmX
chemin = Trim(Dir1Path)
If Right(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If
Dim i As Integer
 For i = 0 To File1.ListCount - 1
   Set itmX = ListView1.ListItems.Add(, , File1.List(i))
   itmX.Text = File1.List(i)
   itmX.SubItems(1) = File1.List(i)
 Next i
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
 Unload Me
End Sub
Private Sub Command8_Click()
Dim i As Integer
Dim r As Long
For i = ListView1.ListItems.Count To 0 Step -1
  If LVIsSelected(ListView1, i) = True Then
   r = LVRemoveItem(ListView1, i)
 End If
Next
End Sub

Private Sub Form_Load()
 Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
 Call SendMessageLong(ListView1.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, True)
' Call SendMessageLong(ListView1.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, True)
 vancien = 0
 vfaufaire = False
End Sub


Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 113 Then
  Call SendMessageLong(ListView1.hWnd, LVM_EDITLABEL, ListView1.SelectedItem.Index - 1, 0)
 End If
End Sub


