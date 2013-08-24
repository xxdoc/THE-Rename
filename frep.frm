VERSION 5.00
Begin VB.Form frep 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Destination Folders"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7635
   ControlBox      =   0   'False
   HelpContextID   =   67
   Icon            =   "frep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Dat&e && Time"
      Height          =   300
      HelpContextID   =   67
      Left            =   4602
      TabIndex        =   6
      ToolTipText     =   "Click to change files attributes while THE Rename copy them"
      Top             =   3105
      Width           =   1100
   End
   Begin VB.CommandButton Command7 
      Caption         =   "A&ttributes"
      Height          =   300
      HelpContextID   =   67
      Left            =   3710
      TabIndex        =   5
      ToolTipText     =   "Click to change files attributes while THE Rename copy them"
      Top             =   3105
      Width           =   855
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   67
      Left            =   7270
      Picture         =   "frep.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Move Up"
      Top             =   990
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   67
      Left            =   7270
      Picture         =   "frep.frx":010E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Move Down"
      Top             =   1470
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Load list"
      Height          =   300
      HelpContextID   =   67
      Left            =   142
      TabIndex        =   1
      ToolTipText     =   "Load destination folders from a list"
      Top             =   3105
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save list..."
      Enabled         =   0   'False
      Height          =   300
      HelpContextID   =   67
      Left            =   1034
      TabIndex        =   2
      ToolTipText     =   "Save list to a text file"
      Top             =   3105
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      Height          =   300
      HelpContextID   =   67
      Left            =   2818
      TabIndex        =   4
      ToolTipText     =   "Delete selected folders from the list"
      Top             =   3105
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   67
      Left            =   6720
      TabIndex        =   8
      ToolTipText     =   "Validate modifications"
      Top             =   3105
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   67
      Left            =   5790
      TabIndex        =   7
      ToolTipText     =   "Cancel all modifications"
      Top             =   3105
      Width           =   855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Add ..."
      Height          =   300
      HelpContextID   =   67
      Left            =   1926
      TabIndex        =   3
      ToolTipText     =   "Select a directory to add to the list"
      Top             =   3105
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   2595
      HelpContextID   =   67
      Left            =   98
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   120
      Width           =   7110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remember, you can also drop files from Windows Explorer to the multiple copy button"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   2790
      Width           =   6000
   End
End
Attribute VB_Name = "frep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdDown_Click()
    ButtonDown List1
End Sub

Private Sub cmdUp_Click()
    ButtonUp List1
End Sub

Private Sub Command1_Click()
Dim szFilename As String, i As Integer, ff As Integer
szFilename = DialogFile(Me.hwnd, 2, "Save as", "copylist.txt", "Text" & Chr$(0) & "*.txt" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "txt")
If Trim$(szFilename) = "" Then Exit Sub
frep.MousePointer = 11
On Error GoTo erreur
ff = FreeFile
Open szFilename For Output As #ff
For i = 0 To List1.ListCount - 1
  Print #ff, List1.List(i)
 Next
Close #ff
frep.MousePointer = 0
Exit Sub

erreur:
 MsgBox "There was an error while saving the list"
 frep.MousePointer = 0
End Sub

Private Sub Command2_Click()
 Dim i As Integer
 If List1.ListCount > 0 Then
    For i = List1.ListCount - 1 To 0 Step -1
        If List1.Selected(i) = True Then
            List1.RemoveItem (i)
        End If
    Next
 End If
 If List1.ListCount > 0 Then
    Command1.Enabled = True
 Else
    Command1.Enabled = False
 End If
End Sub

Private Sub Command3_Click()
 Dim i As Integer
 Dim vmax As Integer
 frep.MousePointer = 11
 VnbRep = List1.ListCount
 For i = 1 To 100
  LesRepertoires(i) = ""
 Next
 vmax = List1.ListCount - 1
 If vmax > 100 Then
    vmax = 100
 End If
 For i = 0 To vmax
  LesRepertoires(i + 1) = List1.List(i)
 Next
 frep.MousePointer = 0
 Unload Me
End Sub

Private Sub Command4_Click()
 Unload Me
End Sub

Private Sub Command5_Click()
Dim szFilename As String, ligne As String, vnb As Integer
Dim vretour As Integer, ff As Integer
szFilename = DialogFile(Me.hwnd, 1, "Open copy list", "copylist.txt", "Text" & Chr$(0) & "*.txt" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "txt")
If Trim$(szFilename) = "" Then
 Exit Sub
End If
If List1.ListCount > 0 Then
 vretour = MsgBox("Current list contains items, do you want to delete them ?", vbOKCancel, "Open an existing list")
 If vretour = vbOK Then
  List1.Clear
 End If
End If
vnb = List1.ListCount
frep.MousePointer = 11
On Error GoTo erreur
ff = FreeFile
Open szFilename For Input As #ff
While Not EOF(ff)
 Line Input #ff, ligne
 vnb = vnb + 1
 If vnb > 100 Then
  GoTo suite
 End If
 ligne = Trim$(ligne)
 If ligne <> "" Then
  List1.AddItem ligne
 Else
  If vnb > 0 Then
   vnb = vnb - 1
  End If
 End If
Wend
suite:
Close #ff
If List1.ListCount > 0 Then
 Command1.Enabled = True
End If
frep.MousePointer = 0
Exit Sub

erreur:
 MsgBox "There was an error while saving the list"
 frep.MousePointer = 0

End Sub

Private Sub Command6_Click()
 Dim szFilename As String
 szFilename = ""
 If List1.ListCount <= 100 Then
  szFilename = BrowseFolder(Me, "Browse for folder:")
  If Len(Trim$(szFilename)) = 0 Then
   Exit Sub
  End If
  List1.AddItem szFilename
 Else
  MsgBox "Error, you are limited to 100 folders"
 End If
 If List1.ListCount > 0 Then
  Command1.Enabled = True
 Else
  Command1.Enabled = False
 End If
End Sub

Private Sub Command7_Click()
    AttrEncours = 3
    attributs.Show 1
End Sub

Private Sub Command8_Click()
    DTEnCours = 3
    fDT.Show 1
End Sub

Private Sub Form_Load()
 Dim i As Integer
 frep.MousePointer = 11
 List1.Clear
 For i = 1 To VnbRep
  List1.AddItem LesRepertoires(i)
 Next
 If List1.ListCount > 0 Then
  Command1.Enabled = True
 End If
 frep.MousePointer = 0
End Sub
Private Sub List1_Click()
 SetListButtons List1, cmdUp, cmdDown
End Sub
