VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form preview 
   AutoRedraw      =   -1  'True
   Caption         =   "Preview"
   ClientHeight    =   4995
   ClientLeft      =   2550
   ClientTop       =   1815
   ClientWidth     =   9660
   HelpContextID   =   53
   Icon            =   "preview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1000
      Left            =   585
      ScaleHeight     =   1005
      ScaleWidth      =   8490
      TabIndex        =   11
      Top             =   3720
      Width           =   8490
      Begin VB.CommandButton Command8 
         Caption         =   "&Rename"
         Height          =   300
         Left            =   2340
         TabIndex        =   8
         Top             =   680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Find long file names"
         Height          =   285
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Find files names greater than 64 characters (path included)"
         Top             =   300
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Validate filenames"
         Height          =   285
         HelpContextID   =   53
         Left            =   1740
         TabIndex        =   2
         ToolTipText     =   "Search for invalid characters in filenames"
         Top             =   300
         Visible         =   0   'False
         WhatsThisHelpID =   233
         Width           =   1560
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Find conflicts"
         Height          =   285
         HelpContextID   =   53
         Left            =   7005
         TabIndex        =   6
         ToolTipText     =   "Find conflicts with existing filenames"
         Top             =   300
         Visible         =   0   'False
         WhatsThisHelpID =   237
         Width           =   1410
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Save as text ..."
         Height          =   285
         HelpContextID   =   53
         Left            =   5535
         TabIndex        =   5
         ToolTipText     =   "Save the list to a text file"
         Top             =   300
         Visible         =   0   'False
         WhatsThisHelpID =   236
         Width           =   1410
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Copy"
         Height          =   285
         HelpContextID   =   53
         Left            =   4440
         TabIndex        =   4
         ToolTipText     =   "Copy the list to the clipboard"
         Top             =   300
         Visible         =   0   'False
         WhatsThisHelpID =   235
         Width           =   1050
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Print"
         Height          =   285
         HelpContextID   =   53
         Left            =   3345
         TabIndex        =   3
         ToolTipText     =   "Print the list"
         Top             =   300
         Visible         =   0   'False
         WhatsThisHelpID =   234
         Width           =   1050
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Ok"
         Default         =   -1  'True
         Height          =   300
         HelpContextID   =   53
         Left            =   4185
         TabIndex        =   7
         ToolTipText     =   "Close this window"
         Top             =   680
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Height          =   240
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   9090
      End
   End
   Begin MSComctlLib.ListView listPreview 
      Height          =   3495
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6165
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Original Name"
         Object.Width           =   7585
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "New Name"
         Object.Width           =   7585
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Problem"
         Object.Width           =   1482
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   4740
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.ListBox listsav 
      Enabled         =   0   'False
      Height          =   255
      Left            =   9840
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   420
   End
End
Attribute VB_Name = "preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oAutoPos As New clsAutoPositioner
Dim ItemX As ListItem
Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Sub Command1_Click()
 Unload Me
End Sub

Private Sub Command2_Click()
Dim i As Long
Dim vnb As Long
On Error GoTo ErrorHandler
preview.MousePointer = 11

Printer.Print
Printer.Print "Preview of rename files in " + Dir1Path + " on " + Format$(Date, "Long Date") + " at " + Format$(Time, "Long Time")
Printer.Print " "
Printer.Print " "
vnb = listPreview.ListItems.Count
For i = 1 To vnb
    Printer.Print listPreview.ListItems.Item(i) & " => " & listPreview.ListItems.Item(i).SubItems(1)
Next
Printer.Print " "
Printer.Print " "
Printer.Print "Total of " + Trim$(Str$(listPreview.ListItems.Count)) + " file(s)"
Printer.EndDoc

preview.MousePointer = 0
Exit Sub

ErrorHandler:
 preview.MousePointer = 0
 MsgBox "There was a problem printing to your printer."
 Exit Sub
End Sub

Private Sub Command3_Click()
 Dim chaine As String
 Dim i As Long
 Dim vnb As Long
 On Error Resume Next
 preview.MousePointer = 11
 chaine = ""
 vnb = listPreview.ListItems.Count
 For i = 1 To vnb
  chaine = chaine + listPreview.ListItems.Item(i) & " => " & listPreview.ListItems.Item(i).SubItems(1) & vbCrLf
 Next
 Clipboard.Clear
 Clipboard.SetText chaine
 chaine = ""
 Beep
 preview.MousePointer = 0
End Sub

Private Sub Command4_Click()
Dim szFilename As String, i As Long, vnb As Long
Dim ff As Integer
szFilename = DialogFile(Me.hWnd, 2, "Save as", "list.txt", "Text" & Chr$(0) & "*.txt" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "txt")

If Trim$(szFilename) = "" Then Exit Sub
preview.MousePointer = 11
ff = FreeFile
Open szFilename For Output As #ff
vnb = listPreview.ListItems.Count
For i = 1 To vnb
  Print #ff, listPreview.ListItems.Item(i) & " => " & listPreview.ListItems.Item(i).SubItems(1)
Next
Close #ff
preview.MousePointer = 0
Beep
End Sub

Private Sub Command5_Click()
 Dim i As Integer
 Dim vnb As Long
 Dim chextr As String
 Dim vnb2 As Long
 preview.MousePointer = 11
 vnb = 0
 listPreview.Visible = False
 For i = 0 To listsav.ListCount - 1
  objFind.flags = LVFI_STRING
  objFind.pSz = listsav.List(i)
  objFind.lParam = 0
  vnb2 = SendMessageAny(RENAME.ListView1.hWnd, LVM_FINDITEM, -1, objFind)
  Set ItemX = listPreview.ListItems.Item(i + 1)
  If vnb2 <> -1 Then
   ItemX.SubItems(2) = "Conflict"
   vnb = vnb + 1
  End If
  vnb2 = SendMessageStr(listsav.hWnd, LB_FINDSTRINGEXACT, 0&, listsav.List(i))
  If vnb2 < i Then
   'If listPreview.Selected(i) = False Then
   If ItemX.SubItems(2) = "" Then
    ItemX.SubItems(2) = "Conflict"
    'listPreview.Selected(i) = True
    vnb = vnb + 1
   End If
  End If
  
  Rem élimination des conflits des noms de fichiers qui vont avoir pour nouveau nom le même nom (l'ancien)
  'chextr = trim$(left$(listPreview.List(i), Instr$(listPreview.List(i), " =>")))
  chextr = Trim$(listPreview.ListItems(i + 1).Text)
  If Trim$(listsav.List(i)) = chextr Then
   'If listPreview.Selected(i) = True Then
   If ItemX.SubItems(2) = "Conflict" Then
    'listPreview.Selected(i) = False
    ItemX.SubItems(2) = ""
    vnb = vnb - 1
   End If
  End If
  
 Next
 
 If listsav.ListCount > 2 Then
 If Trim$(listsav.List(1)) = Trim$(listsav.List(0)) Then
  'If listPreview.Selected(1) = False Then
  If ItemX.SubItems(2) = "" Then
    'listPreview.Selected(1) = True
    ItemX.SubItems(2) = "Conflict"
    vnb = vnb + 1
   End If
  End If
 End If
 
 preview.MousePointer = 0
 listPreview.Visible = True
 If vnb > 0 Then
  Label1.Caption = "Found " + Trim$(Str$(vnb)) + " conflict(s)"
 Else
  Label1.Caption = "No conflicts found"
 End If
End Sub

Private Sub Command6_Click()
' Validate filenames
 Dim i As Integer
 Dim vnb As Long
 preview.MousePointer = 11
 vnb = 0
 listPreview.Visible = False
 For i = 0 To listsav.ListCount - 1
  Set ItemX = listPreview.ListItems.Item(i + 1)
  If ChInterdits(listsav.List(i)) = True Then
    vnb = vnb + 1
    ItemX.SubItems(2) = "File contains invalid character(s)"
  Else
    If Left$(listsav.List(i), 1) = " " Then
        vnb = vnb + 1
        ItemX.SubItems(2) = "First character is a space"
    Else
        ItemX.SubItems(2) = ""
    End If
  End If
  If Len(Trim$(Prefixe(listsav.List(i)))) = 0 Then
    vnb = vnb + 1
    ItemX.SubItems(2) = "Invalid Name - No prefix !"
  End If
 Next
 listPreview.Visible = True
 If vnb > 0 Then
  Label1.Caption = "Founded " + Trim$(Str$(vnb)) + " files with incorrect names. See the note at the side of each file to see why"
 Else
  Label1.Caption = "All files are ok"
 End If
 preview.MousePointer = 0
End Sub

Private Sub Command7_Click()
' Find long file names
 Dim i As Integer
 Dim vnb As Long
 Dim chemin As String
 Dim errmsg As String
 Dim vok As Boolean
 Dim longueur As Integer
 If recursive = False Then ' on n'est pas en recursif, il faut rajouter le chemin
  chemin = AddBackSlash(Trim$(Dir1Path))
 Else ' pas besoin de chemin puisqu'on est en recursif, auquel cas le chemin est déjà là.
   chemin = ""
 End If
 preview.MousePointer = 11
 vnb = 0
 listPreview.Visible = False
 For i = 0 To listsav.ListCount - 1
  vok = False
  Select Case LesOptions.CheckLongFileNameOption
    Case 0  ' On controle tout
          If Len(chemin + listsav.List(i)) > LesOptions.CheckLongFileNameSize Then
            longueur = Len(chemin + listsav.List(i))
            vok = True
            vnb = vnb + 1
          End If
    Case 1  ' On controle le chemin uniquement
        If recursive = True Then
            If Len(ExtractPath(listsav.List(i))) > LesOptions.CheckLongFileNameSize Then
                vok = True
                vnb = vnb + 1
                longueur = Len(ExtractPath(listsav.List(i)))
            End If
        Else
            If Len(chemin) > LesOptions.CheckLongFileNameSize Then
                vok = True
                vnb = vnb + 1
                longueur = Len(chemin)
            End If
        End If
    Case 2  ' On control le nom uniquement
        If Len(Prefixe(listsav.List(i)) & "." & Suffixe(listsav.List(i))) > LesOptions.CheckLongFileNameSize Then
                vok = True
                vnb = vnb + 1
                longueur = Len(Prefixe(listsav.List(i)) & "." & Suffixe(listsav.List(i)))
        End If
  End Select
  Set ItemX = listPreview.ListItems.Item(i + 1)
  If vok Then
   ItemX.SubItems(2) = "Too long, " & Trim$(Str$(longueur)) & " characters"
  Else
   ItemX.SubItems(2) = ""
  End If
  'listPreview.Selected(i) = vok
 Next
 listPreview.Visible = True
  Select Case LesOptions.CheckLongFileNameOption
    Case 0  ' On controle tout
        errmsg = "path included"
    Case 1  ' Path only
        errmsg = "path only"
    Case 2  ' File only
        errmsg = "filename only"
   End Select
 If vnb > 0 Then
  Label1.Caption = "I've found " + Trim$(Str$(vnb)) + " file names whose size is greater than " & LesOptions.CheckLongFileNameSize & " characters (" + errmsg + ")"
 Else
  Label1.Caption = "All files are ok"
 End If
 preview.MousePointer = 0
End Sub

Private Sub Command8_Click()
 Unload Me
 RENAME.StartRename
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 27 Then
  annuler = True
  Unload Me
 End If
End Sub

Private Sub Form_Load()
 m_oAutoPos.AddAssignment Me.Picture1, Me, tCONTAINER_RELATIVE_POS_BOTTOM
 m_oAutoPos.AddAssignment Me.listPreview, Me, tCONTAINER_WIDTH_DELTA_RIGHT
 m_oAutoPos.AddAssignment Me.listPreview, Me, tCONTAINER_HEIGHT_DELTA_BOTTOM
 listPreview.GridLines = LesOptions.PreviewGridLines
 If LesRegles.NumberOfActiveRules > 0 Then
    Label1.Caption = "Warning, THE Rename use rules..."
 End If
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 listPreview.Left = 0
 listPreview.Top = 0
 listPreview.height = Me.ScaleHeight - Picture1.height
 listPreview.width = Me.ScaleWidth
 Picture1.Left = (Me.ScaleWidth / 2) - (Picture1.width / 2)
 m_oAutoPos.RefreshPositions
End Sub

Private Sub listPreview_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub listPreview_Click()
    StatusBar1.SimpleText = "File number " + Trim$(Str$(listPreview.SelectedItem.Index)) + "/" + Trim$(Str$(listPreview.ListItems.Count))
End Sub

Private Sub listPreview_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 38, 40, 36, 35, 33, 34
            StatusBar1.SimpleText = "File number " + Trim$(Str$(listPreview.SelectedItem.Index)) + "/" + Trim$(Str$(listPreview.ListItems.Count))
    End Select
End Sub
