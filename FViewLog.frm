VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form FViewLog 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "View Log File"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8295
   HelpContextID   =   479
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.StatusBar Etat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   6105
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   14
      Left            =   7065
      TabIndex        =   3
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Clear log"
      Height          =   300
      HelpContextID   =   14
      Left            =   5880
      TabIndex        =   2
      ToolTipText     =   "Clear log file's content"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Undo Rename"
      Height          =   300
      Left            =   3240
      TabIndex        =   1
      ToolTipText     =   "Undo the rename operation for the selected files"
      Top             =   5760
      Width           =   1815
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9763
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "On"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "At"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "You renamed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "to"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "On"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "At"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "You renamed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "to"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "FViewLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oAutoPos As New clsAutoPositioner

Private Sub cmdCancel_Click()
    If MsgBox("You are going to delete the log file, are you sure ?", vbYesNo, "Warning") = vbNo Then
        Exit Sub
    Else
        If FileExists(LesOptions.LogFile) Then
            Kill (LesOptions.LogFile)
        End If
        MsgBox "The log file has been deleted"
        ListView1.ListItems.Clear
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
 Dim fileop As New CSHFileOp, i As Long, vnb As Long, vnom1 As String, vnom2 As String, itmX As ListItem
 Dim itmX2 As ListItem, vnbundo As Long
 Dim ff As Integer
 With fileop
    .ParentWnd = hWnd
    .ClearSourceFiles
    .ClearDestFiles
    .AllowUndo = False
    .ConfirmOperation = False
End With
 Me.MousePointer = 11
 vnb = ListView1.ListItems.Count
 List1.Clear
 ListView2.ListItems.Clear
 For i = 1 To vnb
    If ListView1.ListItems(i).Selected = True Then
        Set itmX = ListView1.ListItems(i)
        Set itmX2 = ListView2.ListItems.Add(, , Date$)
        vnom1 = itmX.SubItems(2)     ' Nom d'origine
        vnom2 = itmX.SubItems(3) ' Nouveau nom
        itmX2.SubItems(1) = Time()
        itmX2.SubItems(2) = vnom2
        itmX2.SubItems(3) = vnom1
        Etat.SimpleText = "(UNDO) Rename " + vnom2 + " => " + vnom1
        fileop.AddSourceFile vnom2
        fileop.AddDestFile vnom1
        fileop.RenameFiles
        fileop.ClearSourceFiles
        fileop.ClearDestFiles
        List1.AddItem i
        vnbundo = vnbundo + 1
    End If
Next

Me.MousePointer = 0
If vnbundo > 0 Then
    vnb = List1.ListCount - 1
    For i = vnb To 0 Step -1
        ListView1.ListItems.Remove (Val(List1.List(i)))
    Next
    If MsgBox("Would you like to add log information for operations you have just made ?", vbYesNo, "Add log information ?") = vbYes Then
        Me.MousePointer = 11
        For i = 1 To ListView2.ListItems.Count
            Set itmX = ListView2.ListItems(i)
            Set itmX2 = ListView1.ListItems.Add(, , ListView2.ListItems(i))
            itmX2.SubItems(1) = itmX.SubItems(1)
            itmX2.SubItems(2) = itmX.SubItems(2)
            itmX2.SubItems(3) = itmX.SubItems(3)
        Next
        Me.MousePointer = 0
    End If
End If


If ListView1.ListItems.Count = 0 Then   ' On peut supprimer le fichier .log
    If FileExists(LesOptions.LogFile) Then
        Kill (LesOptions.LogFile)
    End If
    cmdCancel.Enabled = False
    Command1.Enabled = False
Else    ' Il faut reconstruire le fichier .log
    Me.MousePointer = 11
    Etat.SimpleText = "Saving log file...."
    ff = FreeFile
    Open LesOptions.LogFile For Output As #ff
    vnb = ListView1.ListItems.Count
    For i = 1 To vnb
        Set itmX = ListView1.ListItems(i)
        Print #ff, ListView1.ListItems(i).Text & vbTab & itmX.SubItems(1) & vbTab & itmX.SubItems(2) & vbTab & itmX.SubItems(3)
    Next
    Close #ff
    Me.MousePointer = 0
End If

If vnbundo > 0 Then ' On rafraichit la liste des fichiers de la fenêtre principale
    Me.MousePointer = 11
    RENAME.RefreshF5
    Me.MousePointer = 0
End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrGen
    Dim vligne As String, itmX As ListItem, j As Integer
    Dim colonne As ColumnHeader
    Dim ff As Integer
    m_oAutoPos.AddAssignment Me.Command1, Me, tCONTAINER_RELATIVE_POS_BOTTOM
    m_oAutoPos.AddAssignment Me.cmdCancel, Me, tCONTAINER_RELATIVE_POS_BOTTOM
    m_oAutoPos.AddAssignment Me.cmdOK, Me, tCONTAINER_RELATIVE_POS_BOTTOM
    m_oAutoPos.AddAssignment Me.ListView1, Me, tCONTAINER_WIDTH_DELTA_RIGHT
    m_oAutoPos.AddAssignment Me.ListView1, Me, tCONTAINER_HEIGHT_DELTA_BOTTOM
    
    If Trim$(LesOptions.LogFile) = "" Then
        MsgBox "Sorry but you don't use a log file..."
        Unload Me
    End If
    If FileExists(LesOptions.LogFile) Then
        ff = FreeFile
        Open LesOptions.LogFile For Input As #ff
        While Not EOF(ff)
            Line Input #ff, vligne
            Set itmX = ListView1.ListItems.Add(, , GetToken(vligne, vbTab, 1))  ' Date
            itmX.SubItems(1) = GetToken(vligne, vbTab, 2)   ' Heure
            itmX.SubItems(2) = GetToken(vligne, vbTab, 3)   ' Ancien nom
            itmX.SubItems(3) = GetToken(vligne, vbTab, 4)   ' Nouveau nom
        Wend
        Close #ff
        If ListView1.ListItems.Count > 0 Then
            Set colonne = ListView1.ColumnHeaders.Item(1)
            For j = 1 To 4
                Set colonne = ListView1.ColumnHeaders.Item(j)
                AutoSizeColumnHeader ListView1, colonne, True
            Next
        Else
            cmdCancel.Enabled = False
            Command1.Enabled = False
        End If
    End If
    
    If ListView1.ListItems.Count <= 0 Then
        cmdCancel.Enabled = False
        Command1.Enabled = False
    End If
    Exit Sub
    
ErrGen:
    If Err.Number = 53 Then
        Resume Next
    Else
        ErreurGrave "FViewLog:Form_Load"
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    m_oAutoPos.RefreshPositions
End Sub
