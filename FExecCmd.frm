VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FExecCmd 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Execute a command in favorites"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7485
   ControlBox      =   0   'False
   HelpContextID   =   537
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Print preview"
      Height          =   300
      HelpContextID   =   20
      Left            =   2666
      TabIndex        =   14
      ToolTipText     =   "Run command"
      Top             =   4515
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "R"
      Height          =   285
      HelpContextID   =   14
      Left            =   7020
      TabIndex        =   12
      ToolTipText     =   "Select a registered program"
      Top             =   50
      Width           =   285
   End
   Begin VB.Frame Frame1 
      Caption         =   "Run command on "
      Height          =   795
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   7155
      Begin VB.OptionButton Option1 
         Caption         =   "Selected files"
         Height          =   195
         Index           =   1
         Left            =   100
         TabIndex        =   11
         Top             =   480
         Width           =   1635
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All files"
         Height          =   195
         Index           =   0
         Left            =   100
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "..."
      Height          =   285
      HelpContextID   =   14
      Left            =   6600
      TabIndex        =   1
      ToolTipText     =   "Open a program"
      Top             =   50
      Width           =   285
   End
   Begin VB.ListBox List2 
      Height          =   1320
      IntegralHeight  =   0   'False
      Left            =   90
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1320
      Width           =   7320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Preview"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   20
      Left            =   1728
      TabIndex        =   3
      ToolTipText     =   "Preview what commands will be"
      Top             =   4515
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Run"
      Height          =   300
      HelpContextID   =   20
      Left            =   4902
      TabIndex        =   4
      ToolTipText     =   "Run command"
      Top             =   4515
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   300
      HelpContextID   =   20
      Left            =   3964
      TabIndex        =   5
      ToolTipText     =   "Close this window"
      Top             =   4515
      Width           =   855
   End
   Begin VB.ComboBox Text4 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   45
      Width           =   4845
   End
   Begin MSComctlLib.ListView List1 
      Height          =   1710
      Left            =   90
      TabIndex        =   13
      Top             =   2700
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   3016
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Results"
         Object.Width           =   11853
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "List of favorites (check or unchek favorites in witch command must be run)"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   1080
      Width           =   5250
   End
   Begin VB.Label Label2 
      Caption         =   $"FExecCmd.frx":0000
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   6795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type your command"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1440
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mcheckall 
         Caption         =   "Check all"
      End
      Begin VB.Menu muncheckall 
         Caption         =   "Uncheck all"
      End
   End
End
Attribute VB_Name = "FExecCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If ExeCmd = False Then
        Executer True
    Else
        Execute2 True
    End If
End Sub

Private Sub Command2_Click()
    Dim prog As String
    If FProgs.GetProgram(prog) = True Then
        Text4.Text = prog
    End If
End Sub

Private Sub Command3_Click()
    cHist14.AddNewItem Text4.Text
    If ExeCmd = False Then
        Executer False
    Else
        Execute2 False
    End If
End Sub

Private Sub Command4_Click()
    cHist14.AddNewItem Text4.Text
    Unload Me
End Sub

Private Sub Command5_Click()
Dim i As Integer
Dim vnb As Integer
vnb = List1.ListItems.Count

If vnb <= 0 Then
    Exit Sub
End If

Printer.Print
Printer.Print "Preview of your commands on " + Format$(Date, "Long Date") + " at " + Format$(Time, "Long Time")
Printer.Print " "
Printer.Print " "
For i = 1 To vnb
    Printer.Print List1.ListItems(i).Text
Next
Printer.Print " "
Printer.Print "Total of " + Trim$(Str$(vnb)) + " command(s)"
Printer.EndDoc
End Sub

Private Sub Command7_Click()
 Text4.Text = DialogFile(Me.hwnd, 1, "Open", "*.exe", "EXE" & Chr$(0) & "*.exe" & Chr$(0) & "All files" & Chr$(0) & "*.*", App.Path, "exe")
 Text4.SetFocus
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim vnb As Integer
    If ExeCmd = False Then
        vnb = ffavoris.List1.ListCount - 1
        For i = 0 To vnb
            List2.AddItem ffavoris.List1.List(i)
            List2.Selected(i) = True
        Next
        cHist14.sKey = "RunComdOnFav"
    Else    ' C'est une commande à executer sur les fichiers
        cHist14.sKey = "RunComdOnFile"
        FExecCmd.Caption = "Execute a command on file(s)"
        Label3.Visible = False
        List2.Visible = False
        Frame1.Left = List2.Left
        Frame1.Top = List2.Top + 50
        Frame1.Visible = True
        Label2.height = 780
        Label2.Caption = "Remember, you can use %1 for the current drive, %2 for the current folder, %3 for drive and folder, %4 for the last part of the folder, %5 for the filename without path, %6 for the prefix, %7 for the extension, %8 for the filename with it's complete path, %9 for the file's path and %0 for the last part of the file's path."
        List1.Top = Frame1.Top + Frame1.height + 50
        List1.height = List1.height + 500
    End If
    cHist14.Items Text4
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     cHist14.AddNewItem Text4.Text
End Sub

Private Sub Executer(Optional preview As Boolean = False)
    Dim vcmd As String, fav As String, Lect As String, rep As String, vcmd2 As String, vtmp As String
    Dim i As Integer, vnb As Integer, vnb2 As Integer
    Dim itmX As ListItem
    On Error GoTo ErrGen
   
    vnb = List2.ListCount - 1
    vcmd = Text4.Text
    If Trim$(vcmd) = "" Then
        Exit Sub
    End If
    List1.ListItems.Clear
    
    For i = 0 To vnb
        If List2.Selected(i) = True Then
            fav = List2.List(i)
            Lect = Left$(fav, 2)
            rep = fav
            rep = Replace(rep, Lect, "")
            rep = Replace(rep, "\", "")
            vcmd2 = vcmd
            vcmd2 = Replace(vcmd2, "%1", Lect)
            vcmd2 = Replace(vcmd2, "%2", rep)
            vcmd2 = Replace(vcmd2, "%3", fav)
            vnb2 = CharOccurs(fav, "\")
            vtmp = GetToken(fav, "\", vnb2 + 1)
            vcmd2 = Replace(vcmd2, "%4", vtmp)
            Set itmX = List1.ListItems.Add(, , "In " & fav & " => " & vcmd2)
            If preview = False Then
                ChDrive Lect
                ChDir fav
                ExecCmd vcmd2, ""
                Set itmX = List1.ListItems.Add(, , "    Process returns " & RetStatus)
            End If
        End If
    Next
    Exit Sub
    
ErrGen:
    If MsgBox("There was an error, do you want to continue ?", vbYesNo, "Error") = vbYes Then
        Resume Next
    Else
        Exit Sub
    End If

End Sub

Private Sub List1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 Select Case KeyCode
  Case 106  ' * select all
    mcheckall_Click
    List2.SetFocus
  Case 109 ' - unselect
    muncheckall_Click
    List2.SetFocus
 End Select
End Sub

Private Sub List2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu menu
End Sub

Private Sub SelUnsell(Etat As Boolean)
    Dim i As Integer, vnb As Integer
    vnb = List2.ListCount - 1
    For i = 0 To vnb
        List2.Selected(i) = Etat
    Next
End Sub
Private Sub mcheckall_Click()
    SelUnsell True
End Sub

Private Sub muncheckall_Click()
    SelUnsell False
End Sub

Private Sub Execute2(Optional preview As Boolean = False)
    Dim vcmd As String
    Dim i As Long, vnb As Long
    On Error GoTo ErrGen
   
    vcmd = Text4.Text
    If RENAME.ListView1.ListItems.Count <= 0 Then
        Exit Sub
    End If
    
    If Trim$(vcmd) = "" Then
        Exit Sub
    End If
    List1.ListItems.Clear
    
    If Option1(0).Value = True Then ' Executer sur tous les fichiers
        vnb = RENAME.ListView1.ListItems.Count - 1
        For i = 0 To vnb
            Execute22 preview, i, vcmd
        Next
    Else                            ' Executer sur les fichiers sélectionnés
        vnb = LVGetCountSelected(RENAME.ListView1)
        i = LVGetItemSelected(RENAME.ListView1, -1)
        While i <> -1
            Execute22 preview, i, vcmd
            i = LVGetItemSelected(RENAME.ListView1, i)
        Wend
    End If
    
    Exit Sub

ErrGen:
    If MsgBox("There was an error, do you want to continue ?", vbYesNo, "Error") = vbYes Then
        Resume Next
    Else
        Exit Sub
    End If

End Sub

Private Sub Execute22(preview As Boolean, i As Long, vcmd As String)
    On Error GoTo ErrGen
    Dim vnb2 As Integer
    Dim Lect As String, Rep1 As String, vcmd2 As String, ChemCompl1 As String
    Dim LastPart1 As String, LastPart2 As String, Filename As String, Prefix As String, Extension As String
    Dim CompleteFileName As String, CheminFichier As String
    Dim itmX As ListItem

    ' Pour le lecteur et répertoire courant
    Lect = Left$(Dir1Path, 2)
    Rep1 = Dir1Path
    Rep1 = Replace(Rep1, Lect, "")
    ChemCompl1 = Dir1Path
    vnb2 = CharOccurs(ChemCompl1, "\")
    LastPart1 = GetToken(ChemCompl1, "\", vnb2 + 1)
    ' Pour le fichier courant
    Filename = LVGetName(RENAME.ListView1, i)   ' Nom du fichier sans le chemin
    Filename = Prefixe(Filename) & "." & Suffixe(Filename)
    Prefix = Prefixe(Filename)      ' Prefixe
    Extension = Suffixe(Filename)   ' Extension
    If recursive = True Then        ' Nom complet avec le chemin et chemin du fichier
        CompleteFileName = Filename
        CheminFichier = ExtractPath(Filename)
    Else
        CheminFichier = Dir1Path
        CompleteFileName = AddBackSlash(Dir1Path) & Filename
    End If
    vnb2 = CharOccurs(CheminFichier, "\")   ' Dernier partie du nom du répertoire
    LastPart2 = GetToken(CheminFichier, "\", vnb2 + 1)
            
    vcmd2 = vcmd
    vcmd2 = Replace(vcmd2, "%1", Lect)
    vcmd2 = Replace(vcmd2, "%2", Rep1)
    vcmd2 = Replace(vcmd2, "%3", Dir1Path)
    vcmd2 = Replace(vcmd2, "%4", LastPart1)
    vcmd2 = Replace(vcmd2, "%5", Filename)
    vcmd2 = Replace(vcmd2, "%6", Prefix)
    vcmd2 = Replace(vcmd2, "%7", Extension)
    vcmd2 = Replace(vcmd2, "%8", CompleteFileName)
    vcmd2 = Replace(vcmd2, "%9", CheminFichier)
    vcmd2 = Replace(vcmd2, "%0", LastPart2)
    
    Set itmX = List1.ListItems.Add(, , vcmd2)
    If preview = False Then
        ExecCmd vcmd2, ""
        Set itmX = List1.ListItems.Add(, , "    Process returns " & RetStatus)
    End If
Exit Sub
    
ErrGen:
    If MsgBox("There was an error, do you want to continue ?", vbYesNo, "Error") = vbYes Then
        Resume Next
    Else
        Exit Sub
    End If
End Sub

