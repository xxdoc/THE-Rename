VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FSearchFav 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search for files in your favorites"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6930
   ControlBox      =   0   'False
   HelpContextID   =   79
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView List1 
      Height          =   1710
      Left            =   75
      TabIndex        =   2
      Top             =   2460
      Width           =   6720
      _ExtentX        =   11853
      _ExtentY        =   3016
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4665
      Width           =   6930
      _ExtentX        =   12224
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Text4 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   4845
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   300
      HelpContextID   =   20
      Left            =   3518
      TabIndex        =   4
      ToolTipText     =   "Close this window"
      Top             =   4290
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Search"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   20
      Left            =   2558
      TabIndex        =   3
      ToolTipText     =   "Search file(s)"
      Top             =   4290
      Width           =   855
   End
   Begin VB.ListBox List2 
      Height          =   1320
      IntegralHeight  =   0   'False
      Left            =   90
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1035
      Width           =   6720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Type the name of the file you are searching for"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   75
      Width           =   3285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "List of favorites (check or unchek favorites in winch search must be made)"
      Height          =   195
      Left            =   75
      TabIndex        =   5
      Top             =   795
      Width           =   5250
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mcheckall 
         Caption         =   "Check all"
      End
      Begin VB.Menu muncheck 
         Caption         =   "Uncheck all"
      End
   End
End
Attribute VB_Name = "FSearchFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command3_Click()
    cHist14.AddNewItem Text4.Text
    Executer
End Sub

Private Sub Command4_Click()
    cHist14.AddNewItem Text4.Text
    Unload Me
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Dim vnb As Integer
    vnb = ffavoris.List1.ListCount - 1
    For i = 0 To vnb
        List2.AddItem ffavoris.List1.List(i)
        List2.Selected(i) = True
    Next
    cHist14.sKey = "SearchFiles"
    cHist14.Items Text4
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     cHist14.AddNewItem Text4.Text
End Sub

Private Sub Executer()
Dim clsFind As New clsFindFile
Dim strFile As String
Dim afficher As Boolean
Dim vnb As Integer
Dim i As Integer, j As Integer
Dim vnbfounds As Long
Dim vnbfav As Integer
Dim extrait As String, Filtre As String, chemin As String
Dim itmX As ListItem

If Trim$(Text4.Text) = "" Then Exit Sub

vnbfav = List2.ListCount - 1
Filtre = Trim$(Text4.Text)
' Suppression des caractères en trop
If Right$(Filtre, 1) = ";" Then
    Filtre = Left$(Filtre, Len(Filtre) - 1)
End If
If Left$(Filtre, 1) = ";" Then
    Filtre = Mid$(Filtre, 2)
End If
List1.ListItems.Clear
' normalement la condition de filtre est bonne
vnb = CharOccurs(Filtre, ";")
vnb = vnb + 1
Me.MousePointer = vbHourglass
For j = 0 To vnbfav
    If List2.Selected(j) = True Then
        chemin = AddBackSlash(List2.List(j))
        StatusBar1.SimpleText = "Searching in " + chemin
        For i = 1 To vnb    ' Boucle sur les filtres
            extrait = GetToken(Filtre, ";", i)
            strFile = clsFind.Find(chemin & extrait, True)
            Do While Len(strFile)
                afficher = True
                If InStr(Suffixe(extrait), "*") = 0 Then  ' Test de correspondance sur l'intégralité du masque de sélection
                    If UCase$(Suffixe(strFile)) <> UCase$(Suffixe(extrait)) Then
                        afficher = False
                    End If
                End If
                If Trim$(strFile) = "." Or Trim$(strFile) = ".." Then
                    afficher = False
                End If
                If afficher = True Then
                    vnbfounds = vnbfounds + 1
                    Set itmX = List1.ListItems.Add(, , chemin & strFile)
                End If
                strFile = clsFind.FindNext()
            Loop
        Next
    End If  ' Favoris sélectionné **********************************************************************************
Next
Me.MousePointer = vbNormal
StatusBar1.SimpleText = "Match : " + Trim$(Str$(vnbfounds))
End Sub

Private Sub List1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub List1_DblClick()
    FileExecutor Me.hwnd, List1.ListItems(List1.SelectedItem.Index), "Open"
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 13 'Entrée
    FileExecutor Me.hwnd, List1.ListItems(List1.SelectedItem.Index), "Open"
  Case 46 ' Supress
   If Shift = 1 Then
    TemDelete = True
   Else
    TemDelete = False
   End If
   mdelete_Click
 End Select
End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
 On Error Resume Next
 Select Case KeyCode
  Case 106  ' * select all
    mcheckall_Click
    List2.SetFocus
  Case 109 ' - unselect
    muncheck_Click
    List2.SetFocus
 End Select
End Sub

Private Sub List2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu menu
    End If
End Sub
Private Sub mcheckall_Click()
    Dim i As Integer
    Dim vnb As Integer
    vnb = List2.ListCount - 1
    For i = 0 To vnb
        List2.Selected(i) = True
    Next
End Sub

Private Sub muncheck_Click()
    Dim i As Integer
    Dim vnb As Integer
    vnb = List2.ListCount - 1
    For i = 0 To vnb
        List2.Selected(i) = False
    Next
End Sub

Private Sub mdelete_Click()
Dim fileop As New CSHFileOp
Dim i As Long
Dim sItem As String
i = 0
With fileop
    .ParentWnd = hwnd
    .ClearSourceFiles
    .ClearDestFiles
End With

i = LVGetItemSelected(List1, -1)
While i <> -1
 sItem = LVGetName(List1, i)
 fileop.AddSourceFile sItem
 i = LVGetItemSelected(List1, i)
Wend

If TemDelete = True Then
 fileop.AllowUndo = False
 TemDelete = False
End If
If fileop.DeleteFiles Then
    Executer
End If
End Sub

