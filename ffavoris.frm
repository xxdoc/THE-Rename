VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ffavoris 
   AutoRedraw      =   -1  'True
   Caption         =   "Organize your favorites"
   ClientHeight    =   2640
   ClientLeft      =   2070
   ClientTop       =   2925
   ClientWidth     =   7170
   HelpContextID   =   20
   Icon            =   "ffavoris.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8400
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ffavoris.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ffavoris.frx":0466
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ffavoris.frx":0646
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ffavoris.frx":0826
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ffavoris.frx":0A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ffavoris.frx":0B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ffavoris.frx":0CBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ffavoris.frx":0E1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ffavoris.frx":0F76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7170
      _ExtentX        =   12647
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Select a directory to add to your favorites"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Sort the list of favorites"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open the selected favorite in the directory list"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete selected item from your favorites"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search for invalid folders in your favorites"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Run a command in each folder"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search for files in folders"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Recreate deleted folders"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Print the list of your favorites"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   20
      Left            =   6720
      Picture         =   "ffavoris.frx":1152
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Move Down"
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   20
      Left            =   6720
      Picture         =   "ffavoris.frx":1254
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Move Up"
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   2880
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   20
      Left            =   2683
      TabIndex        =   1
      ToolTipText     =   "Cancel any modification"
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   20
      Left            =   3633
      TabIndex        =   2
      ToolTipText     =   "Validate all modifications"
      Top             =   2280
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   1740
      HelpContextID   =   20
      IntegralHeight  =   0   'False
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "This is the list of your favorites"
      Top             =   405
      Width           =   6615
   End
End
Attribute VB_Name = "ffavoris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oAutoPos As New clsAutoPositioner

Private Sub cmdDown_Click()
    ButtonDown List1
End Sub

Private Sub cmdUp_Click()
    ButtonUp List1
End Sub

Private Sub Command3_Click()
 Dim i As Integer
 For i = 0 To List1.ListCount - 1
    fav(i + 1) = List1.List(i)
 Next
 If List1.ListCount < 20 Then
    For i = List1.ListCount + 1 To 20
        fav(i) = ""
    Next
 End If
 For i = 0 To 19
    RENAME.menufav(i).Caption = "&" + Chr$(65 + i) + " " + fav(i + 1)
    RENAME.mnufav(i).Caption = "&" + Chr$(65 + i) + " " + fav(i + 1)
 Next
 Unload Me
End Sub

Private Sub Command4_Click()
 Unload Me
End Sub
Private Sub Form_Load()
 Dim i As Integer
 List1.Clear
 For i = 1 To 20
  If Len(Trim$(fav(i))) > 0 Then
   List1.AddItem fav(i)
  End If
 Next
m_oAutoPos.AddAssignment Me.Command3, Me, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.Command4, Me, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.cmdUp, Me, tCONTAINER_RELATIVE_POS_RIGHT
m_oAutoPos.AddAssignment Me.cmdDown, Me, tCONTAINER_RELATIVE_POS_RIGHT
m_oAutoPos.AddAssignment Me.List1, Me, tCONTAINER_WIDTH_DELTA_RIGHT
m_oAutoPos.AddAssignment Me.List1, Me, tCONTAINER_HEIGHT_DELTA_BOTTOM
End Sub

Private Sub Form_Resize()
    m_oAutoPos.RefreshPositions
End Sub

Private Sub List1_Click()
 SetListButtons List1, cmdUp, cmdDown
End Sub

Private Sub List1_DblClick()
 Ouvrir
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 46 ' Suppr
   SupFavoris
  Case 45 ' Ins
   Ajouter
 End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1  ' Ajouter
            Ajouter
        Case 2  ' Trier
            Trier
        Case 3  ' Ouvrir
            Ouvrir
        Case 4  ' Supprimer
            SupFavoris
        Case 5  ' Valider
            ValidateFav
        Case 7  ' Executer
            FExecCmd.Show 1
            'Executer
        Case 8  ' Rechercher un ou des fichiers dans les favoris
            FSearchFav.Show 1
        Case 9  ' Recréer les répertoires inexistants
            Recreer
        Case 10 ' Imprimer la liste
            ImpListe
    End Select

End Sub

Private Sub Ajouter()
 Dim szFilename As String
 Dim i As Integer
 If List1.ListCount = 20 Then
  MsgBox "Favorites are full, you must delete some of them to add more favorites"
  Exit Sub
 End If
 szFilename = ""
 szFilename = BrowseFolder(Me, "Browse for folder:")
 If Len(Trim$(szFilename)) = 0 Then
  Exit Sub
 End If
 For i = 1 To 20
  If fav(i) = szFilename Then
   Beep
   MsgBox "This directory is already in your favorites"
   Exit Sub
  End If
 Next
 List1.AddItem szFilename
 List1.SetFocus
End Sub

Private Sub Trier()
 Dim i As Integer
 List2.Clear
 List1.Visible = False
 For i = 0 To List1.ListCount - 1
  List2.AddItem List1.List(i)
 Next
 List1.Clear
 For i = 0 To List2.ListCount - 1
  List1.AddItem List2.List(i)
 Next
 List1.Visible = True
 List1.SetFocus
End Sub

Private Sub Ouvrir()
On Error GoTo ErrorHandler
If Len(Trim$(List1.Text)) = 0 Then
 MsgBox "Select a directory please !"
Else
 RENAME.FolderTreeview1(0).SelectedFolder = Trim$(List1.Text)
 Dir1Path = Trim$(List1.Text)
 Command3_Click
End If
Exit Sub
 
ErrorHandler:   ' Error-handling routine.
Select Case Err.Number  ' Evaluate error number.
 Case 76
  MsgBox "Error - This directory doesn't not exist any more !"
 Case 68
  MsgBox "Error - This drive is unavailable !"
End Select
Exit Sub
End Sub

Private Sub SupFavoris()
If Len(Trim$(List1.Text)) = 0 Then
 MsgBox "Select a directory please !"
Else
 List1.RemoveItem (List1.ListIndex)
End If
List1.SetFocus
End Sub
' Procédure chargée de recréer les répertoires qui n'existent plus
Private Sub Recreer()
Dim repertoire As String, unite As String
Dim RepEnCours As String
Dim vnbsupp As Integer, i As Integer
vnbsupp = 0
RepEnCours = CurDir
On Error GoTo ErrorHandler
For i = List1.ListCount - 1 To 0 Step -1
    unite = Left$(Trim$(List1.List(i)), 3)
    repertoire = Mid$(Trim$(List1.List(i)), 4)
    ChDir unite & repertoire
Next
MsgBox "I've created " & vnbsupp & " folder(s)"
ChDir RepEnCours ' et finalement on reviens là où on étais
List1.SetFocus
Exit Sub

ErrorHandler:   ' Error-handling routine.
Select Case Err.Number  ' Evaluate error number.
Case 76
    If MsgBox("This folder does not exist any more, would you like to create it ?" & vbCrLf & List1.List(i), vbYesNo, "Recreate Folders") = vbYes Then
        MkDir List1.List(i)
        vnbsupp = vnbsupp + 1
        Resume Next
    End If
Case 68
    If MsgBox("This folder does not exist any more, would you like to create it ?" & vbCrLf & List1.List(i), vbYesNo, "Recreate Folders") = vbYes Then
        MkDir List1.List(i)
        vnbsupp = vnbsupp + 1
        Resume Next
    End If
End Select
Exit Sub
End Sub
Private Sub ValidateFav()
Dim repertoire As String, unite As String, RepEnCours As String
Dim vnbsupp As Integer
Dim i As Integer
vnbsupp = 0
RepEnCours = CurDir
On Error GoTo ErrorHandler
For i = List1.ListCount - 1 To 0 Step -1
    unite = Left$(Trim$(List1.List(i)), 3)
    repertoire = Mid$(Trim$(List1.List(i)), 4)
    ChDir unite & repertoire
Next
MsgBox "I've deleted " & vnbsupp & " favorite(s)"
ChDir RepEnCours ' et finalement on reviens là où on étais
List1.SetFocus
Exit Sub

ErrorHandler:   ' Error-handling routine.
Select Case Err.Number  ' Evaluate error number.
Case 76
    If MsgBox("This folder does not exist any more, would you like to remove it from your favorites : " + List1.List(i), vbYesNo, "Validate Favorites") = vbYes Then
     List1.RemoveItem (i)
     vnbsupp = vnbsupp + 1
     Resume Next
    End If
Case 68
    If MsgBox("This folder does not exist any more, would you like to remove it from your favorites : " + List1.List(i), vbYesNo, "Validate Favorites") = vbYes Then
     List1.RemoveItem (i)
     vnbsupp = vnbsupp + 1
     Resume Next
    End If
End Select
Exit Sub
End Sub

Private Sub ImpListe()
Dim i As Integer
Dim vnb As Integer
vnb = List1.ListCount - 1

If List1.ListCount <= 0 Then
    Exit Sub
End If

Printer.Print
Printer.Print "List of your favorites on " + Format$(Date, "Long Date") + " at " + Format$(Time, "Long Time")
Printer.Print " "
Printer.Print " "
For i = 0 To vnb
    Printer.Print List1.List(i)
Next
Printer.Print " "
Printer.Print "Total of " + Trim$(Str$(List1.ListCount)) + " favorite(s)"
Printer.EndDoc
End Sub
