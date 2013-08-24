VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form FCmd 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commands"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   HelpContextID   =   249
   Icon            =   "FCmd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "This is the command's description"
      Top             =   2240
      WhatsThisHelpID =   253
      Width           =   6915
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Clear default command"
      Height          =   315
      Left            =   3780
      TabIndex        =   7
      ToolTipText     =   "Don't use any default command"
      Top             =   2640
      WhatsThisHelpID =   258
      Width           =   1875
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3705
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   6720
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Height          =   315
      Left            =   5760
      Picture         =   "FCmd.frx":01CA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Click and you will see some information"
      Top             =   2640
      WhatsThisHelpID =   259
      Width           =   420
   End
   Begin VB.CommandButton Command6 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   315
      Left            =   6225
      TabIndex        =   9
      ToolTipText     =   "Close and save"
      Top             =   2640
      WhatsThisHelpID =   260
      Width           =   795
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Add to list"
      Height          =   315
      Left            =   5460
      TabIndex        =   11
      ToolTipText     =   "Add current command to saved list"
      Top             =   3300
      WhatsThisHelpID =   262
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   75
      TabIndex        =   10
      ToolTipText     =   "This is the current command used by THE Rename"
      Top             =   3300
      WhatsThisHelpID =   261
      Width           =   5295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Modify"
      Height          =   315
      Left            =   5700
      TabIndex        =   2
      ToolTipText     =   "Modify selected command"
      Top             =   1880
      WhatsThisHelpID =   254
      Width           =   915
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Set as default command"
      Height          =   315
      Left            =   1720
      TabIndex        =   6
      ToolTipText     =   "The selected command will be used when you use the ""Free Form"" option"
      Top             =   2640
      WhatsThisHelpID =   257
      Width           =   1995
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete"
      Height          =   315
      Left            =   980
      TabIndex        =   5
      ToolTipText     =   "Remove the selected commands"
      Top             =   2640
      WhatsThisHelpID =   256
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Use &Now"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      ToolTipText     =   "Use the selected command now and close this window"
      Top             =   2640
      WhatsThisHelpID =   255
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "This is the command's description"
      Top             =   1880
      WhatsThisHelpID =   253
      Width           =   5595
   End
   Begin VB.ListBox List1 
      Height          =   1815
      HelpContextID   =   251
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   251
      Width           =   6615
   End
   Begin VB.CommandButton cmdUp 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   20
      Left            =   6690
      Picture         =   "FCmd.frx":0394
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Move Up"
      Top             =   420
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   263
      Width           =   330
   End
   Begin VB.CommandButton cmdDown 
      Enabled         =   0   'False
      Height          =   330
      HelpContextID   =   20
      Left            =   6690
      Picture         =   "FCmd.frx":0496
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Move Down"
      Top             =   915
      UseMaskColor    =   -1  'True
      WhatsThisHelpID =   264
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current command"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   3060
      Width           =   1245
   End
End
Attribute VB_Name = "FCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SIni As New cInifile
Dim modifs As Boolean
Dim EnCours As Boolean
Private Sub cmdDown_Click()
  On Error Resume Next
  Dim nItem As Integer
  List2.ListIndex = List1.ListIndex
  EnCours = True
  With List1
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
    .AddItem .Text, nItem + 2
    .RemoveItem nItem
    .Selected(nItem + 1) = True
  End With
  With List2
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
    .AddItem .Text, nItem + 2
    .RemoveItem nItem
    .Selected(nItem + 1) = True
  End With
  modifs = True
End Sub

Private Sub cmdUp_Click()
  On Error Resume Next
  Dim nItem As Integer
  EnCours = True
  List2.ListIndex = List1.ListIndex
  With List1
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'can't move 1st item up
    .AddItem .Text, nItem - 1
    .RemoveItem nItem + 1
    .Selected(nItem - 1) = True
  End With
  With List2
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'can't move 1st item up
    .AddItem .Text, nItem - 1
    .RemoveItem nItem + 1
    .Selected(nItem - 1) = True
  End With
  modifs = True
'  List1_Click
End Sub

Private Sub Command1_Click()
    If List1.ListIndex = -1 Then Exit Sub
    If Trim$(RENAME.txtlang.Text) <> "" Then
        If MsgBox("Warning, your actual command line contains some text, do you want to replace it with this command ?", vbYesNo, "Warning") = vbNo Then Exit Sub
    End If
    RENAME.txtlang.Text = List1.List(List1.ListIndex)
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim Index As Integer
    Dim i As Integer
    Dim vnb As Integer
    Dim chemin As String
    
    chemin = AppPath + "commands.ini"
    If List1.ListIndex = -1 Then
        Exit Sub
    End If
    If MsgBox("Are you sure you want to delete this command ?" + vbCrLf + List1.List(List1.ListIndex), vbYesNo, "Warning") = vbNo Then Exit Sub
    Me.MousePointer = vbHourglass
    Index = List1.ListIndex
    StatusBar1.SimpleText = "Removing command"
    List1.RemoveItem (Index)
    List2.RemoveItem (Index)
    StatusBar1.SimpleText = "Updating commands file"
    vnb = List1.ListCount
    ' Sauvegarde du nombre de commandes
    With SIni
         .Path = chemin
         .Section = "General"
         .Key = "NumberOfCommands"
         .Value = Trim$(Str$(vnb))
    End With
    ' Suppression des commandes actelles
    With SIni
         .Path = chemin
         .Section = "Commands"
         .DeleteSection
    End With
     ' Suppression des descriptions associées
    With SIni
         .Path = chemin
         .Section = "Descriptions"
         .DeleteSection
    End With
    For i = 0 To vnb - 1
        With SIni
            .Path = chemin
            .Section = "Commands"
            .Key = "Command" & Trim$(Str$(i + 1))
            .Value = List1.List(i)
        End With
        With SIni
            .Path = chemin
            .Section = "Descriptions"
            .Key = "Description" & Trim$(Str$(i + 1))
            .Value = List2.List(i)
        End With
    Next
    StatusBar1.SimpleText = "Command deleted !"
    Text2.Text = ""
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
    Me.MousePointer = vbNormal
End Sub

Private Sub Command3_Click()
    If List1.ListIndex = -1 Then
        Exit Sub
    End If
    LesOptions.DefaultCommand = List1.List(List1.ListIndex)
End Sub

Private Sub Command4_Click()
    Dim Index As Integer
    Dim chemin As String
    
    chemin = AppPath + "commands.ini"
    If List1.ListIndex = -1 Then
        Exit Sub
    End If
    Index = List1.ListIndex
    List2.List(Index) = Text1.Text
    List1.List(Index) = Text3.Text
    With SIni
        .Path = chemin
        .Section = "Descriptions"
        .Key = "Description" & Trim$(Str$(Index + 1))
        .Value = Text1.Text
    End With
    With SIni
        .Path = chemin
        .Section = "Commands"
        .Key = "Command" & Trim$(Str$(Index + 1))
        .Value = Text3.Text
    End With

End Sub

Private Sub Command5_Click()
    Dim i As Integer
    Dim descr As String
    Dim chemin As String
    Dim vnb As Integer
    chemin = AppPath + "commands.ini"
    
    If Trim$(Text2.Text) = "" Then
        Exit Sub
    End If
    ' D'abord on vérifie que la commande n'est pas déjà présente dans la liste
    For i = 0 To List1.ListCount - 1
        If Trim$(List1.List(i)) = Trim$(Text2.Text) Then
            MsgBox "Warning, This command is already in the list!", vbOKOnly, "Warning"
            Exit Sub
        End If
    Next
    ' Elle n'est pas présente donc on peut l'ajouter. On demande sa description
    descr = InputBox("Enter a description for this command", "Description")
    If descr = "" Then
        descr = " "
    End If
    List1.AddItem Text2.Text
    List2.AddItem descr
    Me.MousePointer = vbHourglass
    StatusBar1.SimpleText = "Updating commands file"
    vnb = List1.ListCount
    ' Sauvegarde du nombre de commandes
    With SIni
         .Path = chemin
         .Section = "General"
         .Key = "NumberOfCommands"
         .Value = Trim$(Str$(vnb))
    End With
    ' Suppression des commandes actelles
    With SIni
         .Path = chemin
         .Section = "Commands"
         .DeleteSection
    End With
     ' Suppression des descriptions associées
    With SIni
         .Path = chemin
         .Section = "Descriptions"
         .DeleteSection
    End With
    For i = 0 To vnb - 1
        With SIni
            .Path = chemin
            .Section = "Commands"
            .Key = "Command" & Trim$(Str$(i + 1))
            .Value = List1.List(i)
        End With
        With SIni
            .Path = chemin
            .Section = "Descriptions"
            .Key = "Description" & Trim$(Str$(i + 1))
            .Value = List2.List(i)
        End With
    Next
    StatusBar1.SimpleText = "Command added !"
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
    End If
    Text2.Text = ""
    Me.MousePointer = vbNormal
End Sub

Private Sub Command6_Click()
    Dim i As Integer
    Dim chemin As String
    Dim vnb As Integer
    chemin = AppPath + "commands.ini"
    If modifs = False Then
        Unload Me
    End If
    
    Me.MousePointer = vbHourglass
    StatusBar1.SimpleText = "Updating commands file"
    vnb = List1.ListCount
    ' Sauvegarde du nombre de commandes
    With SIni
         .Path = chemin
         .Section = "General"
         .Key = "NumberOfCommands"
         .Value = Trim$(Str$(vnb))
    End With
    ' Suppression des commandes actelles
    With SIni
         .Path = chemin
         .Section = "Commands"
         .DeleteSection
    End With
     ' Suppression des descriptions associées
    With SIni
         .Path = chemin
         .Section = "Descriptions"
         .DeleteSection
    End With
    For i = 0 To vnb - 1
        With SIni
            .Path = chemin
            .Section = "Commands"
            .Key = "Command" & Trim$(Str$(i + 1))
            .Value = List1.List(i)
        End With
        With SIni
            .Path = chemin
            .Section = "Descriptions"
            .Key = "Description" & Trim$(Str$(i + 1))
            .Value = List2.List(i)
        End With
    Next
    Me.MousePointer = vbNormal
    Unload Me
End Sub

Private Sub Command7_Click()
    MsgBox "Remember, when you are in the text box of the Free Form option, you can use Ctrl + Up Arrow and Ctrl + Down Arrow to navigate through your saved commands. You can also use Ctrl + Home and Ctrl + End to go to the first command and to go to the last command.", vbOKOnly, "Remember"
End Sub

Private Sub Command8_Click()
    LesOptions.DefaultCommand = ""
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim sValue As String
    Dim vnbcmd As Integer
    Dim chemin As String
    EnCours = False
    modifs = False
    chemin = AppPath + "commands.ini"
    sValue = ""
    With SIni
     .Path = chemin
     .Section = "General"
     .Key = "NumberOfCommands"
     sValue = .Value
    End With
    vnbcmd = Val(sValue)
    For i = 1 To vnbcmd
        With SIni
            .Path = chemin
            .Section = "Commands"
            .Key = "Command" & Trim$(Str$(i))
            sValue = .Value
        End With
        List1.AddItem sValue
        With SIni
            .Path = chemin
            .Section = "Descriptions"
            .Key = "Description" & Trim$(Str$(i))
            sValue = .Value
        End With
        List2.AddItem sValue
    Next
    Text2.Text = RENAME.txtlang.Text
End Sub
Private Sub List1_Click()
    If EnCours = False Then
        Text1.Text = List2.List(List1.ListIndex)
        Text3.Text = List1.List(List1.ListIndex)
        List2.ListIndex = List1.ListIndex
        SetListButtons List1, cmdUp, cmdDown
    Else
        EnCours = False
    End If
End Sub
Private Sub List1_DblClick()
    Text3.SetFocus
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then    ' Touche Suppr
        Command2_Click
    End If
End Sub
Private Sub Text1_GotFocus()
    SelAll Text1
End Sub
Private Sub Text2_GotFocus()
    SelAll Text2
End Sub
Private Sub Text3_GotFocus()
    SelAll Text3
End Sub
