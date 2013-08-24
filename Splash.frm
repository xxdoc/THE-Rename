VERSION 5.00
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4470
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3975
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   4470
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LoadMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3780
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   3240
      Top             =   300
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
    Dim datedeb As Single
    Dim datefin As Single
    Dim diff As Long
    datedeb = Timer()
    AppPath = AddBackSlash(App.Path)
    App.HelpFile = AppPath + "therename.hlp"
    
    LoadMsg.Text = "Loading preferences..."
    DoEvents
    OkUseAbbrev = False ' Par défaut on n'utilise pas les abréviations
    UseMP3 = False
    UseVQF = False
    UseCylcic = False
    VnbCyclic = 0
    OptionsCyclic = False
    VnbHistory = 0
    RechPref = False
    RechSuff = False
    FolderOk = False
    Folder1 = 0
    Folder5 = "1"
    Folder2 = 0
    Folder3 = 0
    Folder6 = " "
    Folder4 = 0
    VnbRep = 0
    TemMove = False
    recursive = False
    VancRep = ""
    TemDelete = False
    rafraichir = True
    
    LoadMsg.Text = "Loading jpeg tags..."
    DoEvents
    LoadTags
    
    LoadMsg.Text = "Loading rules..."
    DoEvents
    LesRegles.LoadRulesFromFile AppPath & "Rules.ini"
    
    If GetSetting("THERename", "Donation", "ThanksYou", "") = "" Then
        LoadMsg.ForeColor = &HFF&
        LoadMsg.Text = "Please make a donation / Merci de faire une donation" + vbCrLf + "See the 'Aboux' box / Allez voir la fenêtre 'About'"
        DoEvents
        datefin = Timer()
        diff = (datedeb - datefin) / 1000
        Timer1.Interval = Timer1.Interval - diff
    Else
        DoEvents
        LoadMsg.Text = "Many thanks for your donation !"
        Timer1.Interval = 0
        Unload Me
    End If
End Sub
Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
