VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Begin VB.Form doptions2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options for directory report & Html report "
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5670
   ControlBox      =   0   'False
   HelpContextID   =   473
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   14
      Left            =   4470
      TabIndex        =   12
      ToolTipText     =   "Save settings for ALL sessions"
      Top             =   5580
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   14
      Left            =   3285
      TabIndex        =   11
      ToolTipText     =   "Cancel your selection"
      Top             =   5580
      Width           =   1095
   End
   Begin TabDlg.SSTab Onglet1 
      Height          =   2685
      Left            =   120
      TabIndex        =   16
      Top             =   2730
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   4736
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "MP3"
      TabPicture(0)   =   "doptions2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "MP3List"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "VQF"
      TabPicture(1)   =   "doptions2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "VQFList"
      Tab(1).ControlCount=   2
      Begin MSComctlLib.ListView MP3List 
         Height          =   1980
         Left            =   75
         TabIndex        =   9
         ToolTipText     =   "Press F2 to change the caption"
         Top             =   615
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   3493
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Caption"
            Object.Width           =   4322
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tag"
            Object.Width           =   4322
         EndProperty
      End
      Begin MSComctlLib.ListView VQFList 
         Height          =   1980
         Left            =   -74925
         TabIndex        =   10
         ToolTipText     =   "Press F2 to change the caption"
         Top             =   615
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   3493
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Caption"
            Object.Width           =   4322
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tag"
            Object.Width           =   4322
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To re order elements, use the Alt+Up arrow and Alt+Down arrow"
         Height          =   195
         Left            =   -74520
         TabIndex        =   18
         Top             =   405
         Width           =   4515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To re order elements, use the Alt+Up arrow and Alt+Down arrow"
         Height          =   195
         Left            =   480
         TabIndex        =   17
         Top             =   405
         Width           =   4515
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Elements to include in Html report "
      Height          =   1110
      Left            =   120
      TabIndex        =   15
      Top             =   1485
      Width           =   5445
      Begin VB.CheckBox Check5 
         Caption         =   "Information about MP3 and VQF files"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   810
         Width           =   3075
      End
      Begin VB.CheckBox Check4 
         Caption         =   "File's Attributes"
         Height          =   195
         Left            =   3120
         TabIndex        =   8
         Top             =   540
         Width           =   1455
      End
      Begin VB.CheckBox Check3 
         Caption         =   "File's &Date"
         Height          =   195
         Left            =   3120
         TabIndex        =   7
         Top             =   270
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "File's &Size"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Folder"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Pictures "
      Height          =   600
      Left            =   120
      TabIndex        =   14
      Top             =   795
      Width           =   5445
      Begin VB.CheckBox Check24 
         Caption         =   "Include lin&ks"
         Height          =   195
         Left            =   3120
         TabIndex        =   3
         ToolTipText     =   "Include links to pictures and MP3 files"
         Top             =   270
         Width           =   1200
      End
      Begin VB.CheckBox Check21 
         Caption         =   "Include di&mensions"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "This is available for printed report and HTML report"
         Top             =   240
         Width           =   1650
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Report type "
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5445
      Begin VB.OptionButton Option5 
         Caption         =   "&Long Report"
         Height          =   195
         HelpContextID   =   14
         Index           =   1
         Left            =   3105
         TabIndex        =   1
         ToolTipText     =   "This is only available when you print directory's content"
         Top             =   270
         Width           =   1380
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Sh&ort report"
         Height          =   180
         HelpContextID   =   14
         Index           =   0
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "This is only available when you print directory's content"
         Top             =   270
         Width           =   1155
      End
   End
End
Attribute VB_Name = "doptions2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SIni As New cInifile
Dim chemin As String
Dim TagList(57) As String
Dim VQFTags(9) As String

Private Sub Check5_Click()
    If Check5.Value = 1 Then
        Onglet1.Enabled = True
    Else
        Onglet1.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    Dim vnb As Integer
    Dim vtmp As String
    If Option5(0).Value = True Then
        LesOptions.DirectoryReport = 1
    Else
        LesOptions.DirectoryReport = 2
    End If
    LesOptions.IncludePictInfo = Check21.Value
    LesOptions.IncludeLinks = Check24.Value
    LesOptions.HtmlIncSize = Check2.Value
    LesOptions.HtmlIncDate = Check3.Value
    LesOptions.HtmlIncAttr = Check4.Value
    LesOptions.HtmlIncFolder = Check1.Value
    LesOptions.HtmlIncMusic = Check5.Value
    
    ' Sauvegarde du nombre de tag des MP3
    vnb = MP3List.ListItems.Count
    ' Suppression des commandes actuelles
    With SIni
         .Section = "MP3Tags"
         .DeleteSection
    End With
    With SIni
         .Section = "MP3Tags"
         .Key = "NumberOfTags"
         .Value = Trim$(Str$(vnb))
    End With
    For i = 1 To vnb
        vtmp = "0|"
        If MP3List.ListItems(i).Checked = True Then
            vtmp = "1|"
        End If
        vtmp = vtmp + MP3List.ListItems(i).Text + "|" + Cherche1(MP3List.ListItems(i).SubItems(1))
        With SIni
            .Section = "MP3Tags"
            .Key = "Tag" & Trim$(Str$(i))
            .Value = vtmp
        End With
    Next
    
    ' Sauvegarde du nombre de tag des VQF
    vnb = VQFList.ListItems.Count
    ' Suppression des commandes actelles
    With SIni
         .Section = "VQFTags"
         .DeleteSection
    End With
    With SIni
         .Section = "VQFTags"
         .Key = "NumberOfTags"
         .Value = Trim$(Str$(vnb))
    End With
    For i = 1 To vnb
        vtmp = "0|"
        If VQFList.ListItems(i).Checked = True Then
            vtmp = "1|"
        End If
        vtmp = vtmp + VQFList.ListItems(i).Text + "|" + Cherche2(VQFList.ListItems(i).SubItems(1))
        With SIni
            .Section = "VQFTags"
            .Key = "Tag" & Trim$(Str$(i))
            .Value = vtmp
        End With
    Next
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeTab KeyCode, Shift, Onglet1
    KeyCode = 0
End Sub

Private Sub Form_Load()
Dim itmX As ListItem
Dim vnbcmd As Integer
Dim i As Integer
Dim sValue As String
chemin = AppPath
If LesOptions.DirectoryReport = 1 Then
    Option5(0).Value = True
    Option5(1).Value = False
Else
    Option5(0).Value = False
    Option5(1).Value = True
End If
Check5.Value = LesOptions.HtmlIncMusic
Check21.Value = LesOptions.IncludePictInfo
Check24.Value = LesOptions.IncludeLinks
Check1.Value = LesOptions.HtmlIncFolder
Check2.Value = LesOptions.HtmlIncSize
Check3.Value = LesOptions.HtmlIncDate
Check4.Value = LesOptions.HtmlIncAttr
Check5_Click
VQFTags(1) = "Artist"
VQFTags(2) = "Bitrate"
VQFTags(3) = "Comment"
VQFTags(4) = "Copyright"
VQFTags(5) = "FileSaveAs"
VQFTags(6) = "Mono_Stereo"
VQFTags(7) = "Quality"
VQFTags(8) = "SampleRate"
VQFTags(9) = "Title"

TagList(1) = "Album"
TagList(2) = "Artist"
TagList(3) = "Band"
TagList(4) = "BPM"
TagList(5) = "Comment"
TagList(6) = "Composer"
TagList(7) = "Conductor"
TagList(8) = "ContentGroup"
TagList(9) = "Copyright"
TagList(10) = "Date"
TagList(11) = "EncodedBy"
TagList(12) = "EncryptionMethod"
TagList(13) = "FileOwner"
TagList(14) = "FileType"
TagList(15) = "Genre"
TagList(16) = "GroupIdent"
TagList(17) = "InitialKey"
TagList(18) = "InvolvedPeopleList"
TagList(19) = "ISRC"
TagList(20) = "Language"
TagList(21) = "LinkedInformation"
TagList(22) = "Lyricist"
TagList(23) = "MediaType"
TagList(24) = "MixArtist"
TagList(25) = "NetRadioOwner"
TagList(26) = "NetRadioStation"
TagList(27) = "OriginalAlbum"
TagList(28) = "OriginalArtist"
TagList(29) = "OriginalFilename"
TagList(30) = "OriginalLyricist"
TagList(31) = "OriginalYear"
TagList(32) = "PartOfASet"
TagList(33) = "PlayListDelay"
TagList(34) = "PopulariMeter"
TagList(35) = "Publisher"
TagList(36) = "RecordingDates"
TagList(37) = "SoftwareEncodingSettings"
TagList(38) = "SongLength"
TagList(39) = "SubTitle"
TagList(40) = "SynchronizedLyric"
TagList(41) = "TermsOfUse"
TagList(42) = "Time"
TagList(43) = "Title"
TagList(44) = "TotalTracks"
TagList(45) = "TrackNumber"
TagList(46) = "UnsynchronizedLyric"
TagList(47) = "UserText"
TagList(48) = "wwwArtist"
TagList(49) = "wwwAudioFile"
TagList(50) = "wwwAudioSource"
TagList(51) = "wwwCommercialInfo"
TagList(52) = "wwwCopyright"
TagList(53) = "wwwPayment"
TagList(54) = "wwwPublisher"
TagList(55) = "wwwRadioPage"
TagList(56) = "wwwUserURL"
TagList(57) = "Year"
chemin = chemin + "Music.ini"

' Cas des MP3
With SIni
    .Path = chemin
    .Section = "MP3Tags"
    .Key = "NumberOfTags"
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "MP3Tags"
        .Key = "Tag" & Trim$(Str$(i))
        sValue = .Value
    End With
    Set itmX = MP3List.ListItems.Add(, , GetToken(sValue, "|", 2))
    itmX.SubItems(1) = TagList(Val(GetToken(sValue, "|", 3)))
    If Val(GetToken(sValue, "|", 1)) = 1 Then
        itmX.Checked = True
    End If
Next

' Cas des VQF
With SIni
    .Path = chemin
    .Section = "VQFTags"
    .Key = "NumberOfTags"
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "VQFTags"
        .Key = "Tag" & Trim$(Str$(i))
        sValue = .Value
    End With
    Set itmX = VQFList.ListItems.Add(, , GetToken(sValue, "|", 2))
    itmX.SubItems(1) = VQFTags(Val(GetToken(sValue, "|", 3)))
    If Val(GetToken(sValue, "|", 1)) = 1 Then
        itmX.Checked = True
    End If
Next
End Sub

Private Sub MP3List_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim tAlt As Integer
    On Error Resume Next
    tAlt = GetKeyState(VK_MENU)
    If (tAlt = -127 Or tAlt = -128) And KeyCode = 38 Then ' Déplacer vers le haut
        MoveFilesUp MP3List, doptions2
    End If
 
    If (tAlt = -127 Or tAlt = -128) And KeyCode = 40 Then ' Déplacer vers le bas
        MoveFilesDown MP3List, doptions2
    End If
    
    If KeyCode = 113 Then ' F2
        MP3List.StartLabelEdit
    End If
End Sub
Private Sub VQFList_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim tAlt As Integer
    On Error Resume Next
    tAlt = GetKeyState(VK_MENU)
    If (tAlt = -127 Or tAlt = -128) And KeyCode = 38 Then ' Déplacer vers le haut
        MoveFilesUp VQFList, doptions2
    End If
 
    If (tAlt = -127 Or tAlt = -128) And KeyCode = 40 Then ' Déplacer vers le bas
        MoveFilesDown VQFList, doptions2
    End If
    
    If KeyCode = 113 Then ' F2
        VQFList.StartLabelEdit
    End If
End Sub
Private Function Cherche1(zTexte As String) As String
Dim i As Integer
Dim vnb As Integer
Dim zTxt As String
vnb = UBound(TagList)
zTxt = Trim$(zTexte)
For i = 1 To vnb
    If TagList(i) = zTxt Then
        Cherche1 = Trim$(Str$(i))
        Exit Function
    End If
Next
Cherche1 = "1"
End Function
Private Function Cherche2(zTexte As String) As String
Dim i As Integer
Dim vnb As Integer
Dim zTxt As String
vnb = UBound(VQFTags)
zTxt = Trim$(zTexte)
For i = 1 To vnb
    If VQFTags(i) = zTxt Then
        Cherche2 = Trim$(Str$(i))
        Exit Function
    End If
Next
Cherche2 = "1"
End Function

