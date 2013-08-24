VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Begin VB.Form FExport 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Export tags and information"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5940
   HelpContextID   =   489
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   5940
      Width           =   5940
      _ExtentX        =   10478
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Write header"
      Height          =   195
      Left            =   1298
      TabIndex        =   4
      Top             =   960
      Value           =   1  'Checked
      Width           =   1245
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Invert selection"
      Height          =   315
      Left            =   3717
      TabIndex        =   11
      Top             =   4410
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Unselect all"
      Height          =   315
      Left            =   2348
      TabIndex        =   10
      Top             =   4410
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select &all"
      Height          =   315
      Left            =   979
      TabIndex        =   9
      Top             =   4410
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   14
      Left            =   3023
      TabIndex        =   15
      ToolTipText     =   "Export tags"
      Top             =   5580
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   14
      Left            =   1823
      TabIndex        =   14
      Top             =   5580
      Width           =   1095
   End
   Begin THERename.LabelText LabelText3 
      Height          =   285
      Left            =   2948
      TabIndex        =   3
      Top             =   510
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   503
      Caption         =   "or enter Ascii code"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelWidth      =   1400
      MousePointer    =   0
      Text            =   "59"
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignement  =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2745
      Left            =   90
      TabIndex        =   17
      Top             =   1590
      Width           =   5745
      _ExtentX        =   10134
      _ExtentY        =   4842
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Mp3"
      TabPicture(0)   =   "FExport.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "List1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vqf"
      TabPicture(1)   =   "FExport.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "List2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ogg"
      TabPicture(2)   =   "FExport.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "List3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Exif"
      TabPicture(3)   =   "FExport.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "List4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Wma"
      TabPicture(4)   =   "FExport.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "List5"
      Tab(4).ControlCount=   1
      Begin VB.ListBox List5 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "FExport.frx":008C
         Left            =   -74910
         List            =   "FExport.frx":00B4
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   20
         Top             =   420
         Width           =   5565
      End
      Begin VB.ListBox List4 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "FExport.frx":0121
         Left            =   -74910
         List            =   "FExport.frx":018E
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   420
         Width           =   5565
      End
      Begin VB.ListBox List3 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "FExport.frx":038A
         Left            =   -74910
         List            =   "FExport.frx":03EB
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   420
         Width           =   5565
      End
      Begin VB.ListBox List2 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "FExport.frx":0540
         Left            =   -74910
         List            =   "FExport.frx":0562
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   420
         Width           =   5565
      End
      Begin VB.ListBox List1 
         Height          =   2220
         IntegralHeight  =   0   'False
         ItemData        =   "FExport.frx":05CA
         Left            =   90
         List            =   "FExport.frx":067C
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   420
         Width           =   5565
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Which Files ? "
      Height          =   585
      Left            =   90
      TabIndex        =   16
      Top             =   4860
      Width           =   5745
      Begin VB.OptionButton Option2 
         Caption         =   "All files"
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   12
         Top             =   240
         Value           =   -1  'True
         Width           =   945
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Selected files"
         Height          =   255
         Index           =   1
         Left            =   2175
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enclose strings with """
      Height          =   225
      Left            =   2798
      TabIndex        =   5
      Top             =   960
      Value           =   1  'Checked
      Width           =   1845
   End
   Begin THERename.LabelText LabelText2 
      Height          =   285
      Left            =   1208
      TabIndex        =   2
      Top             =   510
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   503
      Caption         =   "Fields separator"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelWidth      =   1200
      MaxLength       =   1
      MousePointer    =   0
      Text            =   ";"
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command7 
      Caption         =   "..."
      Height          =   285
      HelpContextID   =   14
      Left            =   5460
      TabIndex        =   1
      ToolTipText     =   "Browse for file"
      Top             =   120
      Width           =   285
   End
   Begin THERename.LabelText LabelText1 
      Height          =   285
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   503
      Caption         =   "Filename"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelWidth      =   700
      MousePointer    =   0
      Text            =   ""
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select tags and information to export"
      Height          =   195
      Left            =   1680
      TabIndex        =   18
      Top             =   1320
      Width           =   2580
   End
End
Attribute VB_Name = "FExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrGen
    Dim ff As Integer, i As Integer, i2 As Long, vnb2 As Long, j As Integer, vnb As Integer
    Dim sep As String, FS As String, vnom As String, vext As String, SonInfo As String, chemin As String
    Dim vnbtot As Integer
    ' Vérifications sur les paramètres *********************************************************************************************************
    If Len(LabelText1.Text) = 0 Then
        MsgBox "Please, enter a filename"
        LabelText1.SetFocus
        Exit Sub
    End If
    
    If Len(LabelText2.Text) = 0 Then
        MsgBox "Please, enter a field separator"
        LabelText2.SetFocus
        Exit Sub
    End If

    If Len(LabelText3.Text) = 0 Then
        MsgBox "Please, enter a field separator"
        LabelText3.SetFocus
        Exit Sub
    End If
    
    If List1.SelCount = 0 And List2.SelCount = 0 And List3.SelCount = 0 And List4.SelCount = 0 And List5.SelCount = 0 Then
        MsgBox "Please, select tags and information to export"
        SSTab1.Tab = 0
        List1.SetFocus
        Exit Sub
    End If
    
    ' Début des traitements *******************************************************************************
    Me.MousePointer = vbHourglass
    If Check1.Value = 1 Then
        sep = Chr$(34)
    Else
        sep = ""
    End If
    FS = Chr$(LabelText3.Text)
    vnbtot = List1.SelCount + List2.SelCount + List3.SelCount + List4.SelCount + List5.SelCount
    j = 0
    ff = FreeFile
    Open LabelText1 For Output As #ff
    If Check2.Value = 1 Then    ' Il faut écrire l'entête ****************************************************************************************************************
        vnb = List1.ListCount - 1
        If List1.SelCount > 0 Then
            For i = 0 To vnb
                If List1.Selected(i) = True Then
                    j = j + 1
                    If j >= vnbtot Then FS = ""
                    Print #ff, sep + List1.List(i) + sep + FS;
                End If
            Next
        End If
    
        vnb = List2.ListCount - 1
        If List2.SelCount > 0 Then
            For i = 0 To vnb
                If List2.Selected(i) = True Then
                    j = j + 1
                    If j >= vnbtot Then FS = ""
                    Print #ff, sep + List2.List(i) + sep + FS;
                End If
            Next
        End If
    
        vnb = List3.ListCount - 1
        If List3.SelCount > 0 Then
            For i = 0 To vnb
                If List3.Selected(i) = True Then
                    j = j + 1
                    If j >= vnbtot Then FS = ""
                    Print #ff, sep + List3.List(i) + sep + FS;
                End If
            Next
        End If
        
        vnb = List4.ListCount - 1
        If List4.SelCount > 0 Then
            For i = 0 To vnb
                If List4.Selected(i) = True Then
                    j = j + 1
                    If j >= vnbtot Then FS = ""
                    Print #ff, sep + List4.List(i) + sep + FS;
                End If
            Next
        End If
        
        vnb = List5.ListCount - 1
        If List5.SelCount > 0 Then
            For i = 0 To vnb
                If List5.Selected(i) = True Then
                    j = j + 1
                    If j >= vnbtot Then FS = ""
                    Print #ff, sep + List5.List(i) + sep + FS;
                End If
            Next
        End If
        Print #ff, vbCrLf;   ' Fin de l'entête par 0D0A
    End If  ' Entête ? ***************************************************************************************************************
    
    If recursive = False Then
        chemin = AddBackSlash(Dir1Path)
    Else
        chemin = ""
    End If
    
    ' L'export des données en lui même
    If Option2(0).Value = True Then ' Copy All
        vnb2 = RENAME.ListView1.ListItems.Count - 1
        For i2 = 0 To vnb2  ' Boucle sur les fichiers
            vnom = chemin & LVGetName(RENAME.ListView1, i2)
            StatusBar1.SimpleText = "Processing file " + vnom
            vext = UCase$(Suffixe(vnom))
            FS = Chr$(LabelText3.Text)
            Select Case vext
                Case "MP3"  ' *********************************************************************************************************
                    If List1.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = MusMP3.GetMP3Infos(vnom, False)
                        vnb = List1.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List1.Selected(i) = True Then
                                j = j + 1
                                If j = List1.SelCount Then FS = ""
                                ExportMP3 i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
                    
                Case "OGG"  ' *********************************************************************************************************
                    If List3.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = MusOgg.GetOggInfos(vnom, False)
                        vnb = List3.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List3.Selected(i) = True Then
                                j = j + 1
                                If j = List3.SelCount Then FS = ""
                                ExportOGG i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
                    
                Case "WMA"  ' *********************************************************************************************************
                    If List5.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = MusWMA.GetWMAInfos(vnom, False)
                        vnb = List5.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List5.Selected(i) = True Then
                                j = j + 1
                                If j = List5.SelCount Then FS = ""
                                ExportWMA i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
                    
                Case "JPG", "JPE", "JPEG", "TIF", "TIFF" ' *******************************************************************************************
                    If List4.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = PicEXIF.GetEXIFInfos(vnom, False)
                        vnb = List4.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List4.Selected(i) = True Then
                                j = j + 1
                                If j = List4.SelCount Then FS = ""
                                ExportEXIF i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
                
                Case "VQF"  ' *********************************************************************************************************
                    If List2.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = MusVQF.GetVQFInfos(vnom, False)
                        vnb = List2.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List2.Selected(i) = True Then
                                j = j + 1
                                If j = List2.SelCount Then FS = ""
                                ExportVQF i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
            End Select
        Next
    Else    ' Copy selected ***********************************************************************************************************************
        i2 = LVGetItemSelected(RENAME.ListView1, -1)
        While i2 <> -1
            vnom = chemin & LVGetName(RENAME.ListView1, i2)
            StatusBar1.SimpleText = "Processing file " + vnom
            vext = UCase$(Suffixe(vnom))
            FS = Chr$(LabelText3.Text)
            Select Case vext
                Case "MP3"  ' *********************************************************************************************************
                    If List1.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = MusMP3.GetMP3Infos(vnom, False)
                        vnb = List1.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List1.Selected(i) = True Then
                                j = j + 1
                                If j = List1.SelCount Then FS = ""
                                ExportMP3 i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
                    
                Case "JPG", "JPE", "JPEG", "TIF", "TIFF" ' *******************************************************************************************
                    If List4.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = PicEXIF.GetEXIFInfos(vnom, False)
                        vnb = List4.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List4.Selected(i) = True Then
                                j = j + 1
                                If j = List4.SelCount Then FS = ""
                                ExportEXIF i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
                    
                Case "OGG"  ' *********************************************************************************************************
                    If List3.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = MusOgg.GetOggInfos(vnom, False)
                        vnb = List3.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List3.Selected(i) = True Then
                                j = j + 1
                                If j = List3.SelCount Then FS = ""
                                ExportOGG i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
                    
                Case "WMA"  ' *********************************************************************************************************
                    If List5.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = MusWMA.GetWMAInfos(vnom, False)
                        vnb = List5.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List5.Selected(i) = True Then
                                j = j + 1
                                If j = List5.SelCount Then FS = ""
                                ExportWMA i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
                    
                Case "VQF"  ' *********************************************************************************************************
                    If List2.SelCount > 0 Then  ' Il faut bien exporter les tags
                        SonInfo = ""
                        SonInfo = MusVQF.GetVQFInfos(vnom, False)
                        vnb = List2.ListCount - 1
                        j = 0
                        For i = 0 To vnb
                            If List2.Selected(i) = True Then
                                j = j + 1
                                If j = List2.SelCount Then FS = ""
                                ExportVQF i, ff, sep, FS, vnom
                            End If
                        Next
                        Print #ff, vbCrLf;
                    End If
            End Select
            i2 = LVGetItemSelected(RENAME.ListView1, i2)
        Wend
    End If
    Close #ff
    RENAME.RefreshF5
    Me.MousePointer = vbNormal
    Unload Me
    ' Fin des haricots ************************************************************************************
    Exit Sub
    
ErrGen:
 ErreurGrave "FExport:cmdOk"
 Exit Sub
End Sub

Private Sub Command1_Click()
Select Case SSTab1.Tab
    Case 0
       SelectAll List1
    Case 1
       SelectAll List2
    Case 2
        SelectAll List3
    Case 3
        SelectAll List4
    Case 4
        SelectAll List5
End Select
End Sub

Private Sub Command2_Click()
    Select Case SSTab1.Tab
        Case 0
            UnselectAll List1
        Case 1
            UnselectAll List2
        Case 2
            UnselectAll List3
        Case 3
            UnselectAll List4
        Case 4
            UnselectAll List5
    End Select
End Sub

Private Sub Command3_Click()
Select Case SSTab1.Tab
    Case 0
        InvertSelection List1
    Case 1
        InvertSelection List2
    Case 2
        InvertSelection List3
    Case 3
        InvertSelection List4
    Case 4
        InvertSelection List5
End Select
End Sub

Private Sub Command7_Click()
Dim szFilename As String
szFilename = DialogFile(Me.hWnd, 2, "Select file", "tags.txt", "Text file" & Chr$(0) & "*.txt" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "Text file")
If Trim$(szFilename) = "" Then Exit Sub
LabelText1.Text = szFilename
LabelText1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ChangeTab KeyCode, Shift, SSTab1
    KeyCode = 0
End Sub

Private Sub LabelText2_Change()
    If Len(LabelText2.Text) >= 1 Then
        LabelText3.Text = Asc(LabelText2.Text)
    End If
End Sub

Private Sub SelectAll(lst As ListBox)
    Dim i As Integer, vnb As Integer, Index As Integer
    Index = lst.ListIndex
    lst.Visible = False
    vnb = lst.ListCount - 1
    For i = 0 To vnb
        lst.Selected(i) = True
    Next
    lst.Visible = True
    If Index <> -1 Then lst.ListIndex = Index
    lst.SetFocus
End Sub

Private Sub UnselectAll(lst As ListBox)
    Dim i As Integer, vnb As Integer, Index As Integer
    Index = lst.ListIndex
    lst.Visible = False
    vnb = lst.ListCount - 1
    For i = 0 To vnb
        lst.Selected(i) = False
    Next
    lst.Visible = True
    If Index <> -1 Then lst.ListIndex = Index
    lst.SetFocus
End Sub

Private Sub InvertSelection(lst As ListBox)
    Dim i As Integer, vnb As Integer, Index As Integer
    Index = lst.ListIndex
    lst.Visible = False
    vnb = lst.ListCount - 1
    For i = 0 To vnb
        lst.Selected(i) = Not lst.Selected(i)
    Next
    lst.Visible = True
    If Index <> -1 Then lst.ListIndex = Index
    lst.SetFocus
End Sub
Private Sub ExportWMA(i As Integer, ff As Integer, sep As String, FS As String, vnom As String)
Select Case List5.List(i)
    Case "Album"
        Print #ff, sep + MusWMA.Album + sep + FS;
    Case "Artist"
        Print #ff, sep + MusWMA.Artist + sep + FS;
    Case "BitRate"
        Print #ff, sep + MusWMA.BitRate + sep + FS;
    Case "ChannelMode"
        Print #ff, sep + MusWMA.ChannelMode + sep + FS;
    Case "Comment"
        Print #ff, sep + MusWMA.Comment + sep + FS;
    Case "Duration"
        Print #ff, sep + MusWMA.Duration + sep + FS;
    Case "FileName"
        Print #ff, sep + vnom + sep + FS;
    Case "Genre"
        Print #ff, sep + MusWMA.Genre + sep + FS;
    Case "SampleRate"
        Print #ff, sep + MusWMA.SampleRate + sep + FS;
    Case "Title"
        Print #ff, sep + MusWMA.Title + sep + FS;
    Case "Track"
        Print #ff, sep + MusWMA.Track + sep + FS;
    Case "Year"
        Print #ff, sep + MusWMA.Year + sep + FS;
End Select
End Sub
Private Sub ExportEXIF(i As Integer, ff As Integer, sep As String, FS As String, vnom As String)
Select Case List4.List(i)
    Case "FileName"
        Print #ff, sep + vnom + sep + FS;
    Case "Aperture"
        Print #ff, sep + PicEXIF.Aperture + sep + FS;
    Case "Brightness"
        Print #ff, sep + PicEXIF.Brightness + sep + FS;
    Case "CompressedBitsPerPixel"
        Print #ff, sep + PicEXIF.CompressedBitsPerPixel + sep + FS;
    Case "Copyright"
        Print #ff, sep + PicEXIF.Copyright + sep + FS;
    Case "DateTime"
        Print #ff, sep + PicEXIF.DateTime + sep + FS;
    Case "DateTimeDigitized"
        Print #ff, sep + PicEXIF.DateTimeDigitized + sep + FS;
    Case "DateTimeOriginal"
        Print #ff, sep + PicEXIF.DateTimeOriginal + sep + FS;
    Case "ExposureBias"
        Print #ff, sep + PicEXIF.ExposureBias + sep + FS;
    Case "ExposureProgram"
        Print #ff, sep + PicEXIF.ExposureProgram + sep + FS;
    Case "ExposureTime"
        Print #ff, sep + PicEXIF.ExposureTime + sep + FS;
    Case "FirmwareVersion"
        Print #ff, sep + PicEXIF.FirmwareVersion + sep + FS;
    Case "Flash"
        Print #ff, sep + PicEXIF.Flash + sep + FS;
    Case "FNumber"
        Print #ff, sep + PicEXIF.FNumber + sep + FS;
    Case "FocalLength"
        Print #ff, sep + PicEXIF.FocalLength + sep + FS;
    Case "FocalPlaneResolutionUnit"
        Print #ff, sep + PicEXIF.FocalPlaneResolutionUnit + sep + FS;
    Case "FocalPlaneXResolution"
        Print #ff, sep + PicEXIF.FocalPlaneXResolution + sep + FS;
    Case "FocalPlaneYResolution"
        Print #ff, sep + PicEXIF.FocalPlaneYResolution + sep + FS;
    Case "ImageDescription"
        Print #ff, sep + PicEXIF.ImageDescription + sep + FS;
    Case "ImageHeight"
        Print #ff, sep + PicEXIF.ImageHeight + sep + FS;
    Case "ImageWidth"
        Print #ff, sep + PicEXIF.ImageWidth + sep + FS;
    Case "ISOSpeedRatings"
        Print #ff, sep + PicEXIF.ISOSpeedRatings + sep + FS;
    Case "Make"
        Print #ff, sep + PicEXIF.Make + sep + FS;
    Case "MaxAperture"
        Print #ff, sep + PicEXIF.MaxAperture + sep + FS;
    Case "MeteringMode"
        Print #ff, sep + PicEXIF.MeteringMode + sep + FS;
    Case "Model"
        Print #ff, sep + PicEXIF.Model + sep + FS;
    Case "Orientation"
        Print #ff, sep + PicEXIF.Orientation + sep + FS;
    Case "RelatedSoundFile"
        Print #ff, sep + PicEXIF.RelatedSoundFile + sep + FS;
    Case "ResolutionUnit"
        Print #ff, sep + PicEXIF.ResolutionUnit + sep + FS;
    Case "ShutterSpeed"
        Print #ff, sep + PicEXIF.ShutterSpeed + sep + FS;
    Case "SubjectDistance"
        Print #ff, sep + PicEXIF.SubjectDistance + sep + FS;
    Case "Version"
        Print #ff, sep + PicEXIF.Version + sep + FS;
    Case "WhiteBalance"
        Print #ff, sep + PicEXIF.WhiteBalance + sep + FS;
    Case "XResolution"
        Print #ff, sep + PicEXIF.XResolution + sep + FS;
    Case "YResolution"
        Print #ff, sep + PicEXIF.YResolution + sep + FS;
End Select

End Sub
Private Sub ExportOGG(i As Integer, ff As Integer, sep As String, FS As String, vnom As String)
Select Case List3.List(i)
    Case "Album"
        Print #ff, sep + MusOgg.Album + sep + FS;
    Case "Artist"
        Print #ff, sep + MusOgg.Artist + sep + FS;
    Case "AverageBitrate"
        Print #ff, sep + MusOgg.AverageBitrate + sep + FS;
    Case "Channels"
        Print #ff, sep + MusOgg.Channels + sep + FS;
    Case "Comment"
        Print #ff, sep + MusOgg.Comment + sep + FS;
    Case "CopyRight"
        Print #ff, sep + MusOgg.Copyright + sep + FS;
    Case "Date"
        Print #ff, sep + MusOgg.SaDate + sep + FS;
    Case "Description"
        Print #ff, sep + MusOgg.Description + sep + FS;
    Case "EncoderVersion"
        Print #ff, sep + MusOgg.EncoderVersion + sep + FS;
    Case "FileName"
        Print #ff, sep + vnom + sep + FS;
    Case "Genre"
            Print #ff, sep + MusOgg.Genre + sep + FS;
    Case "ISRC"
        Print #ff, sep + MusOgg.ISRC + sep + FS;
    Case "Length"
        Print #ff, sep + MusOgg.Length + sep + FS;
    Case "Location"
        Print #ff, sep + MusOgg.Location + sep + FS;
    Case "LowerBitrate"
        Print #ff, sep + MusOgg.LowerBitrate + sep + FS;
    Case "NominalBitrate"
        Print #ff, sep + MusOgg.NominalBitrate + sep + FS;
    Case "NumberOfTags"
        Print #ff, sep & MusOgg.NumberOfTags & sep + FS;
    Case "Organization"
        Print #ff, sep + MusOgg.Organization + sep + FS;
    Case "Playtime"
        Print #ff, sep + MusOgg.Playtime + sep + FS;
    Case "SampleRate"
        Print #ff, sep + MusOgg.SampleRate + sep + FS;
    Case "SerialNumber"
        Print #ff, sep + MusOgg.SerialNumber + sep + FS;
    Case "Title"
        Print #ff, sep + MusOgg.Title + sep + FS;
    Case "TotalTracks"
        Print #ff, sep + MusOgg.TotalTracks + sep + FS;
    Case "TrackNumber"
        Print #ff, sep + MusOgg.TrackNumber + sep + FS;
    Case "UpperBitrate"
        Print #ff, sep + MusOgg.UpperBitrate + sep + FS;
    Case "Vendor"
        Print #ff, sep + MusOgg.Vendor + sep + FS;
    Case "Version"
        Print #ff, sep + MusOgg.Version + sep + FS;
    Case "Composer"
        Print #ff, sep + MusOgg.Composer + sep + FS;
    Case "Conductor"
        Print #ff, sep + MusOgg.Conductor + sep + FS;
    Case "Ensemble"
        Print #ff, sep + MusOgg.Ensemble + sep + FS;
    Case "Performer"
        Print #ff, sep + MusOgg.Performer + sep + FS;
End Select
End Sub
Private Sub ExportVQF(i As Integer, ff As Integer, sep As String, FS As String, vnom As String)
Select Case List2.List(i)
    Case "Artist"
        Print #ff, sep + MusVQF.Author + sep + FS;
    Case "Bitrate"
        Print #ff, sep + MusVQF.BitRate + sep + FS;
    Case "Comment"
        Print #ff, sep + MusVQF.Comment + sep + FS;
    Case "Copyright"
        Print #ff, sep + MusVQF.Copyright + sep + FS;
    Case "FileName"
        Print #ff, sep + vnom + sep + FS;
    Case "FileSaveAs"
        Print #ff, sep + MusVQF.SaveAsFilename + sep + FS;
    Case "Mono_Stereo"
        Print #ff, sep + MusVQF.Mono_Stereo + sep + FS;
    Case "Quality"
        Print #ff, sep + MusVQF.Quality + sep + FS;
    Case "SampleRate"
        Print #ff, sep + MusVQF.SampleRate + sep + FS;
    Case "Title"
        Print #ff, sep + MusVQF.Title + sep + FS;
End Select

End Sub
Private Sub ExportMP3(i As Integer, ff As Integer, sep As String, FS As String, vnom As String)
Select Case List1.List(i)
    Case "Album"
        Print #ff, sep + MusMP3.Album + sep + FS;
    Case "Artist"
        Print #ff, sep + MusMP3.Artist + sep + FS;
    Case "Band"
        Print #ff, sep + MusMP3.Band + sep + FS;
    Case "BPM"
        Print #ff, sep + MusMP3.BPM + sep + FS;
    Case "Comment"
        Print #ff, sep + MusMP3.Comment + sep + FS;
    Case "Composer"
        Print #ff, sep + MusMP3.Composer + sep + FS;
    Case "Conductor"
        Print #ff, sep + MusMP3.Conductor + sep + FS;
    Case "ContentGroup"
        Print #ff, sep + MusMP3.ContentGroup + sep + FS;
    Case "Copyright"
        Print #ff, sep + MusMP3.Copyright + sep + FS;
    Case "Date"
        Print #ff, sep + MusMP3.mDate + sep + FS;
    Case "EncodedBy"
        Print #ff, sep + MusMP3.EncodedBy + sep + FS;
    Case "EncryptionMethod"
        Print #ff, sep + MusMP3.EncryptionMethod + sep + FS;
    Case "FileName"
        Print #ff, sep + vnom + sep + FS;
    Case "FileOwner"
        Print #ff, sep + MusMP3.FileOwner + sep + FS;
    Case "FileType"
        Print #ff, sep + MusMP3.FileType + sep + FS;
    Case "Genre"
        Print #ff, sep + MusMP3.Genre + sep + FS;
    Case "GroupIdent"
        Print #ff, sep + MusMP3.GroupIdent + sep + FS;
    Case "InitialKey"
        Print #ff, sep + MusMP3.InitialKey + sep + FS;
    Case "InvolvedPeopleList"
        Print #ff, sep + MusMP3.InvolvedPeopleList + sep + FS;
    Case "ISRC"
        Print #ff, sep + MusMP3.ISRC + sep + FS;
    Case "Language"
        Print #ff, sep + MusMP3.Language + sep + FS;
    Case "LinkedInformation"
        Print #ff, sep + MusMP3.LinkedInformation + sep + FS;
    Case "Lyricist"
        Print #ff, sep + MusMP3.Lyricist + sep + FS;
    Case "MediaType"
        Print #ff, sep + MusMP3.MediaType + sep + FS;
    Case "MixArtist"
        Print #ff, sep + MusMP3.MixArtist + sep + FS;
    Case "NetRadioOwner"
        Print #ff, sep + MusMP3.NetRadioOwner + sep + FS;
    Case "NetRadioStation"
        Print #ff, sep + MusMP3.NetRadioStation + sep + FS;
    Case "OriginalAlbum"
        Print #ff, sep + MusMP3.OriginalAlbum + sep + FS;
    Case "OriginalArtist"
        Print #ff, sep + MusMP3.OriginalArtist + sep + FS;
    Case "OriginalFilename"
        Print #ff, sep + MusMP3.OriginalFilename + sep + FS;
    Case "OriginalLyricist"
        Print #ff, sep + MusMP3.OriginalLyricist + sep + FS;
    Case "OriginalYear"
        Print #ff, sep + MusMP3.OriginalYear + sep + FS;
    Case "PartOfASet"
        Print #ff, sep + MusMP3.PartOfASet + sep + FS;
    Case "PlayListDelay"
        Print #ff, sep + MusMP3.PlayListDelay + sep + FS;
    Case "PopulariMeter"
        Print #ff, sep + MusMP3.PopulariMeter + sep + FS;
    Case "Publisher"
        Print #ff, sep + MusMP3.Publisher + sep + FS;
    Case "RecordingDates"
        Print #ff, sep + MusMP3.RecordingDates + sep + FS;
    Case "SoftwareEncodingSettings"
        Print #ff, sep + MusMP3.SoftwareEncodingSettings + sep + FS;
    Case "SongLength"
        Print #ff, sep + MusMP3.SongLength + sep + FS;
    Case "SubTitle"
        Print #ff, sep + MusMP3.SubTitle + sep + FS;
    Case "SynchronizedLyric"
        Print #ff, sep + MusMP3.SynchronizedLyric + sep + FS;
    Case "TermsOfUse"
        Print #ff, sep + MusMP3.TermsOfUse + sep + FS;
    Case "Time"
        Print #ff, sep + MusMP3.Time + sep + FS;
    Case "Title"
        Print #ff, sep + MusMP3.Title + sep + FS;
    Case "TotalTracks"
        Print #ff, sep + MusMP3.TotalTracks + sep + FS;
    Case "TrackNumber"
        Print #ff, sep + MusMP3.TrackNumber + sep + FS;
    Case "UnsynchronizedLyric"
        Print #ff, sep + MusMP3.UnsynchronizedLyric + sep + FS;
    Case "UserText"
        Print #ff, sep + MusMP3.UserText + sep + FS;
    Case "wwwArtist"
        Print #ff, sep + MusMP3.wwwArtist + sep + FS;
    Case "wwwAudioFile"
        Print #ff, sep + MusMP3.wwwAudioFile + sep + FS;
    Case "wwwAudioSource"
        Print #ff, sep + MusMP3.wwwAudioSource + sep + FS;
    Case "wwwCommercialInfo"
        Print #ff, sep + MusMP3.wwwCommercialInfo + sep + FS;
    Case "wwwCopyright"
        Print #ff, sep + MusMP3.wwwCopyright + sep + FS;
    Case "wwwPayment"
        Print #ff, sep + MusMP3.wwwPayment + sep + FS;
    Case "wwwPublisher"
        Print #ff, sep + MusMP3.wwwPublisher + sep + FS;
    Case "wwwRadioPage"
        Print #ff, sep + MusMP3.wwwRadioPage + sep + FS;
    Case "wwwUserURL"
        Print #ff, sep + MusMP3.wwwUserURL + sep + FS;
    Case "Year"
        Print #ff, sep + MusMP3.Year + sep + FS;
End Select
End Sub
