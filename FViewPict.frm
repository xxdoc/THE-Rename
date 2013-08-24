VERSION 5.00
Begin VB.Form FViewPict 
   Caption         =   "Picture"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6870
   ClipControls    =   0   'False
   HelpContextID   =   80
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin THERename.cpvPicScroll Acdsee2 
      Height          =   1755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3096
      BorderStyle     =   1
      Picture         =   "FViewPict.frx":0000
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      Begin VB.Menu mprint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mclipboard 
         Caption         =   "&Copy to clipboard"
      End
      Begin VB.Menu mwallpaper 
         Caption         =   "&Set as wallpaper"
      End
      Begin VB.Menu msep0 
         Caption         =   "-"
      End
      Begin VB.Menu mzoomim 
         Caption         =   "Zoom &In"
      End
      Begin VB.Menu mzoomout 
         Caption         =   "Zoom &Out"
      End
      Begin VB.Menu mRealSize 
         Caption         =   "&Real Size"
      End
      Begin VB.Menu mstretch 
         Caption         =   "S&tretch"
      End
      Begin VB.Menu mbestfit 
         Caption         =   "Best &Fit"
      End
      Begin VB.Menu msep2 
         Caption         =   "-"
      End
      Begin VB.Menu mfirst 
         Caption         =   "First picture (Home)"
      End
      Begin VB.Menu mprev 
         Caption         =   "Previous picture (PgUp)"
      End
      Begin VB.Menu mnext 
         Caption         =   "Next picture (PgDwn)"
      End
      Begin VB.Menu mlast 
         Caption         =   "Last Picture (End)"
      End
      Begin VB.Menu msep1 
         Caption         =   "-"
      End
      Begin VB.Menu mclose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "FViewPict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Const SPI_SETDESKWALLPAPER = 20
Dim PictName As String

Private Sub Acdsee2_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 39 ' Droite
            Acdsee2.Go gdRight
            Acdsee2.SetFocus
        Case 40 ' Bas
            Acdsee2.Go gdBottom
            Acdsee2.SetFocus
        Case 37 ' Gauche
            Acdsee2.Go gdLeft
            Acdsee2.SetFocus
        Case 38 ' Haut
            Acdsee2.Go gdTop
            Acdsee2.SetFocus
        Case 107    ' +
            Acdsee2.ZoomIn
            Acdsee2.SetFocus
        Case 109    ' -
            Acdsee2.ZoomOut
            Acdsee2.SetFocus
        Case 106    ' *
            Acdsee2.BestFit
            Acdsee2.SetFocus
        Case 111    ' /
            Acdsee2.Stretch
            Acdsee2.SetFocus
    End Select
End Sub

Private Sub Acdsee2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mfile
    End If
End Sub
Private Sub Form_Click()
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim ind As Long, vnb As Long
    Dim vtmp As String, Pref As String, vret As Boolean, chemin As String
    Select Case KeyCode
        Case 39 ' Droite
            Acdsee2.Go gdRight
            Acdsee2.SetFocus
        Case 40 ' Bas
            Acdsee2.Go gdBottom
            Acdsee2.SetFocus
        Case 37 ' Gauche
            Acdsee2.Go gdLeft
            Acdsee2.SetFocus
        Case 38 ' Haut
            Acdsee2.Go gdTop
            Acdsee2.SetFocus
        Case 107    ' +
            Acdsee2.ZoomIn
            Acdsee2.SetFocus
        Case 109    ' -
            Acdsee2.ZoomOut
            Acdsee2.SetFocus
        Case 106    ' *
            Acdsee2.BestFit
            Acdsee2.SetFocus
        Case 111    ' /
            Acdsee2.Stretch
            Acdsee2.SetFocus
        Case 34 ' Page Down
            ind = RENAME.ListView1.SelectedItem.Index
            vnb = RENAME.ListView1.ListItems.Count
            If ind + 1 <= vnb Then
                RENAME.ListView1.ListItems(ind).Selected = False
                RENAME.ListView1.ListItems(ind + 1).Selected = True
                Pref = UCase$(Suffixe(RENAME.ListView1.SelectedItem.Text))
                If Pref = "JPG" Or Pref = "BMP" Or Pref = "GIF" Or Pref = "JPEG" Or Pref = "DIB" Or Pref = "WMF" Or Pref = "EMF" Or Pref = "ICO" Or Pref = "CUR" Then
                    If recursive = False Then
                        vtmp = AddBackSlash(Dir1Path)
                    Else
                        vtmp = ""
                    End If
                    vret = ChargeImage(vtmp & RENAME.ListView1.SelectedItem.Text, RENAME.ListView1.SelectedItem.Text)
                    Set RENAME.Acdsee.Picture = LoadPicture(vtmp & RENAME.ListView1.SelectedItem.Text)
                Else
                    Acdsee2.Clear
                    FViewPict.Caption = RENAME.ListView1.SelectedItem.Text
                End If
            End If
            
        Case 33 ' Page Up
            ind = RENAME.ListView1.SelectedItem.Index
            vnb = RENAME.ListView1.ListItems.Count
            If ind - 1 > 0 Then
                RENAME.ListView1.ListItems(ind).Selected = False
                RENAME.ListView1.ListItems(ind - 1).Selected = True
                Pref = UCase$(Suffixe(RENAME.ListView1.SelectedItem.Text))
                If Pref = "JPG" Or Pref = "BMP" Or Pref = "GIF" Or Pref = "JPEG" Or Pref = "DIB" Or Pref = "WMF" Or Pref = "EMF" Or Pref = "ICO" Or Pref = "CUR" Then
                    If recursive = False Then
                        vtmp = AddBackSlash(Dir1Path)
                    Else
                        vtmp = ""
                    End If
                    vret = ChargeImage(vtmp & RENAME.ListView1.SelectedItem.Text, RENAME.ListView1.SelectedItem.Text)
                    Set RENAME.Acdsee.Picture = LoadPicture(vtmp & RENAME.ListView1.SelectedItem.Text)
                Else
                    Acdsee2.Clear
                    FViewPict.Caption = RENAME.ListView1.SelectedItem.Text
                End If
            End If

        Case 36 ' Home
            ind = RENAME.ListView1.SelectedItem.Index
            RENAME.ListView1.ListItems(ind).Selected = False
            RENAME.ListView1.ListItems(1).Selected = True
            Pref = UCase$(Suffixe(RENAME.ListView1.SelectedItem.Text))
            If Pref = "JPG" Or Pref = "BMP" Or Pref = "GIF" Or Pref = "JPEG" Or Pref = "DIB" Or Pref = "WMF" Or Pref = "EMF" Or Pref = "ICO" Or Pref = "CUR" Then
                If recursive = False Then
                    vtmp = AddBackSlash(Dir1Path)
                Else
                    vtmp = ""
                End If
                vret = ChargeImage(vtmp & RENAME.ListView1.SelectedItem.Text, RENAME.ListView1.SelectedItem.Text)
                Set RENAME.Acdsee.Picture = LoadPicture(vtmp & RENAME.ListView1.SelectedItem.Text)
            Else
                Acdsee2.Clear
                FViewPict.Caption = RENAME.ListView1.SelectedItem.Text
            End If
        
        Case 35 ' End
            ind = RENAME.ListView1.SelectedItem.Index
            RENAME.ListView1.ListItems(ind).Selected = False
            RENAME.ListView1.ListItems(RENAME.ListView1.ListItems.Count).Selected = True
            Pref = UCase$(Suffixe(RENAME.ListView1.SelectedItem.Text))
            If Pref = "JPG" Or Pref = "BMP" Or Pref = "GIF" Or Pref = "JPEG" Or Pref = "DIB" Or Pref = "WMF" Or Pref = "EMF" Or Pref = "ICO" Or Pref = "CUR" Then
                If recursive = False Then
                    vtmp = AddBackSlash(Dir1Path)
                Else
                    vtmp = ""
                End If
                vret = ChargeImage(vtmp & RENAME.ListView1.SelectedItem.Text, RENAME.ListView1.SelectedItem.Text)
                Set RENAME.Acdsee.Picture = LoadPicture(vtmp & RENAME.ListView1.SelectedItem.Text)
            Else
                Acdsee2.Clear
                FViewPict.Caption = RENAME.ListView1.SelectedItem.Text
            End If
        
        
        Case 27 ' Esc
            Unload Me
        
        Case 13 ' Entrée
            chemin = AddBackSlash(Trim$(Dir1Path))
            FileExecutor Me.hwnd, chemin + RENAME.ListView1.ListItems(RENAME.ListView1.SelectedItem.Index), "Open"
        
        Case 46 ' Suppr
    End Select
End Sub
Public Function ChargeImage(NomImage As String, Optional Title As String) As Boolean
    On Error Resume Next
    If Not IsMissing(Title) Then
        Me.Caption = NomImage
    End If
    PictName = NomImage
    Me.MousePointer = vbHourglass
    Set Acdsee2.Picture = LoadPicture(NomImage)
    Acdsee2.Left = 0
    Acdsee2.Top = 0
    Acdsee2.width = Me.ScaleWidth
    Acdsee2.height = Me.ScaleHeight
    Me.MousePointer = vbNormal
    Me.Show vbModal
    ChargeImage = True
End Function
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mfile
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Acdsee2.Visible = False
    Acdsee2.Left = 0
    Acdsee2.Top = 0
    Acdsee2.width = Me.ScaleWidth
    Acdsee2.height = Me.ScaleHeight
    Acdsee2.Visible = True
    Acdsee2.Refresh
End Sub

Private Sub mbestfit_Click()
    Acdsee2.BestFit
    Acdsee2.SetFocus
End Sub

Private Sub mclipboard_Click()
    Acdsee2.CopyToClipboard
End Sub

Private Sub mclose_Click()
    Unload Me
End Sub

Private Sub mfirst_Click()
    SendKeys "{HOME}"
End Sub

Private Sub mlast_Click()
    SendKeys "{END}"
End Sub

Private Sub mnext_Click()
    SendKeys "{PGDN}"
End Sub

Private Sub mprev_Click()
    SendKeys "{PGUP}"
End Sub

Private Sub mprint_Click()
    FViewPict.PrintForm
End Sub
Function SetWallPaper(sFile As String) As Boolean
    SetWallPaper = (SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, sFile, 0) <> 0)
End Function

Private Sub mRealSize_Click()
    Acdsee2.ZoomReal
    Acdsee2.SetFocus
End Sub

Private Sub mstretch_Click()
    Acdsee2.Stretch
    Acdsee2.SetFocus
End Sub

Private Sub mwallpaper_Click()
    SetWallPaper PictName
End Sub

Private Sub mzoomim_Click()
    Acdsee2.ZoomIn
    Acdsee2.SetFocus
End Sub

Private Sub mzoomout_Click()
    Acdsee2.ZoomOut
    Acdsee2.SetFocus
End Sub
