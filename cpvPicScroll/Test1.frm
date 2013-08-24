VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "*\A..\..\..\..\..\ARCHIV~1\MICROS~1\VBCARL~1\CONTROLS\CPVPIC~1\cpvPicScroll.vbp"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PicScroll OCX 2.0 demo"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Drawing 
      Caption         =   "Drawing..."
      Height          =   390
      Left            =   6000
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   5880
      Width           =   975
   End
   Begin PicScroll.cpvPicScroll cpvPicScroll1 
      Height          =   4050
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7144
      BorderStyle     =   1
      Picture         =   "Test1.frx":0000
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   6465
      Top             =   5025
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load..."
      Height          =   390
      Left            =   135
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4605
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   390
      Left            =   1110
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4605
      Width           =   975
   End
   Begin VB.CommandButton Stretch 
      Caption         =   "Stretch"
      Height          =   390
      Left            =   3120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5115
      Width           =   975
   End
   Begin VB.ComboBox cmbScrollBars 
      Height          =   300
      ItemData        =   "Test1.frx":001C
      Left            =   5265
      List            =   "Test1.frx":0026
      Style           =   2  'Dropdown List
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   405
      Width           =   1725
   End
   Begin VB.CheckBox chkMouseScrolling 
      Caption         =   "MouseScrolling"
      Height          =   195
      Left            =   5265
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   795
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Go to"
      Height          =   1650
      Left            =   5265
      TabIndex        =   13
      Top             =   1140
      Width           =   1725
      Begin VB.CommandButton Go 
         Caption         =   "á"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   660
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   255
         Width           =   420
      End
      Begin VB.CommandButton Go 
         Caption         =   "â"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   660
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1095
         Width           =   420
      End
      Begin VB.CommandButton Go 
         Caption         =   "ß"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   240
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   675
         Width           =   420
      End
      Begin VB.CommandButton Go 
         Caption         =   "à"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   1080
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   675
         Width           =   420
      End
      Begin VB.CommandButton Go 
         Caption         =   "w"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   660
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   675
         Width           =   420
      End
      Begin VB.CommandButton Go 
         Caption         =   "ã"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   5
         Left            =   240
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   255
         Width           =   420
      End
      Begin VB.CommandButton Go 
         Caption         =   "ä"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   6
         Left            =   1080
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   255
         Width           =   420
      End
      Begin VB.CommandButton Go 
         Caption         =   "å"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   7
         Left            =   240
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1095
         Width           =   420
      End
      Begin VB.CommandButton Go 
         Caption         =   "æ"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   8
         Left            =   1080
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1095
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Scroll [cursor keys]"
      Height          =   1650
      Left            =   5265
      TabIndex        =   8
      Top             =   2805
      Width           =   1725
      Begin VB.CommandButton Scroll 
         Caption         =   "á"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   660
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   255
         Width           =   420
      End
      Begin VB.CommandButton Scroll 
         Caption         =   "ß"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   675
         Width           =   420
      End
      Begin VB.CommandButton Scroll 
         Caption         =   "à"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   1080
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   675
         Width           =   420
      End
      Begin VB.CommandButton Scroll 
         Caption         =   "â"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   660
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1095
         Width           =   420
      End
   End
   Begin VB.CommandButton ZoomIn 
      Caption         =   "Zoom in"
      Height          =   390
      Left            =   135
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5115
      Width           =   975
   End
   Begin VB.CommandButton ZoomOut 
      Caption         =   "Zoom out"
      Height          =   390
      Left            =   1110
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5115
      Width           =   975
   End
   Begin VB.CommandButton ZoomReal 
      Caption         =   "100%"
      Height          =   390
      Left            =   2085
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5115
      Width           =   975
   End
   Begin VB.CommandButton FeedBest 
      Caption         =   "BestFit"
      Height          =   390
      Left            =   4095
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5115
      Width           =   975
   End
   Begin VB.CommandButton SaveTo 
      Caption         =   "Save to..."
      Height          =   390
      Left            =   4095
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton PasteFromClipboard 
      Caption         =   "Paste from Clipboard"
      Height          =   390
      Left            =   2085
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1950
   End
   Begin VB.CommandButton CopyToClipboard 
      Caption         =   "Copy to Clipboard"
      Height          =   390
      Left            =   135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1950
   End
   Begin VB.Label Label1 
      Caption         =   "ScrollBars"
      Height          =   210
      Left            =   5265
      TabIndex        =   33
      Top             =   195
      Width           =   1125
   End
   Begin VB.Label lblZoomFactor 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   4440
      TabIndex        =   32
      Top             =   4590
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Zoom percent :"
      Height          =   180
      Left            =   2940
      TabIndex        =   31
      Top             =   4590
      Width           =   1140
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   3810
      TabIndex        =   30
      Top             =   4815
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Picture size :"
      Height          =   180
      Left            =   2940
      TabIndex        =   29
      Top             =   4815
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "PicScroll OCX 2.0"
      Height          =   180
      Left            =   135
      TabIndex        =   28
      Top             =   195
      Width           =   1650
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   135
      X2              =   6975
      Y1              =   5745
      Y2              =   5745
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   135
      X2              =   6975
      Y1              =   5760
      Y2              =   5760
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ——————————————————————————
' cpvPicScroll OCX 2.0  Demo
' by Carles P.V., 2001
' ——————————————————————————
' E-mail carles_pv@terra.es
' ——————————————————————————






Dim ScrollKeyPressed As Boolean



Private Sub Form_Load()

    cmbScrollBars.ListIndex = cpvPicScroll1.ScrollBars
    chkMouseScrolling = IIf(cpvPicScroll1.MouseScrolling, 1, 0)
    
End Sub

Private Sub cmdLoad_Click()

    CommonDialog.Filter = "*.bmp;*.jpg|*.bmp;*.jpg"
    CommonDialog.DefaultExt = "*.bmp"
    
    CommonDialog.Flags = cdlOFNHideReadOnly Or _
                         cdlOFNPathMustExist Or _
                         cdlOFNOverwritePrompt Or _
                         cdlOFNNoReadOnlyReturn
                         
    CommonDialog.DialogTitle = "Select a file"
    CommonDialog.CancelError = True
    
    On Error GoTo CancelOpen
    CommonDialog.ShowOpen
    
    DoEvents
    Set cpvPicScroll1.Picture = LoadPicture(CommonDialog.FileName)
    
CancelOpen:
    
End Sub

Private Sub cmdClear_Click()
    cpvPicScroll1.Clear
    lblZoomFactor = ""
    lblSize = ""
End Sub

Private Sub cmbScrollBars_Click()
    cpvPicScroll1.ScrollBars = cmbScrollBars.ListIndex
End Sub

Private Sub chkMouseScrolling_Click()
    cpvPicScroll1.MouseScrolling = chkMouseScrolling
End Sub

Private Sub cpvPicScroll1_ButtonClick()
    MsgBox "Hi!, 'btnUserCommand' clicked."
End Sub

Private Sub Go_Click(Index As Integer)
    Select Case Index
        Case 0: 'Top
            cpvPicScroll1.Go gdTop
        Case 1: 'Bottom
            cpvPicScroll1.Go gdBottom
        Case 2: 'Left
            cpvPicScroll1.Go gdLeft
        Case 3: 'Right
            cpvPicScroll1.Go gdRight
        Case 4: 'Center
            cpvPicScroll1.Go gdCenter
        Case 5: 'Top-Left
            cpvPicScroll1.Go gdTopLeft
        Case 6: 'Top-Right
            cpvPicScroll1.Go gdTopRight
        Case 7: 'Bottom-Left
            cpvPicScroll1.Go gdBottomLeft
        Case 8: 'Bottom-Right
            cpvPicScroll1.Go gdBottomRight
    End Select
End Sub

Private Sub Scroll_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ScrollKeyPressed = True
    Select Case Index
        Case 0: 'scrUp
            Do
                cpvPicScroll1.Scroll sdcUp, 1
                DoEvents
            Loop Until ScrollKeyPressed = False
        Case 1: 'scrDown
            Do
                cpvPicScroll1.Scroll sdDown, 1
                DoEvents
            Loop Until ScrollKeyPressed = False
        Case 2: 'scrLeft
            Do
                cpvPicScroll1.Scroll sdLeft, 1
                DoEvents
            Loop Until ScrollKeyPressed = False
        Case 3: 'scrRight
            Do
                cpvPicScroll1.Scroll sdRight, 1
                DoEvents
            Loop Until ScrollKeyPressed = False
    End Select
End Sub

Private Sub Scroll_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    ScrollKeyPressed = False
End Sub

Private Sub ZoomIn_Click()
    cpvPicScroll1.ZoomIn
End Sub

Private Sub ZoomOut_Click()
    cpvPicScroll1.ZoomOut
End Sub

Private Sub ZoomReal_Click()
    cpvPicScroll1.ZoomReal
End Sub

Private Sub Stretch_Click()
    cpvPicScroll1.Stretch
    lblZoomFactor = ""
End Sub

Private Sub FeedBest_Click()
    cpvPicScroll1.BestFit
    lblZoomFactor = ""
End Sub

Private Sub cpvPicScroll1_PictureSizeChanged()
    lblZoomFactor = cpvPicScroll1.ZoomPercent & " %"
    lblSize = Round(cpvPicScroll1.PictureWidth * cpvPicScroll1.ZoomPercent / 100) & _
              " x " & _
              Round(cpvPicScroll1.PictureHeight * cpvPicScroll1.ZoomPercent / 100)
End Sub

' ******************************************************************************************

Private Sub SaveTo_Click()

    CommonDialog.Filter = "*.bmp|*.bmp"
    CommonDialog.Flags = cdlOFNHideReadOnly Or _
                         cdlOFNPathMustExist Or _
                         cdlOFNOverwritePrompt Or _
                         cdlOFNNoReadOnlyReturn
                         
    CommonDialog.DialogTitle = "Save pre-bitmap as"
    CommonDialog.CancelError = True
    
    On Error GoTo CancelSave
    CommonDialog.ShowSave
    
    cpvPicScroll1.SaveTo CommonDialog.FileName
    
CancelSave:

End Sub

Private Sub CopyToClipboard_Click()
    cpvPicScroll1.CopyToClipboard
End Sub

Private Sub PasteFromClipboard_Click()
    cpvPicScroll1.PasteFromClipboard
End Sub

' ******************************************************************************************

Private Sub Drawing_Click()
    Form2.Show vbModal
End Sub
