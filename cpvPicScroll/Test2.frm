VERSION 5.00
Object = "*\AcpvPicScroll.vbp"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Drawing"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5235
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ZoomOut 
      Caption         =   "Zoom out"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4080
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1575
      Width           =   975
   End
   Begin VB.CommandButton ZoomIn 
      Caption         =   "Zoom in"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4080
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1140
      Width           =   975
   End
   Begin PicScroll.cpvPicScroll cpvPicScroll1 
      Height          =   3660
      Left            =   195
      TabIndex        =   1
      Top             =   195
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   6456
      Appearance      =   0
      BackColor       =   8421504
      BackColor       =   8421504
      BorderStyle     =   1
      Picture         =   "Test2.frx":0000
      ScrollBars      =   1
   End
   Begin VB.CommandButton Draw 
      Caption         =   "Draw something"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4080
      TabIndex        =   0
      Top             =   195
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Drawing code extracted from:
    
    
    'KPD-Team 2000
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net


Option Explicit

Const PS_DOT = 2
Const PS_SOLID = 0
Const RGN_AND = 1
Const RGN_COPY = 5
Const RGN_OR = 2
Const RGN_XOR = 3
Const RGN_DIFF = 4
Const HS_DIAGCROSS = 5

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type

Private Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long



Private Sub Draw_Click()

    Dim hHBr As Long, R As RECT, hFRgn As Long, hRRgn As Long, hRPen As Long, LP As LOGPEN
    Dim hFFBrush As Long, mIcon As Long, Cnt As Long
    'Set the rectangle's values
    SetRect R, 0, 0, cpvPicScroll1.PictureWidth, cpvPicScroll1.PictureHeight
    'Create a new brush
    hHBr = CreateHatchBrush(HS_DIAGCROSS, vbRed)
    'Draw a frame
    FrameRect cpvPicScroll1.hdc, R, hHBr
    'Draw a rounded rectangle
    hFRgn = CreateRoundRectRgn(0, 0, cpvPicScroll1.PictureWidth, cpvPicScroll1.PictureHeight, (cpvPicScroll1.PictureWidth / 3) * 2, (cpvPicScroll1.PictureHeight / 3) * 5)
    'Draw a frame
    FrameRgn cpvPicScroll1.hdc, hFRgn, hHBr, cpvPicScroll1.PictureWidth, cpvPicScroll1.PictureHeight
    'Invert a region
    InvertRgn cpvPicScroll1.hdc, hFRgn
    'Move our region
    OffsetRgn hFRgn, 10, 10
    'Create a new region
    hRRgn = CreateRectRgnIndirect(R)
    'Combine our two regions
    CombineRgn hRRgn, hFRgn, hRRgn, RGN_XOR
    'Draw a frame
    FrameRgn cpvPicScroll1.hdc, hRRgn, hHBr, cpvPicScroll1.PictureWidth, cpvPicScroll1.PictureHeight
    'Crete a new pen
    hRPen = CreatePen(PS_SOLID, 5, vbBlue)
    'Select our pen into the form's device context and delete the old pen
    DeleteObject SelectObject(cpvPicScroll1.hdc, hRPen)
    'Draw a rectangle
    Rectangle cpvPicScroll1.hdc, cpvPicScroll1.PictureWidth / 2 - 25, cpvPicScroll1.PictureHeight / 2 - 25, cpvPicScroll1.PictureWidth / 2 + 25, cpvPicScroll1.PictureHeight / 2 + 25
    'Delete our pen
    DeleteObject hRPen
    LP.lopnStyle = PS_DOT
    LP.lopnColor = vbGreen
    'Create a new pen
    hRPen = CreatePenIndirect(LP)
    'Select our pen into the form's device context
    SelectObject cpvPicScroll1.hdc, hRPen
    'Draw a rounded rectangle
    RoundRect cpvPicScroll1.hdc, cpvPicScroll1.PictureWidth / 2 - 25, cpvPicScroll1.PictureHeight / 2 - 25, cpvPicScroll1.PictureWidth / 2 + 25, cpvPicScroll1.PictureHeight / 2 + 25, 50, 50
    'Create a new solid brush
    hFFBrush = CreateSolidBrush(vbYellow)
    'Select this brush into our form's device context
    SelectObject cpvPicScroll1.hdc, hFFBrush
    'Floodfill our form
    FloodFill cpvPicScroll1.hdc, cpvPicScroll1.PictureWidth / 2, cpvPicScroll1.PictureHeight / 2, vbBlue
    'Delete our brush
    DeleteObject hFFBrush
    'Create a new solid brush
    hFFBrush = CreateSolidBrush(vbMagenta)
    'Select our solid brush into our form's device context
    SelectObject cpvPicScroll1.hdc, hFFBrush
    'Clean up
    DeleteObject hFFBrush
    DeleteObject hRPen
    DeleteObject hRRgn
    DeleteObject hFRgn
    DeleteObject hHBr
    
    '**** Refresh
    cpvPicScroll1.Refresh

End Sub




Private Sub ZoomIn_Click()
        cpvPicScroll1.ZoomIn
End Sub

Private Sub ZoomOut_Click()
        cpvPicScroll1.ZoomOut
End Sub
