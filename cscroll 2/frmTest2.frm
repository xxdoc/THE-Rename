VERSION 5.00
Begin VB.Form frmScrollDemo2 
   Caption         =   "Scroll Demo 2 - Adds Scroll Bars to a Control"
   ClientHeight    =   3735
   ClientLeft      =   4020
   ClientTop       =   3480
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3735
   ScaleWidth      =   5895
   Begin VB.Frame fraInfo 
      Caption         =   "Information"
      Height          =   3675
      Left            =   4260
      TabIndex        =   2
      Top             =   0
      Width           =   1575
      Begin VB.Label lblInfo 
         Caption         =   "The scroll bars are added to a VB picture box control (picScrollBox) and the client is a child picture box control (picClient)."
         Height          =   1815
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   2100
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmTest2.frx":0000
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.PictureBox picScrollBox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   60
      ScaleHeight     =   3555
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   60
      Width           =   4155
      Begin VB.PictureBox picClient 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   7530
         Left            =   60
         Picture         =   "frmTest2.frx":008A
         ScaleHeight     =   7530
         ScaleWidth      =   11010
         TabIndex        =   1
         Top             =   60
         Width           =   11010
      End
   End
End
Attribute VB_Name = "frmScrollDemo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' frmScrollDemo2
' ---------------------------------------------------------------------------
' Copyright ® 1998 Steve McMahon (steve@dogma.demon.co.uk)
' Visit vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ---------------------------------------------------------------------------
'
' Description:
' Demonstrates adding scroll bars to a picture box.  The client area of
' the picture box is implemented as another picture box, which is moved
' in response to the scroll bar positions.  Note that the VB properties
' ScaleHeight and ScaleWidth adjust to the size excluding the scroll bars.
' ===========================================================================

Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1

Private Sub Form_Load()
   ' Set up scroll bars:
   Set m_cScroll = New cScrollBars
   m_cScroll.Create picScrollBox.hwnd
   ' Initialise client to top,left
   picClient.Move 0, 0
End Sub

Private Sub Form_Resize()
   picScrollBox.Move picScrollBox.left, picScrollBox.top, Me.ScaleWidth - picScrollBox.left * 3 - fraInfo.Width, Me.ScaleHeight - picScrollBox.top * 2
   fraInfo.Move picScrollBox.left * 2 + picScrollBox.Width, fraInfo.top, fraInfo.Width, Me.ScaleHeight - fraInfo.top * 2
End Sub

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
   m_cScroll_Scroll eBar
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)
   If (eBar = efsHorizontal) Then
      picClient.left = -Screen.TwipsPerPixelX * m_cScroll.Value(eBar)
   Else
      picClient.top = -Screen.TwipsPerPixelY * m_cScroll.Value(eBar)
   End If
End Sub

Private Sub picScrollBox_Resize()
Dim lHeight As Long
Dim lWidth As Long
Dim lProportion As Long

   lHeight = (picClient.Height - picScrollBox.ScaleHeight) \ Screen.TwipsPerPixelY
   If (lHeight > 0) Then
      lProportion = lHeight \ (picScrollBox.ScaleHeight \ Screen.TwipsPerPixelY) + 1
      m_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
      m_cScroll.Max(efsVertical) = lHeight
      m_cScroll.Visible(efsVertical) = True
   Else
      m_cScroll.Visible(efsVertical) = False
      picClient.top = 0
   End If
   
   lWidth = (picClient.Width - picScrollBox.ScaleWidth) \ Screen.TwipsPerPixelX
   If (lWidth > 0) Then
      lProportion = lWidth \ (picScrollBox.ScaleWidth \ Screen.TwipsPerPixelX) + 1
      m_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
      m_cScroll.Max(efsHorizontal) = lWidth
      m_cScroll.Visible(efsHorizontal) = True
   Else
      m_cScroll.Visible(efsHorizontal) = False
      picClient.left = 0
   End If

End Sub
