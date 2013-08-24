VERSION 5.00
Begin VB.Form frmScrollDemo3 
   Caption         =   "Virtual Control Scroll Bar Demonstration"
   ClientHeight    =   4260
   ClientLeft      =   2250
   ClientTop       =   2190
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
   ScaleHeight     =   4260
   ScaleWidth      =   5895
   Begin VB.PictureBox picVirtualGrid 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
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
      TabIndex        =   2
      Top             =   60
      Width           =   4155
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Information"
      Height          =   4215
      Left            =   4200
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      Begin VB.Label lblInfo 
         Caption         =   $"frmTest3.frx":0000
         Height          =   1995
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   2100
         Width           =   1395
      End
      Begin VB.Label lblInfo 
         Caption         =   $"frmTest3.frx":009E
         Height          =   1815
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmScrollDemo3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' frmScrollDemo3
' ---------------------------------------------------------------------------
' Copyright ® 1998 Steve McMahon (steve@dogma.demon.co.uk)
' Visit vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ---------------------------------------------------------------------------
'
' Description:
' Demonstrates adding scroll bars to a picture box.  The client area of
' the picture box is drawn directly, and the view area is moved
' in response to the scroll bar positions.  Note that the VB properties
' ScaleHeight and ScaleWidth adjust to the size excluding the scroll bars.
' ===========================================================================

Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1

' Virtual "grid":
Private m_iCols As Long
Private m_iRows As Long
Private m_iColWidth As Long
Private m_iRowHeight As Long

Private Sub DrawGrid()
Dim lCol As Long
Dim lRow As Long
Dim lStartCol As Long
Dim lX As Long
Dim lStartX As Long
Dim lY As Long

   '
   ' NOTE: This grid will need *some* work if you want it to work
   ' properly!  You will need to eliminate the flicker by drawing
   ' rows onto a hidden picture box and then using PaintPicture to
   ' load them into the view.
   ' Use API calls rather than VB drawing code to improve speed.
   '
   With picVirtualGrid
      ' Erase backdrop:
      picVirtualGrid.Line (0, 0)-(.ScaleWidth, .ScaleHeight), .BackColor, BF
      ' Draw the grid:
      lCol = 1
      lRow = 1
      If (m_cScroll.Visible(efsHorizontal)) Then
         lX = -m_cScroll.Value(efsHorizontal) * Screen.TwipsPerPixelX
      End If
      If (m_cScroll.Visible(efsVertical)) Then
         lY = -m_cScroll.Value(efsVertical) * Screen.TwipsPerPixelY
      End If
      lStartX = lX
      Do
         If (lY + m_iRowHeight > 0) Then
            Do
               If (lX + m_iColWidth > 0) Then
                  If (lStartCol = 0) Then
                     lStartCol = lCol
                     lStartX = lX
                  End If
                  picVirtualGrid.Line (lX, lY)-(lX + m_iColWidth, lY + m_iRowHeight), &HC0C0C0, B
                  picVirtualGrid.CurrentX = lX + 3 * Screen.TwipsPerPixelX
                  picVirtualGrid.CurrentY = lY + Screen.TwipsPerPixelY
                  picVirtualGrid.Print "Row:" & lRow & ",Col:" & lCol
               End If
               lCol = lCol + 1
               lX = lX + m_iColWidth
            Loop While lCol <= m_iCols And lX < .ScaleWidth
            lCol = lStartCol
            lX = lStartX
         End If
         lRow = lRow + 1
         lY = lY + m_iRowHeight
      Loop While lRow <= m_iRows And lY < .ScaleHeight
   End With
End Sub

Private Sub Form_Load()
   ' Set up scroll bars:
   Set m_cScroll = New cScrollBars
   m_cScroll.Create picVirtualGrid.hwnd
   ' Set up the grid:
   m_iRows = 512
   m_iCols = 9
   m_iColWidth = 84 * Screen.TwipsPerPixelX
   m_iRowHeight = 16 * Screen.TwipsPerPixelY
   m_cScroll.SmallChange(efsHorizontal) = 48
   m_cScroll.SmallChange(efsVertical) = 16
   picVirtualGrid_Resize
End Sub

Private Sub Form_Resize()
   picVirtualGrid.Move picVirtualGrid.left, picVirtualGrid.top, Me.ScaleWidth - picVirtualGrid.left * 3 - fraInfo.Width, Me.ScaleHeight - picVirtualGrid.top * 2
   fraInfo.Move picVirtualGrid.left * 2 + picVirtualGrid.Width, fraInfo.top, fraInfo.Width, Me.ScaleHeight - fraInfo.top * 2
End Sub

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
   m_cScroll_Scroll eBar
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)
   DrawGrid
End Sub

Private Sub picVirtualGrid_Paint()
   DrawGrid
End Sub

Private Sub picVirtualGrid_Resize()
Dim lHeight As Long
Dim lWidth As Long
Dim lProportion As Long
   
   ' Pixels are the minimum change size for a screen object.
   ' Therefore we set the scroll bars in pixels.
   
   lHeight = (m_iRows * m_iRowHeight - picVirtualGrid.ScaleHeight) \ Screen.TwipsPerPixelY
   If (lHeight > 0) Then
      lProportion = lHeight \ (picVirtualGrid.ScaleHeight \ Screen.TwipsPerPixelY) + 1
      m_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
      m_cScroll.Max(efsVertical) = lHeight
      m_cScroll.Visible(efsVertical) = True
   Else
      m_cScroll.Visible(efsVertical) = False
   End If
   
   lWidth = (m_iCols * m_iColWidth - picVirtualGrid.ScaleWidth) \ Screen.TwipsPerPixelX
   If (lWidth > 0) Then
      lProportion = lWidth \ (picVirtualGrid.ScaleWidth \ Screen.TwipsPerPixelX) + 1
      m_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
      m_cScroll.Max(efsHorizontal) = lWidth
      m_cScroll.Visible(efsHorizontal) = True
   Else
      m_cScroll.Visible(efsHorizontal) = False
   End If

End Sub
