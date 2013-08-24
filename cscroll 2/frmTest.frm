VERSION 5.00
Begin VB.Form frmScrollDemo 
   Caption         =   "Scroll Demo 1 - Adds Scroll Bars to a Form"
   ClientHeight    =   5055
   ClientLeft      =   3210
   ClientTop       =   2175
   ClientWidth     =   6615
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
   ScaleHeight     =   5055
   ScaleWidth      =   6615
   Begin VB.TextBox txtDemo 
      Height          =   315
      Index           =   0
      Left            =   1140
      TabIndex        =   5
      Text            =   "TestItem0"
      Top             =   60
      Width           =   1995
   End
   Begin VB.CommandButton cmdPictureTest 
      Caption         =   "&Control Demo"
      Height          =   315
      Left            =   3420
      TabIndex        =   4
      Top             =   540
      Width           =   1335
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Information"
      Height          =   2115
      Left            =   3300
      TabIndex        =   2
      Top             =   1260
      Width           =   1635
      Begin VB.Label lblInfo 
         Caption         =   $"frmTest.frx":0000
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdVirtualTest 
      Caption         =   "&Virtual Demo"
      Height          =   315
      Left            =   3420
      TabIndex        =   1
      Top             =   900
      Width           =   1335
   End
   Begin VB.PictureBox picClient 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   960
      Width           =   1155
   End
   Begin VB.Image imgvbAccel 
      Height          =   330
      Left            =   3450
      Picture         =   "frmTest.frx":0087
      Top             =   105
      Width           =   1275
   End
   Begin VB.Label lblDemo 
      Caption         =   "Demo 0"
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblVBAccel 
      BackColor       =   &H00000066&
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   3300
      TabIndex        =   7
      Top             =   60
      Width           =   1635
   End
End
Attribute VB_Name = "frmScrollDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' frmScrollDemo
' ---------------------------------------------------------------------------
' Copyright ® 1998 Steve McMahon (steve@dogma.demon.co.uk)
' Visit vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ---------------------------------------------------------------------------
'
' Description:
' Demonstrates adding scroll bars to a form.  All the controls on the form
' are added to a picture box, which is moved in response to the scroll bar
' positions, allowing a scrollable viewport.  When both horizontal and
' vertical scroll bars are shown, VB automatically adds a sizing box for
' the form.  Neat!  Note also that the VB properties ScaleHeight and
' ScaleWidth adjust to the size excluding the scroll bars.
' ===========================================================================

Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1

Private Sub cmdPictureTest_Click()
   frmScrollDemo2.Show
End Sub

Private Sub cmdVirtualTest_Click()
   frmScrollDemo3.Show
End Sub

Private Sub Form_Load()
Dim i As Long
Dim ctl As Control

   ' Set up scroll bars:
   Set m_cScroll = New cScrollBars
   m_cScroll.Create Me.hwnd
   
   ' To make it easier to design the form,
   ' we place all the controls on the form,
   ' then switch them into the client box
   ' at run-time.
   On Error Resume Next
   For Each ctl In Controls
      If Not ctl Is picClient Then
         If ctl.Container Is Me Then
            Set ctl.Container = picClient
         End If
      End If
   Next ctl
   
   ' Create something in the viewport:
   picClient.BorderStyle = 0
   For i = 1 To 50
      Load lblDemo(i)
      Load txtDemo(i)
      lblDemo(i).top = lblDemo(i - 1).top + lblDemo(i - 1).Height + 2 * Screen.TwipsPerPixelY
      lblDemo(i).Caption = "Demo" & i
      lblDemo(i).Visible = True
      txtDemo(i).top = lblDemo(i).top
      txtDemo(i).Text = "TestItem" & i
      txtDemo(i).Visible = True
   Next i
   picClient.Move 0, 0, fraInfo.left + fraInfo.Width + 2 * Screen.TwipsPerPixelY, lblDemo(lblDemo.UBound).top + lblDemo(0).Height + 2 * Screen.TwipsPerPixelY
   
End Sub

Private Sub Form_Resize()
Dim lHeight As Long
Dim lWidth As Long
Dim lProportion As Long
   
   ' Pixels are the minimum change size for a screen object.
   ' Therefore we set the scroll bars in pixels.

   lHeight = (picClient.Height - Me.ScaleHeight) \ Screen.TwipsPerPixelY
   If (lHeight > 0) Then
      lProportion = lHeight \ (Me.ScaleHeight \ Screen.TwipsPerPixelY) + 1
      m_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
      m_cScroll.Max(efsVertical) = lHeight
      m_cScroll.Visible(efsVertical) = True
   Else
      m_cScroll.Visible(efsVertical) = False
   End If
   
   lWidth = (picClient.Width - Me.ScaleWidth) \ Screen.TwipsPerPixelX
   If (lWidth > 0) Then
      lProportion = lWidth \ (Me.ScaleWidth \ Screen.TwipsPerPixelX) + 1
      m_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
      m_cScroll.Max(efsHorizontal) = lWidth
      m_cScroll.Visible(efsHorizontal) = True
   Else
      m_cScroll.Visible(efsHorizontal) = False
   End If
   
End Sub

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
   If (m_cScroll.Visible(eBar)) Then
      If (eBar = efsHorizontal) Then
         picClient.left = -m_cScroll.Value(eBar) * Screen.TwipsPerPixelX
      Else
         picClient.top = -m_cScroll.Value(eBar) * Screen.TwipsPerPixelY
      End If
   Else
      picClient.Move 0, 0
   End If
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)
   m_cScroll_Change eBar
End Sub

