VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Brought to you by Brad Martinez
'   http://www.mvps.org/ccrp/
'   news://news.mvps.org/ccrp.foldertreeview

' =================================================
' A fairly rudimentary, yet handy, vertical splitter class.
' =================================================

Private Enum CBoolean
  CFalse = 0
  CTrue = 1
End Enum

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As GWL_nIndex) As Long

Private Const WS_BORDER = &H800000
Private Const WS_EX_CLIENTEDGE = &H200

Private Enum GWL_nIndex
'  GWL_WNDPROC = (-4)
'  GWL_HWNDPARENT = (-8)
'  GWL_ID = (-12)
  GWL_STYLE = (-16)
  GWL_EXSTYLE = (-20)
'  GWL_USERDATA = (-21)
End Enum

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As CBoolean) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long  ' lpPoint As POINTAPI) As Long

Private Const HWND_DESKTOP = 0

Private Type POINTAPI   ' pt
  x As Long
  y As Long
End Type

Private Type RECT   ' rct
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As GetSystemMetrics_codes) As Long
Private Enum GetSystemMetrics_codes
'  SM_CXSCREEN = 0
'  SM_CYSCREEN = 1
'  SM_CXVSCROLL = 2
'  SM_CYHSCROLL = 3
'  SM_CYCAPTION = 4
  SM_CXBORDER = 5
'  SM_CYBORDER = 6
'  SM_CXDLGFRAME = 7
'  SM_CYDLGFRAME = 8
'  SM_CYVTHUMB = 9
'  SM_CXHTHUMB = 10
'  SM_CXICON = 11
'  SM_CYICON = 12
'  SM_CXCURSOR = 13
'  SM_CYCURSOR = 14
'  SM_CYMENU = 15
'  SM_CXFULLSCREEN = 16
'  SM_CYFULLSCREEN = 17
'  SM_CYKANJIWINDOW = 18
'  SM_MOUSEPRESENT = 19
'  SM_CYVSCROLL = 20
'  SM_CXHSCROLL = 21
'  SM_DEBUG = 22
'  SM_SWAPBUTTON = 23
'  SM_RESERVED1 = 24
'  SM_RESERVED2 = 25
'  SM_RESERVED3 = 26
'  SM_RESERVED4 = 27
'  SM_CXMIN = 28
'  SM_CYMIN = 29
'  SM_CXSIZE = 30
'  SM_CYSIZE = 31
  SM_CXFRAME = 32
'  SM_CYFRAME = 33
  SM_CXMINTRACK = 34
'  SM_CYMINTRACK = 35
'  SM_CXDOUBLECLK = 36
'  SM_CYDOUBLECLK = 37
'  SM_CXICONSPACING = 38
'  SM_CYICONSPACING = 39
'  SM_MENUDROPALIGNMENT = 40
'  SM_PENWINDOWS = 41
'  SM_DBCSENABLED = 42
'  SM_CMOUSEBUTTONS = 43
'  SM_CXFIXEDFRAME = SM_CXDLGFRAME           ' win40 name change
'  SM_CYFIXEDFRAME = SM_CYDLGFRAME           ' win40 name change
  SM_CXSIZEFRAME = SM_CXFRAME               ' win40 name change
'  SM_CYSIZEFRAME = SM_CYFRAME               ' win40 name change
'  SM_SECURE = 44
'  SM_CXEDGE = 45
'  SM_CYEDGE = 46
'  SM_CXMINSPACING = 47
'  SM_CYMINSPACING = 48
'  SM_CXSMICON = 49
'  SM_CYSMICON = 50
'  SM_CYSMCAPTION = 51
'  SM_CXSMSIZE = 52
'  SM_CYSMSIZE = 53
'  SM_CXMENUSIZE = 54
'  SM_CYMENUSIZE = 55
'  SM_ARRANGE = 56
'  SM_CXMINIMIZED = 57
'  SM_CYMINIMIZED = 58
'  SM_CXMAXTRACK = 59
'  SM_CYMAXTRACK = 60
'  SM_CXMAXIMIZED = 61
'  SM_CYMAXIMIZED = 62
'  SM_NETWORK = 63
'  SM_CLEANBOOT = 67
'  SM_CXDRAG = 68
'  SM_CYDRAG = 69
'  SM_SHOWSOUNDS = 70
'  SM_CXMENUCHECK = 71 ' Use instead of GetMenuCheckMarkDimensions()!
'  SM_CYMENUCHECK = 72
'  SM_SLOWMACHINE = 73
'  SM_MIDEASTENABLED = 74
'  SM_MOUSEWHEELPRESENT = 75
'  SM_CMETRICS = 76
End Enum

' ==========================================================
' Required in the client for proper splitter functionality:
'
' The Parent (Form, UserControl), with the following default property settings
' changed to:
'    KeyPreview = True
'    ScaleMode = vbPixels
'
' Two child control windows in the parent, both having an hWnd property,
' and similar Appearance and BorderStyle settings.

' A Frame control, with the following default property settings changed to:
'    Appearance = 0   ' flat
'    BorderStyle = vbBSNone
'    MouseIcon = LoadPicture(App.Path & "\VSplit.cur")
'
' ==========================================================

' The minimum allowable window width of the controls, in pixels.
Private m_cxMinChild As Long

' The splitter bar's offset, in pixels.
Private m_cxOffset As Long

' Splitter bar's fixed width, in pixels
Private m_cxSplit As Long

' Splitter bar's current x position, relative to the
' parent form's client coords, in pixels
Private m_xSplit As Long

' The parent's current ScaleWidth and ScaleHeight, in pixels
Private m_cxClient As Long
Private m_cyClient As Long

Private m_fTrackSplit As Boolean
Private m_xSplitStart As Long

' The desktop's hDC. Used by PatBlt to draw the splitter bar, and is
' used as a flag indicating when the spilitter is currently being drawn.
Private m_hdcDesktop As Long

Private m_hwndParent As Long
Private m_Ctrl1 As Control
Private m_Ctrl2 As Control
Private m_fraSplit As Frame
'

Private Sub Class_Terminate()
  If m_hdcDesktop Then
    Call ReleaseDC(HWND_DESKTOP, m_hdcDesktop)
  End If
  Set m_Ctrl1 = Nothing
  Set m_Ctrl2 = Nothing
  Set m_fraSplit = Nothing
End Sub

Private Function HasHWnd(ctrl As Control) As Boolean
  On Error GoTo Out
  HasHWnd = ctrl.hWnd
Out:
  If Err Then MsgBox ctrl & " must be a window (have a window handle)."
End Function

' ================================================================
' Public members

Public Sub SetControls(hwndParent As Long, Ctrl1 As Control, Ctrl2 As Control, fraSplit As Frame)
  m_hwndParent = 0
  Set m_Ctrl1 = Nothing
  Set m_Ctrl2 = Nothing
  Set m_fraSplit = Nothing
  
  If HasHWnd(Ctrl1) Then
    If HasHWnd(Ctrl2) Then
    
      m_hwndParent = hwndParent
      Set m_Ctrl1 = Ctrl1
      Set m_Ctrl2 = Ctrl2
      Set m_fraSplit = fraSplit
  
      ' Setup the splitter frame control. It serves 2 purposes here:
      '   - Shows the splitter cursor when the mouse is over it
      '   - Allows the splitter cursor to appear in a 2 pixel wider area than
      '     the distance betwen the two controls.
      With m_fraSplit
    '    .Appearance = 0   ' flat
        .BackColor = vbActiveBorder
    '    .BorderStyle = vbBSNone
        .MousePointer = vbSizeWE
        .ZOrder   ' put the frame on top
      End With

      Call FillSplitterMetrics
    
    End If
  End If
    
End Sub

Public Property Get Left() As Long
  Left = m_xSplit
End Property

Public Property Let Left(x As Long)
  m_fraSplit.Left = x
  Call ParentResize(True)
End Property

Public Property Get Width() As Byte
  Width = m_cxSplit
End Property

Public Property Let Width(cx As Byte)
  If cx Then
    m_cxSplit = cx
    Call ParentResize(True)
  End If
End Property

Public Sub ParentResize(fUpdateCtrls As Boolean)
  Dim rc As RECT
  On Error GoTo Out
  If m_hwndParent Then
    Call GetClientRect(m_hwndParent, rc)
    m_cxClient = rc.Right - rc.Left
    m_cyClient = rc.Bottom - rc.Top
    ' Calculate the splitter bar's left coordinate
    m_xSplit = Min(Max((m_cxMinChild), m_fraSplit.Left), m_cxClient - (m_cxSplit + m_cxMinChild))
    Call MoveControls(fUpdateCtrls)
  End If
Out:
End Sub

Public Property Get Splitting() As Boolean
  Splitting = m_hdcDesktop
End Property

Public Property Get TrackSplit() As Boolean
  TrackSplit = m_fTrackSplit
End Property

Public Property Let TrackSplit(fTrackSplit As Boolean)
  m_fTrackSplit = fTrackSplit
End Property

' Initiates splitter bar dragging.
' the X and Y params are the frame's internal clent coords, in TWIPS!!!

Public Sub BeginSplit(x As Single, y As Single)
  
  If m_hwndParent Then
  
    ' Get the desktop's DC (also used as a flag indicating we're currently splitting)
    m_hdcDesktop = GetDC(HWND_DESKTOP)
    If m_hdcDesktop Then
          
      ' Calculate the splitter bar's initial position. The splitter's left coordinate
      ' will never be less than the first control's minimum control width (-1 pixel to
      ' cover the first control's right border), and never be more than the second control's
      ' left coordinate (minus the splitter's width + 1 pixel to cover the second control's
      ' left border). There's a 1 pixel slop when the form is at it's minimum width...
      m_xSplit = Min(Max((m_cxMinChild), m_fraSplit.Left + (x \ Screen.TwipsPerPixelX)), _
                             m_cxClient - (m_cxSplit + m_cxMinChild))
      
      If m_fTrackSplit Then
        m_xSplitStart = m_xSplit
      Else
        ' Draw the splitter bar on top of everything.
        Call DrawSplitter(m_xSplit, 0, m_cxSplit, m_cyClient)
      End If
      
    End If   ' m_hdcDesktop
  End If   ' m_hwndParent
  
End Sub

' Moves the splitter bar
' the X and Y params are the frame's internal clent coords, in TWIPS!!!

Public Sub MoveSplitter(x As Single, y As Single)
  Dim cx As Long
  
  If m_hdcDesktop Then
    
    If (m_fTrackSplit = False) Then
      ' Erase the previous splitter bar.
      Call DrawSplitter(m_xSplit, 0, m_cxSplit, m_cyClient)
    End If
    
    ' Calculate the splitter bar's new position.
    m_xSplit = Min(Max((m_cxMinChild), m_fraSplit.Left + (x \ Screen.TwipsPerPixelX)), _
                           m_cxClient - (m_cxSplit + m_cxMinChild))
  
    If m_fTrackSplit Then
      Call MoveControls(True)
    Else
      ' Draw the splitter bar's new position.
      Call DrawSplitter(m_xSplit, 0, m_cxSplit, m_cyClient)
    End If
    
  End If
End Sub

' Finishes splitter bar moving and resizes the form's child windows.
' the X and Y params are the frame's internal clent coords, in TWIPS!!!

Public Sub EndSplit(x As Single, y As Single, fUpdateCtrls As Boolean)
  Dim wSplitRight As Long
  
  If m_hdcDesktop Then
    
    If (m_fTrackSplit = False) Then
      ' Erase the previous bar.
      Call DrawSplitter(m_xSplit, 0, m_cxSplit, m_cyClient)
    End If
    
    ' Release and clear the desktop's hDC variable.
    Call ReleaseDC(HWND_DESKTOP, m_hdcDesktop)
    m_hdcDesktop = 0
    
    ' Calculate the splitter's final position.
    m_xSplit = Min(Max((m_cxMinChild), m_fraSplit.Left + (x \ Screen.TwipsPerPixelX)), _
                           m_cxClient - (m_cxSplit + m_cxMinChild))
    
    ' Move the controls to their new positions.
    Call MoveControls(fUpdateCtrls)
    
  End If
End Sub

' Cancels any current splitting operation.

Public Sub CancelSplit()

  If m_hdcDesktop Then

    ' Remove the splitter pointer.
    Call ReleaseCapture
    
    ' Restore the splitter's original position.
    If m_fTrackSplit Then
      m_xSplit = m_xSplitStart
      Call MoveControls(True)
    Else
      ' Erase the previous splitter bar.
      Call DrawSplitter(m_xSplit, 0, m_cxSplit, m_cyClient)
    End If
    
    ' Release and clear the desktop's hDC variable.
    Call ReleaseDC(HWND_DESKTOP, m_hdcDesktop)
    m_hdcDesktop = 0
  
  End If

End Sub

' ================================================================
' Private calls

' Sets the splitter's width to match the window sizing border width (which will be a
' width of 2 WS_BORDER's greater than the distance between the controls. Also
' sets a minimum width for the controls.

' A top level window's default minimum width (SM_CXMINTRACK) minus the widths
' of it's left and right sizing borders (SM_CXSIZEFRAME), leaves a constant minimum
' client width of 104 pixels.

Private Sub FillSplitterMetrics()
  Dim cxSizeFrame As Long
  Dim cxBorder As Long
  Dim cxMinTrack As Long
  Dim dwStyle As Long
  
  If m_hwndParent Then
    
    ' Get the top level window sizing border width (default is 4 pixels)
    cxSizeFrame = GetSystemMetrics(SM_CXSIZEFRAME)
    
    ' Get the window border width (WS_BORDER, default is 1 pixel)
    cxBorder = GetSystemMetrics(SM_CXBORDER)
  
    ' Set the splitter bar's width (the top level window sizing border
    ' width minus the window border width, default is 3 pixels)
    If (m_cxSplit = 0) Then
      m_cxSplit = cxSizeFrame - cxBorder
    End If
    
    ' Set the control offset. The splitter will cover a border width
    ' of the each control only when they have both of their
    ' WS_BORDER and WS_EX_CLIENTEDGE style bits set.
    ' (Appearance and BorderStyle properties respectively).
    ' We'll just check the first control, and assume the second
    ' is set the same...
    If ((GetWindowLong(m_Ctrl1.hWnd, GWL_STYLE) And WS_BORDER) = WS_BORDER) And _
       ((GetWindowLong(m_Ctrl1.hWnd, GWL_EXSTYLE) And WS_EX_CLIENTEDGE) = WS_EX_CLIENTEDGE) Then
  '  If (m_Ctrl1.Appearance = 3D) And (m_Ctrl1.BorderStyle = FixedSingle) Then
      m_cxOffset = cxBorder
    End If
    
   ' Get the top level window minimum tracking width.
    cxMinTrack = GetSystemMetrics(SM_CXMINTRACK)
    
    ' Calculate the control's minimum horizontal tracking width.
    ' (usually around 50 pixels)
    m_cxMinChild = (cxMinTrack - ((cxSizeFrame * 2) + m_cxSplit)) \ 2
  
  End If   ' m_hwndParent
  
End Sub

Private Sub DrawSplitter(x As Long, y As Long, cx As Long, cy As Long)
  Dim pt As POINTAPI
  pt.x = x
  pt.y = y
  Call ClientToScreen(m_hwndParent, pt)
  Call PatBlt(m_hdcDesktop, pt.x, pt.y, cx, cy, vbDstInvert)
End Sub

Private Sub MoveControls(fUpdateCtrls As Boolean)
  Static xSplitPrev As Long
  Dim wSplitRight As Long
  On Error GoTo Out

  ' ======================================================
  ' Resize the first control.
  
  ' If the splitter's new position is less than it's previous position...
  If m_xSplit < xSplitPrev Then
    ' Resize m_Ctrl1 first and repaint it if specified.
    m_Ctrl1.Move 0, 0, m_xSplit + m_cxOffset, m_cyClient
    If fUpdateCtrls Then Call UpdateWindow(m_Ctrl1.hWnd)
  Else
    ' Resize m_Ctrl2 first and repaint it if specified.
    wSplitRight = m_xSplit + m_cxSplit
    m_Ctrl2.Move wSplitRight - m_cxOffset, 0, (m_cxClient - wSplitRight) + m_cxOffset, m_cyClient
    If fUpdateCtrls Then Call UpdateWindow(m_Ctrl2.hWnd)
  End If
    
  ' ======================================================
  ' Move the splitter bar (frame).
  
  ' Make the frame invisible before the move, or the frame's
  ' shadow may be left on the controls when splitting.
  m_fraSplit.Visible = False
  
  ' Center the frame between the two controls making it the full
  ' height of the parent. The frame's Move method is really slow...
'  m_fraSplit.Move m_xSplit, 0, m_cxSplit, m_cyClient
  Call MoveWindow(m_fraSplit.hWnd, m_xSplit, 0, m_cxSplit, m_cyClient, CTrue)
  
  ' Make the frame visible again and repaint it if specified.
  m_fraSplit.Visible = True
  If fUpdateCtrls Then Call UpdateWindow(m_fraSplit.hWnd)

  ' ======================================================
  ' Resize the second control.
  
  If m_xSplit < xSplitPrev Then
    ' Resize m_Ctrl2 and repaint it if specified.
    wSplitRight = m_xSplit + m_cxSplit
    m_Ctrl2.Move wSplitRight - m_cxOffset, 0, (m_cxClient - wSplitRight) + m_cxOffset, m_cyClient
    If fUpdateCtrls Then Call UpdateWindow(m_Ctrl2.hWnd)
  Else
    ' Resize m_Ctrl1 and repaint it if specified.
    m_Ctrl1.Move 0, 0, m_xSplit + m_cxOffset, m_cyClient
    If fUpdateCtrls Then Call UpdateWindow(m_Ctrl1.hWnd)
  End If
  
Out:
  ' Save the splitter's current position.
  xSplitPrev = m_xSplit
  
End Sub

' Returns the larger of the two passed params

Private Function Max(param1 As Long, param2 As Long) As Long
  If param1 > param2 Then Max = param1 Else Max = param2
End Function

' Returns the smaller of the two passed params

Private Function Min(param1 As Long, param2 As Long) As Long
  If param1 < param2 Then Min = param1 Else Min = param2
End Function
