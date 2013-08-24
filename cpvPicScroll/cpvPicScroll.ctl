VERSION 5.00
Begin VB.UserControl cpvPicScroll 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4020
   ScaleHeight     =   186
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ToolboxBitmap   =   "cpvPicScroll.ctx":0000
   Begin VB.PictureBox iLoaded 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   1560
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1575
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox iSave 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   3075
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1020
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton btnUC 
      Height          =   195
      Left            =   2415
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2415
      Width           =   195
   End
   Begin VB.HScrollBar hSB 
      Height          =   195
      Left            =   0
      Max             =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2415
      Width           =   2415
   End
   Begin VB.VScrollBar vSB 
      Height          =   2415
      Left            =   2415
      Max             =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   195
   End
   Begin VB.PictureBox iScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   0
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   -15
      Width           =   2340
   End
   Begin VB.Image cRelease 
      Height          =   480
      Left            =   3060
      Picture         =   "cpvPicScroll.ctx":0312
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cGrab 
      Height          =   480
      Left            =   3060
      Picture         =   "cpvPicScroll.ctx":061C
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image iTo 
      Height          =   750
      Left            =   2970
      Top             =   915
      Visible         =   0   'False
      Width           =   750
   End
End
Attribute VB_Name = "cpvPicScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'#  ——————————————————
'#  PicScroll OCX v2.0
'#  Carles P.V. - 2002
'#  ——————————————————
'#  carles_pv@terra.es
'#  ——————————————————

'  05/01/2002
'  ———————————————————————————————————————————————————————————————————
'  Not great changes...
'  1. Use of StretchBlt instead of PaintPicture
'  2. Now, CopyToClipboard "saves" sized image
'  3. Added hDC property and Refresh method (Allows API drawing)
'  4. Event raises fixed
'  5. BestFit fixed
'  6. ... something more
'
'  Obviously, it still can be improved:
'  If anyone knows how to zoom/scroll picture without changing
'  image size...
'  ———————————————————————————————————————————————————————————————————



Option Explicit

'#  Constants
Public Enum AppearanceCts
            [Flat]
            [3D]
End Enum

Public Enum BorderStyleCts
            [None]
            [Fixed Single]
End Enum

Public Enum GoDirectionCts
            [gdTop]
            [gdBottom]
            [gdLeft]
            [gdRight]
            [gdCenter]
            [gdTopLeft]
            [gdTopRight]
            [gdBottomLeft]
            [gdBottomRight]
End Enum

Public Enum ScrollBarCts
            [sbAutomatic]
            [sbNone]
End Enum

Public Enum ScrollDirectionCts
            [sdcUp]
            [sdDown]
            [sdLeft]
            [sdRight]
End Enum



'# Private Zoom variables
Private PT As POINTAPI                      '# Cursor position
Private tmpPt As POINTAPI                   '# Temp. cursor position (anchor point)
Private ScrollingPicture As Boolean         '# flag
Private pKeyPressed As Boolean              '# flag
Private Zoom(0 To 14) As Integer            '# Zoom coeficients array
Private zi As Byte                          '# Zoom array index

'#  Default Property Values:
Private Const m_def_BarsWidth = 13         '# [Pixels]
Private Const m_def_MouseScrolling = True  '# Enable mouse scrolling
Private Const m_def_ScrollBars = 0         '# Automatic mode

'#  Property Variables:
Private m_ScrollBars As ScrollBarCts
Private m_MouseScrolling As Boolean
Private m_BarsWidth As Integer

'#  Event Declarations:
Public Event ButtonClick()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event PictureScrolled()
Public Event PictureSizeChanged()
'#  ===========================================================================
'#  Init/Read/Write properties
'#  ===========================================================================
Private Sub UserControl_InitProperties()

        m_BarsWidth = m_def_BarsWidth
        m_MouseScrolling = m_def_MouseScrolling
        m_ScrollBars = m_def_ScrollBars
        
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        m_BarsWidth = PropBag.ReadProperty("BarsWidth", m_def_BarsWidth)
        m_MouseScrolling = PropBag.ReadProperty("MouseScrolling", m_def_MouseScrolling)
        m_ScrollBars = PropBag.ReadProperty("ScrollBars", m_def_ScrollBars)
        
        UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
        UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
        UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
        UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
        
        iScroll.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
        iLoaded.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
        
        Set Picture = PropBag.ReadProperty("Picture", Nothing)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

        With PropBag
            .WriteProperty "Appearance", UserControl.Appearance, 1
            .WriteProperty "BackColor", iScroll.BackColor, &H8000000F
            .WriteProperty "BackColor", UserControl.BackColor, &H8000000F
            .WriteProperty "BarsWidth", m_BarsWidth, m_def_BarsWidth
            .WriteProperty "BorderStyle", UserControl.BorderStyle, 0
            .WriteProperty "Enabled", UserControl.Enabled, True
            .WriteProperty "MouseScrolling", m_MouseScrolling, m_def_MouseScrolling
            .WriteProperty "Picture", iLoaded, Nothing
            .WriteProperty "ScrollBars", m_ScrollBars, m_def_ScrollBars
        End With
    
End Sub

'#  ===========================================================================
'#  UserControl
'#  ===========================================================================

Private Sub UserControl_Initialize()
            
        Zoom(0) = 5
        Zoom(1) = 10
        Zoom(2) = 20
        Zoom(3) = 30
        Zoom(4) = 40
        Zoom(5) = 50
        Zoom(6) = 60
        Zoom(7) = 70
        Zoom(8) = 80
        Zoom(9) = 90
        Zoom(10) = 100
        Zoom(11) = 125
        Zoom(12) = 150
        Zoom(13) = 175
        Zoom(14) = 200
        
        zi = 10 '#  100% (default)
    
End Sub

Private Sub UserControl_Show()
        ZoomReal
End Sub

Private Sub UserControl_Resize()

        If ScaleWidth < m_BarsWidth * 2 Then UserControl.width = (m_BarsWidth * 2) * 15
        If ScaleHeight < m_BarsWidth * 2 Then UserControl.height = (m_BarsWidth * 2) * 15
            
        Select Case m_ScrollBars
        
            Case 0: '# Automatic
                 If iScroll.width > ScaleWidth And iLoaded <> 0 Then hSB.Visible = True Else hSB.Visible = False
                 If iScroll.height > ScaleHeight And iLoaded <> 0 Then vSB.Visible = True Else vSB.Visible = False
            
            Case 1: '# None
                 hSB.Visible = False
                 vSB.Visible = False
                
        End Select
        
        btnUC.Visible = IIf(hSB.Visible Or vSB.Visible, True, False)
        
        '# Resize & locate scroll bars ...
        hSB.Move 0, ScaleHeight - m_BarsWidth, ScaleWidth - m_BarsWidth, m_BarsWidth
        vSB.Move ScaleWidth - m_BarsWidth, 0, m_BarsWidth, ScaleHeight - m_BarsWidth
        '# ... set max values ...
        hSB.max = IIf(iLoaded <> 0, iScroll.width, hSB.width) - hSB.width
        vSB.max = IIf(iLoaded <> 0, iScroll.height, vSB.height) - vSB.height
        '# ... set LargeChange values ...
        hSB.LargeChange = ScaleWidth
        vSB.LargeChange = ScaleHeight
        '# ... and readjust max values
        If Not hSB.Visible Then vSB.max = vSB.max - m_BarsWidth
        If Not vSB.Visible Then hSB.max = hSB.max - m_BarsWidth
        
        '# Resize & locate user command button ...
        btnUC.Move hSB.width, vSB.height, m_BarsWidth, m_BarsWidth
         
        With iScroll
        
             '# Center picture
             If iLoaded <> 0 Then
                 hSB = hSB.max / 2
                 vSB = vSB.max / 2
             Else
                 .width = ScaleWidth
                 .height = ScaleHeight
                 iLoaded.width = ScaleWidth
                 iLoaded.height = ScaleHeight
             End If
            
             '# Set MousePointer
             If MouseScrolling And (hSB.Visible Or _
                                    vSB.Visible Or _
                                   .width > ScaleWidth Or _
                                   .height > ScaleHeight) Then
               .MousePointer = vbCustom
               .MouseIcon = cRelease
             Else
               .MousePointer = vbDefault
             End If
            
            .Visible = True
        
        End With
        
End Sub

Private Sub UserControl_Terminate()
        Set iLoaded = Nothing
        Set iScroll = Nothing
        Set iTo = Nothing
        Set iSave = Nothing
End Sub

'#  ===========================================================================
'#  ScrollBars
'#  ===========================================================================

Private Sub hSB_GotFocus()
        iScroll.SetFocus
End Sub

Private Sub hSB_Change()
        iScroll.Left = -hSB
        RaiseEvent PictureScrolled
End Sub

Private Sub hSB_Scroll()
        hSB_Change
End Sub

Private Sub vSB_GotFocus()
        iScroll.SetFocus
End Sub

Private Sub vSB_Change()
        iScroll.Top = -vSB
        RaiseEvent PictureScrolled
End Sub

Private Sub vSB_Scroll()
        vSB_Change
End Sub

'#  ===========================================================================
'#  btnUserCommand
'#  ===========================================================================

Private Sub btnUC_Click()
        iScroll.SetFocus
        RaiseEvent ButtonClick
End Sub

'#  ===========================================================================
'#  Properties
'#  ===========================================================================

'#  Appearance ****************************************************************
Public Property Get Appearance() As AppearanceCts
       Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceCts)
       UserControl.Appearance() = New_Appearance
       iScroll.BackColor = BackColor
       iLoaded.BackColor = BackColor
       Refresh
       PropertyChanged "Appearance"
End Property

'#  BackColor *****************************************************************
Public Property Get BackColor() As OLE_COLOR
       BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
       UserControl.BackColor = New_BackColor
       iScroll.BackColor = New_BackColor
       iLoaded.BackColor = New_BackColor
       Refresh
       PropertyChanged "BackColor"
End Property

'#  B a r s W i d t h *********************************************************
Public Property Get BarsWidth() As Integer
       BarsWidth = m_BarsWidth
End Property

Public Property Let BarsWidth(ByVal New_BarsWidth As Integer)
       If New_BarsWidth > 13 Then
          m_BarsWidth = New_BarsWidth
       Else
          m_BarsWidth = 13
       End If
       UserControl_Resize
       PropertyChanged "BarsWidth"
End Property

'#  BorderStyle ***************************************************************
Public Property Get BorderStyle() As BorderStyleCts
       BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleCts)
       UserControl.BorderStyle() = New_BorderStyle
       Refresh
       PropertyChanged "BorderStyle"
End Property

'#  Enabled *******************************************************************
Public Property Get Enabled() As Boolean
       Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
       UserControl.Enabled() = New_Enabled
       PropertyChanged "Enabled"
End Property

'#  MouseScrolling ************************************************************
Public Property Get MouseScrolling() As Boolean
       MouseScrolling = m_MouseScrolling
End Property

Public Property Let MouseScrolling(ByVal New_MouseScrolling As Boolean)
       m_MouseScrolling = New_MouseScrolling
       PropertyChanged "MouseScrolling"
       UserControl_Resize
End Property

'#  Picture *******************************************************************
Public Property Get Picture() As StdPicture
       Set Picture = iLoaded
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
       Set iScroll = Nothing
       hSB.Visible = False
       vSB.Visible = False
       btnUC.Visible = False
       Set iLoaded = New_Picture
       ZoomReal
       PropertyChanged "Picture"
End Property

'#  PictureWidth **************************************************************
Public Property Get PictureWidth() As Integer
       PictureWidth = iLoaded.width
End Property

'#  PictureHeight *************************************************************
Public Property Get PictureHeight() As Integer
       PictureHeight = iLoaded.height
End Property

'#  ScrollBars ****************************************************************
Public Property Get ScrollBars() As ScrollBarCts
       ScrollBars = m_ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarCts)
       m_ScrollBars = New_ScrollBars
       PropertyChanged "ScrollBars"
       UserControl_Resize
End Property

'#  ZoomPercent ***************************************************************
Public Property Get ZoomPercent()
       ZoomPercent = Zoom(zi)
End Property

'#  ===========================================================================
'#  UserControl Events
'#  ===========================================================================

Private Sub iScroll_Click()
        RaiseEvent Click
End Sub

Private Sub iScroll_DblClick()
        RaiseEvent DblClick
End Sub

Private Sub iScroll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = vbLeftButton Then iScroll.MouseIcon = cGrab
        ScrollingPicture = True
        GetCursorPos PT
        tmpPt = PT
        RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub iScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        If (Button <> vbLeftButton) Or _
           (iLoaded = 0) Or _
           (MouseScrolling = False) Or _
           (ScrollingPicture = False) Then Exit Sub
        
        GetCursorPos PT
            
        If iScroll.width > ScaleWidth Then
           If (PT.X - tmpPt.X) > 0 Then
               If hSB - (PT.X - tmpPt.X) > 0 Then
                  hSB = hSB - (PT.X - tmpPt.X)
               Else
                  hSB = 0
               End If
           Else
               If hSB - (PT.X - tmpPt.X) < hSB.max Then
                  hSB = hSB - (PT.X - tmpPt.X)
               Else
                  hSB = hSB.max
               End If
           End If
        End If
        
        If iScroll.height > ScaleHeight Then
           If (PT.Y - tmpPt.Y) > 0 Then
               If vSB - (PT.Y - tmpPt.Y) > 0 Then
                  vSB = vSB - (PT.Y - tmpPt.Y)
               Else
                  vSB = 0
               End If
           Else
               If vSB - (PT.Y - tmpPt.Y) < vSB.max Then
                  vSB = vSB - (PT.Y - tmpPt.Y)
               Else
                  vSB = vSB.max
               End If
           End If
        End If
          
        tmpPt = PT
        
        RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

Private Sub iScroll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        iScroll.MouseIcon = cRelease
        ScrollingPicture = False
        RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub iScroll_KeyDown(KeyCode As Integer, Shift As Integer)
 
        pKeyPressed = True
    
        Select Case KeyCode
               Case 38 '# Up arrow
                    Do
                         DoEvents
                         If vSB > 0 Then vSB = vSB - 1 Else vSB = 0: Exit Do
                    Loop Until pKeyPressed = False
        
               Case 40 '# Down arrow
                    Do
                         DoEvents
                         If vSB < vSB.max Then vSB = vSB + 1 Else vSB = vSB.max: Exit Do
                    Loop Until pKeyPressed = False
        
               Case 37 '# Left arrow
                    Do
                         DoEvents
                         If hSB > 0 Then hSB = hSB - 1 Else hSB = 0: Exit Do
                    Loop Until pKeyPressed = False
        
               Case 39 '# Right arrow
                    Do
                         DoEvents
                         If hSB < hSB.max Then hSB = hSB + 1 Else hSB = hSB.max: Exit Do
                    Loop Until pKeyPressed = False
        End Select
        
        RaiseEvent KeyDown(KeyCode, Shift)
    
End Sub

Private Sub iScroll_KeyPress(KeyAscii As Integer)
        RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub iScroll_KeyUp(KeyCode As Integer, Shift As Integer)
        pKeyPressed = False
        RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'#  ===========================================================================
'#  UserControl Methods
'#  ===========================================================================

'#  CopyToClipboard ***********************************************************
Public Sub CopyToClipboard()

       Clipboard.Clear
        
       With iScroll
            iTo = .Image
            iSave.width = .width
            iSave.height = .height
            iSave.PaintPicture iTo, 0, 0, .ScaleWidth, .ScaleHeight, _
                                    0, 0, .ScaleWidth, .ScaleHeight, vbSrcCopy
       End With
      
       iSave = iSave.Image
       Set iTo = Nothing
       Clipboard.SetData iSave
    
End Sub

'#  PasteFromClipboard ********************************************************
Public Sub PasteFromClipboard()
       Set iLoaded = Clipboard.GetData
       ZoomReal
End Sub

'#  Clear *********************************************************************
Public Sub Clear()
       Set iLoaded = Nothing
       Set iScroll = Nothing
       Set iTo = Nothing
       Set iSave = Nothing
       UserControl_Resize
End Sub

'#  Go ************************************************************************
Public Sub Go(Direction As GoDirectionCts)
    
       iScroll.Visible = False
        
       Select Case Direction
              Case 0: '# goTop
                   hSB = hSB.max / 2
                   vSB = 0
                   
              Case 1: '# goBottom
                   hSB = hSB.max / 2
                   vSB = vSB.max
                   
              Case 2: '# goLeft
                   hSB = 0
                   vSB = vSB.max / 2
                
              Case 3: '# goRight
                   hSB = hSB.max
                   vSB = vSB.max / 2
                
              Case 4: '# goCenter
                   hSB = hSB.max / 2
                   vSB = vSB.max / 2
                
              Case 5: '# goTopLeft
                   hSB = 0
                   vSB = 0
                
              Case 6: '# goTopRight
                   hSB = hSB.max
                   vSB = 0
                
              Case 7: '# goBottomLeft
                   hSB = 0
                   vSB = vSB.max
                
              Case 8: '# goBottomRight
                   hSB = hSB.max
                   vSB = vSB.max
        End Select
        
        iScroll.Visible = True

End Sub

'#  BestFit *******************************************************************
Public Sub BestFit()
     
       If iLoaded = 0 Then Exit Sub
       
       zi = 10             '# Reset zoom to 100%
       Dim cW As Single    '# Width coef.
       Dim cH As Single    '# Height coef.
       Dim W As Integer    '# Final Width
       Dim H As Integer    '# Final Height
        
       If (iLoaded.height <> ScaleHeight) Or (iLoaded.width <> ScaleWidth) Then
        
           cH = ScaleHeight / iLoaded.height
           cW = ScaleWidth / iLoaded.width
            
           If cW < cH Then
              W = ScaleWidth
              H = Int(iLoaded.height * cW)
           Else
              H = ScaleHeight
              W = Int(iLoaded.width * cH)
           End If
            
       Else
        
           W = iLoaded.width
           H = iLoaded.height
            
       End If
        
       ChangeSize W, H
    
End Sub

'#  SaveTo ********************************************************************
Public Sub SaveTo(Path As String)
    
       With iScroll
            iTo = .Image
            iSave.width = .width
            iSave.height = .height
            iSave.PaintPicture iTo, 0, 0, .ScaleWidth, .ScaleHeight, _
                                    0, 0, .ScaleWidth, .ScaleHeight, vbSrcCopy
       End With
        
       iSave = iSave.Image
       Set iTo = Nothing
        
       SavePicture iSave, Path
        
End Sub

'# Scroll *********************************************************************
Public Sub Scroll(Direction As ScrollDirectionCts, Pixels As Integer)
    
       If iLoaded = 0 Then Exit Sub
       If Pixels < 1 Then Err.Raise 380
        
       Select Case Direction
              Case 0: '# scrUp
                   If iScroll.height <= ScaleHeight Then Exit Sub
                   If vSB - Pixels > 0 Then vSB = vSB - Pixels Else vSB = 0
                     
              Case 1: '# scrDown
                   If iScroll.height <= ScaleHeight Then Exit Sub
                   If vSB + Pixels < vSB.max Then vSB = vSB + Pixels Else vSB = vSB.max
                     
              Case 2: '# scrLeft
                   If iScroll.width <= ScaleWidth Then Exit Sub
                   If hSB - Pixels > 0 Then hSB = hSB - Pixels Else hSB = 0
                 
              Case 3: '# scrRight
                   If iScroll.width <= ScaleWidth Then Exit Sub
                   If hSB + Pixels < hSB.max Then hSB = hSB + Pixels Else hSB = hSB.max
       End Select
        
End Sub

'#  ScrollLoop ****************************************************************
Public Sub ScrollLoop(Direction As ScrollDirectionCts, Pixels As Integer)
    
       If iLoaded = 0 Then Exit Sub
       If Pixels < 1 Then Err.Raise 380
       
       Select Case Direction
              Case 0: '# scrUp
                   If iScroll.height <= ScaleHeight Then Exit Sub
                   If vSB - Pixels > 0 Then vSB = vSB - Pixels Else vSB = vSB.max
                    
              Case 1: '# scrDown
                   If iScroll.height <= ScaleHeight Then Exit Sub
                   If vSB + Pixels < vSB.max Then vSB = vSB + Pixels Else vSB = 0
                    
              Case 2: '# scrLeft
                   If iScroll.width <= ScaleWidth Then Exit Sub
                   If hSB - Pixels > 0 Then hSB = hSB - Pixels Else hSB = hSB.max
                
              Case 3: '# scrRight
                   If iScroll.width <= ScaleWidth Then Exit Sub
                   If hSB + Pixels < hSB.max Then hSB = hSB + Pixels Else hSB = 0
       End Select
       
End Sub

'#  Stretch *******************************************************************
Public Sub Stretch()
       If iLoaded = 0 Then Exit Sub
       zi = 10
       ChangeSize ScaleWidth, ScaleHeight
End Sub

'#  ZoomIn ********************************************************************
Public Sub ZoomIn()
       If iLoaded = 0 Then Exit Sub
       If zi < 14 Then
          zi = zi + 1
          ChangeSize iLoaded.width * Zoom(zi) / 100, iLoaded.height * Zoom(zi) / 100
       End If
End Sub

'#  ZoomOut *******************************************************************
Public Sub ZoomOut()
       If iLoaded = 0 Then Exit Sub
       If zi > 0 Then
          zi = zi - 1
          ChangeSize iLoaded.width * Zoom(zi) / 100, iLoaded.height * Zoom(zi) / 100
       End If
End Sub

'#  ZoomReal ******************************************************************
Public Sub ZoomReal()
       If iLoaded = 0 Then Exit Sub
       zi = 10
       ChangeSize iLoaded.width * 1, iLoaded.height * 1
End Sub

'#
'#
'#

'#  ChangeSize ****************************************************************
Private Sub ChangeSize(newWidth As Integer, newHeight As Integer)

        Screen.MousePointer = vbHourglass
        
        With iScroll
        
            .Visible = False
            .width = newWidth
            .height = newHeight
            
             On Error Resume Next
             SetStretchBltMode iScroll.hdc, 3
             StretchBlt .hdc, 0, 0, newWidth, newHeight, iLoaded.hdc, 0, 0, iLoaded.width, iLoaded.height, SRCCOPY
                           
            If Err Then
               Err.Clear
               ZoomReal
            End If
            
        End With
        UserControl_Resize
            
        Screen.MousePointer = vbDefault
        RaiseEvent PictureSizeChanged

End Sub


'# If you want to use it to paint...

'#  hDC ***********************************************************************
Public Property Get hdc() As Long
Attribute hdc.VB_Description = "Devuelve un controlador (de Microsoft Windows) al contexto de dispositivo del objeto."
       hdc = iLoaded.hdc
End Property

'#  RefreshPaints *************************************************************
Public Sub Refresh()
       StretchBlt iScroll.hdc, 0, 0, iScroll.width, iScroll.height, iLoaded.hdc, 0, 0, iLoaded.width, iLoaded.height, SRCCOPY
       iScroll.Refresh
End Sub
