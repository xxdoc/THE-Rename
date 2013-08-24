VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDTPickerCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------------------
'   Name:           CDTPickerCtl (CDTPICKERCTL.CLS)
'   Type:           Control wrapper class
'   Description:    Wraps the DateTime (DT) Picker common control
'
'   Author:         Klaus H. Probst [kprobst@vbbox.com]
'   URL:            http://www.vbbox.com/
'   Copyright:      This work is copyright © 1998-2001, Klaus H. Probst
'   Usage:          You may use this code as you see fit, provided that you assume all
'                   responsibilities for doing so.
'   Distribution:   If you intend to distribute the file(s) that make up this sample to
'                   any WWW site, online service, electronic bulletin board system (BBS),
'                   CD or any other electronic or physical media, you must notify me in
'                   advance to obtain my express permission.
'
'   Dependencies:
'
'       (bas)       General declarations, CommCtrls module and the defs module for this
'                   control class
'
'
'--------------------------------------------------------------------------------------------
Option Explicit
DefLng A-Z

'///////////////////////////////////////////////////////////////////////////////////////
'// Private data
'//
Private m_hWnd As Long              '// Handle to the wrapped control
Private m_hFont As Long             '// Rendering font handle
Private m_lplf As LOGFONT           '// Used to maintain the HFONT
Private m_fOwned As Boolean         '// Owned vs. attached flag


'///////////////////////////////////////////////////////////////////////////////////////
'// Public Implementation (control-specific)
'//

'///////////////////////////////////////////////////////////////////////////////////////
'// Methods
'//

'   <SDK>
'   DWORD DateTimePick_GetSystemtime(HWND m_hWnd, LPSYSTEMTIME pst)
'       returns GDT_NONE if "none" is selected (DTS_SHOWNONE only)
'       returns GDT_VALID and modifies *pst to be the currently selected value
'   </SDK>
'
'   You can also use the friendlier CurSel() properties
'
Friend Function GetSystemTime(ByRef ust As SYSTEMTIME)
    
    GetSystemTime = SendMessage(m_hWnd, DTM_GETSYSTEMTIME, 0, ust)

End Function

'   <SDK>
'   BOOL SetSystemtime(HWND m_hWnd, DWORD gd, LPSYSTEMTIME pst)
'       if gd==GDT_NONE, sets datetimepick to None (DTS_SHOWNONE only)
'       if gd==GDT_VALID, sets datetimepick to *pst
'       returns TRUE on success, FALSE on error (such as bad params)
'   </SDK>
'
Friend Function SetSystemTime(ByVal gd As Long, ByRef ust As SYSTEMTIME)
    
    SetSystemTime = SendMessage(m_hWnd, DTM_SETSYSTEMTIME, gd, ust)

End Function

'   <SDK>
'   BOOL SetFormat(HWND m_hWnd, LPCTSTR sz)
'       sets the display formatting string to sz (see GetDateFormat and GetTimeFormat for valid formatting chars)
'       NOTE: 'X' is a valid formatting character which indicates that the application
'       will determine how to display information. Such apps must support DTN_WMKEYDOWN,
'       DTN_FORMAT, and DTN_FORMATQUERY.
'
'   Possible formats
'   ----------------------------------------------------------------------
'      "d"    The one- or two-digit day.
'      "dd"   The two-digit day. Single-digit day values are preceded by a zero.
'      "ddd"  The three-character weekday abbreviation.
'      "dddd" The full weekday name.
'      "h"    The one- or two-digit hour in 12-hour format.
'      "hh"   The two-digit hour in 12-hour format. Single-digit values are preceded by a zero.
'      "H"    The one- or two-digit hour in 24-hour format.
'      "HH"   The two-digit hour in 24-hour format. Single-digit values are preceded by a zero.
'      "m"    The one- or two-digit minute.
'      "mm"   The two-digit minute. Single-digit values are preceded by a zero.
'      "M"    The one- or two-digit month number.
'      "MM"   The two-digit month number. Single-digit values are preceded by a zero.
'      "MMM"  The three-character month abbreviation.
'      "MMMM" The full month name.
'      "t"    The one-letter AM/PM abbreviation (that is, AM is displayed as "A").
'      "tt"   The two-letter AM/PM abbreviation (that is, AM is displayed as "AM").
'      "yy"   The last two digits of the year (that is, 1996 would be displayed as "96").
'      "yyyy" The full year (that is, 1996 would be displayed as "1996").
'
'   </SDK>
'
Friend Function SetFormat(ByVal Format As String) As Long
    
    SetFormat = SendMessage(m_hWnd, DTM_SETFORMAT, 0, ByVal Format)

End Function

'   <SDK>
'   DWORD GetRange(HWND m_hWnd, LPSYSTEMTIME rgst)
'       modifies rgst[0] to be the minimum ALLOWABLE systemtime (or 0 if no minimum)
'       modifies rgst[1] to be the maximum ALLOWABLE systemtime (or 0 if no maximum)
'       returns GDTR_MIN|GDTR_MAX if there is a minimum|maximum limit
'   </SDK>
'
Friend Function GetRange(ByRef aust() As SYSTEMTIME)
    
    GetRange = SendMessage(m_hWnd, DTM_GETRANGE, 0, aust(0))

End Function

'   <SDK>
'   BOOL SetRange(HWND m_hWnd, DWORD gdtr, LPSYSTEMTIME rgst)
'       if GDTR_MIN, sets the minimum ALLOWABLE systemtime to rgst[0], otherwise removes minimum
'       if GDTR_MAX, sets the maximum ALLOWABLE systemtime to rgst[1], otherwise removes maximum
'       returns TRUE on success, FALSE on error (such as invalid parameters)
'   </SDK>
'
Friend Function SetRange(ByVal ArraySize As Long, aust() As SYSTEMTIME)
    
    SetRange = SendMessage(m_hWnd, DTM_SETRANGE, ArraySize, aust(0))

End Function


'///////////////////////////////////////////////////////////////////////////////////////
'// Properties
'//


'   <SDK>
'   HWND GetMonthCal(HWND m_hWnd)
'        returns the HWND of the MonthCal popup window. Only valid
'        between DTN_DROPDOWN and DTN_CLOSEUP notifications.
'   </SDK>
'
'   You can also use and instance of the CMonthCalCtl class and
'   attach it to the HWND returned by this method
'
Friend Property Get MonthCal()
    
    MonthCal = SendMessage(m_hWnd, DTM_GETMONTHCAL, 0, ByVal 0&)

End Property

'
'   Set/return the internal MonthCal color scheme
'
Friend Property Let MonthCalColor(ByVal Color As MCC_ColorCodes, ByVal rValue As Long)
    
    Call SendMessage(m_hWnd, DTM_SETMCCOLOR, Color, ByVal rValue)

End Property

Friend Property Get MonthCalColor(ByVal Color As MCC_ColorCodes) As Long
    
    MonthCalColor = SendMessage(m_hWnd, DTM_GETMCCOLOR, Color, ByVal 0&)

End Property

'
'   Sets/returns the font used by the MonthCal dropdown.
'   The passed value must be an HFONT handle.
'
Friend Property Let MonthCalFont(ByVal rValue As Long)
    
    Call SendMessage(m_hWnd, DTM_SETMCFONT, rValue, ByVal 1)

End Property

Friend Property Get MonthCalFont()
    
    MonthCalFont = SendMessage(m_hWnd, DTM_GETMCFONT, 0, ByVal 0&)

End Property

'
'   Friendly prop wrapper for setting the control's format.
'   See the SetFormat() prototype above for information on
'   how the different format strings work. Sorry, no l-value
'
Friend Property Let Format(ByVal rValue As String)

    Call SendMessage(m_hWnd, DTM_SETFORMAT, 0, ByVal rValue)

End Property


'
'   Friendly prop wrappers to set/get the current
'   date. The l-value proto will recognize the passed
'   value as being not a date (pass Null or Empty or
'   whatever) and will tell the control to set the current
'   date to GDT_NONE (if the control supports that by having
'   the DTS_SHOWNONE style). Accordingly, The r-value proto will
'   return an empty variant if the current selection is GDT_NONE.
'
Friend Property Get CurSel() As Variant

    Dim ust As SYSTEMTIME
    Dim vDate As Variant
    Dim lResult As Long
        
    lResult = SendMessage(m_hWnd, DTM_GETSYSTEMTIME, 0, ust)
    If (lResult = GDT_VALID) Then
        With ust
            
            '// See if we're a DTS_TIMEFORMAT control and return only
            '// the time. Otherwise return everything
            If (GetWindowLong(m_hWnd, GWL_STYLE) And DTS_TIMEFORMAT) Then
                vDate = TimeSerial(.wHour, .wMinute, .wSecond)
            Else
                vDate = DateSerial(.wYear, .wMonth, .wDay) & " " & TimeSerial(.wHour, .wMinute, .wSecond)
            End If
            
        End With
    
    Else '// No selection (or an error)
        vDate = Empty

    End If
    
    CurSel = vDate

End Property

Friend Property Let CurSel(ByVal rValue As Variant)

    Dim ust As SYSTEMTIME

    If (IsDate(rValue)) Then
        With ust
        
            If (GetWindowLong(m_hWnd, GWL_STYLE) And DTS_TIMEFORMAT) Then
                .wHour = Hour(rValue)
                .wMinute = Minute(rValue)
                .wSecond = Second(rValue)
            Else
                '// Could fill these - doesn't matter. The control ignores them anyway
                .wYear = Year(rValue)
                .wMonth = Month(rValue)
                .wDay = Day(rValue)
            End If
            
        End With

        Call SendMessage(m_hWnd, DTM_SETSYSTEMTIME, GDT_VALID, ust)
    
    Else
        
        Call SendMessage(m_hWnd, DTM_SETSYSTEMTIME, GDT_NONE, ByVal 0&)
    
    End If


End Property



'///////////////////////////////////////////////////////////////////////////////////////
'// Public Implementation (stock Common control)
'//

'///////////////////////////////////////////////////////////////////////////////////////
'// Methods
'//

'
'   I assume this is the same as the toolbar's TB_SETPARENT message,
'   where a new HWND is specified for reception of the control's
'   WM_NOTIFY messages. The message is in COMMCTRL.H but is not
'   documented anywhere in MSDN. I don't know what the return value
'   is supposed to be, but I'd guess it either returns the previous
'   HWND or TRUE/FALSE on success or failure.
'
Friend Function SetNotifyWindow(ByVal hWndNew As Long) As Long

    SetNotifyWindow = SendMessage(m_hWnd, CCM_SETNOTIFYWINDOW, hWndNew, ByVal 0&)

End Function


'///////////////////////////////////////////////////////////////////////////////////////
'// Properties
'//

'
'   Set/return the expected behavior version of individual controls.
'   Note that this is supported only in version 5 or higer, and
'   is set only for the control receiving the message.
'   This message is especially useful for custom draw handling
'   that involves changing rendering fonts on the fly.
'
Friend Property Get CCVersion() As Long
    
    CCVersion = SendMessage(m_hWnd, CCM_GETVERSION, 0, ByVal 0&)

End Property

Friend Property Let CCVersion(ByVal rValue As Long)
    
    Call SendMessage(m_hWnd, CCM_SETVERSION, rValue, ByVal 0&)

End Property

'
'   Set/return the Unicode format the control uses. This message
'   is only useful for Win9x systems with 4.71 or higher installed.
'
Friend Property Get CCUnicodeFormat() As Boolean

    CCUnicodeFormat = (SendMessage(m_hWnd, CCM_GETUNICODEFORMAT, 0, ByVal 0&) <> 0)

End Property
Friend Property Let CCUnicodeFormat(ByVal rValue As Boolean)

    Call SendMessage(m_hWnd, CCM_SETUNICODEFORMAT, Abs(rValue), ByVal 0&)

End Property



'///////////////////////////////////////////////////////////////////////////////////////
'// Public Implementation (stock window/object/font)
'//

'///////////////////////////////////////////////////////////////////////////////////////
'// Properties
'//

'
'   Sets/returns the face name used to render
'   the control's font.
'
Friend Property Let FontName(ByVal rValue As String)

    If (Len(rValue) > 0) Then
        m_lplf.lfFaceName = ""
        Call lstrcpyA(m_lplf.lfFaceName, rValue)
        If (m_hFont <> 0) Then Call SetFont
    
    End If

End Property

Friend Property Get FontName() As String
    
    FontName = Left$(m_lplf.lfFaceName, lstrlenA(m_lplf.lfFaceName))

End Property

'
'   Sets/returns the size of the font used by the
'   control (in standard points)
'
Friend Property Let FontSize(ByVal rValue As Long)

    If ((rValue >= 4) And (rValue <= 30)) Then
        m_lplf.lfHeight = rValue
        If m_hFont <> 0 Then Call SetFont

    End If

End Property

Friend Property Get FontSize() As Long
    
    FontSize = m_lplf.lfHeight

End Property


'
'   Sets/returns whether or not the control's
'   font will be rendered in a bold ("heavy")
'   style.
'
Friend Property Let FontBold(ByVal rValue As Boolean)

    If rValue = True Then
        m_lplf.lfWeight = FW_BOLD
    Else
        m_lplf.lfWeight = FW_NORMAL
    End If

    If m_hFont <> 0 Then Call SetFont

End Property

Friend Property Get FontBold() As Boolean
    
    FontBold = (m_lplf.lfWeight <= FW_BOLD)

End Property


'
'   Sets/returns whether or not the control's
'   font will be rendered in an italized
'   style.
'
Friend Property Let FontItalic(ByVal rValue As Boolean)
    
    m_lplf.lfItalic = Abs(rValue)
    If m_hFont <> 0 Then Call SetFont

End Property

Friend Property Get FontItalic() As Boolean
    
    FontItalic = (m_lplf.lfItalic = 1)

End Property

'
'   Returns the windows handle of the
'   wrapped control.
'
Friend Property Get hWnd() As Long
    
    hWnd = m_hWnd

End Property

'
'   Enables/disables the wrapped control
'
Friend Property Get Enabled() As Boolean
    
    Enabled = (IsWindowEnabled(m_hWnd) = 1)

End Property

Friend Property Let Enabled(ByVal rValue As Boolean)
    
    Call EnableWindow(m_hWnd, Abs(rValue))

End Property

'
'   Shows/hides the wrapped control
'
Friend Property Let Visible(ByVal rValue As Boolean)
    
    If (rValue = True) Then
        Call ShowWindow(m_hWnd, SW_SHOW)
    Else
        Call ShowWindow(m_hWnd, SW_HIDE)
    End If

End Property

Friend Property Get Visible() As Boolean
    
    Visible = (IsWindowVisible(m_hWnd) <> 0)

End Property

'///////////////////////////////////////////////////////////////////////////////////////
'// Methods
'//

'
'   Destroys the wrapped control handle. While the
'   class does this automatically when it is terminated,
'   it's always a good idea to call this method explicitly
'   when you no longer use the control.
'
Friend Sub Destroy()

    If (m_fOwned And (m_hWnd <> 0)) Then
        Call DestroyWindow(m_hWnd)
        Call DeleteObject(m_hFont)
        m_hFont = 0
        m_hWnd = 0
        m_fOwned = False
    End If

End Sub

'
'   Attaches this instance of the class wrapper
'   to an existing control.
'
Friend Sub Attach(ByVal hWndNew As Long)

    Debug.Assert IsWindow(hWndNew)

    Call Destroy
    
    m_hWnd = hWndNew
    m_fOwned = False

End Sub

'
'   Repositions the wrapped control within its
'   parent window.
'
Friend Sub Move(ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, Optional ByVal Repaint As Boolean = True)
    
    Call MoveWindow(m_hWnd, x, y, cx, cy, Abs(Repaint))

End Sub

'
'   Sets the input focus to the wrapped control
'
Friend Sub SetFocus()

    Call win32_SetFocus(m_hWnd)

End Sub


'
'   Creates a new control of the type wrapped by this class.
'   In most cases, the hWndParent must be specified, along with
'   at least the WS_CHILD and WS_VISIBLE styles.
'
Friend Function Create(ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal hWndParent As Long, ByVal Style As Long, Optional ByVal ExStyle As Long = 0) As Long

    '// Sanity check
    Debug.Assert (m_hWnd = 0)
    
    If (m_hWnd = 0) Then
    
        m_hWnd = CreateWindowEx(ExStyle, DATETIMEPICK_CLASS, vbNullString, _
                                Style, _
                                x, y, cx, cy, _
                                hWndParent, 0, App.hInstance, ByVal 0&)
    
        If (m_hWnd <> 0) Then
            m_fOwned = True
            Call SetFont
            Create = m_hWnd
    
        End If
    
    End If

End Function


'///////////////////////////////////////////////////////////////////////////////////////
'// Private Implementation (stock window/object/font)
'//

'
'   Used internally to apply changes made
'   to the control's font.
'
Private Sub SetFont()

    Dim pDC As Long
    Dim lOldSize As Long
    
    If m_hWnd <> 0 Then
    
        If m_hFont <> 0 Then _
            Call DeleteObject(m_hFont)
        
        With m_lplf
            .lfCharSet = DEFAULT_CHARSET
            pDC = GetDC(m_hWnd)
            If pDC Then
                lOldSize = .lfHeight
                .lfHeight = -MulDiv(.lfHeight, GetDeviceCaps(pDC, LOGPIXELSY), 72)
                Call ReleaseDC(m_hWnd, pDC)
            End If
        End With
        
        m_hFont = CreateFontIndirect(m_lplf)
        
        If (m_hFont <> 0) Then
            Call SendMessage(m_hWnd, WM_SETFONT, m_hFont, ByVal 1&)
        End If
        
        m_lplf.lfHeight = lOldSize
    
    End If

End Sub


'///////////////////////////////////////////////////////////////////////////////////////
'// Construction/Destruction
'//

Private Sub Class_Initialize()

    Call CommCtl_Initialize(ICC_DATE_CLASSES)
    With m_lplf
        .lfFaceName = ""
        Call lstrcpyA(.lfFaceName, "MS Sans Serif")
        .lfHeight = 8
    End With

End Sub

Private Sub Class_Terminate()

    Call Destroy

End Sub


