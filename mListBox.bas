Attribute VB_Name = "mListBox"
Option Explicit

Public Const LB_GETHORIZONTALEXTENT = &H193
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const DT_CALCRECT = &H400
Public Const SM_CXVSCROLL = 2

'Public Type RECT
'   Left As Long
'   Top As Long
'   Right As Long
'   Bottom As Long
'End Type

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
