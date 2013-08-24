Attribute VB_Name = "Flat"
Public Const WM_USER = &H400
Public Const TB_SETSTYLE = WM_USER + 56
Public Const TB_GETSTYLE = WM_USER + 57
Public Const TBSTYLE_FLAT = &H800

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Function MakeFlat(Tlb As Toolbar)
  Dim style As Long
  Dim hToolbar As Long
  Dim r As Long
  hToolbar = FindWindowEx(Tlb.hWnd, 0&, "ToolbarWindow32", vbNullString)
  style = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
   If style And TBSTYLE_FLAT Then
    style = style Xor TBSTYLE_FLAT
   Else: style = style Or TBSTYLE_FLAT
   End If
   r = SendMessageLong(hToolbar, TB_SETSTYLE, 0, style)
   Tlb.Refresh
End Function
