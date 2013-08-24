Attribute VB_Name = "MOnTop"
'In General Declarations:
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long) As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public Sub OnTop(bOnTop As Boolean)
'When ready, call:
Dim i As Long
If bOnTop = True Then
    i = SetWindowPos(RENAME.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS)
Else
    i = SetWindowPos(RENAME.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS)
End If
End Sub

