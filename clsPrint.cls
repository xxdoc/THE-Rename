VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Program Name          clsPrint
' Program Description   Class module which enables printing to file

Option Explicit

' Variable to hold original printer and port (from Win.ini)
Private strPrinterHold  As String
    
    ' API used to read current printer and port
    Private Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
    
    ' API used to set current printer and port
    Private Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String) As Integer
    
    ' API used to notify Windows when printer and port is changed
    Private Declare Function PostMessageByString Lib "User" Alias "PostMessage" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer

' Constants used by PostMessage
Private Const HWND_BROADCAST = -1
Private Const WM_WININICHANGE = &H1A
    

Public Sub PrintToFile()
                        
    On Local Error Resume Next
    
    ' Hold the current printer and port
    strPrinterHold = PrinterDefault
    
    ' Do the necessary Win.INI mumbo jumbo
    ' to make FILE: the current port
    Dim strHold As String
    
    ' Keeping default printer, just gonna change port
    strHold = Left(strPrinterHold, _
                   PosOfLast(",", strPrinterHold)) _
            & "FILE:"
    PrinterChange strHold
    DoEvents
    
End Sub

Public Sub PrinterRestorePort()
    
    ' DESCRIPTION:
        ' Uses the printer & port held when PrintToFile
        ' was called to restore the original printer
        ' and port for future printing.
    
    DoEvents
    PrinterChange strPrinterHold
    
End Sub

Private Function PrinterDefault() As String
    
    ' DESCRIPTION:
        ' Function which returns the default Windows printer including
        ' the "real" printer name and port being used.
        
    On Local Error Resume Next
    
    Dim intReturnValue          As Integer
    Dim strDefaultPrinter       As String
    
    strDefaultPrinter = String$(256, 0)
    intReturnValue = GetProfileString("windows", _
                                      "device", _
                                      "", _
                                      strDefaultPrinter, _
                                      2047)
    PrinterDefault = strDefaultPrinter

End Function

Private Function PrinterChange( _
                         ByVal PrinterIDNameAndPort As String) _
                               As Boolean
    
    ' DESCRIPTION:
        ' Function which updates the system to the new printer selection
        ' and notifies all Windows apps that the printer has been changed
        
    On Local Error Resume Next
    
    Dim intReturnValue          As Integer
    Dim lngReturnValue          As Long
    
    PrinterChange = True
    If Len(PrinterIDNameAndPort) > 0 Then
        intReturnValue = WriteProfileString("windows", _
                                            "device", _
                                            PrinterIDNameAndPort)
            
        DoEvents
        ' Notify all applications that printer has changed
        lngReturnValue = PostMessageByString(HWND_BROADCAST, _
                                             WM_WININICHANGE, _
                                             0, _
                                             "windows")
        DoEvents
    Else
        PrinterChange = False
    End If
    DoEvents
    
End Function

Private Function PosOfLast( _
                     ByVal strCharacter As String, _
                     ByVal strInString As String) _
                           As Integer
    
    ' DESCRIPTION:
        ' This function returns the last position within the string of the
        ' specified character(s).
        
    On Local Error Resume Next
    
    Dim intPosition         As Integer
    Dim intPosAcc           As Integer
    
    Do
        If InStr(strInString, strCharacter) = 0 Then Exit Do
        intPosition = InStr(strInString, strCharacter)
        intPosAcc = intPosAcc + intPosition
        strInString = Right(strInString, Len(strInString) - intPosition)
    Loop
    
    PosOfLast = intPosAcc

End Function


