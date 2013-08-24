Attribute VB_Name = "ELine"
Option Explicit
'*************************************************************************
'* Function: EtchedLine(frmEtch As Form, ByVal intX1 As Integer, ByVal intY1 As Integer, ByVal intLength As Integer)
'*************************************************************************
'* Description: Draws an 'etched' line upon the specified form starting
'*              at the X,Y location passed in and of the specified length.
'*              Coordinates are in the current ScaleMode of the passed
'*              in form.
'*************************************************************************
'* Parameters: [frmEtch] - form to draw the line upon
'*             [intX1] - starting horizontal of line
'*             [intY1] - starting vertical of line
'*             [intLength] - length of the line
'*************************************************************************
Public Sub EtchedLine(frmEtch As Form, ByVal intX1 As Integer, ByVal intY1 As Integer, ByVal intLength As Integer)
On Error Resume Next
    frmEtch.Line (intX1, intY1)-(intX1 + intLength, intY1), vb3DShadow
    frmEtch.Line (frmEtch.CurrentX + 5, intY1 + 12)-(intX1 - 5, intY1 + 12), vb3DHighlight
End Sub
