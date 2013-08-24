VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoPositioner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_oAssignments As New Collection



Public Function AddAssignment(ctl As Object, _
                                ctlRelative As Object, _
                                tPosType As tPOSITION_TYPE)
                                
                                
Dim x As New clsAutoPositionerItem
Set x.oCTL = ctl
Set x.oREL = ctlRelative
x.tPosType = tPosType

Select Case x.tPosType
    Case tCONTAINER_RELATIVE_POS_RIGHT:
            x.lValue = x.oREL.Width - x.oCTL.Left
    Case tCONTAINER_RELATIVE_POS_BOTTOM:
            x.lValue = x.oREL.Height - x.oCTL.Top
            
    Case tCONTAINER_WIDTH_DELTA_RIGHT:
            x.lValue = x.oREL.Width - (x.oCTL.Left + x.oCTL.Width)
    Case tCONTAINER_HEIGHT_DELTA_BOTTOM:
            x.lValue = x.oREL.Height - (x.oCTL.Top + x.oCTL.Height)
            
    Case tCONTROL_RELATIVE_SAME_POS_VERTICAL:
            x.lValue = x.oCTL.Left - x.oREL.Left
    Case tCONTROL_RELATIVE_SAME_POS_HORIZONTAL:
            x.lValue = x.oCTL.Top - x.oREL.Top
            
End Select
m_oAssignments.Add x

End Function
Public Function RefreshPositions()
Dim i As Long
Dim x As clsAutoPositionerItem
Dim erg As Long
For i = 1 To m_oAssignments.Count
    Set x = m_oAssignments.Item(i)
    Select Case x.tPosType
        Case tCONTAINER_RELATIVE_POS_RIGHT:
            erg = x.oREL.Width - x.lValue
            If (erg > 0) Then x.oCTL.Left = erg
        Case tCONTAINER_RELATIVE_POS_BOTTOM:
            erg = x.oREL.Height - x.lValue
            If (erg > 0) Then x.oCTL.Top = erg
            
        Case tCONTAINER_WIDTH_DELTA_RIGHT:
            erg = x.oREL.Width - x.oCTL.Left - x.lValue
            If (erg > 0) Then x.oCTL.Width = erg
        Case tCONTAINER_HEIGHT_DELTA_BOTTOM:
            erg = x.oREL.Height - x.oCTL.Top - x.lValue
            If (erg > 0) Then
                x.oCTL.Height = erg
            Else
                erg = erg
            End If
            
        Case tCONTROL_RELATIVE_SAME_POS_VERTICAL:
            erg = x.oREL.Left + x.lValue
            x.oCTL.Left = erg
        Case tCONTROL_RELATIVE_SAME_POS_HORIZONTAL:
            erg = x.oREL.Top + x.lValue
            x.oCTL.Top = erg
    End Select
Next
        







End Function

                                

