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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim m_oAssignments As New Collection



Public Function AddAssignment(ctl As Object, _
                                ctlRelative As Object, _
                                tPosType As tPOSITION_TYPE)
                                
                                
Dim X As New clsAutoPositionerItem
Set X.oCTL = ctl
Set X.oREL = ctlRelative
X.tPosType = tPosType

Select Case X.tPosType
    Case tCONTAINER_RELATIVE_POS_RIGHT:
            X.lValue = X.oREL.width - X.oCTL.Left
    Case tCONTAINER_RELATIVE_POS_BOTTOM:
            X.lValue = X.oREL.height - X.oCTL.Top
            
    Case tCONTAINER_WIDTH_DELTA_RIGHT:
            X.lValue = X.oREL.width - (X.oCTL.Left + X.oCTL.width)
    Case tCONTAINER_HEIGHT_DELTA_BOTTOM:
            X.lValue = X.oREL.height - (X.oCTL.Top + X.oCTL.height)
            
    Case tCONTROL_RELATIVE_SAME_POS_VERTICAL:
            X.lValue = X.oCTL.Left - X.oREL.Left
    Case tCONTROL_RELATIVE_SAME_POS_HORIZONTAL:
            X.lValue = X.oCTL.Top - X.oREL.Top
            
End Select
m_oAssignments.Add X

End Function
Public Function RefreshPositions()
Dim i As Long
Dim X As clsAutoPositionerItem
Dim erg As Long
For i = 1 To m_oAssignments.Count
    Set X = m_oAssignments.Item(i)
    Select Case X.tPosType
        Case tCONTAINER_RELATIVE_POS_RIGHT:
            erg = X.oREL.width - X.lValue
            If (erg > 0) Then X.oCTL.Left = erg
        Case tCONTAINER_RELATIVE_POS_BOTTOM:
            erg = X.oREL.height - X.lValue
            If (erg > 0) Then X.oCTL.Top = erg
            
        Case tCONTAINER_WIDTH_DELTA_RIGHT:
            erg = X.oREL.width - X.oCTL.Left - X.lValue
            If (erg > 0) Then X.oCTL.width = erg
        Case tCONTAINER_HEIGHT_DELTA_BOTTOM:
            erg = X.oREL.height - X.oCTL.Top - X.lValue
            If (erg > 0) Then
                X.oCTL.height = erg
            Else
                erg = erg
            End If
            
        Case tCONTROL_RELATIVE_SAME_POS_VERTICAL:
            erg = X.oREL.Left + X.lValue
            X.oCTL.Left = erg
        Case tCONTROL_RELATIVE_SAME_POS_HORIZONTAL:
            erg = X.oREL.Top + X.lValue
            X.oCTL.Top = erg
    End Select
Next
End Function

                                

