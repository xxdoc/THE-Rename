VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAutoPositionerItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Enum tPOSITION_TYPE
    tCONTAINER_RELATIVE_POS_RIGHT
    tCONTAINER_RELATIVE_POS_BOTTOM
    tCONTAINER_WIDTH_DELTA_RIGHT
    tCONTAINER_HEIGHT_DELTA_BOTTOM
    tCONTROL_RELATIVE_SAME_POS_VERTICAL
    tCONTROL_RELATIVE_SAME_POS_HORIZONTAL
End Enum


Public oCTL As Object
Public oREL As Object
Public tPosType As tPOSITION_TYPE
Public lValue As tPOSITION_TYPE
