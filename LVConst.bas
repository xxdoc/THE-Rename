Attribute VB_Name = "LVConst"
Option Explicit
DefLng A-Z
Public Const LVM_FIRST As Long = &H1000
Public Const LVFI_PARAM As Long = 1
Public Const LVFI_STRING As Long = &H2
Public Const LVIF_STATE As Long = 8
Public Const LVIF_TEXT As Long = 1
Public Const LVIS_SELECTED As Long = &H2
Public Const LVM_DELETEITEM As Long = (LVM_FIRST + 8)
Public Const LVM_FINDITEM As Long = LVM_FIRST + 13
Public Const LVM_GETITEMSTATE As Long = LVM_FIRST + 44
Public Const LVM_GETITEMTEXT As Long = LVM_FIRST + 45
Public Const LVM_GETNEXTITEM As Long = (LVM_FIRST + 12)
Public Const LVM_GETSELECTEDCOUNT As Long = (LVM_FIRST + 50)
Public Const LVM_SETITEMSTATE As Long = LVM_FIRST + 43
Public Const LVM_SORTITEMS As Long = LVM_FIRST + 48
Public Const LVNI_SELECTED As Long = &H2
