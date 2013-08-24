Attribute VB_Name = "MTypes"
Option Explicit
Public Const MAX_PATH = 260

Public Type POINT
  X As Long
  Y As Long
End Type

Public Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Public Type SYSTEMTIME
  wYear             As Integer
  wMonth            As Integer
  wDayOfWeek        As Integer
  wDay              As Integer
  wHour             As Integer
  wMinute           As Integer
  wSecond           As Integer
  wMilliseconds     As Integer
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes  As Long
  ftCreationTime    As FILETIME
  ftLastAccessTime  As FILETIME
  ftLastWriteTime   As FILETIME
  nFileSizeHigh     As Long
  nFileSizeLow      As Long
  dwReserved0       As Long
  dwReserved1       As Long
  cFileName         As String * MAX_PATH
  cAlternate        As String * 14
End Type


Public Type LV_FINDINFO
  flags As Long
  pSz As String
  lParam As Long
  PT As POINT
  vkDirection As Long
End Type

Public Type LV_ITEM
    Mask As Long
    iItem As Long
    iSubItem As Long
    State As Long
    StateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type

Public Type SHITEMID
  cb      As Long
  abID    As Byte
End Type

Public Type ITEMIDLIST
  mkid    As SHITEMID
End Type
