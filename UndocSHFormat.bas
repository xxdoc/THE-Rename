Attribute VB_Name = "UndocSHFormat"
Option Explicit
Private Type OSVERSIONINFO
      dwOSVersionInfoSize As Long
      dwMajorVersion As Long
      dwMinorVersion As Long
      dwBuildNumber As Long
      dwPlatformId As Long
      szCSDVersion As String * 128
End Type

  '------------------------------------------------------
  'Determines if the current OS is WinNT.
  'Tested in the form load event only
   Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
   'Const VER_PLATFORM_WIN32s = 0
   'Const VER_PLATFORM_WIN32_WINDOWS = 1
   Private Const VER_PLATFORM_WIN32_NT = 2
  '------------------------------------------------------
  'APIs declarations for the format function

  'Normally, these below would be declared as constants,
  'if the values between Win95/98 and NT were the same (see
  'the commented-out const values below).
   Public SHFD_FORMAT_QUICK As Long
   Public SHFD_FORMAT_FULL As Long
   Public SHFD_FORMAT_SYSONLY As Long
  
  'However, because VB lacks a conditional #If variable
  'to distinguish between Win95/98 and NT, they are DIMmed
  'here and set in the Form_Load routine after GetVersionEx
  'has been called.
      
  '------------------------------------------------------
  'support APIs to populate the combo box
  'with the available drives
  
   Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
   Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
 
   Public Const DRIVE_REMOVABLE = 2
   Public Const DRIVE_FIXED = 3
   Public Const DRIVE_REMOTE = 4
   Public Const DRIVE_CDROM = 5
   Public Const DRIVE_RAMDISK = 6
Public Function IsWinNT() As Boolean
 'Returns True if the current operating
 'system is WinNT
   Dim osvi As OSVERSIONINFO
   osvi.dwOSVersionInfoSize = Len(osvi)
   GetVersionEx osvi
   IsWinNT = (osvi.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

