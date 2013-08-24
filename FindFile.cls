VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFindFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
'*********************************************************************
' FindFile.cls - Encapsulates the Win32 FindFile functions
'*********************************************************************
Option Explicit
' Attribute constants which differ from VB
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_NORMAL = &H80
' Win32 API constants required by FindFile
Private Const INVALID_HANDLE_VALUE = -1
' Win32 data types (or structs) required by FindFile
Private Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        FileSizeHigh As Long
        FileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

'*********************************************************************
' Win32 API calls required by this class
'*********************************************************************
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'*********************************************************************
' clsFindFiles private member variables
'*********************************************************************
Private mlngFile As Long
Private mstrDateFormat As String
Private mstrUnknownDateText As String
Private mwfdFindData As WIN32_FIND_DATA

'*********************************************************************
' Public interface for setting the format string used for dates
'*********************************************************************
Public Property Let Dateformat(strDateFormat As String)
    mstrDateFormat = strDateFormat
End Property

'*********************************************************************
' Public interface for setting the string used when the date for a
' file is unknown
'*********************************************************************
Public Property Let UnknownDateText(strUnknownDateText As String)
    mstrUnknownDateText = strUnknownDateText
End Property
'*********************************************************************
' Returns the file attributes for the current file
'*********************************************************************
Public Property Get FileAttributes() As Long
    If mlngFile Then FileAttributes = mwfdFindData.dwFileAttributes
End Property
'*********************************************************************
' Returns true if the compress bit is set for the current file
'*********************************************************************
Public Property Get IsCompressed() As Boolean
    If mlngFile Then IsCompressed = mwfdFindData.dwFileAttributes And FILE_ATTRIBUTE_COMPRESSED
End Property
'*********************************************************************
' Returns the value of the Normal attribute bit for dwFileAttributes
'*********************************************************************
Public Property Get NormalAttribute() As Long
   NormalAttribute = FILE_ATTRIBUTE_NORMAL
End Property
Public Function FileSize() As Long
    If mlngFile Then FileSize = mwfdFindData.FileSizeHigh + mwfdFindData.FileSizeLow
End Function

'*********************************************************************
' Primary method in this class for finding the FIRST matching file in
' a directory that matches the path &|or pattern in strFile
'*********************************************************************
Public Function Find(strFile As String, Optional blnShowError As Boolean) As String
    ' If you are already searching, then end the current search
    If mlngFile Then
        If blnShowError Then
            If MsgBox("Cancel the current search?", vbYesNo Or vbQuestion) = vbNo Then Exit Function
        End If
        ' Call cleanup routines before beginning new search
        EndFind
    End If
    ' Find the first file matching the search pattern in strFile
    mlngFile = FindFirstFile(strFile, mwfdFindData)
    ' Check to see if FindFirstFile failed
    If mlngFile = INVALID_HANDLE_VALUE Then
        mlngFile = 0
        ' If blnShowError, then display a default error message
        If blnShowError Then
 '           MsgBox strFile & " could not be found!", vbExclamation
        ' Otherwise raise a user-defined error with a default err msg
        Else
'            Err.Raise vbObjectError + 5000, "clsFindFile_Find", strFile & " could not be found!"
        End If
        Exit Function
    End If
    ' Return the found filename without any nulls
    Find = Left$(mwfdFindData.cFileName, InStr(mwfdFindData.cFileName, Chr$(0)) - 1)
End Function
' Call this function until it returns "" to get the remaining files
Public Function FindNext() As String
    ' Exit if no files have been found
    If mlngFile = 0 Then Exit Function
    ' Be sure to clear the contents of cFileName before each call to
    ' avoid garbage characters from being returned in your string.
    mwfdFindData.cFileName = Space$(MAX_PATH)
    ' If another file is found, then return it. Otherwise EndFind.
    If FindNextFile(mlngFile, mwfdFindData) Then
        FindNext = Left$(mwfdFindData.cFileName, InStr(mwfdFindData.cFileName, Chr$(0)) - 1)
    Else
        EndFind
    End If
End Function
' A private helper method which is called internally to close the
' FindFile handle and clear mlngFile to end a FindFile operation.
Private Sub EndFind()
    FindClose mlngFile
    mlngFile = 0
End Sub
' Return the short name of a found file (default = long filename)
Public Function GetShortName() As String
    Dim strShortFileName As String
    ' If no current file, then exit
    If mlngFile = 0 Then Exit Function
    ' Get the short filename (without trailing nulls)
    strShortFileName = Left$(mwfdFindData.cAlternate, InStr(mwfdFindData.cAlternate, Chr$(0)) - 1)
    ' If there is no short filename info, then strShortFilename will
    ' equal null (because of the (- 1) above)
    If Len(strShortFileName) = 0 Then
        ' If no short filename, then its already a short filename so
        ' set strShortFileName = .cFileName.
        strShortFileName = Left$(mwfdFindData.cFileName, InStr(mwfdFindData.cFileName, Chr$(0)) - 1)
    End If
    ' Return the short filename
    GetShortName = strShortFileName
End Function
' Return the date the current file was created. If the optional args
' are provided, then they will be set = to date and time values.
Public Function GetCreationDate(Optional datDate As Date, Optional datTime As Date) As String
    If mlngFile = 0 Then Exit Function
    '*****************************************************************
    ' If dwHighDateTime, then Win32 couldn't determine the date so
    ' return the unknown string. "Unknown" is the default.  Set this
    ' value to something else by using the UnknownDateText property.
    '*****************************************************************
    If mwfdFindData.ftCreationTime.dwHighDateTime = 0 Then
       If recursive = False Then
          GetCreationDate = Format$(FileDateTime(AddBackSlash(Dir1Path) & mwfdFindData.cFileName), LesOptions.DisplayDate)
        Else  ' On est en recherche récursive
          GetCreationDate = Format$(FileDateTime(AddBackSlash(lechemin) & mwfdFindData.cFileName), LesOptions.DisplayDate)
        End If
        Exit Function
    End If
    '*****************************************************************
    ' Get the time (in the current local/time zone)
    '*****************************************************************
    With GetSystemTime(mwfdFindData.ftCreationTime)
        '*************************************************************
        ' If datDate was provided, then set it to a date serial
        '*************************************************************
        datDate = DateSerial(.wYear, .wMonth, .wDay)
        '*************************************************************
        ' If datTime was provided, then set it to a time serial
        '*************************************************************
        datTime = TimeSerial(.wHour, .wMinute, .wSecond)
        '*************************************************************
        ' Use datDate and datTime as local variables (even if they
        ' weren't passed ByRef in the optional args) to create a
        ' a valid date/time value.  Return the date/time formatted
        ' using the default format of "m/d/yy h:nn:ss AM/PM" or
        ' the user-defined value which was set using the DateFormat
        ' property.
        '*************************************************************
        'GetCreationDate = format$(datDate + datTime, mstrDateFormat)
        GetCreationDate = Format$(datDate + datTime, LesOptions.DisplayDate)
    End With
End Function
'*********************************************************************
' Similar to GetCreationDate.  See GetCreationDate for comments.
'*********************************************************************
Public Function GetLastAccessDate(Optional datDate As Date, Optional datTime As Date) As String
    
    If mlngFile = 0 Then Exit Function
    
    If mwfdFindData.ftLastAccessTime.dwHighDateTime = 0 Then
        GetLastAccessDate = mstrUnknownDateText
        Exit Function
    End If
    
    With GetSystemTime(mwfdFindData.ftLastAccessTime)
        datDate = DateSerial(.wYear, .wMonth, .wDay)
        datTime = TimeSerial(.wHour, .wMinute, .wSecond)
        'GetLastAccessDate = format$(datDate + datTime, mstrDateFormat)
        GetLastAccessDate = Format$(datDate + datTime, LesOptions.DisplayDate)
    End With
    
End Function
'*********************************************************************
' Similar to GetCreationDate.  See GetCreationDate for comments.
'*********************************************************************
Public Function GetLastWriteDate(Optional datDate As Date, Optional datTime As Date) As String
    
    If mlngFile = 0 Then Exit Function
    
    If mwfdFindData.ftLastWriteTime.dwHighDateTime = 0 Then
        GetLastWriteDate = mstrUnknownDateText
        Exit Function
    End If
    
    With GetSystemTime(mwfdFindData.ftLastWriteTime)
        datDate = DateSerial(.wYear, .wMonth, .wDay)
        datTime = TimeSerial(.wHour, .wMinute, .wSecond)
        'GetLastWriteDate = format$(datDate + datTime, mstrDateFormat)
        GetLastWriteDate = Format$(datDate + datTime, LesOptions.DisplayDate)
    End With
    
End Function
'*********************************************************************
' Takes a FILETIME and converts it into the local system time
'*********************************************************************
Private Function GetSystemTime(ftmFileTime As FILETIME) As SYSTEMTIME
    Dim ftmLocalTime As FILETIME
    Dim stmSystemTime As SYSTEMTIME
    FileTimeToLocalFileTime ftmFileTime, ftmLocalTime
    FileTimeToSystemTime ftmLocalTime, stmSystemTime
    GetSystemTime = stmSystemTime
End Function
'*********************************************************************
' Sets the default values for private members when this object is
' created
'*********************************************************************
Private Sub Class_Initialize()
    mstrUnknownDateText = "Unknown"
    mstrDateFormat = "m/d/yy h:nn:ss AM/PM"
End Sub
'*********************************************************************
' Ends any open finds, if necessary
'*********************************************************************
Private Sub Class_Terminate()
    If mlngFile Then EndFind
End Sub
