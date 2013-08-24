VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSHFileOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
'SHFileOp - VB program that demonstrates using SHFileOperation
'Copyright (c) 1997 SoftCircuits Programming (R)
'Redistributed by Permission.
'
'Starting with Windows 95, the Windows API provides the SHFileOperation
'function. This function provides many options for copying, moving,
'renaming and deleting files. For lengthy operations, a dialog box is
'displayed that indicates the current operation and progress. This is
'exactly the same progress dialog box displayed by the Windows Explorer
'because the Windows Explorer also calls SHFileOperation.
'
'The CSHFileOp class encapsulates SHFileOperation making it much
'easier to use. Not only does it provide much easier access to the many
'options but performs some special work so that SHFILEOPSTRUCT can be
'fully exploited by VB. Most available options are implemented as class
'properties and methods. For example, set the AllowUndo property to
'True to cause deleted files to be sent to the Recycle Bin. View the
'comments within the source code for information about these options.
'Refer to the Windows API documentation for additional information.
'
'After specifying the source and destination files and setting any
'options you want, call the CopyFiles, MoveFiles, DeleteFiles or
'RenameFiles methods to perform the operation. Each of these methods
'returns True if all operations completely successfully, or False if
'there was an error or the user canceled an operation. (Note: You do
'not normally need to display error information because Windows will
'notify the user of any tasks that could not be completed.)
'
'This program may be distributed on the condition that it is
'distributed in full and unchanged, and that no fee is charged for
'such distribution with the exception of reasonable shipping and media
'charged. In addition, the code in this program may be incorporated
'into your own programs and the resulting programs may be distributed
'without payment of royalties.
'
'This example program was provided by:
' SoftCircuits Programming
' http://www.softcircuits.com
' P.O. Box 16262
' Irvine, CA 92623
Option Explicit
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

'Default property values
Private Const DEFAULT_ALLOWUNDO = True
Private Const DEFAULT_CONFIRMMAKEDIR = True
Private Const DEFAULT_CONFIRMOPERATION = True
Private Const DEFAULT_CUSTOMTEXT = ""
Private Const DEFAULT_INCLUDEDIRECTORIES = True
Private Const DEFAULT_PARENTWND = 0
Private Const DEFAULT_RENAMEONCOLLISION = True
Private Const DEFAULT_SILENTMODE = True

'SHFileOperation declarations
Private Const FO_MOVE = 1
Private Const FO_COPY = 2
Private Const FO_DELETE = 3
Private Const FO_RENAME = 4

Private Const FOF_MULTIDESTFILES = &H1
Private Const FOF_SILENT = &H4
Private Const FOF_RENAMEONCOLLISION = &H8
Private Const FOF_NOCONFIRMATION = &H10
'Private Const FOF_WANTMAPPINGHANDLE = &H20
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_FILESONLY = &H80
Private Const FOF_SIMPLEPROGRESS = &H100
Private Const FOF_NOCONFIRMMKDIR = &H200

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, pFrom As Any, ByVal lCount As Long)

'Private variables
Private m_bAllowUndo As Boolean
Private m_bConfirmMakeDir As Boolean
Private m_bConfirmOperation As Boolean
Private m_sCustomText As String
Private m_bIncludeDirectories As Boolean
Private m_hParentWnd As Long
Private m_bRenameOnCollision As Boolean
Private m_bSilentMode As Boolean

Private m_SourceFiles As New Collection
Private m_DestFiles As New Collection

'=== Public Properties ======================================
'Sets if Windows stores undo information (if possible)
Public Property Let AllowUndo(bAllowUndo As Boolean)
    m_bAllowUndo = bAllowUndo
End Property

'Gets if Windows stores undo information (if possible)
Public Property Get AllowUndo() As Boolean
    AllowUndo = m_bAllowUndo
End Property

'Sets if is user is prompted before creating any needed directories
Public Property Let ConfirmMakeDir(bConfirmMakeDir As Boolean)
    m_bConfirmMakeDir = bConfirmMakeDir
End Property

'Gets if is user is prompted before creating any needed directories
Public Property Get ConfirmMakeDir() As Boolean
    ConfirmMakeDir = m_bConfirmMakeDir
End Property

'Sets if user is prompted for confirmation
Public Property Let ConfirmOperation(bConfirmOperation As Boolean)
    m_bConfirmOperation = bConfirmOperation
End Property

'Gets if user is prompted for confirmation
Public Property Get ConfirmOperation() As Boolean
    ConfirmOperation = m_bConfirmOperation
End Property

'Set custom text displayed in SHFileOperation dialog box
'If text is "" the dialog box displays the name of each file
Public Property Let CustomText(sCustomText As String)
    m_sCustomText = sCustomText
End Property

'Get custom text displayed in SHFileOperation dialog box
Public Property Get CustomText() As String
    CustomText = m_sCustomText
End Property

'Sets if operation affects directories if wildcards (*.*) are specified
Public Property Let IncludeDirectories(bIncludeDirectories As Boolean)
    m_bIncludeDirectories = bIncludeDirectories
End Property

'Gets if operation affects directories if wildcards (*.*) are specified
Public Property Get IncludeDirectories() As Boolean
    IncludeDirectories = m_bIncludeDirectories
End Property

'Set the parent window of the SHFileOperation dialog box
Public Property Let ParentWnd(hParentWnd As Long)
    m_hParentWnd = hParentWnd
End Property

'Get the parent window of the SHFileOperation dialog box
Public Property Get ParentWnd() As Long
    ParentWnd = m_hParentWnd
End Property

'Specifies if destination should be renamed if file with same name already exists
Public Property Let RenameOnCollision(bRenameOnCollision As Boolean)
    m_bRenameOnCollision = bRenameOnCollision
End Property

'Gets if destination should be renamed if
'file with same name already exists
Public Property Get RenameOnCollision() As Boolean
    RenameOnCollision = m_bRenameOnCollision
End Property

'Sets if the SHFileOperation dialog box should be hidden
Public Property Let SilentMode(bSilentMode As Boolean)
    m_bSilentMode = bSilentMode
End Property

'Gets if the SHFileOperation dialog box should be hidden
Public Property Get SilentMode() As Boolean
    SilentMode = m_bSilentMode
End Property


'=== Public Methods ========================================

'Moves the source files to the location specified by the
'destination. Returns false if any operations were not
'completed successfully (whether due to errors or user
'actions). If an error occurred, the user has already been
'notified.
Public Function MoveFiles() As Boolean
    MoveFiles = DoOperation(FO_MOVE)
End Function

'Copies the source files to the location specified by the
'destination. Returns false if any operations were not
'completed successfully (whether due to errors or user
'actions). If an error occurred, the user has already been
'notified.
Public Function CopyFiles() As Boolean
    CopyFiles = DoOperation(FO_COPY)
End Function

'Deletes the source files. Returns false if any operations
'were not completed successfully (whether due to errors or
'user actions). If an error occurred, the user has already
'been notified.
Public Function DeleteFiles() As Boolean
    DeleteFiles = DoOperation(FO_DELETE)
End Function

'Renames the source files. Returns false if any operations
'were not completed successfully (whether due to errors or
'user actions). If an error occurred, the user has already
'been notified.
Public Function RenameFiles() As Boolean
    RenameFiles = DoOperation(FO_RENAME)
End Function

'Resets the list of source files
Public Sub ClearSourceFiles()
    Set m_SourceFiles = Nothing
End Sub

'Adds a file to the list of source files
Public Sub AddSourceFile(sFilename As String)
    m_SourceFiles.Add sFilename
End Sub

'Resets the list of destination files
Public Sub ClearDestFiles()
    Set m_DestFiles = Nothing
End Sub

'Adds a file to the list of destination files
Public Sub AddDestFile(sFilename As String)
    m_DestFiles.Add sFilename
End Sub


'=== Private Operations ======================================

'Initialize properties to default values
Private Sub Class_Initialize()
    m_bAllowUndo = DEFAULT_ALLOWUNDO
    m_bConfirmMakeDir = DEFAULT_CONFIRMMAKEDIR
    m_bConfirmOperation = DEFAULT_CONFIRMOPERATION
    m_sCustomText = DEFAULT_CUSTOMTEXT
    m_bIncludeDirectories = DEFAULT_INCLUDEDIRECTORIES
    m_hParentWnd = DEFAULT_PARENTWND
    m_bRenameOnCollision = DEFAULT_RENAMEONCOLLISION
    m_bSilentMode = DEFAULT_SILENTMODE
End Sub

'Execute call to SHFileOperation
Private Function DoOperation(wFunc As Integer) As Boolean
    Dim i As Long, ptr As Long
    Dim shfo As SHFILEOPSTRUCT
    Dim ByteArray() As Byte
    Dim buff1() As Byte, buff2() As Byte, buff3() As Byte
    
    If wFunc = FO_RENAME And LesOptions.ChineseKorean = 1 Then ' Il faut changer de méthode pour renommer les fichers
        Name m_SourceFiles(1) As m_DestFiles(1)
        DoOperation = True
        Exit Function
    End If
    
    'Parent window of dialog box--just use 0
    shfo.hWnd = m_hParentWnd
    'Operation to perform
    shfo.wFunc = wFunc
    'Operation flags
    shfo.fFlags = 0
    If m_bAllowUndo Then
        shfo.fFlags = shfo.fFlags Or FOF_ALLOWUNDO
    End If
    If m_bSilentMode Then
        shfo.fFlags = shfo.fFlags Or FOF_SILENT
    End If
    If m_bRenameOnCollision Then
        shfo.fFlags = shfo.fFlags Or FOF_RENAMEONCOLLISION
    End If
    If Not m_bConfirmOperation Then
        shfo.fFlags = shfo.fFlags Or FOF_NOCONFIRMATION
    End If
    If Not m_bConfirmMakeDir Then
        shfo.fFlags = shfo.fFlags Or FOF_NOCONFIRMMKDIR
    End If
    If Not m_bIncludeDirectories Then
        shfo.fFlags = shfo.fFlags Or FOF_FILESONLY
    End If
    If Len(m_sCustomText) > 0 Then
        ' Modif du 11/10/2001
        shfo.lpszProgressTitle = m_sCustomText & vbNullChar & vbNullChar
        shfo.fFlags = shfo.fFlags Or FOF_SIMPLEPROGRESS
    End If
    'Build 'From' string
    If m_SourceFiles.Count = 0 Then
        Err.Raise vbObjectError + 1000, , "No source files specified file operation"
    End If
    For i = 1 To m_SourceFiles.Count
        shfo.pFrom = shfo.pFrom & m_SourceFiles(i) & Chr$(0)
    Next
    ' AJout du 11/10/2001
    ' Multiple names must be null-separated. The list of names must be double null-terminated.
    shfo.pFrom = shfo.pFrom & Chr$(0)
    
    'Build 'To' string
    For i = 1 To m_DestFiles.Count
        shfo.pTo = shfo.pTo & m_DestFiles(i) & Chr$(0)
    Next
    ' Multiple names must be null-separated. The list of names must be double null-terminated.
    ' Ajout du 11/10/2001
    shfo.pTo = shfo.pTo & Chr$(0)
    
    'Test if more than one destination files
    If m_DestFiles.Count > 1 Then
        shfo.fFlags = shfo.fFlags Or FOF_MULTIDESTFILES
    End If
    'Note: Windows packs the SHFILEOPSTRUCT structure but
    '32-bit Visual Basic does not. Therefore, all members
    'following the two-byte fFlags member are offset by
    '2 bytes. To deal with this, we copy structure members
    'to a byte array with the proper alignment and pass
    'the byte array to SHFileOperation.
    ReDim ByteArray(LenB(shfo) - 2)
    CopyMemory ByteArray(0), shfo.hWnd, Len(shfo.hWnd)
    CopyMemory ByteArray(4), shfo.wFunc, Len(shfo.wFunc)
    'Variable-length strings require extra work
    buff1 = StrConv(shfo.pFrom & Chr$(0), vbFromUnicode)
    ptr = VarPtr(buff1(0))
    CopyMemory ByteArray(8), ptr, LenB(ptr)
    buff2 = StrConv(shfo.pTo & Chr$(0), vbFromUnicode)
    ptr = VarPtr(buff2(0))
    CopyMemory ByteArray(12), ptr, LenB(ptr)
    CopyMemory ByteArray(16), shfo.fFlags, Len(shfo.fFlags)
    CopyMemory ByteArray(18), shfo.fAnyOperationsAborted, Len(shfo.fAnyOperationsAborted)
    CopyMemory ByteArray(22), shfo.hNameMappings, Len(shfo.hNameMappings)
    buff3 = StrConv(shfo.lpszProgressTitle & Chr$(0), vbFromUnicode)
    ptr = VarPtr(buff3(0))
    CopyMemory ByteArray(26), ptr, LenB(ptr)
    'Call SHFileOperation
    i = SHFileOperation(ByteArray(0))
    'Retrieve fAnyOperationsAborted flag
    CopyMemory shfo.fAnyOperationsAborted, ByteArray(18), Len(shfo.fAnyOperationsAborted)
    'Return True if SHFileOperation succeeded and no operations aborted
    DoOperation = Not CBool(i Or shfo.fAnyOperationsAborted)
End Function


