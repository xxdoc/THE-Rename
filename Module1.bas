Attribute VB_Name = "Module1"
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function OpenFile% Lib "Kernel" (ByVal lpFileName$, lpReOpenBuff As OFSTRUCT, ByVal wStyle%)

'OFSTRUCT structure used by the OpenFile API function
Type OFSTRUCT            '136 bytes in length
         cBytes As String * 1
         fFixedDisk As String * 1
         nErrCode As Integer
         reserved As String * 4
         szPathName As String * 128
End Type
     
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_MOVE = &H1
Public Const FO_RENAME = &H4

Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_FILESONLY = &H80
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_SILENT = &H4
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_WANTMAPPINGHANDLE = &H20


Private Type SHFILEOPSTRUCT
   hwnd As Long
   wFunc As Long
   pFrom As String
   pTo As String
   fFlags As Integer
   fAnyOperationsAborted As Boolean
   hNameMappings As Long
   lpszProgressTitle As String
End Type

Public Function CopyThisFile(Filename As String, ToDir As String)
    On Error Resume Next
    Dim FileStruct As SHFILEOPSTRUCT
    Dim X As Long
    Dim P As Boolean
    P = DoesFileExist(Filename)
    If P = True Then
        FileStruct.pFrom = Filename
        FileStruct.pTo = ToDir
        FileStruct.fFlags = FOF_NOCONFIRMMKDIR + FOF_NOCONFIRMATION + FOF_SILENT
        FileStruct.wFunc = FO_COPY
        X = SHFileOperation(FileStruct)
    Else
        Err.Raise Copy_Error, "FileOperations::CopyThisFile", Err.Description
    End If
End Function

Public Function DoesFileExist(NameOfFile As String) As Boolean
    Dim X As Long
    Dim wStyle As Long
    Dim Buffer As OFSTRUCT
    Dim IsThere As Long

    If (Len(NameOfFile) = 0) Or (InStr(NameOfFile, "*") > 0) Or (InStr(NameOfFile, "?") > 0) Then
        DoesFileExist = False
        Exit Function
    End If
    On Error GoTo NoSuchFile
    X = OpenFile(NameOfFile, Buffer, OF_EXIST)
    If X < 0 Then
        GoTo CheckForError
        Else
        DoesFileExist = True
        Exit Function
    End If

CheckForError:
    X = Buffer.nErrCode
    If X = 3 Then
        GoTo NoSuchFile
    End If

NoSuchFile:
    DoesFileExist = False

ExitFileExist:
    On Error GoTo 0
End Function


Public Function DeleteThisFile(Filename As String)
    On Error Resume Next
    Dim FileStruct As SHFILEOPSTRUCT
    Dim X As Long
    Dim P As Boolean
    P = DoesFileExist(Filename)
    If P = True Then
        FileStruct.pFrom = Filename
        FileStruct.fFlags = FOF_SILENT + FOF_ALLOWUNDO + FOF_NOCONFIRMATION
        FileStruct.wFunc = FO_DELETE
        X = SHFileOperation(FileStruct)
    Else
        Err.Raise Del_Error, "FileOperations::DeleteThisFile", Err.Description
    End If
End Function

Public Function MoveThisFile(Filename As String, DestName As String)
    On Error Resume Next
    Dim FileStruct As SHFILEOPSTRUCT
    Dim P As Boolean
    Dim X As Long

    P = DoesFileExist(Filename)
    If P = True Then
        FileStruct.pFrom = Filename
        FileStruct.pTo = DestName
        FileStruct.fFlags = FOF_SILENT + FOF_NOCONFIRMATION
        FileStruct.wFunc = FO_MOVE
        X = SHFileOperation(FileStruct)
    Else
        Err.Raise Move_Error, "FileOperations::MoveThisFile", Err.Description
    End If
End Function

Public Function RenameThisFile(Filename As String, Target As String)
    On Error Resume Next
    Dim FileStruct As SHFILEOPSTRUCT
    Dim P As Boolean
    Dim X As Long
    P = DoesFileExist(Filename)
Rem    If P = True Then
Rem        FileStruct.hwnd = Me.hwnd
        FileStruct.pFrom = Filename + Chr$(0) + Chr$(0)
        FileStruct.pTo = Target + Chr$(0) + Chr$(0)
        FileStruct.fFlags = FOF_ALLOWUNDO
        FileStruct.wFunc = FO_RENAME
        X = SHFileOperation(FileStruct)
Rem    Else
Rem        Err.Raise Rename_Error, "FileOperation::RenameThisFile", Err.Description
Rem    End If
End Function


