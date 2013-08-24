Attribute VB_Name = "MCreateDir"
Option Explicit
Private Const INVALID_HANDLE_VALUE = -1
Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Function CreateNestedFoldersByPath(ByVal completeDirectory As String) As Integer
  'creates nested directories on the drive included in the path by parsing the final
  'directory string into a directory array, and looping through each to create the final path.
  'The path could be passed to this method as a pre-filled array, reducing the code.
   Dim SA As SECURITY_ATTRIBUTES
   Dim drivePart As String
   Dim newDirectory  As String
   Dim Item As String
   Dim sfolders() As String
   Dim Pos As Integer
   Dim X As Integer
   completeDirectory = AddBackSlash(completeDirectory)
   Pos = InStr(completeDirectory, ":")

   If Pos Then
         drivePart = GetPart(completeDirectory, "\")
   Else: drivePart = ""
   End If
   Do Until completeDirectory = ""
     Item = GetPart(completeDirectory, "\")
     ReDim Preserve sfolders(0 To X) As String
     If X = 0 Then Item = drivePart & Item
     sfolders(X) = Item
     X = X + 1 'increment the array counter
   Loop
   X = -1
   Do
      X = X + 1
      newDirectory = newDirectory & sfolders(X)
      SA.nLength = LenB(SA)
      Call CreateDirectory(newDirectory, SA)
   Loop Until X = UBound(sfolders)
   CreateNestedFoldersByPath = X + 1
End Function
Private Function GetPart(startStrg As String, delimiter As String) As String
  Dim c As Integer
  Dim Item As String
  c = 1
  Do
    If Mid$(startStrg, c, 1) = delimiter Then
      Item = Mid$(startStrg, 1, c)
      startStrg = Mid$(startStrg, c + 1, Len(startStrg))
      GetPart = Item
      Exit Function
    End If
    c = c + 1
  Loop
End Function
