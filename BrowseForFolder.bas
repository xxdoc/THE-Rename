Attribute VB_Name = "BrowseForFolder"
Option Explicit
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hWnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long     'Optional parameter
    lpClass       As String   'Optional parameter
    hkeyClass     As Long     'Optional parameter
    dwHotKey      As Long     'Optional parameter
    hIcon         As Long     'Optional parameter
    hProcess      As Long     'Optional parameter
End Type

Private Type BROWSEINFO
  hOwner          As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfn            As Long
  lParam          As Long
  iImage          As Long
End Type

'Public Const BIF_DONTGOBELOWDOMAIN = &H2
'Public Const BIF_STATUSTEXT = &H4
'Public Const BIF_RETURNFSANCESTORS = &H8
'Public Const BIF_BROWSEFORCOMPUTER = &H1000
'Public Const BIF_BROWSEFORPRINTER = &H2000
'Public pidl As Long
'Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Rem déclarations pour les propriétés
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Public Function BrowseFolder(feuille As Form, Optional titre As String = "Browse for folder") As String
 Dim bi As BROWSEINFO, pidl As Long, Path As String, Pos As Integer
 bi.hOwner = feuille.hWnd
 bi.pidlRoot = 0&
 If titre = "" Then
  bi.lpszTitle = "Select Folder"
 Else
  bi.lpszTitle = titre
 End If
 bi.ulFlags = BIF_RETURNONLYFSDIRS
 pidl = SHBrowseForFolder(bi)
 Path = Space$(MAX_PATH)
 If SHGetPathFromIDList(ByVal pidl, ByVal Path) Then
  Pos = InStr(Path, Chr$(0))
  BrowseFolder = Left$(Path, Pos - 1)
 Else
  BrowseFolder = ""
 End If
 Call CoTaskMemFree(pidl)
End Function
Public Sub ShowProperties(Filename As String, feuille As Form)
   Dim SEI As SHELLEXECUTEINFO, R As Long
   With SEI
      .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
      .hWnd = feuille.hWnd
      .lpVerb = "properties"
      .lpFile = Filename
      .lpParameters = vbNullChar
      .lpDirectory = vbNullChar
      .nShow = 0
      .hInstApp = 0
      .lpIDList = 0
      .cbSize = Len(SEI)
   End With
   R = ShellExecuteEX(SEI)
End Sub
