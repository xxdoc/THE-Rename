Attribute VB_Name = "CommDlgs"
'//
'// Common Dialogs Module
'//
'// Description:
'// Provides wrapper functions into the various Windows OS common dialog boxes
'//
'// ***************************************************************
'// *  Go to Dragon's VB Code Corner for more useful sourcecode:  *
'// *  http://personal.inet.fi/cool/dragon/vb/                    *
'// ***************************************************************
'//
'// Author of this module: Unknown
'//

Option Explicit
'//
'// Win32s (Private Functions for Wrappers Below)
'//
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWnd As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'Public Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGSTRUC) As Long
'Public Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
'Public Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
'Public Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
'Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'//
'// Win32s (Public)
'//
'Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
'Declare Function HTMLHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwnd As Long, ByVal szFilename As String, ByVal dwCommand As Long, ByVal dwData As Any) As Long

'//
'// Constants (Public for Print Dialog Box)
'//
'Public Const PD_NOSELECTION = &H4
'Public Const PD_DISABLEPRINTTOFILE = &H80000
'Public Const PD_PRINTTOFILE = &H20
'Public Const PD_RETURNDC = &H100
'Public Const PD_RETURNDEFAULT = &H400
'Public Const PD_RETURNIC = &H200
'Public Const PD_SELECTION = &H1
'Public Const PD_SHOWHELP = &H800
'Public Const PD_NOPAGENUMS = &H8
'Public Const PD_PAGENUMS = &H2
'Public Const PD_ALLPAGES = &H0
'Public Const PD_COLLATE = &H10
'Public Const PD_HIDEPRINTTOFILE = &H100000

'//
'// Constants (Private)
'//
'Private Const FW_BOLD = 700
'Private Const GMEM_MOVEABLE = &H2
'Private Const GMEM_ZEROINIT = &H40
'Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
'Private Const OFN_ALLOWMULTISELECT = &H200
'Private Const OFN_CREATEPROMPT = &H2000
'Private Const OFN_ENABLEHOOK = &H20
'Private Const OFN_ENABLETEMPLATE = &H40
'Private Const OFN_ENABLETEMPLATEHANDLE = &H80
'Private Const OFN_EXPLORER = &H80000
'Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
'Private Const OFN_LONGNAMES = &H200000
'Private Const OFN_NOCHANGEDIR = &H8
'Private Const OFN_NODEREFERENCELINKS = &H100000
'Private Const OFN_NOLONGNAMES = &H40000
'Private Const OFN_NONETWORKBUTTON = &H20000
'Private Const OFN_NOREADONLYRETURN = &H8000
'Private Const OFN_NOTESTFILECREATE = &H10000
'Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
'Private Const OFN_READONLY = &H1
'Private Const OFN_SHAREAWARE = &H4000
'Private Const OFN_SHAREFALLTHROUGH = 2
'Private Const OFN_SHARENOWARN = 1
'Private Const OFN_SHAREWARN = 0
'Private Const OFN_SHOWHELP = &H10
'Private Const PD_ENABLEPRINTHOOK = &H1000
'Private Const PD_ENABLEPRINTTEMPLATE = &H4000
'Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
'Private Const PD_ENABLESETUPHOOK = &H2000
'Private Const PD_ENABLESETUPTEMPLATE = &H8000
'Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
'Private Const PD_NONETWORKBUTTON = &H200000
'Private Const PD_PRINTSETUP = &H40
'Private Const PD_USEDEVMODECOPIES = &H40000
'Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
'Private Const PD_NOWARNING = &H80
'Private Const CF_ANSIONLY = &H400&
'Private Const CF_APPLY = &H200&
'Private Const CF_BITMAP = 2
'Private Const CF_PRINTERFONTS = &H2
'Private Const CF_PRIVATEFIRST = &H200
'Private Const CF_PRIVATELAST = &H2FF
'Private Const CF_RIFF = 11
'Private Const CF_SCALABLEONLY = &H20000
'Private Const CF_SCREENFONTS = &H1
'Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'Private Const CF_DIB = 8
'Private Const CF_DIF = 5
'Private Const CF_DSPBITMAP = &H82
'Private Const CF_DSPENHMETAFILE = &H8E
'Private Const CF_DSPMETAFILEPICT = &H83
'Private Const CF_DSPTEXT = &H81
'Private Const CF_EFFECTS = &H100&
'Private Const CF_ENABLEHOOK = &H8&
'Private Const CF_ENABLETEMPLATE = &H10&
'Private Const CF_ENABLETEMPLATEHANDLE = &H20&
'Private Const CF_ENHMETAFILE = 14
'Private Const CF_FIXEDPITCHONLY = &H4000&
'Private Const CF_FORCEFONTEXIST = &H10000
'Private Const CF_GDIOBJFIRST = &H300
'Private Const CF_GDIOBJLAST = &H3FF
'Private Const CF_INITTOLOGFONTSTRUCT = &H40&
'Private Const CF_LIMITSIZE = &H2000&
'Private Const CF_METAFILEPICT = 3
'Private Const CF_NOFACESEL = &H80000
'Private Const CF_NOVERTFONTS = &H1000000
'Private Const CF_NOVECTORFONTS = &H800&
'Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
'Private Const CF_NOSCRIPTSEL = &H800000
'Private Const CF_NOSIMULATIONS = &H1000&
'Private Const CF_NOSIZESEL = &H200000
'Private Const CF_NOSTYLESEL = &H100000
'Private Const CF_OEMTEXT = 7
'Private Const CF_OWNERDISPLAY = &H80
'Private Const CF_PALETTE = 9
'Private Const CF_PENDATA = 10
'Private Const CF_SCRIPTSONLY = CF_ANSIONLY
'Private Const CF_SELECTSCRIPT = &H400000
'Private Const CF_SHOWHELP = &H4&
'Private Const CF_SYLK = 4
'Private Const CF_TEXT = 1
'Private Const CF_TIFF = 6
'Private Const CF_TTONLY = &H40000
'Private Const CF_UNICODETEXT = 13
'Private Const CF_USESTYLE = &H80&
'Private Const CF_WAVE = 12
'Private Const CF_WYSIWYG = &H8000
'Private Const CFERR_CHOOSEFONTCODES = &H2000
'Private Const CFERR_MAXLESSTHANMIN = &H2002
'Private Const CFERR_NOFONTS = &H2001
'Private Const CC_ANYCOLOR = &H100
'Private Const CC_CHORD = 4
'Private Const CC_CIRCLES = 1
'Private Const CC_ELLIPSES = 8
'Private Const CC_ENABLEHOOK = &H10
'Private Const CC_ENABLETEMPLATE = &H20
'Private Const CC_ENABLETEMPLATEHANDLE = &H40
'Private Const CC_FULLOPEN = &H2
'Private Const CC_INTERIORS = 128
'Private Const CC_NONE = 0
'Private Const CC_PIE = 2
'Private Const CC_PREVENTFULLOPEN = &H4
'Private Const CC_RGBINIT = &H1
'Private Const CC_ROUNDRECT = 256 '
'Private Const CC_SHOWHELP = &H8
'Private Const CC_SOLIDCOLOR = &H80
'Private Const CC_STYLED = 32
'Private Const CC_WIDE = 16
'Private Const CC_WIDESTYLED = 64
'Private Const CCERR_CHOOSECOLORCODES = &H5000
'Private Const LOGPIXELSY = 90
'Private Const CCHDEVICENAME = 32
'Private Const CCHFORMNAME = 32
'Private Const SIMULATED_FONTTYPE = &H8000
'Private Const PRINTER_FONTTYPE = &H4000
'Private Const SCREEN_FONTTYPE = &H2000
'Private Const BOLD_FONTTYPE = &H100
'Private Const ITALIC_FONTTYPE = &H200
'Private Const REGULAR_FONTTYPE = &H400
'Private Const WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1)
'Private Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
'Private Const SHAREVISTRING = "commdlg_ShareViolation"
'Private Const FILEOKSTRING = "commdlg_FileNameOK"
'Private Const COLOROKSTRING = "commdlg_ColorOK"
'Private Const SETRGBSTRING = "commdlg_SetRGBColor"
'Private Const FINDMSGSTRING = "commdlg_FindReplace"
'Private Const HELPMSGSTRING = "commdlg_help"
'Private Const CD_LBSELNOITEMS = -1
'Private Const CD_LBSELCHANGE = 0
'Private Const CD_LBSELSUB = 1
'Private Const CD_LBSELADD = 2
'Private Const NOERROR = 0
'Private Const CSIDL_DESKTOP = &H0
'Private Const CSIDL_PROGRAMS = &H2
'Private Const CSIDL_CONTROLS = &H3
'Private Const CSIDL_PRINTERS = &H4
'Private Const CSIDL_PERSONAL = &H5
'Private Const CSIDL_FAVORITES = &H6
'Private Const CSIDL_STARTUP = &H7
'Private Const CSIDL_RECENT = &H8
'Private Const CSIDL_SENDTO = &H9
'Private Const CSIDL_BITBUCKET = &HA
'Private Const CSIDL_STARTMENU = &HB
'Private Const CSIDL_DESKTOPDIRECTORY = &H10
'Private Const CSIDL_DRIVES = &H11
'Private Const CSIDL_NETWORK = &H12
'Private Const CSIDL_NETHOOD = &H13
'Private Const CSIDL_FONTS = &H14
'Private Const CSIDL_TEMPLATES = &H15
'Private Const BIF_RETURNONLYFSDIRS = &H1
'Private Const BIF_DONTGOBELOWDOMAIN = &H2
'Private Const BIF_STATUSTEXT = &H4
'Private Const BIF_RETURNFSANCESTORS = &H8
'Private Const BIF_BROWSEFORCOMPUTER = &H1000
'Private Const BIF_BROWSEFORPRINTER = &H2000
'Private Const HWND_BROADCAST = &HFFFF&
'Private Const WM_WININICHANGE = &H1A

'//
'// ByteToString Function
'//
'// Description:
'// Converts an array of bytes into a string
'//
'// Syntax:
'// StrVar = ByteToString(ARRAY)
'//
'// Example:
'// szBuf = BytesToString(aChars(10))
'//
'Private Function ByteToString(aBytes() As Byte) As String
'    Dim dwBytePoint As Long, dwByteVal As Long, szOut As String
'    dwBytePoint = LBound(aBytes)
'    While dwBytePoint <= UBound(aBytes)
'        dwByteVal = aBytes(dwBytePoint)
'        If dwByteVal = 0 Then
'            ByteToString = szOut
'            Exit Function
'        Else
'            szOut = szOut & Chr$(dwByteVal)
'        End If
'        dwBytePoint = dwBytePoint + 1
'    Wend
'    ByteToString = szOut
'End Function
'//
'// DialogFile Function
'//
'// Description:
'// Displays the File Open/Save As common dialog boxes.
'//
'// Syntax:
'// StrVar = DialogFile(hWnd, IntVar, StrVar, StrVar, StrVar, StrVar, StrVar)
'//
'// Example:
'// szFilename = DialogFile(Me.hWnd, 1, "Open", "MyFileName.doc", "Documents" & chr$(0) & "*.doc" & chr$(0) & "All files" & chr$(0) & "*.*", App.Path, "doc")
'//
'// Please note that the szFilter var works a bit differently
'// from the filter property associated with the common dialog
'// control. Instead of separating the differents parts of the
'// string with pipe chars, |, you should use null chars, chr$(0),
'// as separators.
Public Function DialogFile(hWnd As Long, wMode As Integer, szDialogTitle As String, szFilename As String, szFilter As String, szDefDir As String, szDefExt As String) As String
    Dim X As Long, OFN As OPENFILENAME, szFile As String
    OFN.lStructSize = Len(OFN)
    OFN.hWnd = hWnd
    OFN.lpstrTitle = szDialogTitle
    OFN.lpstrFile = szFilename & String$(250 - Len(szFilename), 0)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = String$(255, 0)
    OFN.nMaxFileTitle = 255
    OFN.lpstrFilter = szFilter
    OFN.nFilterIndex = 1
    OFN.lpstrInitialDir = szDefDir
    OFN.lpstrDefExt = szDefExt

   If wMode = 1 Then
        OFN.flags = OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST
        X = GetOpenFileName(OFN)
    Else
        OFN.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        X = GetSaveFileName(OFN)
    End If

    If X <> 0 Then
        '// If Instr$(OFN.lpstrFileTitle, Chr$(0)) > 0 Then
        '//     szFileTitle = Left$(OFN.lpstrFileTitle, Instr$(OFN.lpstrFileTitle, Chr$(0)) - 1)
        '// End If
        If InStr(OFN.lpstrFile, Chr$(0)) > 0 Then
            szFile = Left$(OFN.lpstrFile, InStr(OFN.lpstrFile, Chr$(0)) - 1)
        End If
        '// OFN.nFileOffset is the number of characters from the beginning of the
        '// full path to the start of the file name
        '// OFN.nFileExtension is the number of characters from the beginning of the
        '// full path to the file's extention, including the (.)
        '// MsgBox "File Name is " & szFileTitle & Chr$(13) & Chr$(10) & "Full path and file is " & szFile, , "Open"
        '// DialogFile = szFile & "|" & szFileTitle
        DialogFile = szFile
    Else
        DialogFile = ""
    End If
End Function
'//
'// DialogPrint Function
'//
'// Description:
'// Displays the Print common dialog box and returns a structure containing user entered
'// information from the common dialog box.
'//
'// Syntax:
'// PRINTPROPS = DialogPrint(hWnd, BOOL, DWORD)
'//
'// Example:
'// Dim PP As PRINTPROPS
'// PP = DialogPrint(Me.hWnd, True, PD_PAGENUMS or PD_SELECTION or PD_SHOWHELP)
'//
'Public Function DialogPrint(hWnd As Long, bPages As Boolean, flags As Long) As PRINTPROPS
'    Dim DM As DEVMODE, PD As PRINTDLGSTRUC
'    Dim lpDM As Long, wNull As Integer, szDevName As String
'    PD.lStructSize = Len(PD)
'    PD.hWnd = hWnd
'    PD.hDevMode = 0
'    PD.hDevNames = 0
'    PD.hDC = 0
'    PD.flags = flags
'    PD.nFromPage = 0
'    PD.nToPage = 0
'    PD.nMinPage = 0
'    If bPages Then PD.nMaxPage = bPages - 1
'    PD.nCopies = 0
'    DialogPrint.Cancel = True
'    If PrintDlg(PD) Then
'        lpDM = GlobalLock(PD.hDevMode)
'        CopyMemory DM, ByVal lpDM, Len(DM)
'        lpDM = GlobalUnlock(PD.hDevMode)
'        DialogPrint.Cancel = False
'        DialogPrint.Device = left$(DM.dmDeviceName, Instr$(DM.dmDeviceName, chr$(0)) - 1)
'        DialogPrint.FromPage = 0
'        DialogPrint.ToPage = 0
'        DialogPrint.All = True
'        If PD.flags And PD_PRINTTOFILE Then DialogPrint.file = True Else DialogPrint.file = False
'        If PD.flags And PD_COLLATE Then DialogPrint.Collate = True Else DialogPrint.Collate = False
'        If PD.flags And PD_PAGENUMS Then
'            DialogPrint.Pages = True
'            DialogPrint.All = False
'            DialogPrint.FromPage = PD.nFromPage
'            DialogPrint.ToPage = PD.nToPage
'        Else
'            DialogPrint.Pages = False
'        End If
'        If PD.flags And PD_SELECTION Then
'            DialogPrint.Selection = True
'            DialogPrint.All = False
'        Else
'            DialogPrint.Pages = False
'        End If
'        If PD.nCopies = 1 Then
'            DialogPrint.Copies = DM.dmCopies
'        End If
'        DialogPrint.DM = DM
'    End If
'End Function
