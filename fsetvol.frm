VERSION 5.00
Begin VB.Form fsetvol 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set volume label"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   ControlBox      =   0   'False
   HelpContextID   =   63
   Icon            =   "fsetvol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Text1 
      Height          =   315
      Left            =   68
      TabIndex        =   1
      ToolTipText     =   "Limited to 11 characters"
      Top             =   1080
      Width           =   4290
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   63
      Left            =   1103
      TabIndex        =   2
      ToolTipText     =   "Close this window"
      Top             =   1485
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   63
      Left            =   2273
      TabIndex        =   3
      ToolTipText     =   "Rename selected drive"
      Top             =   1485
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      HelpContextID   =   63
      Left            =   68
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select drive"
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "serial"
      Height          =   195
      Left            =   1680
      TabIndex        =   6
      ToolTipText     =   "Drive's serial number"
      Top             =   840
      Width           =   2700
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Type new name"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select drive to rename"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1590
   End
End
Attribute VB_Name = "fsetvol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cHist13 As New cHistory
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function SetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
 Dim R As Long
 cHist13.AddNewItem Text1.Text
 R = SetVolumeLabel(Left$(Combo1.Text, 3), Text1.Text)
 Unload Me
End Sub
Private Sub Combo1_Change()
 infolect
End Sub
Private Sub Combo1_Click()
 infolect
End Sub

Private Sub Form_Load()
 cHist13.sKey = "VolumeLabel"
 cHist13.Items Text1
 LoadAvailableDrives
 Combo1.ListIndex = 0
 If IsWinNT Then
      SHFD_FORMAT_FULL = 0    'full format
      SHFD_FORMAT_QUICK = 1   'quick format
   Else 'it's Win95/98
      SHFD_FORMAT_QUICK = 0   'quick format
      SHFD_FORMAT_FULL = 1    'full format
      SHFD_FORMAT_SYSONLY = 2 'copies system files only
   End If
End Sub

Private Function GetDrvstr(currDrive As String) As String
 Select Case GetDriveType(currDrive)
      Case 0, 1: GetDrvstr = "(Unknow Drive Type)"
      
      Case DRIVE_REMOVABLE:
           Select Case Left$(currDrive, 1)
              Case "A", "B": GetDrvstr = "Floppy drive"
              Case Else:     GetDrvstr = "Removable drive"
           End Select
      
      Case DRIVE_FIXED:     GetDrvstr = "Fixed Hard drive"
      Case DRIVE_REMOTE:    GetDrvstr = "Remote drive"
      Case DRIVE_CDROM:     GetDrvstr = "CD ROM"
      Case DRIVE_RAMDISK:   GetDrvstr = "Ram disk"
   End Select
End Function
Private Sub LoadAvailableDrives()
  Dim R As Long, lpBuffer  As String * 256, longueur As Long
  Dim n&
  On Error GoTo ErrGen
  lpBuffer = Space$(256)
  longueur = Len(lpBuffer)
  R = GetLogicalDriveStrings(longueur, lpBuffer)
  If R Then
    Do
     n = InStr(lpBuffer, Chr$(0))
     If n > 1 Then
      Combo1.AddItem Left$(lpBuffer, n - 1) & "   " & GetDrvstr(Left$(lpBuffer, n - 1))
      lpBuffer = Mid$(lpBuffer, n + 1)
     End If
    Loop Until n <= 1
   End If
    
Exit Sub
ErrGen:
ErreurGrave "fsetvol.LoadAvailableDrives"
End Sub
Private Sub infolect()
 Dim volbuf$, sysname$
 Dim VolumeSN&, sysflags&, componentlength&
 Dim res&
 Dim HiWord As Long
 Dim HiHexStr As String
 Dim LoWord As Long
 Dim LoHexStr As String
 volbuf$ = String$(256, 0)
 sysname$ = String$(256, 0)
 res = GetVolumeInformation(Left$(Combo1.Text, 3), volbuf$, 255, VolumeSN&, componentlength, sysflags, sysname$, 255)
 If res = 0 Then
   Text1.Text = ""
 Else
  Text1.Text = volbuf$
  HiWord = GetHiWord(VolumeSN&) And &HFFFF&
   LoWord = GetLoWord(VolumeSN&) And &HFFFF&
'   HiHexStr = Format$(Hex(HiWord), "0000")
'   LoHexStr = Format$(Hex(LoWord), "0000")
   HiHexStr = Hex$(HiWord)
   LoHexStr = Hex$(LoWord)
   Label3 = HiHexStr & "-" & LoHexStr
 End If
End Sub

Private Function GetHiWord(dw As Long) As Integer
    If dw And &H80000000 Then
          GetHiWord = (dw \ 65535) - 1
    Else: GetHiWord = dw \ 65535
    End If
End Function

Private Function GetLoWord(dw As Long) As Integer
    If dw And &H8000& Then
          GetLoWord = &H8000 Or (dw And &H7FFF&)
    Else: GetLoWord = dw And &HFFFF&
    End If
End Function
Private Sub Text1_GotFocus()
    SelAll Text1
End Sub
