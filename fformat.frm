VERSION 5.00
Begin VB.Form fformat 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Format drive/floppy"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   ControlBox      =   0   'False
   HelpContextID   =   62
   Icon            =   "fformat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   62
      Left            =   1103
      TabIndex        =   1
      ToolTipText     =   "Close this window"
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   62
      Left            =   2295
      TabIndex        =   2
      ToolTipText     =   "Format selected drive"
      Top             =   840
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      HelpContextID   =   62
      Left            =   68
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Select drive"
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select drive to format"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1500
   End
End
Attribute VB_Name = "fformat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SHFD_CAPACITY_DEFAULT = 0 'default drive capacity
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim R As Long
   Dim drvToFormat As Integer
   drvToFormat = Combo1.ListIndex
   DoEvents
   R = SHFormatDrive(Me.hWnd, drvToFormat, SHFD_CAPACITY_DEFAULT, SHFD_FORMAT_QUICK)
End Sub

Private Sub Form_Load()
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
      Case 0, 1: GetDrvstr = "(Unknown Drive Type)"
      
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
ErreurGrave "fformat.LoadAvailableDrives"
End Sub

