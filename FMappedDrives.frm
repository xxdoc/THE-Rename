VERSION 5.00
Begin VB.Form FMappedDrives 
   Caption         =   "Mapped Drives"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   ControlBox      =   0   'False
   HelpContextID   =   161
   Icon            =   "FMappedDrives.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "&Map Network drive..."
      Height          =   300
      HelpContextID   =   161
      Left            =   2528
      TabIndex        =   2
      ToolTipText     =   "Map network drive"
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Disconnect Network drive..."
      Height          =   300
      HelpContextID   =   161
      Left            =   98
      TabIndex        =   1
      ToolTipText     =   "Disconnect Network drive"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   161
      Left            =   4598
      TabIndex        =   3
      ToolTipText     =   "Close this window"
      Top             =   2880
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   2595
      HelpContextID   =   161
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Refresh"
      Height          =   300
      Left            =   6240
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   195
   End
End
Attribute VB_Name = "FMappedDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_oAutoPos As New clsAutoPositioner
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Const NO_ERROR = 0
Private Const ERROR_BAD_DEVICE = 1200&
Private Const ERROR_NOT_CONNECTED = 2250&
Private Const ERROR_MORE_DATA = 234
Private Const ERROR_CONNECTION_UNAVAIL = 1201&
Private Const ERROR_NO_NETWORK = 1222&
Private Const ERROR_EXTENDED_ERROR = 1208&
Private Sub Command1_Click()
   Dim d As Integer
   Dim sRet As String
   List1.Clear
   For d = Asc("C") To Asc("Z")
      sRet = DriveLetterToUNC(Chr$(d))
      If Len(sRet) Then
         List1.AddItem Chr$(d) & ": --> " & sRet
      End If
   Next
End Sub

Private Function DriveLetterToUNC(ByVal DriveLetter As String) As String
   Dim nRet As Long
   Dim Drv As String, Dbg As String, Buffer As String
   Dim BufLen As Long
   Const MAX_PATH = 260

   If Len(DriveLetter) Then
      Drv = UCase$(Left$(DriveLetter, 1)) & ":"
      Buffer = Space$(MAX_PATH)
      BufLen = Len(Buffer)
      nRet = WNetGetConnection(Drv, Buffer, BufLen)
      If nRet = ERROR_MORE_DATA Then
         Buffer = Space$(BufLen)
         nRet = WNetGetConnection(Drv, Buffer, BufLen)
      End If

      If nRet = NO_ERROR Then
         DriveLetterToUNC = Left$(Buffer, InStr(Buffer, vbNullChar) - 1)
      Else
         Select Case nRet
            Case ERROR_BAD_DEVICE
               Dbg = Drv & " --> "
               Dbg = Dbg & "The specified device name is invalid."
            Case ERROR_NOT_CONNECTED
               Dbg = Drv & " --> "
               Dbg = Dbg & "This network connection does not exist."
            Case ERROR_CONNECTION_UNAVAIL
               Dbg = Drv & " --> "
               Dbg = Dbg & "The device is not currently connected but it is a remembered connection."
            Case ERROR_NO_NETWORK
               Dbg = "The network is not present or not started."
            Case ERROR_MORE_DATA
               Dbg = "Buffer is too small!"
            Case ERROR_EXTENDED_ERROR
               Dbg = "An error has occurred, call WNetGetLastError."
            Case Else
               Dbg = "Unknown error code: " & nRet
         End Select
      End If
   End If
End Function

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Command3_Click()
 Dim R As Long
 R = WNetDisconnectDialog(RENAME.hwnd, RESOURCETYPE_DISK)
End Sub

Private Sub Command4_Click()
 Dim R As Long
 R = WNetConnectionDialog(RENAME.hwnd, RESOURCETYPE_DISK)
End Sub

Private Sub Form_Load()
    Command1_Click
    m_oAutoPos.AddAssignment Me.Command2, Me, tCONTAINER_RELATIVE_POS_BOTTOM
    m_oAutoPos.AddAssignment Me.Command3, Me, tCONTAINER_RELATIVE_POS_BOTTOM
    m_oAutoPos.AddAssignment Me.Command4, Me, tCONTAINER_RELATIVE_POS_BOTTOM
    m_oAutoPos.AddAssignment Me.List1, Me, tCONTAINER_WIDTH_DELTA_RIGHT
    m_oAutoPos.AddAssignment Me.List1, Me, tCONTAINER_HEIGHT_DELTA_BOTTOM
End Sub

Private Sub Form_Resize()
    m_oAutoPos.RefreshPositions
End Sub
