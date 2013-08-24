VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EF59A10B-9BC4-11D3-8E24-44910FC10000}#10.0#0"; "VBALEDIT.OCX"
Begin VB.Form FCode 
   AutoRedraw      =   -1  'True
   Caption         =   "Your program"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin vbalEdit.vbalRichEdit Text1 
      Height          =   5655
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   9975
      Version         =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      ViewMode        =   0
      AutoURLDetect   =   0   'False
      ScrollBars      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   6525
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7497
            MinWidth        =   7497
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
            Object.ToolTipText     =   "Line and Column"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "14:16"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1931
            MinWidth        =   1940
            TextSave        =   "28/11/1999"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   495
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   9015
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FCode.frx":0000
         Left            =   2760
         List            =   "FCode.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   135
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Functions and variables"
         Height          =   195
         Left            =   840
         TabIndex        =   1
         Top             =   195
         Width           =   1680
      End
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu msave 
         Caption         =   "Save"
      End
      Begin VB.Menu msaveas 
         Caption         =   "Save as..."
      End
   End
End
Attribute VB_Name = "FCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer
For i = 0 To RENAME.listcmd.ListCount - 1
 Combo1.AddItem RENAME.listcmd.List(i)
Next i
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 Frame1.width = Me.ScaleWidth - 80
 Text1.top = Frame1.top + Frame1.height + 100
 Text1.width = Me.ScaleWidth - Text1.left - 20
 Text1.height = Me.ScaleHeight - StatusBar1.height - Frame1.height - 100
End Sub

Private Sub mopen_Click()
 Text1.LoadFromFile "c:\autoexec.bat", SF_TEXT
End Sub
