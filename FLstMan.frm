VERSION 5.00
Begin VB.Form FLstMan 
   AutoRedraw      =   -1  'True
   Caption         =   "Edit the list manually"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   3195
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   600
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   4575
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1163
      ScaleHeight     =   375
      ScaleWidth      =   2355
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2355
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   300
         HelpContextID   =   14
         Left            =   1200
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Save the list"
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   300
         HelpContextID   =   14
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Don't save the list"
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter the new name, a tabulation and the old name"
      Height          =   195
      Left            =   540
      TabIndex        =   4
      Top             =   60
      Width           =   3615
   End
   Begin VB.Menu medit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mcut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mcopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mpaste 
         Caption         =   "&Paste"
      End
   End
End
Attribute VB_Name = "FLstMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim vtexte As String
    Dim ch1 As String, ch2 As String, vligne As String
    Dim i As Integer, vnb As Integer, longueur As Integer
    Dim itmX As ListItem
    vtexte = Text1.Text
    longueur = Len(vtexte)
    Me.MousePointer = vbHourglass
    RENAME.ListView2.ListItems.Clear
    RENAME.ListView2.Visible = False
    For i = 1 To longueur
        If Mid$(vtexte, i, 1) = Chr$(13) Then
            If Trim$(vligne) <> "" Then
                ch1 = GetToken(vligne, vbTab, 1)
                ch2 = GetToken(vligne, vbTab, 2)
                Set itmX = RENAME.ListView2.ListItems.Add(, , ch1)
                itmX.Text = ch1
                itmX.SubItems(1) = ch2
            End If
            vligne = ""
        Else
            If Mid$(vtexte, i, 1) <> Chr$(10) Then vligne = vligne + Mid$(vtexte, i, 1)
        End If
    Next
    RENAME.ListView2.Visible = True
    Me.MousePointer = vbNormal
    Unload Me
End Sub
Private Sub Form_Load()
    Dim i As Long
    Dim vnb As Long
    vnb = RENAME.ListView2.ListItems.Count
    Text1.Text = ""
    Me.MousePointer = vbHourglass
    Text1.Visible = False
    For i = 1 To vnb
        Text1.Text = Text1.Text + RENAME.ListView2.ListItems(i).Text + vbTab + RENAME.ListView2.ListItems(i).SubItems(1) & vbCrLf
    Next
    Text1.Visible = True
    Me.MousePointer = vbNormal
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Label1.Left = (Me.ScaleWidth - Label1.width) / 2
    Label1.Top = 10
    Picture1.Left = (Me.ScaleWidth / 2) - (Picture1.width / 2)
    Picture1.Top = Me.ScaleHeight - Picture1.height
    Text1.Top = 200
    Text1.Left = 10
    Text1.width = Me.ScaleWidth - 10
    Text1.height = (Me.ScaleHeight - Picture1.height - Label1.height) - 70
End Sub

