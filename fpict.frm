VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FPICT 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pictures Information"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5715
   ControlBox      =   0   'False
   HelpContextID   =   78
   Icon            =   "fpict.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1725
      TabIndex        =   6
      ToolTipText     =   "Don't use pictures information"
      Top             =   3450
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2940
      TabIndex        =   7
      ToolTipText     =   "Use pictures information with these parameters"
      Top             =   3450
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3315
      Left            =   120
      TabIndex        =   8
      Top             =   60
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   5847
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Exif"
      TabPicture(0)   =   "fpict.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "List1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdclear"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   210
         TabIndex        =   0
         Top             =   825
         Width           =   4380
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "Cl&ear"
         Height          =   330
         Left            =   4695
         TabIndex        =   1
         ToolTipText     =   "Clear command line"
         Top             =   800
         Width           =   615
      End
      Begin VB.ListBox List1 
         Height          =   1230
         ItemData        =   "fpict.frx":0028
         Left            =   180
         List            =   "fpict.frx":0095
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1890
         Width           =   4920
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   4950
         TabIndex        =   9
         Top             =   1260
         Width           =   4950
         Begin VB.OptionButton Option2 
            Caption         =   "Replace prefix"
            Height          =   255
            Index           =   2
            Left            =   3375
            TabIndex        =   4
            ToolTipText     =   "Replace prefix"
            Top             =   0
            Width           =   1470
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Add to the left"
            Height          =   255
            Index           =   0
            Left            =   15
            TabIndex        =   2
            ToolTipText     =   "Add information to the left"
            Top             =   0
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Add to the right"
            Height          =   255
            Index           =   1
            Left            =   1650
            TabIndex        =   3
            ToolTipText     =   "Add information to the right"
            Top             =   0
            Width           =   1440
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "List of available information"
         Height          =   195
         Left            =   210
         TabIndex        =   11
         Top             =   1635
         Width           =   1905
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select commands to extract information from your pictures"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   525
         Width           =   4065
      End
   End
End
Attribute VB_Name = "FPICT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    PicEXIF.UseEXIF = False
    Unload Me
End Sub

Private Sub cmdclear_Click()
    Text1.Text = ""
    Text1.SetFocus
End Sub

Private Sub cmdOK_Click()
    With PicEXIF
        .UseEXIF = True
        .Rule = Text1.Text
        .PlaceWhereToPut = 2
        If Option2(0).Value = True Then
            .PlaceWhereToPut = 0
        ElseIf Option2(0).Value = True Then
            .PlaceWhereToPut = 1
        End If
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = PicEXIF.Rule
    Option2(PicEXIF.PlaceWhereToPut).Value = True
End Sub

Private Sub List1_DblClick()
    InsertTextInTextBox Text1, List1
    Text1.SetFocus
End Sub
