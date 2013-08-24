VERSION 5.00
Begin VB.Form ffolder 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add folder's name"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3510
   ControlBox      =   0   'False
   HelpContextID   =   232
   Icon            =   "ffolder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Reset"
      Height          =   300
      Left            =   68
      TabIndex        =   11
      ToolTipText     =   "Reset values"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Position "
      Height          =   1095
      Left            =   68
      TabIndex        =   19
      Top             =   2640
      Width           =   3375
      Begin VB.OptionButton Option3 
         Caption         =   "Replace prefix"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Replace prefix with it's path"
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Add to the right"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Add file's path to the right of the prefix"
         Top             =   465
         Width           =   1440
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Add to the left"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Add file's path to the left of the prefix"
         Top             =   240
         Value           =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "\"
      Height          =   975
      Left            =   68
      TabIndex        =   17
      Top             =   1560
      Width           =   3375
      Begin VB.PictureBox Picture3 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         ScaleHeight     =   615
         ScaleWidth      =   2655
         TabIndex        =   18
         Top             =   240
         Width           =   2655
         Begin VB.OptionButton Option2 
            Caption         =   "Delete ""\"""
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   5
            ToolTipText     =   "This will delete the \ character contain in the file's path"
            Top             =   0
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Replace ""\"" with :"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   6
            ToolTipText     =   "Enter here the character to use to replace \"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1725
            TabIndex        =   7
            Text            =   " "
            ToolTipText     =   "Enter here the character to use to replace \"
            Top             =   230
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Name "
      Height          =   1455
      Left            =   68
      TabIndex        =   14
      Top             =   0
      Width           =   3375
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   120
         ScaleHeight     =   1125
         ScaleWidth      =   2415
         TabIndex        =   15
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton Option1 
            Caption         =   "Use complete folder's name"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   0
            ToolTipText     =   "THE Rename will uses the complete file's path"
            Top             =   0
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Use # of levels"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   1
            ToolTipText     =   "Use only some elements of the path"
            Top             =   260
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1560
            TabIndex        =   2
            Text            =   "1"
            ToolTipText     =   "Enter the number of levels to return"
            Top             =   260
            Width           =   495
         End
         Begin VB.PictureBox Picture2 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   240
            ScaleHeight     =   615
            ScaleWidth      =   1335
            TabIndex        =   16
            Top             =   560
            Width           =   1335
            Begin VB.OptionButton Option0 
               Caption         =   "From left"
               Enabled         =   0   'False
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   3
               ToolTipText     =   "Return elements from the left"
               Top             =   0
               Value           =   -1  'True
               Width           =   975
            End
            Begin VB.OptionButton Option0 
               Caption         =   "From right"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   4
               ToolTipText     =   "Return elements from the right"
               Top             =   260
               Width           =   1095
            End
         End
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1208
      TabIndex        =   12
      ToolTipText     =   "Don't use folder's name"
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2348
      TabIndex        =   13
      ToolTipText     =   "Use folder's name with these parameters"
      Top             =   3840
      Width           =   1095
   End
End
Attribute VB_Name = "ffolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
 FolderOk = False
 Unload Me
End Sub

Private Sub cmdOK_Click()
 If Val(Text1.Text) <= 0 And Option1(1).Value = True Then
    MsgBox "Please enter a valid value"
    Text1.SetFocus
    Exit Sub
 End If
 
 If Option1(0).Value = True Then
  Folder1 = 0
 Else
  Folder1 = 1
 End If
 
 If Option0(0).Value = True Then
  Folder2 = 0
 Else
  Folder2 = 1
 End If
 
 If Option2(0).Value = True Then
  Folder3 = 0
 Else
  Folder3 = 1
 End If
 
 If Option3(0).Value = True Then
  Folder4 = 0
 Else
  If Option3(1).Value = True Then
   Folder4 = 1
  Else
   Folder4 = 2
  End If
 End If
 Folder5 = Text1.Text
 Folder6 = Text2.Text
 FolderOk = True
 Unload Me
End Sub

Private Sub Command1_Click()
 Folder1 = 0
 Folder2 = 0
 Folder3 = 0
 Folder4 = 0
 Folder5 = ""
 Folder6 = ""
 Option1(Folder1).Value = True
 Option0(Folder2).Value = True
 Option2(Folder3).Value = True
 Option3(Folder4).Value = True
 Text1.Text = Folder5
 Text2.Text = Folder6
End Sub

Private Sub Form_Load()
 Option1(Folder1).Value = True
 Option0(Folder2).Value = True
 Option2(Folder3).Value = True
 Option3(Folder4).Value = True
 Text1.Text = Folder5
 Text2.Text = Folder6
End Sub

Private Sub Option1_Click(Index As Integer)
 If Index = 0 Then
  Text1.Enabled = False
  Option0(0).Enabled = False
  Option0(1).Enabled = False
 Else
  Text1.Enabled = True
  Option0(0).Enabled = True
  Option0(1).Enabled = True
 End If
End Sub

Private Sub Option2_Click(Index As Integer)
 If Index = 0 Then
  Text2.Enabled = False
 Else
  Text2.Enabled = True
 End If
End Sub

Private Sub Text1_GotFocus()
    SelAll Text1
End Sub

Private Sub Text2_Change()
 CharInterdits Text2.Text
End Sub

Private Sub Text2_GotFocus()
    SelAll Text2
End Sub
