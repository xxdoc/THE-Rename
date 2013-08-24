VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "CCRPFTV6.OCX"
Object = "{06D5A045-D511-11D3-9875-BB56A32B4523}#1.0#0"; "PROPBRWS.OCX"
Begin VB.Form Form5 
   Caption         =   "v"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   7635
   Begin VB.PictureBox Ongl2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6150
      Left            =   7665
      ScaleHeight     =   6150
      ScaleWidth      =   7530
      TabIndex        =   6
      Top             =   360
      Visible         =   0   'False
      Width           =   7530
      Begin PB.PropertyBrowser Pb2 
         Height          =   6090
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   10742
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CatFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NameWidth       =   75
      End
   End
   Begin VB.PictureBox Ongl1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   210
      ScaleHeight     =   6135
      ScaleWidth      =   7500
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   7500
      Begin MSComctlLib.ListView LV1 
         Height          =   6090
         Left            =   2955
         TabIndex        =   4
         Top             =   0
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   10742
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin CCRPFolderTV6.FolderTreeview FTV1 
         Height          =   6090
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   10742
         IntegralHeight  =   0   'False
      End
   End
   Begin MSComctlLib.TabStrip TS1 
      Height          =   6495
      Left            =   0
      TabIndex        =   0
      Top             =   345
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11456
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Files && Folders"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Viewer"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   6975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   13421772
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":0554
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":0AA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":0FFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":1550
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":1AA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":1FF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":254C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":2AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":2FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":3548
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":3A9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":3FF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":4544
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":4A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":4FEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":5540
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":5A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":5FE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":653C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      WhatsThisHelpID =   174
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select all"
            Description     =   "Select all files"
            Object.ToolTipText     =   "Select all files in current directory"
            Object.Tag             =   "Select all files in current directory"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "unselect"
            Description     =   "Unselect files from current directory"
            Object.ToolTipText     =   "Unselect files from current directory"
            Object.Tag             =   "Unselect files from current directory"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "invert selection"
            Description     =   "Invert selection"
            Object.ToolTipText     =   "Invert selection"
            Object.Tag             =   "Invert selection"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step"
            Description     =   "Enter a step value for selection"
            Object.ToolTipText     =   "Enter a step value for selection"
            Object.Tag             =   "Enter a step value for selection"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DropFiles"
            Description     =   "Drop files to copy them, click to configure"
            Object.ToolTipText     =   "Drop files to copy them, click to configure"
            Object.Tag             =   "Drop files to copy them, click to configure"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Recursive mode"
            Description     =   "Recursive mode"
            Object.ToolTipText     =   "Recursive mode"
            Object.Tag             =   "Recursive mode"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up one level"
            Description     =   "Up one level"
            Object.ToolTipText     =   "Up one level"
            Object.Tag             =   "Up one level"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Root directory"
            Description     =   "Root directory"
            Object.ToolTipText     =   "Root directory"
            Object.Tag             =   "Root directory"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Show history of your moves"
            Description     =   "Show history of your moves"
            Object.ToolTipText     =   "Show history of your moves"
            Object.Tag             =   "Show history of your moves"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add to your favorites"
            Object.ToolTipText     =   "Add to your favorites"
            Object.Tag             =   "Add to your favorites"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "organize your favorites"
            Object.ToolTipText     =   "Organize your favorites"
            Object.Tag             =   "organize your favorites"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "First favorite"
            Object.ToolTipText     =   "First favorite"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previous favorite"
            Object.ToolTipText     =   "previous favorite"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next favorite"
            Object.ToolTipText     =   "Next favorite"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Last favorite"
            Object.ToolTipText     =   "Last favorite"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
      Begin VB.ComboBox Combo5 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "Form5.frx":6A90
         Left            =   6075
         List            =   "Form5.frx":6A92
         TabIndex        =   2
         ToolTipText     =   "Type a file filter and press enter or select a filter from the list"
         Top             =   20
         WhatsThisHelpID =   182
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    TS1.width = Form5.ScaleWidth
    TS1.height = Form5.ScaleHeight - Toolbar1.height - 50

End Sub

Private Sub TS1_Click()
'MsgBox TS1.SelectedItem.Index
Select Case TS1.SelectedItem.Index
    Case 1
        Ongl1.left = 45
        Ongl1.top = 45
        Ongl1.width = TS1.width - 50
        Ongl1.height = TS1.height - 200
        Ongl1.Visible = True
        Ongl2.Visible = False
    Case 2
        Ongl2.left = 45
        Ongl2.top = 45
        Ongl2.width = TS1.width - 50
        Ongl2.height = TS1.height - 200
        Ongl2.Visible = True
        Ongl1.Visible = False
End Select
End Sub
