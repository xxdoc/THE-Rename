VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form rename 
   AutoRedraw      =   -1  'True
   Caption         =   "THE Rename - Freeware by Hervé Thouzard - Version 2.1.4"
   ClientHeight    =   7095
   ClientLeft      =   1590
   ClientTop       =   3150
   ClientWidth     =   10950
   Icon            =   "rename.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10950
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   11940
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":04E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox panelcmd 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   120
      ScaleHeight     =   1920
      ScaleWidth      =   5715
      TabIndex        =   19
      Top             =   7620
      Visible         =   0   'False
      Width           =   5715
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   4815
         TabIndex        =   122
         Top             =   60
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Clear command line"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Organize your commands"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Parms"
                     Object.Tag             =   "Parameters"
                     Text            =   "Parameters..."
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.ListBox Combo6 
         Height          =   810
         IntegralHeight  =   0   'False
         Left            =   1140
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   390
         Visible         =   0   'False
         Width           =   3195
      End
      Begin MSComctlLib.TreeView TV1 
         Height          =   1230
         Left            =   45
         TabIndex        =   121
         Top             =   540
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2170
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   500
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command6 
         Height          =   300
         HelpContextID   =   168
         Left            =   5280
         Picture         =   "rename.frx":06BE
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Set options for Cyclic selection"
         Top             =   1200
         UseMaskColor    =   -1  'True
         WhatsThisHelpID =   198
         Width           =   300
      End
      Begin VB.ListBox listcmd 
         Height          =   1230
         ItemData        =   "rename.frx":07A8
         Left            =   45
         List            =   "rename.frx":07AA
         Sorted          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "If you need help about a command, press the F1 key"
         Top             =   540
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtlang 
         Height          =   330
         HelpContextID   =   41
         Left            =   45
         OLEDropMode     =   2  'Automatic
         TabIndex        =   20
         ToolTipText     =   "Type your command or pick one from the list (see the help file for more information)"
         Top             =   45
         WhatsThisHelpID =   206
         Width           =   4770
      End
      Begin THERename.LabelText cmdtxt1 
         Height          =   285
         Left            =   2310
         TabIndex        =   22
         ToolTipText     =   "Enter begin's value"
         Top             =   1215
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   503
         Caption         =   "Counter's intial value"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   1600
         MousePointer    =   0
         Text            =   "1"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignement  =   1
      End
      Begin THERename.LabelText cmdtxt2 
         Height          =   285
         Left            =   4380
         TabIndex        =   23
         ToolTipText     =   "Enter increment value (1 for example)"
         Top             =   1215
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         Caption         =   "Step"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   500
         MousePointer    =   0
         Text            =   "1"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignement  =   1
      End
   End
   Begin TabDlg.SSTab TabGen 
      Height          =   6260
      Left            =   4800
      TabIndex        =   26
      Top             =   480
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Folders"
      TabPicture(0)   =   "rename.frx":07AC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FolderTreeview1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Naming Rules"
      TabPicture(1)   =   "rename.frx":07C8
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameDroite"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tags"
      TabPicture(2)   =   "rename.frx":07E4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LvMP3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Pictures"
      TabPicture(3)   =   "rename.frx":0800
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Acdsee"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Text"
      TabPicture(4)   =   "rename.frx":081C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text1"
      Tab(4).ControlCount=   1
      Begin THERename.MyFrame Frame1 
         Height          =   2595
         Left            =   120
         TabIndex        =   73
         Top             =   420
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4577
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         Caption         =   "Prefix"
         ShowBorderInDesignMode=   0   'False
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "rename.frx":0838
            Left            =   60
            List            =   "rename.frx":083A
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   118
            ToolTipText     =   "Scroll to select action to perform on prefix"
            Top             =   300
            WhatsThisHelpID =   175
            Width           =   5175
         End
         Begin VB.PictureBox PanelPrefix 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   60
            ScaleHeight     =   1815
            ScaleWidth      =   5775
            TabIndex        =   75
            Top             =   720
            Width           =   5775
            Begin VB.CommandButton Command10 
               Caption         =   "Reset"
               Height          =   300
               HelpContextID   =   110
               Left            =   4800
               TabIndex        =   76
               ToolTipText     =   "Reset all selections to their default's value"
               Top             =   1500
               WhatsThisHelpID =   181
               Width           =   870
            End
            Begin TabDlg.SSTab SSTab1 
               Height          =   1455
               Left            =   0
               TabIndex        =   77
               Top             =   0
               Width           =   5745
               _ExtentX        =   10134
               _ExtentY        =   2566
               _Version        =   393216
               Style           =   1
               Tabs            =   8
               TabsPerRow      =   8
               TabHeight       =   520
               TabCaption(0)   =   "Counter"
               TabPicture(0)   =   "rename.frx":083C
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "onglcounter"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Check3"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).ControlCount=   2
               TabCaption(1)   =   "Size"
               TabPicture(1)   =   "rename.frx":0858
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Check5"
               Tab(1).Control(1)=   "Picture5"
               Tab(1).ControlCount=   2
               TabCaption(2)   =   "Date"
               TabPicture(2)   =   "rename.frx":0874
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Picture6"
               Tab(2).Control(1)=   "Check6"
               Tab(2).ControlCount=   2
               TabCaption(3)   =   "Time"
               TabPicture(3)   =   "rename.frx":0890
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Picture7"
               Tab(3).Control(1)=   "Check7"
               Tab(3).ControlCount=   2
               TabCaption(4)   =   "Folder's name"
               TabPicture(4)   =   "rename.frx":08AC
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "Command19"
               Tab(4).ControlCount=   1
               TabCaption(5)   =   "Text "
               TabPicture(5)   =   "rename.frx":08C8
               Tab(5).ControlEnabled=   0   'False
               Tab(5).Control(0)=   "Command8"
               Tab(5).Control(1)=   "Option1(0)"
               Tab(5).Control(2)=   "Text2"
               Tab(5).Control(3)=   "Picture1"
               Tab(5).Control(4)=   "Option1(1)"
               Tab(5).Control(5)=   "Text14"
               Tab(5).ControlCount=   6
               TabCaption(6)   =   "Pictures"
               TabPicture(6)   =   "rename.frx":08E4
               Tab(6).ControlEnabled=   0   'False
               Tab(6).Control(0)=   "Command1"
               Tab(6).Control(1)=   "Picture8"
               Tab(6).Control(2)=   "Check1"
               Tab(6).ControlCount=   3
               TabCaption(7)   =   "Sounds"
               TabPicture(7)   =   "rename.frx":0900
               Tab(7).ControlEnabled=   0   'False
               Tab(7).Control(0)=   "Command5"
               Tab(7).ControlCount=   1
               Begin VB.CommandButton Command1 
                  Caption         =   "Pictures Tags..."
                  Height          =   300
                  Left            =   -70620
                  TabIndex        =   119
                  ToolTipText     =   "Use EXIF tags"
                  Top             =   360
                  Width           =   1275
               End
               Begin VB.CommandButton Command19 
                  Caption         =   "Options..."
                  Height          =   375
                  Left            =   -72720
                  TabIndex        =   117
                  Top             =   600
                  WhatsThisHelpID =   195
                  Width           =   1455
               End
               Begin VB.TextBox Text14 
                  Height          =   285
                  Left            =   -73200
                  TabIndex        =   116
                  ToolTipText     =   "Add text to the prefix"
                  Top             =   825
                  Visible         =   0   'False
                  WhatsThisHelpID =   200
                  Width           =   2055
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Add text"
                  Height          =   195
                  Index           =   1
                  Left            =   -74880
                  TabIndex        =   115
                  ToolTipText     =   "Click to add or remove a text to the prefix"
                  Top             =   870
                  WhatsThisHelpID =   199
                  Width           =   960
               End
               Begin VB.PictureBox Picture1 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   -71040
                  ScaleHeight     =   255
                  ScaleWidth      =   1335
                  TabIndex        =   112
                  Top             =   900
                  Width           =   1340
                  Begin VB.OptionButton Option2 
                     Caption         =   "Begin"
                     Height          =   195
                     Index           =   0
                     Left            =   0
                     TabIndex        =   114
                     ToolTipText     =   "Text will be added to the left of prefix"
                     Top             =   0
                     Value           =   -1  'True
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   735
                  End
                  Begin VB.OptionButton Option2 
                     Caption         =   "End"
                     Height          =   195
                     Index           =   1
                     Left            =   740
                     TabIndex        =   113
                     ToolTipText     =   "Text will be added at the end of prefix"
                     Top             =   0
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   615
                  End
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Left            =   -73200
                  TabIndex        =   111
                  ToolTipText     =   "Current prefix will be replace with this text"
                  Top             =   480
                  Visible         =   0   'False
                  WhatsThisHelpID =   197
                  Width           =   2775
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Replace with text"
                  Height          =   195
                  Index           =   0
                  Left            =   -74880
                  TabIndex        =   110
                  ToolTipText     =   "Click to replace or not file's prefix with a text"
                  Top             =   525
                  WhatsThisHelpID =   196
                  Width           =   1635
               End
               Begin VB.CheckBox Check5 
                  Caption         =   "Add file's size"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   109
                  ToolTipText     =   "Click to add file's size"
                  Top             =   360
                  WhatsThisHelpID =   189
                  Width           =   1320
               End
               Begin VB.PictureBox Picture5 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -74640
                  ScaleHeight     =   240
                  ScaleWidth      =   5055
                  TabIndex        =   105
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   190
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   4
                     Left            =   1650
                     TabIndex        =   108
                     ToolTipText     =   "File's size will be added to the right"
                     Top             =   0
                     WhatsThisHelpID =   190
                     Width           =   1440
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   3
                     Left            =   15
                     TabIndex        =   107
                     ToolTipText     =   "File's size will be added to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   190
                     Width           =   1320
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     Index           =   5
                     Left            =   3375
                     TabIndex        =   106
                     ToolTipText     =   "File's size will replace prefix"
                     Top             =   0
                     WhatsThisHelpID =   190
                     Width           =   1470
                  End
               End
               Begin VB.PictureBox Picture6 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -74640
                  ScaleHeight     =   240
                  ScaleWidth      =   5055
                  TabIndex        =   101
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   192
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   8
                     Left            =   15
                     TabIndex        =   104
                     ToolTipText     =   "File's date will be added to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   192
                     Width           =   1365
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   7
                     Left            =   1650
                     TabIndex        =   103
                     ToolTipText     =   "File's date will be added to the right"
                     Top             =   0
                     WhatsThisHelpID =   192
                     Width           =   1410
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     Index           =   6
                     Left            =   3375
                     TabIndex        =   102
                     ToolTipText     =   "File's date will replace prefix"
                     Top             =   0
                     WhatsThisHelpID =   192
                     Width           =   1575
                  End
               End
               Begin VB.CheckBox Check6 
                  Caption         =   "Add file's date"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   100
                  ToolTipText     =   "Click to add file's date"
                  Top             =   360
                  WhatsThisHelpID =   191
                  Width           =   1365
               End
               Begin VB.PictureBox Picture7 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   -74640
                  ScaleHeight     =   285
                  ScaleWidth      =   4695
                  TabIndex        =   96
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   194
                  Width           =   4695
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     Index           =   11
                     Left            =   3375
                     TabIndex        =   99
                     ToolTipText     =   "File's time will replace prefix"
                     Top             =   0
                     WhatsThisHelpID =   194
                     Width           =   2295
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   10
                     Left            =   1650
                     TabIndex        =   98
                     ToolTipText     =   "File's time will be added to the right"
                     Top             =   0
                     WhatsThisHelpID =   194
                     Width           =   1440
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   9
                     Left            =   15
                     TabIndex        =   97
                     ToolTipText     =   "File's time will be added to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   194
                     Width           =   1365
                  End
               End
               Begin VB.CheckBox Check7 
                  Caption         =   "Add file's time"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   95
                  ToolTipText     =   "Click to add file's time"
                  Top             =   360
                  WhatsThisHelpID =   193
                  Width           =   1320
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "Add a counter"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   94
                  ToolTipText     =   "Click to add a counter"
                  Top             =   360
                  UseMaskColor    =   -1  'True
                  WhatsThisHelpID =   183
                  Width           =   1365
               End
               Begin VB.PictureBox Picture8 
                  BorderStyle     =   0  'None
                  Height          =   660
                  Left            =   -74880
                  ScaleHeight     =   660
                  ScaleWidth      =   5475
                  TabIndex        =   90
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   203
                  Width           =   5475
                  Begin VB.TextBox Text11 
                     Height          =   285
                     HelpContextID   =   14
                     Left            =   200
                     TabIndex        =   120
                     Text            =   "%w%x%h%"
                     ToolTipText     =   "%w% represents picture's with, %h% represents height. Unit is pixel"
                     Top             =   50
                     Width           =   4995
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     HelpContextID   =   162
                     Index           =   15
                     Left            =   3735
                     TabIndex        =   93
                     ToolTipText     =   "Replace prefix"
                     Top             =   400
                     WhatsThisHelpID =   203
                     Width           =   2295
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     HelpContextID   =   162
                     Index           =   16
                     Left            =   2010
                     TabIndex        =   92
                     ToolTipText     =   "Add to the right"
                     Top             =   400
                     WhatsThisHelpID =   203
                     Width           =   1395
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     HelpContextID   =   162
                     Index           =   17
                     Left            =   375
                     TabIndex        =   91
                     ToolTipText     =   "Add to the left"
                     Top             =   400
                     Value           =   -1  'True
                     WhatsThisHelpID =   203
                     Width           =   1365
                  End
               End
               Begin VB.CheckBox Check1 
                  Caption         =   "Add picture's width and height"
                  Height          =   255
                  HelpContextID   =   162
                  Left            =   -74880
                  TabIndex        =   89
                  Top             =   360
                  WhatsThisHelpID =   202
                  Width           =   2445
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "Options..."
                  Height          =   375
                  Left            =   -72720
                  TabIndex        =   88
                  Top             =   600
                  WhatsThisHelpID =   204
                  Width           =   1455
               End
               Begin VB.CommandButton Command8 
                  Height          =   300
                  HelpContextID   =   168
                  Left            =   -69650
                  Picture         =   "rename.frx":091C
                  Style           =   1  'Graphical
                  TabIndex        =   87
                  ToolTipText     =   "Use cyclic selection"
                  Top             =   495
                  UseMaskColor    =   -1  'True
                  Visible         =   0   'False
                  WhatsThisHelpID =   198
                  Width           =   300
               End
               Begin VB.PictureBox onglcounter 
                  AutoRedraw      =   -1  'True
                  BorderStyle     =   0  'None
                  Height          =   705
                  Left            =   120
                  ScaleHeight     =   705
                  ScaleWidth      =   5475
                  TabIndex        =   78
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   5475
                  Begin VB.PictureBox Picture2 
                     BorderStyle     =   0  'None
                     Height          =   255
                     Left            =   210
                     ScaleHeight     =   255
                     ScaleWidth      =   4815
                     TabIndex        =   80
                     Top             =   420
                     WhatsThisHelpID =   188
                     Width           =   4815
                     Begin VB.OptionButton Option3 
                        Caption         =   "Replace prefix"
                        Height          =   255
                        Index           =   2
                        Left            =   3375
                        TabIndex        =   83
                        ToolTipText     =   "Counter will replace file's prefixe"
                        Top             =   0
                        WhatsThisHelpID =   188
                        Width           =   2295
                     End
                     Begin VB.OptionButton Option3 
                        Caption         =   "Add to the right"
                        Height          =   255
                        Index           =   1
                        Left            =   1665
                        TabIndex        =   82
                        ToolTipText     =   "Counter will be added to the right of prefix"
                        Top             =   0
                        WhatsThisHelpID =   188
                        Width           =   1395
                     End
                     Begin VB.OptionButton Option3 
                        Caption         =   "Add to the left"
                        Height          =   255
                        Index           =   0
                        Left            =   15
                        TabIndex        =   81
                        ToolTipText     =   "Counter will be added to the left of prefix"
                        Top             =   0
                        Value           =   -1  'True
                        WhatsThisHelpID =   188
                        Width           =   1455
                     End
                  End
                  Begin VB.ComboBox Combo3 
                     Height          =   315
                     ItemData        =   "rename.frx":0A06
                     Left            =   4095
                     List            =   "rename.frx":0A1C
                     Style           =   2  'Dropdown List
                     TabIndex        =   79
                     Top             =   35
                     WhatsThisHelpID =   187
                     Width           =   1300
                  End
                  Begin THERename.LabelText Text3 
                     Height          =   285
                     Left            =   180
                     TabIndex        =   84
                     ToolTipText     =   "Enter begin's value"
                     Top             =   36
                     Width           =   1005
                     _ExtentX        =   1773
                     _ExtentY        =   503
                     Caption         =   "Begin"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     LabelWidth      =   550
                     MousePointer    =   0
                     Text            =   "1"
                     BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     TextAlignement  =   1
                  End
                  Begin THERename.LabelText Text4 
                     Height          =   285
                     Left            =   1455
                     TabIndex        =   85
                     ToolTipText     =   "Enter increment value (1 for example)"
                     Top             =   35
                     Width           =   945
                     _ExtentX        =   1667
                     _ExtentY        =   503
                     Caption         =   "Step"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     LabelWidth      =   500
                     MousePointer    =   0
                     Text            =   "1"
                     BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     TextAlignement  =   1
                  End
                  Begin THERename.LabelText Text5 
                     Height          =   285
                     Left            =   2730
                     TabIndex        =   86
                     ToolTipText     =   "Enter number of digits for the counter"
                     Top             =   35
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   503
                     Caption         =   "Digits"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     LabelWidth      =   550
                     MousePointer    =   0
                     Text            =   "4"
                     BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     TextAlignement  =   1
                  End
               End
            End
         End
         Begin VB.CommandButton Command3 
            Caption         =   "="
            Height          =   315
            Left            =   5340
            TabIndex        =   74
            ToolTipText     =   "Use the same option for prefix and extension"
            Top             =   300
            WhatsThisHelpID =   177
            Width           =   255
         End
      End
      Begin THERename.MyFrame Frame2 
         Height          =   2400
         Left            =   120
         TabIndex        =   35
         Top             =   3060
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4233
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         Caption         =   "Extension"
         ShowBorderInDesignMode=   0   'False
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   60
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   72
            ToolTipText     =   "Scroll to select action to perform on extension"
            Top             =   300
            WhatsThisHelpID =   176
            Width           =   5175
         End
         Begin VB.PictureBox PanelExt 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1650
            Left            =   30
            ScaleHeight     =   1650
            ScaleWidth      =   5775
            TabIndex        =   37
            Top             =   720
            Width           =   5775
            Begin VB.CommandButton Command11 
               Caption         =   "Reset"
               Height          =   300
               HelpContextID   =   110
               Left            =   4800
               TabIndex        =   38
               ToolTipText     =   "Reset all selections to their default's value"
               Top             =   1350
               WhatsThisHelpID =   181
               Width           =   870
            End
            Begin TabDlg.SSTab SSTab2 
               Height          =   1305
               Left            =   60
               TabIndex        =   39
               Top             =   0
               Width           =   5625
               _ExtentX        =   9922
               _ExtentY        =   2302
               _Version        =   393216
               Style           =   1
               Tabs            =   5
               TabsPerRow      =   5
               TabHeight       =   520
               TabCaption(0)   =   "Counter"
               TabPicture(0)   =   "rename.frx":0A55
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Check11"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "onglcounter2"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).ControlCount=   2
               TabCaption(1)   =   "Size"
               TabPicture(1)   =   "rename.frx":0A71
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Picture13"
               Tab(1).Control(1)=   "Check12"
               Tab(1).ControlCount=   2
               TabCaption(2)   =   "Date"
               TabPicture(2)   =   "rename.frx":0A8D
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Picture14"
               Tab(2).Control(1)=   "Check13"
               Tab(2).ControlCount=   2
               TabCaption(3)   =   "Time"
               TabPicture(3)   =   "rename.frx":0AA9
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Picture4"
               Tab(3).Control(1)=   "Check4"
               Tab(3).ControlCount=   2
               TabCaption(4)   =   "Text "
               TabPicture(4)   =   "rename.frx":0AC5
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "Picture3"
               Tab(4).Control(0).Enabled=   0   'False
               Tab(4).Control(1)=   "Option4(1)"
               Tab(4).Control(1).Enabled=   0   'False
               Tab(4).Control(2)=   "Option4(0)"
               Tab(4).Control(2).Enabled=   0   'False
               Tab(4).Control(3)=   "Text15"
               Tab(4).Control(3).Enabled=   0   'False
               Tab(4).Control(4)=   "Text8"
               Tab(4).Control(4).Enabled=   0   'False
               Tab(4).ControlCount=   5
               Begin VB.PictureBox onglcounter2 
                  AutoRedraw      =   -1  'True
                  BorderStyle     =   0  'None
                  Height          =   580
                  Left            =   120
                  ScaleHeight     =   585
                  ScaleWidth      =   5475
                  TabIndex        =   63
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   5475
                  Begin VB.ComboBox Combo4 
                     Height          =   315
                     ItemData        =   "rename.frx":0AE1
                     Left            =   4095
                     List            =   "rename.frx":0AF7
                     Style           =   2  'Dropdown List
                     TabIndex        =   68
                     Top             =   35
                     WhatsThisHelpID =   187
                     Width           =   1300
                  End
                  Begin VB.PictureBox Picture11 
                     BorderStyle     =   0  'None
                     Height          =   255
                     Left            =   165
                     ScaleHeight     =   255
                     ScaleWidth      =   5055
                     TabIndex        =   64
                     Top             =   360
                     WhatsThisHelpID =   188
                     Width           =   5055
                     Begin VB.OptionButton Option3 
                        Caption         =   "Replace extension"
                        Height          =   255
                        Index           =   24
                        Left            =   3375
                        TabIndex        =   67
                        ToolTipText     =   "Counter will replace file's prefixe"
                        Top             =   0
                        WhatsThisHelpID =   188
                        Width           =   2295
                     End
                     Begin VB.OptionButton Option3 
                        Caption         =   "Add to the right"
                        Height          =   255
                        Index           =   25
                        Left            =   1665
                        TabIndex        =   66
                        ToolTipText     =   "Counter will be added to the right of extension"
                        Top             =   0
                        WhatsThisHelpID =   188
                        Width           =   1440
                     End
                     Begin VB.OptionButton Option3 
                        Caption         =   "Add to the left"
                        Height          =   255
                        Index           =   26
                        Left            =   15
                        TabIndex        =   65
                        ToolTipText     =   "Counter will be added to the left of extension"
                        Top             =   0
                        Value           =   -1  'True
                        WhatsThisHelpID =   188
                        Width           =   1365
                     End
                  End
                  Begin THERename.LabelText Text16 
                     Height          =   285
                     Left            =   180
                     TabIndex        =   69
                     ToolTipText     =   "Enter begin's value"
                     Top             =   35
                     Width           =   1005
                     _ExtentX        =   1773
                     _ExtentY        =   503
                     Caption         =   "Begin"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     LabelWidth      =   550
                     MousePointer    =   0
                     Text            =   "1"
                     BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     TextAlignement  =   1
                  End
                  Begin THERename.LabelText Text17 
                     Height          =   285
                     Left            =   1455
                     TabIndex        =   70
                     ToolTipText     =   "Enter increment value (1 for example)"
                     Top             =   35
                     Width           =   945
                     _ExtentX        =   1667
                     _ExtentY        =   503
                     Caption         =   "Step"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     LabelWidth      =   500
                     MousePointer    =   0
                     Text            =   "1"
                     BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     TextAlignement  =   1
                  End
                  Begin THERename.LabelText Text18 
                     Height          =   285
                     Left            =   2730
                     TabIndex        =   71
                     ToolTipText     =   "Enter number of digits for the counter"
                     Top             =   35
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   503
                     Caption         =   "Digits"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     LabelWidth      =   550
                     MousePointer    =   0
                     Text            =   "4"
                     BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     TextAlignement  =   1
                  End
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Left            =   -73200
                  TabIndex        =   62
                  ToolTipText     =   "Current extension will be replace with this text"
                  Top             =   480
                  Visible         =   0   'False
                  WhatsThisHelpID =   197
                  Width           =   2775
               End
               Begin VB.TextBox Text15 
                  Height          =   285
                  Left            =   -73200
                  TabIndex        =   61
                  ToolTipText     =   "Add a text to the extension"
                  Top             =   840
                  Visible         =   0   'False
                  WhatsThisHelpID =   200
                  Width           =   2055
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "Replace with text"
                  Height          =   195
                  Index           =   0
                  Left            =   -74880
                  TabIndex        =   60
                  ToolTipText     =   "Click to replace or not file's extension with a text"
                  Top             =   525
                  WhatsThisHelpID =   196
                  Width           =   1635
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "Add text"
                  Height          =   195
                  Index           =   1
                  Left            =   -74880
                  TabIndex        =   59
                  ToolTipText     =   "Click to add or remove a text to the extension"
                  Top             =   870
                  WhatsThisHelpID =   199
                  Width           =   945
               End
               Begin VB.PictureBox Picture3 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -71070
                  ScaleHeight     =   240
                  ScaleWidth      =   1455
                  TabIndex        =   56
                  Top             =   900
                  Width           =   1455
                  Begin VB.OptionButton Option5 
                     Caption         =   "Begin"
                     Height          =   195
                     Index           =   0
                     Left            =   0
                     TabIndex        =   58
                     ToolTipText     =   "Text will be added to the left of extension"
                     Top             =   0
                     Value           =   -1  'True
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   735
                  End
                  Begin VB.OptionButton Option5 
                     Caption         =   "End"
                     Height          =   195
                     Index           =   1
                     Left            =   720
                     TabIndex        =   57
                     ToolTipText     =   "Text will be added at the end of extension"
                     Top             =   0
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   615
                  End
               End
               Begin VB.CheckBox Check11 
                  Caption         =   "Add a counter"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   55
                  ToolTipText     =   "Click to add a counter"
                  Top             =   360
                  WhatsThisHelpID =   183
                  Width           =   1320
               End
               Begin VB.PictureBox Picture13 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -74640
                  ScaleHeight     =   240
                  ScaleWidth      =   5055
                  TabIndex        =   51
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   190
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   29
                     Left            =   15
                     TabIndex        =   54
                     ToolTipText     =   "Click to add file's size to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   190
                     Width           =   1365
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   28
                     Left            =   1650
                     TabIndex        =   53
                     ToolTipText     =   "Click to add file's size to the right"
                     Top             =   0
                     WhatsThisHelpID =   190
                     Width           =   1395
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace extension"
                     Height          =   255
                     Index           =   27
                     Left            =   3375
                     TabIndex        =   52
                     ToolTipText     =   "Click to replace file's extension with file's size"
                     Top             =   0
                     WhatsThisHelpID =   190
                     Width           =   2295
                  End
               End
               Begin VB.PictureBox Picture14 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   -74640
                  ScaleHeight     =   285
                  ScaleWidth      =   5055
                  TabIndex        =   47
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   192
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace extension"
                     Height          =   255
                     Index           =   32
                     Left            =   3375
                     TabIndex        =   50
                     ToolTipText     =   "Click to replace file's extension with file's date"
                     Top             =   0
                     WhatsThisHelpID =   192
                     Width           =   2295
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   31
                     Left            =   1650
                     TabIndex        =   49
                     ToolTipText     =   "Click to add file's date to the right"
                     Top             =   0
                     WhatsThisHelpID =   192
                     Width           =   1440
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   30
                     Left            =   15
                     TabIndex        =   48
                     ToolTipText     =   "Click to add file's date to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   192
                     Width           =   1365
                  End
               End
               Begin VB.CheckBox Check13 
                  Caption         =   "Add file's date"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   46
                  ToolTipText     =   "Click to add file's date"
                  Top             =   360
                  WhatsThisHelpID =   191
                  Width           =   1320
               End
               Begin VB.PictureBox Picture4 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -74640
                  ScaleHeight     =   240
                  ScaleWidth      =   5055
                  TabIndex        =   42
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   194
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   14
                     Left            =   15
                     TabIndex        =   45
                     ToolTipText     =   "File's time will be added to to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   194
                     Width           =   1365
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   13
                     Left            =   1650
                     TabIndex        =   44
                     ToolTipText     =   "File's time will be added to to the right"
                     Top             =   0
                     WhatsThisHelpID =   194
                     Width           =   1440
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace extension"
                     Height          =   255
                     Index           =   12
                     Left            =   3375
                     TabIndex        =   43
                     ToolTipText     =   "Click to replace extension will file's time"
                     Top             =   0
                     WhatsThisHelpID =   194
                     Width           =   2295
                  End
               End
               Begin VB.CheckBox Check4 
                  Caption         =   "Add file's time"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   41
                  ToolTipText     =   "Click to add file's time"
                  Top             =   360
                  WhatsThisHelpID =   193
                  Width           =   1275
               End
               Begin VB.CheckBox Check12 
                  Caption         =   "Add file's size"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   40
                  ToolTipText     =   "Click to add file's size"
                  Top             =   360
                  WhatsThisHelpID =   189
                  Width           =   1320
               End
            End
         End
         Begin VB.CommandButton Command4 
            Caption         =   "="
            Height          =   315
            Left            =   5340
            TabIndex        =   36
            ToolTipText     =   "Use the same option for prefix and extension"
            Top             =   300
            WhatsThisHelpID =   177
            Width           =   255
         End
      End
      Begin THERename.cpvPicScroll Acdsee 
         Height          =   4635
         Left            =   -74940
         TabIndex        =   31
         ToolTipText     =   "Double click to open the picture viewer"
         Top             =   360
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   8176
         BorderStyle     =   1
         Picture         =   "rename.frx":0B30
      End
      Begin VB.TextBox Text1 
         Height          =   4335
         Left            =   -74940
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         ScrollBars      =   3  'Both
         TabIndex        =   30
         Top             =   360
         Width           =   5835
      End
      Begin MSComctlLib.ListView LvMP3 
         Height          =   5820
         Left            =   -74940
         TabIndex        =   29
         Top             =   360
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   10266
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tag"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox FrameDroite 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   120
         ScaleHeight     =   5775
         ScaleWidth      =   5955
         TabIndex        =   27
         Top             =   400
         Width           =   5955
         Begin THERename.MyFrame Frame3 
            Height          =   705
            Left            =   0
            TabIndex        =   32
            Top             =   5040
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1244
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            Caption         =   "Sample"
            ShowBorderInDesignMode=   0   'False
            Begin VB.Label laides 
               BackStyle       =   0  'Transparent
               Height          =   210
               Left            =   120
               TabIndex        =   34
               ToolTipText     =   "Sample of what extension will be"
               Top             =   450
               WhatsThisHelpID =   178
               Width           =   5655
            End
            Begin VB.Label laidep 
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   120
               TabIndex        =   33
               ToolTipText     =   "Sample of what prefix will be"
               Top             =   240
               WhatsThisHelpID =   178
               Width           =   5655
            End
         End
      End
      Begin CCRPFolderTV6.FolderTreeview FolderTreeview1 
         Height          =   5820
         Index           =   0
         Left            =   -74940
         TabIndex        =   28
         Top             =   360
         WhatsThisHelpID =   172
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   10266
         IntegralHeight  =   0   'False
      End
   End
   Begin VB.PictureBox PanelList 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4275
      Left            =   5895
      ScaleHeight     =   4275
      ScaleWidth      =   5730
      TabIndex        =   8
      Top             =   7440
      Visible         =   0   'False
      Width           =   5730
      Begin VB.CommandButton Command2 
         Caption         =   "Edit..."
         Height          =   300
         Left            =   4980
         TabIndex        =   123
         ToolTipText     =   "Edit the list to change names manually"
         Top             =   3900
         WhatsThisHelpID =   217
         Width           =   750
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3855
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "Press F2 to edit filename"
         Top             =   0
         WhatsThisHelpID =   212
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   6800
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "New Name"
            Object.Width           =   5080
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Original name"
            Object.Width           =   5080
         EndProperty
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Copy all"
         Height          =   300
         Left            =   4190
         TabIndex        =   13
         ToolTipText     =   "Copy all files from files list"
         Top             =   3900
         WhatsThisHelpID =   217
         Width           =   750
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Copy selected"
         Height          =   300
         Left            =   2900
         TabIndex        =   12
         ToolTipText     =   "Copy selected files from the files list"
         Top             =   3900
         WhatsThisHelpID =   216
         Width           =   1250
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Remove"
         Height          =   300
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   "Remove selected items from list"
         Top             =   3900
         WhatsThisHelpID =   213
         Width           =   800
      End
      Begin VB.CommandButton Command12 
         Caption         =   "&Open a list"
         Height          =   300
         Left            =   840
         TabIndex        =   10
         ToolTipText     =   "Open an existing list"
         Top             =   3900
         WhatsThisHelpID =   214
         Width           =   980
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Sa&ve list..."
         Height          =   300
         Left            =   1860
         TabIndex        =   9
         ToolTipText     =   "Save list to a text file"
         Top             =   3900
         WhatsThisHelpID =   215
         Width           =   1000
      End
   End
   Begin VB.PictureBox paneltext 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   11760
      ScaleHeight     =   1365
      ScaleWidth      =   5295
      TabIndex        =   16
      Top             =   7560
      Visible         =   0   'False
      Width           =   5295
      Begin THERename.LabelText Text9 
         Height          =   285
         Left            =   0
         TabIndex        =   17
         ToolTipText     =   "Enter begin's value"
         Top             =   80
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   503
         Caption         =   "Number of characters to take from file"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         LabelWidth      =   2700
         MousePointer    =   0
         Text            =   "25"
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignement  =   1
      End
      Begin VB.Label Label10 
         Caption         =   $"rename.frx":0B4C
         Height          =   930
         Left            =   0
         TabIndex        =   18
         Top             =   405
         Width           =   5265
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      WhatsThisHelpID =   174
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select all"
            Description     =   "Select all files"
            Object.ToolTipText     =   "Select all files in current directory"
            Object.Tag             =   "Select all files in current directory"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Fie1"
                  Text            =   "Fie1"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "File2"
                  Object.Tag             =   "Fie2"
                  Text            =   "Fie2"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "File3"
                  Object.Tag             =   "Fie3"
                  Text            =   "Fie3"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "unselect"
            Description     =   "Unselect files from current directory"
            Object.ToolTipText     =   "Unselect files from current directory"
            Object.Tag             =   "Unselect files from current directory"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "invert selection"
            Description     =   "Invert selection"
            Object.ToolTipText     =   "Invert selection"
            Object.Tag             =   "Invert selection"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step"
            Description     =   "Enter a step value for selection"
            Object.ToolTipText     =   "Enter a step value for selection"
            Object.Tag             =   "Enter a step value for selection"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Start"
            Description     =   "Start to rename files"
            Object.ToolTipText     =   "Start to rename files"
            Object.Tag             =   "Start to rename files"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Description     =   "Preview what the filenames will be"
            Object.ToolTipText     =   "Preview what the filenames will be"
            Object.Tag             =   "Preview what the filenames will be"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Manually"
            Description     =   "Enables you to rename files manually"
            Object.ToolTipText     =   "Enables you to rename files manually (F2)"
            Object.Tag             =   "Enables you to rename files manually"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DropFiles"
            Description     =   "Drop files to copy them, click to configure"
            Object.ToolTipText     =   "Drop files to copy them, click to configure"
            Object.Tag             =   "Drop files to copy them, click to configure"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Recursive mode"
            Description     =   "Recursive mode"
            Object.ToolTipText     =   "Show files in sub folders"
            Object.Tag             =   "Show files in sub folders"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up one level"
            Description     =   "Up one level"
            Object.ToolTipText     =   "Up one level"
            Object.Tag             =   "Up one level"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Root directory"
            Description     =   "Root directory"
            Object.ToolTipText     =   "Root directory"
            Object.Tag             =   "Root directory"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Show history of your moves"
            Description     =   "Show history of your moves"
            Object.ToolTipText     =   "Show history of your moves"
            Object.Tag             =   "Show history of your moves"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add to your favorites"
            Object.ToolTipText     =   "Add to your favorites"
            Object.Tag             =   "Add to your favorites"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "organize your favorites"
            Object.ToolTipText     =   "Organize your favorites"
            Object.Tag             =   "organize your favorites"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "First favorite"
            Object.ToolTipText     =   "First favorite"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previous favorite"
            Object.ToolTipText     =   "previous favorite"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next favorite"
            Object.ToolTipText     =   "Next favorite"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Last favorite"
            Object.ToolTipText     =   "Last favorite"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help Pointer"
            Object.ToolTipText     =   "Enable Help Pointer"
            Object.Tag             =   "Enable Help Pointer"
            ImageIndex      =   20
         EndProperty
      EndProperty
      OLEDropMode     =   1
      Begin VB.ComboBox Combo5 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "rename.frx":0C6F
         Left            =   7755
         List            =   "rename.frx":0C71
         TabIndex        =   1
         ToolTipText     =   "Type a file filter and press enter or select a filter from the list"
         Top             =   20
         WhatsThisHelpID =   182
         Width           =   1455
      End
   End
   Begin VB.ListBox lhistory 
      Enabled         =   0   'False
      Height          =   285
      IntegralHeight  =   0   'False
      Left            =   12000
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ListBox List3 
      Height          =   255
      ItemData        =   "rename.frx":0C73
      Left            =   12000
      List            =   "rename.frx":0C7A
      TabIndex        =   6
      Top             =   1380
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ListBox List2 
      Height          =   255
      ItemData        =   "rename.frx":0C8F
      Left            =   12195
      List            =   "rename.frx":0C96
      TabIndex        =   5
      Top             =   -45
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "rename.frx":0CAE
      Left            =   12000
      List            =   "rename.frx":0CB5
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1320
   End
   Begin MSComctlLib.StatusBar état 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   6825
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12832
            MinWidth        =   1058
            Object.ToolTipText     =   "Program information"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Key             =   "Position"
            Object.ToolTipText     =   "File that is rename"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Object.ToolTipText     =   "Number of files in directory"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Object.ToolTipText     =   "Number of files you have selected"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   1411
            TextSave        =   "04/12/2002"
            Object.ToolTipText     =   "system date"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   979
            MinWidth        =   970
            TextSave        =   "19:38"
            Object.ToolTipText     =   "system time"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   285
      Left            =   12000
      TabIndex        =   15
      Top             =   780
      Visible         =   0   'False
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   503
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Action"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Created"
         Object.Width           =   1586
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Attrib."
         Object.Width           =   706
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12000
      Top             =   1980
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
            Picture         =   "rename.frx":0CD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":122D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":1781
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":1CD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":2229
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":277D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":2CD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":3225
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":3779
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":3CCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":4221
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":4775
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":4CC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":521D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":5771
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":5CC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":6219
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":676D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":6CC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rename.frx":7215
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5805
      Left            =   60
      TabIndex        =   0
      Top             =   480
      WhatsThisHelpID =   173
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10239
      View            =   3
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Created"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Attrib."
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Type"
         Object.Width           =   706
      EndProperty
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      HelpContextID   =   13
      Begin VB.Menu mopenset 
         Caption         =   "&Open settings..."
         HelpContextID   =   50
         Shortcut        =   ^O
      End
      Begin VB.Menu msave 
         Caption         =   "&Save settings"
         HelpContextID   =   51
         Shortcut        =   ^S
      End
      Begin VB.Menu msaveas 
         Caption         =   "Save settings &as..."
         HelpContextID   =   52
         Shortcut        =   {F12}
      End
      Begin VB.Menu msep29 
         Caption         =   "-"
      End
      Begin VB.Menu mprintdir 
         Caption         =   "&Print directory"
         HelpContextID   =   57
         Shortcut        =   ^P
      End
      Begin VB.Menu HTMLReport 
         Caption         =   "Generate &HTML report..."
         HelpContextID   =   248
      End
      Begin VB.Menu mexport 
         Caption         =   "&Export tags && information..."
         HelpContextID   =   489
      End
      Begin VB.Menu mfilefind 
         Caption         =   "File &Find..."
         HelpContextID   =   60
         Shortcut        =   ^K
      End
      Begin VB.Menu mgreat 
         Caption         =   "Find bi&ggest counter"
         HelpContextID   =   44
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu mend 
         Caption         =   "E&xit"
         HelpContextID   =   17
         Index           =   0
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu medit 
      Caption         =   "&Edit"
      HelpContextID   =   143
      Begin VB.Menu msearchpref 
         Caption         =   "Search and replace in &prefix..."
         HelpContextID   =   110
         Shortcut        =   {F3}
      End
      Begin VB.Menu msearchext 
         Caption         =   "Search and replace in &extension..."
         HelpContextID   =   110
         Shortcut        =   {F4}
      End
      Begin VB.Menu mglob 
         Caption         =   "&Global Search && Replace..."
         HelpContextID   =   110
         Shortcut        =   ^H
      End
      Begin VB.Menu mregrenam2 
         Caption         =   "RegE&xp rename..."
         Enabled         =   0   'False
         HelpContextID   =   303
      End
      Begin VB.Menu msep11 
         Caption         =   "-"
      End
      Begin VB.Menu mabrev 
         Caption         =   "A&bbreviations..."
         HelpContextID   =   266
      End
      Begin VB.Menu mrules 
         Caption         =   "&Rules..."
         HelpContextID   =   472
         Shortcut        =   ^R
      End
      Begin VB.Menu marrays 
         Caption         =   "Arra&ys..."
         Visible         =   0   'False
      End
      Begin VB.Menu msep50 
         Caption         =   "-"
      End
      Begin VB.Menu mundo 
         Caption         =   "&Undo last rename"
         HelpContextID   =   16
         Shortcut        =   ^Z
      End
      Begin VB.Menu msep6 
         Caption         =   "-"
      End
      Begin VB.Menu mdatetime 
         Caption         =   "Change file's &date and time..."
         HelpContextID   =   42
         Shortcut        =   ^D
      End
      Begin VB.Menu mattrib 
         Caption         =   "Change files a&ttributes..."
         HelpContextID   =   38
      End
      Begin VB.Menu msep58 
         Caption         =   "-"
      End
      Begin VB.Menu m1selectAll 
         Caption         =   "Select &All"
         HelpContextID   =   149
         Shortcut        =   ^A
      End
      Begin VB.Menu M1Unselect 
         Caption         =   "U&nselect"
         HelpContextID   =   150
      End
      Begin VB.Menu M1Invert 
         Caption         =   "&Invert selection"
         HelpContextID   =   151
      End
      Begin VB.Menu M1Step 
         Caption         =   "&Step"
         HelpContextID   =   152
      End
      Begin VB.Menu msep501 
         Caption         =   "-"
      End
      Begin VB.Menu mfold 
         Caption         =   "&Folders"
         HelpContextID   =   477
         Begin VB.Menu mgo1 
            Caption         =   "Go to the &first folder in the hierarchy"
            HelpContextID   =   477
            Shortcut        =   +{F1}
         End
         Begin VB.Menu mgonext 
            Caption         =   "Go to the &next folder in the hierarchy"
            HelpContextID   =   477
            Shortcut        =   +{F2}
         End
         Begin VB.Menu mgoprev 
            Caption         =   "Go to the &previous folder in the hierarchy"
            HelpContextID   =   477
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mgolast 
            Caption         =   "Go to the l&ast folder in the hierarchy"
            HelpContextID   =   477
            Shortcut        =   +{F4}
         End
      End
   End
   Begin VB.Menu mview 
      Caption         =   "&View"
      HelpContextID   =   157
      Begin VB.Menu Mrefresh 
         Caption         =   "Refres&h"
         HelpContextID   =   158
         Shortcut        =   {F5}
      End
      Begin VB.Menu msep30 
         Caption         =   "-"
      End
      Begin VB.Menu moptions 
         Caption         =   "Op&tions..."
         HelpContextID   =   4
         Shortcut        =   ^T
      End
      Begin VB.Menu minfos 
         Caption         =   "&Information..."
         HelpContextID   =   15
         Shortcut        =   ^I
      End
      Begin VB.Menu msep99 
         Caption         =   "-"
      End
      Begin VB.Menu mhistory 
         Caption         =   "&History..."
         HelpContextID   =   59
      End
      Begin VB.Menu mviewlog 
         Caption         =   "&Log file..."
         HelpContextID   =   479
      End
      Begin VB.Menu mchangetab 
         Caption         =   "&Change tab"
         HelpContextID   =   285
         Shortcut        =   {F9}
      End
      Begin VB.Menu mprevioustab 
         Caption         =   "&Previous tab"
         HelpContextID   =   75
         Shortcut        =   +{F9}
      End
      Begin VB.Menu mviewtabs 
         Caption         =   "&Tabs"
         Begin VB.Menu mviewmp3tab 
            Caption         =   "&Tags tab"
         End
         Begin VB.Menu mviewpicturetab 
            Caption         =   "&Pictures tab"
         End
         Begin VB.Menu mviewtexttab 
            Caption         =   "Te&xt tab"
         End
      End
   End
   Begin VB.Menu mrun 
      Caption         =   "&Run"
      HelpContextID   =   148
      Begin VB.Menu M2Start 
         Caption         =   "&Start"
         HelpContextID   =   153
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu m2preview 
         Caption         =   "&Preview"
         HelpContextID   =   154
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu m2manually 
         Caption         =   "&Rename manually"
         HelpContextID   =   155
      End
      Begin VB.Menu msep59 
         Caption         =   "-"
      End
      Begin VB.Menu mexecmd 
         Caption         =   "&Execute a command on file(s)..."
         HelpContextID   =   537
      End
      Begin VB.Menu m2recursive 
         Caption         =   "Recursive &mode (show files in sub folders)"
         HelpContextID   =   156
      End
   End
   Begin VB.Menu mdisk 
      Caption         =   "&Disk"
      HelpContextID   =   61
      Begin VB.Menu mshformat 
         Caption         =   "&Format drive/floppy..."
         HelpContextID   =   62
         Shortcut        =   ^Y
      End
      Begin VB.Menu msetvolabel 
         Caption         =   "&Set volume label..."
         HelpContextID   =   63
         Shortcut        =   ^L
      End
      Begin VB.Menu msepdsk1 
         Caption         =   "-"
      End
      Begin VB.Menu mmap 
         Caption         =   "&Map Network drive..."
         HelpContextID   =   64
         Shortcut        =   ^M
      End
      Begin VB.Menu mdisconnect 
         Caption         =   "&Disconnect Network drive..."
         HelpContextID   =   65
      End
      Begin VB.Menu mmapped_drives 
         Caption         =   "Mapped Dri&ves"
         HelpContextID   =   161
      End
   End
   Begin VB.Menu mfavorites 
      Caption         =   "Favor&ites"
      HelpContextID   =   18
      Begin VB.Menu madddirectory 
         Caption         =   "Add this directory to your favorites"
         HelpContextID   =   19
         Shortcut        =   ^E
      End
      Begin VB.Menu morganyze 
         Caption         =   "Organize your favorites..."
         HelpContextID   =   20
         Shortcut        =   ^G
      End
      Begin VB.Menu msep 
         Caption         =   "-"
      End
      Begin VB.Menu menufav 
         Caption         =   "mnufav"
         Index           =   0
      End
   End
   Begin VB.Menu mcontextuel 
      Caption         =   "Contextuel"
      Visible         =   0   'False
      Begin VB.Menu maction 
         Caption         =   "A&ctions"
         Begin VB.Menu mdelete 
            Caption         =   "&Delete file(s)"
         End
         Begin VB.Menu mopen 
            Caption         =   "&Open"
         End
         Begin VB.Menu mpropertyes 
            Caption         =   "P&roperties"
         End
         Begin VB.Menu mprint 
            Caption         =   "&Print"
         End
      End
      Begin VB.Menu mcreatefold 
         Caption         =   "Create folders with names..."
      End
      Begin VB.Menu mcreatfoldman 
         Caption         =   "Create a list of folders manually..."
      End
      Begin VB.Menu msepchg2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuswap 
         Caption         =   "Swap filenames"
         Visible         =   0   'False
      End
      Begin VB.Menu mregrename 
         Caption         =   "Re&gExp rename..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mchg1 
         Caption         =   "Change files date and time now..."
      End
      Begin VB.Menu mchg2 
         Caption         =   "Changes files attributes now..."
      End
      Begin VB.Menu msepchg 
         Caption         =   "-"
      End
      Begin VB.Menu mexec 
         Caption         =   "E&xecute dialog"
      End
      Begin VB.Menu mexplorer 
         Caption         =   "&Explorer"
      End
      Begin VB.Menu msep5 
         Caption         =   "-"
      End
      Begin VB.Menu msearch 
         Caption         =   "Searc&h..."
      End
      Begin VB.Menu mremdisp 
         Caption         =   "Remove from display"
      End
      Begin VB.Menu madd 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mcopy 
         Caption         =   "Cop&y name(s)..."
      End
      Begin VB.Menu msepcont1 
         Caption         =   "-"
      End
      Begin VB.Menu mmove 
         Caption         =   "Move in list"
         Begin VB.Menu mfirst 
            Caption         =   "&First"
         End
         Begin VB.Menu mmiddle 
            Caption         =   "&Middle"
         End
         Begin VB.Menu mlast 
            Caption         =   "&Last"
         End
      End
   End
   Begin VB.Menu mbag 
      Caption         =   "&Bin"
      HelpContextID   =   66
      Begin VB.Menu mcopybag 
         Caption         =   "&Copy to bin"
         HelpContextID   =   66
         Shortcut        =   ^C
      End
      Begin VB.Menu maddbag 
         Caption         =   "Add to bin"
         HelpContextID   =   66
         Shortcut        =   {F6}
      End
      Begin VB.Menu mbagsep4 
         Caption         =   "-"
      End
      Begin VB.Menu mcutbag 
         Caption         =   "Cu&t to bin"
         HelpContextID   =   66
         Shortcut        =   ^X
      End
      Begin VB.Menu mcutadditive 
         Caption         =   "Cut additive"
         HelpContextID   =   66
         Shortcut        =   {F7}
      End
      Begin VB.Menu mbagsep3 
         Caption         =   "-"
      End
      Begin VB.Menu mpastebag 
         Caption         =   "Paste from bin"
         HelpContextID   =   66
         Shortcut        =   ^V
      End
      Begin VB.Menu mpastekeep 
         Caption         =   "Paste and keep"
         HelpContextID   =   66
         Shortcut        =   {F8}
      End
      Begin VB.Menu mbagsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mclearbag 
         Caption         =   "Cl&ear bin"
         HelpContextID   =   66
         Shortcut        =   ^B
      End
      Begin VB.Menu mviewbag 
         Caption         =   "&View bin..."
         HelpContextID   =   66
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mhelp 
      Caption         =   "&Help"
      Begin VB.Menu mindex 
         Caption         =   "&Index"
      End
      Begin VB.Menu msep15 
         Caption         =   "-"
      End
      Begin VB.Menu mapropos 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mcontext2 
      Caption         =   "contexteul 2"
      Visible         =   0   'False
      Begin VB.Menu mgendir 
         Caption         =   "Directory"
         Begin VB.Menu mrendirect 
            Caption         =   "&Rename directory..."
         End
         Begin VB.Menu mmakedir 
            Caption         =   "Ne&w directory..."
         End
         Begin VB.Menu mdgroupe 
            Caption         =   "Create a &group of directories..."
         End
         Begin VB.Menu mdelrep 
            Caption         =   "&Delete directory"
         End
         Begin VB.Menu mchgattdir 
            Caption         =   "&Change attributes..."
         End
      End
      Begin VB.Menu mgoto 
         Caption         =   "&Go to..."
      End
      Begin VB.Menu mprop2 
         Caption         =   "Proper&ties"
      End
      Begin VB.Menu msep55 
         Caption         =   "-"
      End
      Begin VB.Menu maddfavorites 
         Caption         =   "&Add to your favorites"
      End
      Begin VB.Menu myourfav 
         Caption         =   "Your &favorites"
         Begin VB.Menu mnufav 
            Caption         =   "mnufav"
            Index           =   0
         End
      End
      Begin VB.Menu mmdrives 
         Caption         =   "Dri&ves"
         Begin VB.Menu mnudrives 
            Caption         =   "mnudrives"
            Index           =   0
         End
      End
      Begin VB.Menu msep31 
         Caption         =   "-"
      End
      Begin VB.Menu mstartup 
         Caption         =   "Set as startup director&y"
      End
      Begin VB.Menu mdosprompthere 
         Caption         =   "Command prompt &here"
      End
      Begin VB.Menu mcopy2 
         Caption         =   "Copy &name"
      End
   End
   Begin VB.Menu m3contextuel 
      Caption         =   "Free Form"
      Visible         =   0   'False
      Begin VB.Menu m3prefix 
         Caption         =   "Prefix"
         Begin VB.Menu m3cmdprefix 
            Caption         =   "cmdprefix"
            Index           =   0
         End
      End
      Begin VB.Menu m3extension 
         Caption         =   "Extension"
         Begin VB.Menu m3cmdextension 
            Caption         =   "cmdextension"
            Index           =   0
         End
      End
      Begin VB.Menu m3general 
         Caption         =   "General"
         Begin VB.Menu mlang 
            Caption         =   "langage"
            Index           =   0
         End
      End
      Begin VB.Menu mmusic 
         Caption         =   "Music"
         Begin VB.Menu mimusic 
            Caption         =   "music"
            Index           =   0
         End
      End
      Begin VB.Menu myourcmd 
         Caption         =   "Your commands"
         Begin VB.Menu yourcmd 
            Caption         =   "yourcmd"
            Index           =   0
         End
      End
      Begin VB.Menu mpict0 
         Caption         =   "Pictures"
         Begin VB.Menu cmdpictures 
            Caption         =   "cmdpictures"
            Index           =   0
         End
      End
   End
   Begin VB.Menu m4contextuel 
      Caption         =   "Contextuel 4"
      Visible         =   0   'False
      Begin VB.Menu mrnremoveall 
         Caption         =   "Remove All"
      End
      Begin VB.Menu mrnsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnewformlist 
         Caption         =   "Open a list containing only new names..."
      End
      Begin VB.Menu msaveonlynewnames 
         Caption         =   "Save only new names..."
      End
      Begin VB.Menu mrnsep3 
         Caption         =   "-"
      End
      Begin VB.Menu mpastenewnames 
         Caption         =   "Paste new names from clipboard"
      End
      Begin VB.Menu mrnclipboard 
         Caption         =   "Copy to clipboard"
      End
   End
   Begin VB.Menu mgenhistory 
      Caption         =   "History"
      Visible         =   0   'False
      Begin VB.Menu mnuhistory 
         Caption         =   "mnuhistory"
         Index           =   0
      End
   End
   Begin VB.Menu MPicture 
      Caption         =   "Picture"
      Visible         =   0   'False
      Begin VB.Menu mpict 
         Caption         =   "Zoom &In"
         Index           =   0
      End
   End
End
Attribute VB_Name = "RENAME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Déclarations diverses
Private Const VK_SHIFT = &H10
Private Const VK_CONTROL = &H11
Private Const HELP_CONTEXT = &H1          '  Display topic in ulTopic
Private K() As cControlFlater   ' Classe pour avoir des contrôles 'Flat'
Dim acpreview As Boolean
Dim CptFichier As Long        ' Utilisé pour la commande <copy>, compteur de copie de fichier
Dim VancFichier As String     ' Utilisé quand on renomme un fichier directement sur le listview
Dim LeCancel As Boolean
Dim NomSettings As String
Dim aOuvrir As Boolean
Dim GlobalLoad As Boolean
Dim m_oAutoPos As New clsAutoPositioner
Dim cFichier As New cFile
Dim LaPosSauve As Integer
Dim LongSauve As Integer
Dim SIni As New cInifile
Dim VnbTokensPr As Integer  ' nombre de tokens du préfix
Dim VnbTokensEx As Integer  ' nombre de tokens de l'extension
Dim VnbTokensFo As Integer  ' nombre de tokens du répertoire
Dim TablTokensPr(256) As String ' Le tableau contenant les tokens du prefix
Dim TablTokensEx(256) As String ' Le tableau contenant les tokens de l'extension
Dim TablTokensFo(256) As String ' Le tableau contenant les tokens du répertoire
Dim PTablTokensPr(256) As String ' Le tableau contenant les positions des tokens du prefix
Dim PTablTokensEx(256) As String ' Le tableau contenant les positions des tokens de l'extension
Dim PTablTokensFo(256) As String ' Le tableau contenant les positions des tokens du répertoire
Dim itmX As ListItem
Dim OptionVis As Boolean
Dim CurrentCommand As Integer    ' Commande courant pour le free form
Dim MaxCommand As Integer        ' Nombre maxi de commandes dans le fichier
' ***************************************************************************************************
' Paramètres cachés
' ***************************************************************************************************
Dim RefreshRate As Integer  ' Nombre de fichiers au bout duquel THE Rename redonne la main à windows pour la file d'attente des messages
Dim DisplayRenMsg As Integer ' Faut il afficher des infos sur les fichiers en cours de renommage ?
Dim ShowPreviewList As Integer ' Faut il masquer la liste de preview pendant le renommage ?
' ***************************************************************************************************
' Les tableaux pour le Free Form
' ***************************************************************************************************
Dim hlplang(265) As Integer ' Contient les numéros de topics du fichier d'aide
Dim langage(265) As String  ' Contient le nom des commandes
Dim LngCmd(265, 2) As Integer ' Premier indice= longueur de la commande à tester (0= pas de test de longeur donc pas de paramètres), 2ième indice=nombre de paramètres de la commande
Dim vnbcmd As Integer  ' Nombre de commandes dans le langage
Dim commandes(300, 6) As String ' Contient les commandes et les paramètres
Dim TemShift As Boolean
Dim FavEncours As Integer
Dim PbFtv1 As Boolean
Private m_cMRU As New cMRUFileList
Private Declare Function SHExecuteDlg Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Reserved1 As Long, ByVal Reserved2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal flags As Long) As Long
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_LWIN = &H5B
Private Const KEYEVENTF_KEYUP = &H2
' Pour éviter le piège des Ctrl C, Ctrl V, Ctrl X et autres quand on est en édition de noms de fichiers sur le listview principal
Dim Piège1 As Boolean
Rem les tableaux pour toutes les commandes possibles
Dim optionp(21) As String ' Options pour le préfixe
Dim options(15) As String ' Options pour le suffixe
Rem Les tableaux des aides pour les commandes
Dim aidep(21) As String
Dim aides(15) As String
Dim vnboptionp As Integer
Dim vnboptions As Integer
Dim letat As Boolean
Private Sub DropFiles()
 frep.Show 1
End Sub
Private Sub ExitLang()
 RENAME.MousePointer = 0
 If acpreview = True Then
  Unload preview
  End If
End Sub
' Fonction pour renommer à partir d'une liste
Private Function FRenameList(chemin As String, temoin1 As Integer, temoin2 As Integer, temoin3 As Integer) As Long
    Dim OldName As String, NewName As String
    Dim multitache As Long, vnbtot As Long, vnb As Long, i As Long
    Dim prog11 As String, prog22 As String
    Dim fileop As New CSHFileOp
    fileop.ConfirmOperation = False
    vnbtot = ListView1.ListItems.Count
    multitache = 0
    vnb = 0
 
    For i = 0 To ListView2.ListItems.Count - 1
        NewName = LVGetName(ListView2, i)
        OldName = LVGetItemName(ListView2, i, 1)
        vnb = vnb + 1
        multitache = multitache + 1  ' Permet à l'écran de se rafraichir et de redonner la main
        If multitache = 10 Then
            multitache = 0
            DoEvents
        End If
        If Annuler = True Then  ' Gestion de l'arrêt du programme
            Annuler = False
            ExitLang
            remplissage
            Exit Function
        End If
        fileop.AddSourceFile chemin + OldName
        If LesOptions.RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
            NewName = RemIllegals(NewName)
        End If
        If LesOptions.RemoveStartingSpaces = 1 Then  ' Il faut supprimer les espaces de début dans le nom
            NewName = LTrim$(NewName)
        End If
  
        fileop.AddDestFile chemin + NewName
        état.Panels(1).Text = "Rename " + OldName + " to " + NewName
        état.Panels(2).Text = Trim$(Str$(vnb)) + "/" + Trim$(Str$(vnbtot))
        If acpreview = False Then  ' On n'est pas en preview
            If temoin1 = 1 Then ' Undofile
                Print #1, "ren " + Chr$(34) + NewName + Chr$(34) + " " + Chr$(34) + OldName + Chr$(34)
            End If
            If temoin3 = 1 Then ' Logfile
                Print #3, Str$(Date) + " " + Str$(Time) + " => " + chemin + OldName + " is rename to " + chemin + NewName
            End If
            If temoin2 = 1 Then ' Batch
                Print #2, "ren " + Chr$(34) + OldName + Chr$(34) + " " + Chr$(34) + NewName + Chr$(34)
            Else
                ' Les ajouts dans les listes sont pour l'UNDO
                List2.AddItem chemin + OldName ' Nom d'origine.
                List3.AddItem chemin + NewName   ' Nom d'arrivée.
                lhistory.AddItem Trim$(Str$(Time())) + "|" + chemin + "|" + OldName + "|" + NewName ' Historique
                If Len(Trim$(LesOptions.prog1)) <> 0 Then ' Lancer un programme avant de renommer le fichier
                    prog11 = LesOptions.prog1
                    ExecCmd prog11, chemin + OldName
                End If
                If LesOptions.CopyRename = True Then  ' Renommer les fichier et copier
                    If Not fileop.RenameFiles Then
                    End If
                Else ' On copie les fichiers, on ne les renomme pas
                    If Not fileop.CopyFiles Then
                    End If
                End If
            End If
            If Len(Trim$(LesOptions.prog2)) <> 0 Then ' Lancer un programme après avoir renommé le fichier
                prog22 = LesOptions.prog2
                ExecCmd prog22, chemin + NewName
            End If
            DT1.SetFileDateTime (chemin + NewName)
            Attr1.ChangeAttr (chemin + NewName)
        Else ' On est en preview *********************************************************************
            Set itmX = preview.listPreview.ListItems.Add(, , OldName)
            itmX.SubItems(1) = NewName
            preview.listsav.AddItem NewName
        End If ' Preview ou pas ?
        fileop.ClearSourceFiles
        fileop.ClearDestFiles
    Next
    FRenameList = vnbtot
End Function

Private Sub InvertSelection()
 Dim i As Long
 Dim vnb As Long
 letat = True
 vnb = ListView1.ListItems.Count - 1
 RENAME.MousePointer = 11
 ListView1.Visible = False
 For i = 0 To vnb
  If LVIsSelected(ListView1, i) = True Then
    LVSetItemNotSelected ListView1, i
  Else
   LVSetItemSelected ListView1, i
  End If
 Next
 RENAME.MousePointer = 0
 état.Panels(4).Text = Trim$(Str$(LVGetCountSelected(ListView1)))
 ListView1.Visible = True
 letat = False
 ListView1.SetFocus
End Sub

Private Sub MoveRoot()
 If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase$(Left$(Dir1Path, 1))) > 0 And Mid$(Dir1Path, 2, 1) = ":" And Mid$(Dir1Path, 3, 1) = "\" Then
  TemMove = False
  FolderTreeview1(0).Visible = False
  FolderTreeview1(0).SelectedFolder = Left$(Dir1Path, 3)
  Dir1Path = Left$(Dir1Path, 3)
  FolderTreeview1(0).Visible = True
  mundo.Enabled = False
  List2.Clear
  List3.Clear
 End If
End Sub

Private Sub MoveUp()
 mundo.Enabled = False
 List2.Clear
 List3.Clear
 TemMove = False
 FolderTreeview1(0).SetFocus
 SendKeys "{BACKSPACE}"
End Sub

Private Sub PreviewRename()
 acpreview = True
 StartRename
 acpreview = False
End Sub

Private Sub RenameManually()
Dim i As Long, chemin As String, vnb As Integer, vnb2 As Long
Dim vnbtot As Integer, fileop As New CSHFileOp
Dim vfichier As String, sItem As String
chemin = AddBackSlash(Trim$(Dir1Path))
If Recursive = True Then
 chemin = ""
End If

Rem Les listes pour l'undo
List2.Clear
List3.Clear
With fileop
    .ParentWnd = hWnd
    .ClearSourceFiles
    .ClearDestFiles
    .AllowUndo = False
    .ConfirmOperation = False
End With
vnb = 0
vnbtot = 0
vnb = LVGetCountSelected(ListView1)
vnbtot = vnb

If vnb = 0 Then
    MsgBox "You must select files before renaming them !"
    Exit Sub
End If

vnb = 0
i = LVGetItemSelected(ListView1, -1)
While i <> -1
 sItem = LVGetName(ListView1, i)
 vnb = vnb + 1
 vfichier = InputBox("Rename " + sItem + Chr$(13) + Chr$(10) + "Type ;end; to stop the process", "File " + Trim$(Str$(vnb)) + "/" + Trim$(Str$(vnbtot)), sItem)
 If Trim$(vfichier) <> "" Then
  If Trim$(UCase$(vfichier)) = ";END;" Then
   GoTo fin
  End If
  fileop.AddSourceFile chemin + sItem
  If LesOptions.RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
   vfichier = RemIllegals(vfichier)
  End If
  If LesOptions.RemoveStartingSpaces = 1 Then  ' Il faut supprimer les espaces de début de nom de fichier
    vfichier = LTrim$(vfichier)
  End If
  fileop.AddDestFile chemin + vfichier
Rem pour le undo
  List2.AddItem chemin + sItem        ' Nom d'origine.
  List3.AddItem chemin + vfichier     ' Nom d'arrivée.
  lhistory.AddItem Trim$(Str$(Time())) & "|" & chemin & "|" & sItem & "|" & vfichier ' Historique
  With fileop
    .RenameFiles
    .ClearSourceFiles
    .ClearDestFiles
  End With
 End If
 i = LVGetItemSelected(ListView1, i)
Wend
fin:
vnb2 = remplissage()
état.Panels(1).Text = "Ok"
With LesOptions
    .LastUseDate = Date
    .LastUseTime = Time
    .NumberOFiles = vnbtot
    .LastDirectory = Dir1Path
End With
' *** pour le undo ***
If List2.ListCount > 0 Then
 mundo.Enabled = True
End If
End Sub

Private Sub SelectAll()
 Dim i As Long, vnb As Long
 vnb = 0
 letat = True
 ListView1.Visible = False
 RENAME.MousePointer = 11
 For i = 0 To ListView1.ListItems.Count - 1
   LVSetItemSelected ListView1, i
   vnb = vnb + 1
 Next
 RENAME.MousePointer = 0
 état.Panels(4).Text = Trim$(Str$(vnb))
 ListView1.Visible = True
 letat = False
 
 If OptionVis = False Then
    If ListView1.Visible = True Then
        ListView1.SetFocus
    End If
 End If
End Sub

Public Sub StartRename()
    Dim Buffer As String * 256
    Dim IsFile As Boolean
    Dim Pos1 As Long, Pos2 As Long
    Dim sataille As Long, satailles As String
    Dim SonInfo As String
    Dim sesattr As Integer, simginfo As String, sesattrs As String
    Dim TemoinTrop As Boolean ' Temoin indiquant s'il y a trop de fichiers à traiter pour les listbox
    Dim i As Long, vnb As Long, val1 As Long, val2 As Long, val3 As Integer, val4 As Long, val5 As Long, val6 As Integer, valtempo As Long, longueur As Integer, vpos As Integer
    Dim val7 As Long, val8 As Integer
    Dim vprefixe As String, multitache As Long, vsuffixe As String
    Dim vformat1 As Integer, vformat2 As Integer
    Dim vnomfichier As String, vNomComplet As String, vnomorig As String
    Dim fileop As New CSHFileOp, vnbtot As Long, vaction1 As Integer
    Dim vaction2 As Integer, letexte As String, nboctets As Integer
    Dim vnomdesti As String, ChaineTempo As String, chemin As String
    Dim chemin2 As String, vnbparam As Integer
    Dim vMontre As Boolean
    Dim sItem As String, vtmp1 As String, vtmp2 As String
    Dim PRPrevFileOldName As String
    Dim PRPrevFileNewName As String
    Dim EXPrevFileOldName As String
    Dim EXPrevFileNewName As String
    Rem **** Les dim pour le langage de commandes ************************
    Dim vTemCopy As Boolean, vrai As Boolean, copie As String
    Dim TokenVoulu As Integer
    Dim litteral As String, vnb1 As Integer, vnb2 As Integer
    Dim cmdencours As Integer
    Dim chainetempo2 As String, chainetempo3 As String, laboucle As Integer
    Dim ChaineTempo4 As String
    Dim laboucle2 As Long, laboucle3 As Integer, cmdprefix As String, LaBoucle4 As Integer
    Dim cmdextension As String, vnom As String, vtempo As String
    Dim CompteurFichier As Long
    Dim ps As New clsParseString
    Rem **** Les dim pour le langage de commandes ************************

    Rem ******** Variables pour les abréviations *************************
    Dim ij As Integer, str3 As Integer, vnbcount As Integer
    Dim Match As Boolean
    Dim str1 As String, str2 As String
    Rem ******************************************************************
    Dim vtmpGlob As String
    Rem **** Pour éviter le piège des octets lus alors qu'il n'y en a pas assez à lire
    Dim Temoin11 As Boolean
    Rem Nom des programmes à lancer avant après et à la fin du traitement
    Dim prog11 As String, prog22 As String, prog33 As String
    Dim temoin1 As Integer ' Undofile
    Dim temoin2 As Integer ' Batch
    Dim temoin3 As Integer ' Logfile
    Rem *** Les Dim pour les noms aléatoires ******
    Dim LongRandom As Integer, LetterRandom As Long, VraiRandom As Boolean, NomRandom As String, BoucleRandom As Integer

    fileop.ConfirmOperation = False
    TemoinTrop = False

    chemin2 = AddBackSlash(Trim$(Dir1Path))
    multitache = 0
    CptFichier = 1
    Annuler = False
    vTemCopy = False

    If Len(LesOptions.UndoFile) > 0 Then
        temoin1 = 1
        Close #1
        On Error GoTo Erreur1
        If ExtractPath(LesOptions.UndoFile) = LesOptions.UndoFile Then
            Open chemin2 + LesOptions.UndoFile For Output As #1
        Else
            Open LesOptions.UndoFile For Output As #1
        End If
    Else
        temoin1 = 0
    End If

    If Len(LesOptions.batch) > 0 Then
        temoin2 = 1
        Close #2
        On Error GoTo Erreur2
        If ExtractPath(LesOptions.batch) = LesOptions.batch Then
            Open chemin2 + LesOptions.batch For Output As #2
        Else
            Open LesOptions.batch For Output As #2
        End If
    Else
        temoin2 = 0
    End If

    If Len(LesOptions.LogFile) > 0 Then
        temoin3 = 1
        Close #3
        On Error GoTo Erreur3
        If ExtractPath(LesOptions.LogFile) = LesOptions.LogFile Then
            Open chemin2 + LesOptions.LogFile For Append As #3
        Else
            Open LesOptions.LogFile For Append As #3
        End If
    Else
        temoin3 = 0
    End If

    On Error GoTo ErrGen
    If vaction1 <> 13 Then
        nboctets = Val(Text9.Text)
    End If
    vaction1 = 0
    vaction2 = 0

    vformat1 = Combo3.ListIndex
    vformat2 = Combo4.ListIndex

    vnb = 0
    vprefixe = ""
    vsuffixe = ""

    ' Récupération du nombre de fichiers sélectionnés dans la liste
    vnb = LVGetCountSelected(ListView1)
    vnbtot = vnb
    If vnb = 0 Then
        MsgBox "Error, You must select files before to rename them !"
        ExitLang
        Exit Sub
    End If

    If vnb > 32736 Then
        TemoinTrop = True
        If MsgBox("WARNING, When there are more than 32736 files to rename, THE Rename can't use Preview, History and UNDO. Are you sure you want to go on ?", vbYesNo, "!! WARNING, IMPORTANT !!") = vbNo Then
            RENAME.MousePointer = 0
            If acpreview Then
                Unload preview
            End If
            Exit Sub
        End If
    End If

    If LesOptions.UseHistory Then ' on vérifie s'il n'y aura pas trop de fichiers pour l'historique
        If TemoinTrop = False Then
            If lhistory.ListCount + vnb > 32736 Then
                MsgBox "Warning, History is full, i'm emptying it", vbOKOnly, "History"
                lhistory.Clear
            End If
        End If
    End If

    val1 = Val(Text3.Text)    ' valeur de départ
    val2 = Val(Text4.Text)    ' incrément
    val3 = Val(Text5.Text)    ' nb digits
    val4 = Val(Text16.Text)   ' valeur de départ
    val5 = Val(Text17.Text)   ' incrément
    val6 = Val(Text18.Text)   ' nb digits

    ' Si on demande un compteur en lettres et qu'on à mis une valeur de départ à
    ' zéro alors on la passe à 1.
    If vformat1 = 4 Then
        If val1 = 0 Then
            val1 = 1
        End If
    End If

    If vformat2 = 4 Then
        If val4 = 0 Then
            val4 = 1
        End If
    End If

    If val2 = 0 Then
        MsgBox "Step for counter is invalid, change it."
        ExitLang
        Text4.SetFocus
        Exit Sub
    End If
    If LesOptions.CompleCounters = 1 Then
        If val3 = 0 Then
            MsgBox "number of digits for counter is invalid, change it."
            ExitLang
            Text5.SetFocus
            Exit Sub
        End If
    End If

    If val5 = 0 Then
        MsgBox "Step for counter is invalid, change it."
        ExitLang
        Text17.SetFocus
        Exit Sub
    End If
    
    If val6 = 0 Then
        MsgBox "number of digits for counter is invalid, change it."
        ExitLang
        Text18.SetFocus
        Exit Sub
    End If

    ' Recherche des actions à effectuer sur le programme
    For i = 1 To vnboptionp
        If Trim$(Combo1.List(Combo1.ListIndex)) = Trim$(optionp(i)) Then
            vaction1 = i
            Exit For
        End If
    Next
    For i = 1 To vnboptions
        If Trim$(Combo2.List(Combo2.ListIndex)) = Trim$(options(i)) Then
            vaction2 = i
            Exit For
        End If
    Next

    If vaction1 = 0 Or vaction2 = 0 Then
        MsgBox "Internal Error 1, please contact me !"
        ExitLang
        Exit Sub
    End If

    If vaction1 = 11 Or vaction1 = 13 Then ' On vérifie que le nombre maxi de caractères est bon
        If nboctets <= 0 Or nboctets > 256 Then
            MsgBox "Numbers of characters to take from file is not valid. Change it !"
            ExitLang
            If vaction1 = 11 Then
                Text9.SetFocus
            End If
            Exit Sub
        End If
    End If

    Rem On vérifie que les compteurs seront bons *****************************
    If Recursive = False Then
        If LesOptions.CompleCounters = 1 Then
            valtempo = val1 - 1
            If Check3.Value = 1 Then
                For i = 1 To vnb
                    valtempo = valtempo + val2
                Next
                If vformat1 = 0 Then
                    ChaineTempo = Trim$(Str$(valtempo))
                Else
                    If vformat1 = 1 Then
                        ChaineTempo = Trim$(Hex$(valtempo))
                    Else
                        ChaineTempo = Trim$(Oct$(valtempo))
                    End If
                End If
                If Len(ChaineTempo) > val3 Then
                    MsgBox "Counter for prefix will not have enough digits to rename all files. Change counter digits"
                    ExitLang
                    Exit Sub
                End If
            End If
            valtempo = val4 - 1
            If Check4.Value = 1 Then
                For i = 1 To vnb
                    valtempo = valtempo + val5
                Next
                If vformat2 = 0 Then
                    ChaineTempo = Trim$(Str$(valtempo))
                Else
                    If vformat2 = 1 Then
                        ChaineTempo = Trim$(Hex$(valtempo))
                    Else
                        ChaineTempo = Trim$(Oct$(valtempo))
                    End If
                End If
                If Len(ChaineTempo) > val6 Then
                    MsgBox "Counter for extension will not have enough digits to rename all files. Change counter digits"
                    ExitLang
                    Exit Sub
                End If
            End If
        End If
    End If ' Test sur les compteurs seulement si on n'est pas en mode récursif

    vnb = 0

    Rem ***************************************************************************
    Rem On prépare les listes pour l'undo
    List2.Clear
    List3.Clear

    Rem ***************************************************************************
    With fileop
        .ParentWnd = hWnd
        .ConfirmOperation = False
        .ClearSourceFiles
        .ClearDestFiles
    End With

    Rem ********** Gestion du preview *********************************************
    If acpreview Then
        If TemoinTrop = False Then
            With preview
                .listPreview.ListItems.Clear
                .listsav.Clear
                .Command1.Visible = False
            End With
            If ShowPreviewList = 0 Then
                preview.listPreview.Visible = False
            End If
            preview.Show 0
        End If
    End If

    ' *********************************************************************************************************************************************************************
    ' ****************** Analyse en cas de langage ************************************************************************************************************************
    ' *********************************************************************************************************************************************************************
    If vaction1 = 13 Then
        If Len(Trim$(txtlang.Text)) = 0 Then
            MsgBox "Warning, expression is empty !"
            ExitLang
            txtlang.SetFocus
            Exit Sub
        End If
        ClearCommands
        txtlang.Text = Trim$(txtlang.Text)
        copie = UCase$(txtlang.Text)
        vnb1 = CharOccurs(copie, "<")
        vnb2 = CharOccurs(copie, ">")
        If vnb1 <> vnb2 Then
            If vnb1 > vnb2 Then
                MsgBox "Error, number of '>' is different from number of '<'"
                ExitLang
                txtlang.SetFocus
                Exit Sub
            Else
                MsgBox "Error, number of '<' is different from number of '>'"
                ExitLang
                txtlang.SetFocus
                Exit Sub
            End If
            Exit Sub
        End If
 
        ' Théoriquement il n'y a plus d'erreurs, on peut commencer à lancer l'analyse
        cmdencours = 0
        longueur = Len(txtlang.Text)
        i = 1
        While i <= longueur
            If Mid$(copie, i, 1) = "<" Then ' C'est une commande *******************************
                ChaineTempo = Mid$(copie, i)
                ChaineTempo4 = Mid$(txtlang.Text, i)
                vnb1 = At(ChaineTempo, ">", 1)
                ChaineTempo = Mid$(ChaineTempo, 1, vnb1)
                ChaineTempo4 = Mid$(ChaineTempo4, 1, vnb1)
                ' Bon, on sait qu'on est sur une commande, on l'a, il ne reste plus qu'à savoir laquelle c'est
                cmdencours = cmdencours + 1
                If cmdencours > 300 Then
                    MsgBox "Error, your expression is too long"
                    ExitLang
                    txtlang.SetFocus
                    Exit Sub
                End If
                If Left$(ChaineTempo, 2) <> "<D" And Left$(ChaineTempo, 2) <> "<X" And Left$(ChaineTempo, 2) <> "<O" Then ' ce n'est pas une commande de compteur
zsuite1:
                    vrai = False
                    For vnb2 = 1 To vnbcmd
                        If LngCmd(vnb2, 1) = 0 Then ' c'est une commande sans paramètres
                            If UCase$(Trim$(langage(vnb2))) = UCase$(Trim$(ChaineTempo)) Then
                                vrai = True
                                Exit For
                            End If
                        Else ' C'est une commande AVEC paramètres *******************************
                            If Left$(UCase$(Trim$(langage(vnb2))), LngCmd(vnb2, 1)) = Left$(UCase$(Trim$(ChaineTempo)), LngCmd(vnb2, 1)) Then
                                If InStr(ChaineTempo, ",") <> 0 Then    ' La commande saisie par l'utilisateur contient un paramètre
                                    vtmp1 = langage(vnb2)
                                    If At(UCase$(Trim$(langage(vnb2))), ",", 1) = 0 Then
                                        vtmp1 = Left$(vtmp1, Len(vtmp1) - 1) + "," + ">"
                                    End If
                                    If Left$(ChaineTempo, At(ChaineTempo, ",", 1)) = Left$(UCase$(Trim$(vtmp1)), At(UCase$(Trim$(vtmp1)), ",", 1)) Then
                                        vrai = True
                                        Exit For
                                    End If
                                Else            ' La commande saisie par l'utilisateur ne contient pas de paramètre(s)
                                    If UCase$(Trim$(langage(vnb2))) = ChaineTempo Then
                                        vrai = True
                                        Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If vrai = False Then
                        If Left$(ChaineTempo, 2) = "<P" Or Left$(ChaineTempo, 2) = "<p" Or Left$(ChaineTempo, 2) = "<E" Or Left$(ChaineTempo, 2) = "<e" Or Left$(ChaineTempo, 2) = "<f" Or Left$(ChaineTempo, 2) = "<F" Then
                            If At(ChaineTempo, ",", 1) <> 0 Then ' On a mis un ou des paramètres
                                If IsNumeric(Mid$(ChaineTempo, 3, At(ChaineTempo, ",", 1) - 1)) Then
                                    GoTo zsuite2
                                End If
                            Else    ' Pas de paramètres
                                If IsNumeric(Mid$(ChaineTempo, 3, Len(ChaineTempo) - 3)) Then
                                    GoTo zsuite2
                                End If
                            End If
                        End If
                        MsgBox "Error, " + ChaineTempo + " is not a valid command !"
                        ExitLang
                        txtlang.SetFocus
                        Exit Sub
                    End If
zsuite2:
                    cmdencours = cmdencours + 1
                    If cmdencours > 300 Then
                        MsgBox "Error, your expression is too long"
                        ExitLang
                        txtlang.SetFocus
                        Exit Sub
                    End If
                    If vnb2 = vnbcmd + 1 Then   ' Exception pour les commandes de tokens
                        commandes(cmdencours, 1) = "-1"     ' L'indice de la commande
                    Else
                        commandes(cmdencours, 1) = Trim$(Str$(vnb2))     ' L'indice de la commande
                    End If
                    If vnb2 = 32 Then ' C'est la commande <COPY>
                        If UCase$(Left$(Trim$(txtlang.Text), 9)) <> "<COPYFILE" Then
                            MsgBox "Error, the <copyfile> command MUST be the first command !"
                            ExitLang
                            txtlang.SetFocus
                            Exit Sub
                        End If
                        vTemCopy = True
                    End If
    
                    If vnb2 = 58 Then ' C'est la commande <CyclicSelection> on vérifie qu'il y a des textes à utiliser
                        If VnbCyclic = 0 Then
                            MsgBox "Warning, you are using the command <CyclicSelection> but you did not entered any text...", vbOKOnly, "Warning !"
                            ExitLang
                            txtlang.SetFocus
                            Exit Sub
                        End If
                    End If
                    If vnb2 <= vnbcmd Then   ' Exception pour les commandes de tokens
                        If LngCmd(vnb2, 1) = 0 Then ' c'est une commande sans paramètres
                            commandes(cmdencours, 2) = ""                  ' Commande sans paramètre
                        Else ' C'est une commande AVEC paramètres
                            vnbparam = CharOccurs(ChaineTempo, ",")
                            If LngCmd(vnb2, 2) > 0 Then ' Commande n'acceptant pas un nombre variable de paramètres
                                If vnbparam <> LngCmd(vnb2, 2) Then
                                    MsgBox "Error, Invalid number of parameters for the command " + ChaineTempo + ", change it. There should be " + Str$(LngCmd(vnb2, 2)) + " parameter(s)"
                                    ExitLang
                                    txtlang.SetFocus
                                    Exit Sub
                                End If ' Test sur la validité du nombre de paramètres
                            Else    ' commande acceptant un nombre variable d'arguments
                                If vnbparam > Abs(LngCmd(vnb2, 2)) Then
                                    MsgBox "Error, the command " + ChaineTempo + " only accepts a maximum of " + Trim$(Str$(Abs(LngCmd(vnb2, 2)))) + " parameters"
                                    ExitLang
                                    txtlang.SetFocus
                                    Exit Sub
                                Else
                                End If
                            End If
                            chainetempo3 = Left$(ChaineTempo4, Len(ChaineTempo4) - 1)
                            For laboucle3 = 1 To vnbparam ' Insertion des paramètres dans le tableau de commandes
                                commandes(cmdencours, 1 + laboucle3) = Replace(Trim$(GetToken(chainetempo3, ",", laboucle3 + 1)), "\w", " ", , , vbTextCompare)
                                If Len(commandes(cmdencours, 1 + laboucle3)) = 0 Then
                                    MsgBox "Error for the command " + chainetempo3 + " parameter #" + Trim$(Str$(laboucle3)) + " is empty !"
                                    ExitLang
                                    txtlang.SetFocus
                                    Exit Sub
                                End If
                                If vnb2 = 32 Then ' Commande <COPY>
                                    CptFichier = Val(commandes(cmdencours, 1 + laboucle3))
                                End If
                            Next
                            If vnb2 = 155 Then  ' Commande <CounterEx>
                                val7 = Val(commandes(cmdencours, 3))
                                val8 = Val(commandes(cmdencours, 4))
                            End If
                        End If ' Commande avec ou sans paramètres ?
                    Else ' Commande de token, il faut la mémoriser comme paramètre de facon à connaitre le token voulu
                        commandes(cmdencours, 2) = ChaineTempo
                    End If
                Else ' c'est une commande de compteur ******************************************************************************************************************************
                    vrai = False
                    If Left$(ChaineTempo, 4) = "<OGG" Then  ' exception pour les commandes <Ogg..> qui sans le test sont prises pour des commandes de compteur en octal style <ooo>
                        GoTo zsuite1
                    End If
                    If Left$(ChaineTempo, 2) = "<D" Then
                        chainetempo2 = "<DDDDD>"
                    Else
                        If Left$(ChaineTempo, 2) = "<X" Then
                            chainetempo2 = "<XXXXX>"
                        Else
                            chainetempo2 = "<OOOOO>"
                        End If
                    End If
                    For vnb2 = 1 To vnbcmd
                        If UCase$(Trim$(langage(vnb2))) = UCase$(Trim$(chainetempo2)) Then
                            vrai = True
                            Exit For
                        End If
                    Next
                    If vrai = False Then
                        MsgBox "Error, " + ChaineTempo + " is not a valid command !"
                        ExitLang
                        txtlang.SetFocus
                        Exit Sub
                    End If
                    cmdencours = cmdencours + 1
                    If cmdencours > 300 Then
                        MsgBox "Error, your expression is too long, there are more than 300 commands !"
                        ExitLang
                        txtlang.SetFocus
                        Exit Sub
                    End If
                    commandes(cmdencours, 1) = Trim$(Str$(vnb2))                  ' L'indice de la commande
                    commandes(cmdencours, 2) = Trim$(Str$(Len(ChaineTempo) - 2))  ' Commande sans paramètre
                End If
                i = vnb1 + Len(Left$(copie, i))
            Else ' C'est un litteral ***************************************************************
                vrai = False
                litteral = ""
                While vrai = False And i <= longueur
                    If Mid$(txtlang.Text, i, 1) <> "<" Then
                        litteral = litteral + Mid$(txtlang.Text, i, 1) ' Pour les litteraux, il faut prendre le texte original, pas celui qui a été passé en majuscules
                    Else
                        vrai = True
                    End If
    
                    If Not vrai Then
                        If i <= longueur Then ' si i reste inférieur à longueur et si on n'a pas déjà demandé à s'arrêter
                            i = i + 1
                        Else ' On arrive en fin de chaine
                            vrai = False
                        End If
                    End If
                Wend
                cmdencours = cmdencours + 1
                commandes(cmdencours, 1) = "0"      ' 0 indique un litteral
                commandes(cmdencours, 2) = litteral ' le texte du litteral
            End If
        Wend
        Rem réglages des paramètres pour les compteurs
        val1 = Val(cmdtxt1.Text)    ' valeur de départ
        val2 = Val(cmdtxt2.Text)    ' incrément
    End If

    CompteurCyclic = 0
' *************************************************************************************************************************************************
' **********        Traitement sur les fichiers         ********************************************************************************************
' *************************************************************************************************************************************************
    chemin = AddBackSlash(Trim$(Dir1Path))
    RENAME.MousePointer = 11

    If vaction1 = 16 Then ' Rename from a list
        vnbtot = FRenameList(chemin, temoin1, temoin2, temoin3)
        GoTo zsuite
    End If

    i = LVGetItemSelected(ListView1, -1)
    While i <> -1
        If Annuler Then  ' Gestion de l'arrêt du programme
            Annuler = False
            ExitLang
            remplissage
            Exit Sub
        End If
  
        sItem = LVGetName(ListView1, i)
        vnb = vnb + 1
        multitache = multitache + 1  ' Permet à l'écran de se rafraichir et de redonner la main
        If RefreshRate <> -1 Then
            If multitache = RefreshRate Then
                multitache = 0
                DoEvents
            End If
        End If
   
        vprefixe = Prefixe(sItem) ' Le préfixe uniquement
        vsuffixe = Suffixe(sItem) ' Le suffixe uniquement
        vnomfichier = sItem ' Nom complet du fichier avec le chemin s'il existe
        vnomorig = vprefixe + "." + vsuffixe
        If Recursive Then
            chemin = ExtractPath(sItem) ' Si on est en récursif, il faut récupérer le chemin du fichier
        End If
        vNomComplet = chemin + vprefixe + "." + vsuffixe
        CompteurFichier = CompteurFichier + 1
        If LVGetItemName(ListView1, i, 4) = "File" Then
            IsFile = True
        Else
            IsFile = False
        End If
        ' **** Les règles ****************************************************************************
        If LesRegles.NumberOfActiveRules > 0 Then   ' On ne fait les tests que s'il y a des règles d'actives
            Set cFichier = Nothing ' On prépare les infos sur le fichier
            If IsFile Then
                cFichier.SetFileName vNomComplet, True  ' Fichier
            Else
                cFichier.SetFileName vNomComplet, False ' Répertoire
            End If
            If Not LesRegles.TestRules(cFichier) Then ' On fait les tests par rapport aux règles
                GoTo fin
            End If
        End If
        ' ********************************************************************************************
      
        If RechGlob Then
            If LesOptions.SearchAndReplace = 0 Or LesOptions.SearchAndReplace = 2 Then
                rech3.SourceString = vprefixe + "." + vsuffixe
                vtmpGlob = rech3.BeginSearchAndReplace
                vprefixe = Prefixe(vtmpGlob)
                vsuffixe = Suffixe(vtmpGlob)
                rech3.SourceString = vprefixe + "." + vsuffixe
                vtmpGlob = rech3.BeginReplaceCharacters
                vprefixe = Prefixe(vtmpGlob)
                vsuffixe = Suffixe(vtmpGlob)
            End If
        End If
    
        If RechPref Then
            If LesOptions.SearchAndReplace = 0 Or LesOptions.SearchAndReplace = 2 Then
                rech1.SourceString = vprefixe
                vprefixe = rech1.BeginSearchAndReplace
                rech1.SourceString = vprefixe
                vprefixe = rech1.BeginReplaceCharacters
            End If
        End If
   
        ' Abréviations
        If OkUseAbbrev Then ' il faut utiliser les abbréviations
            If LesOptions.SearchAndReplace = 0 Or LesOptions.SearchAndReplace = 2 Then
                For ij = 1 To CollAbrev.Count   ' Boucle sur toutes les abbréviations de la collection
                    If GetToken(CollAbrev.Item(ij), Chr$(254), 7) = "yes" Then ' on utilise des expressions régulières
                        'On Error Resume Next ********** Laisser tel quel
                        Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                            Case "YES"          ' on recherche partout (préfixe et suffixe)
                                str2 = vprefixe + "." + vsuffixe
                            Case "PREFIX"      ' Préfixe uniquement
                                str2 = vprefixe
                            Case "EXTENSION"    ' Extension uniquement
                                str2 = vsuffixe
                        End Select
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 4) = "yes" Then
                            vnbcount = 9999
                        Else
                            vnbcount = Val(GetToken(CollAbrev.Item(ij), Chr$(254), 4))
                        End If
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 5) = "yes" Then
                            str3 = 1
                        Else
                            str3 = 0
                        End If
                        Match = RegSub(str2, GetToken(CollAbrev.Item(ij), Chr$(254), 1), GetToken(CollAbrev.Item(ij), Chr$(254), 2), str1, str3, vnbcount, 0, 0)
                        If Match Then
                            Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                                Case "YES"          ' on recherche partout (préfixe et suffixe)
                                    vprefixe = Prefixe(str1)
                                    vsuffixe = Suffixe(str1)
                                Case "PREFIX"      ' Préfixe uniquement
                                    vprefixe = str1
                                Case "EXTENSION"    ' Extension uniquement
                                    vsuffixe = str1
                            End Select
                        End If
                    Else ' on n'utilise pas d'expression régulière, recherche "normale"
                        Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                            Case "YES"          ' on recherche partout (préfixe et suffixe)
                                str2 = vprefixe + "." + vsuffixe
                            Case "PREFIX"      ' Préfixe uniquement
                                str2 = vprefixe
                            Case "EXTENSION"    ' Extension uniquement
                                str2 = vsuffixe
                        End Select
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 4) = "yes" Then
                            vnbcount = 9999
                        Else
                            vnbcount = Val(GetToken(CollAbrev.Item(ij), Chr$(254), 4))
                        End If
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 5) = "yes" Then
                            str3 = vbBinaryCompare
                        Else
                            str3 = vbTextCompare
                        End If
                        str1 = Replace(str2, GetToken(CollAbrev.Item(ij), Chr$(254), 1), GetToken(CollAbrev.Item(ij), Chr$(254), 2), Val(GetToken(CollAbrev.Item(ij), Chr$(254), 3)), vnbcount, str3)
                        Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                            Case "YES"          ' on recherche partout (préfixe et suffixe)
                                vprefixe = Prefixe(str1)
                                vsuffixe = Suffixe(str1)
                            Case "PREFIX"      ' Préfixe uniquement
                                vprefixe = str1
                            Case "EXTENSION"    ' Extension uniquement
                                vsuffixe = str1
                        End Select
                    End If
                Next
            End If
        End If
   
        If LesOptions.RestartCounter = 1 Then
            If chemin <> VancRep Then
                val1 = Val(Text3.Text)    ' valeur de départ
                val4 = Val(Text16.Text)   ' valeur de départ
                VancRep = chemin
            End If
        ElseIf LesOptions.RestartCounter = 2 Then
            If Left$(chemin, At(chemin, "\", LesOptions.LevelRestart)) <> VancRep Then
                val1 = Val(Text3.Text)    ' valeur de départ
                val4 = Val(Text16.Text)   ' valeur de départ
                VancRep = Left$(chemin, At(chemin, "\", LesOptions.LevelRestart))
            End If
        End If

        Rem ***************** Action à effectuer sur le préfixe *********************
        Select Case vaction1
            Case 1 ' Mettre en majuscules
                vprefixe = UCase$(vprefixe)
            Case 2 ' Mettre en minuscules
                vprefixe = LCase$(vprefixe)
            Case 3 ' Inverser majuscules/minucules
                vprefixe = ToggleCase(vprefixe)
            Case 6 ' Inverser les lettres
                vprefixe = StrReverse(vprefixe)
            Case 5 ' Capitalize all words
                vprefixe = MyStrConv(vprefixe)
            Case 4 ' Conserver le préfixe
                vprefixe = vprefixe
            Case 7 ' Remplacer par la date système
                vprefixe = FmtDate(Now)
            Case 8 ' Remplacer par l'heure système
                vprefixe = Menage(FmtHeure(Time()))
            Case 9 ' Remplacer par la date + l'heure système
                vprefixe = FmtDate(Now) + Menage(FmtHeure(Time()))
    
            Case 10 ' Modifier le préfixe
                If Option1(0).Value Then ' remplacer par un texte fixe
                    vprefixe = Text2.Text
                    If UseCylcic Then
                        CompteurCyclic = CompteurCyclic + 1
                        If CompteurCyclic > VnbCyclic Then
                            CompteurCyclic = 1
                        End If
                        Select Case PlacementCyclic
                            Case 0 ' Add to the right
                                vprefixe = vprefixe + LesCyclic(CompteurCyclic)
                            Case 1 ' Add to the left
                                vprefixe = LesCyclic(CompteurCyclic) + vprefixe
                            Case 2 ' replace prefix
                                vprefixe = LesCyclic(CompteurCyclic)
                        End Select
                    End If ' Selection cyclique
                Else ' ajouter un texte fixe ***************************************
                    If Option1(1).Value Then
                        If Option2(0).Value Then  ' au début
                            vprefixe = Text14.Text + LTrim$(vprefixe)
                        Else ' à la fin
                            vprefixe = RTrim$(vprefixe) + Text14.Text
                        End If
                        If UseCylcic Then
                            CompteurCyclic = CompteurCyclic + 1
                            If CompteurCyclic > VnbCyclic Then
                                CompteurCyclic = 1
                            End If
                            Select Case PlacementCyclic
                                Case 0 ' Add to the right
                                    vprefixe = vprefixe + LesCyclic(CompteurCyclic)
                                Case 1 ' Add to the left
                                    vprefixe = LesCyclic(CompteurCyclic) + vprefixe
                                Case 2 ' replace prefix
                                    vprefixe = LesCyclic(CompteurCyclic)
                            End Select
                        End If ' Selection cyclique
                    End If
                End If
      
                If Check3.Value = 1 Then ' ajouter un compteur
                    If Option3(0).Value Then   ' Rajouter à gauche
                        vprefixe = Compteur(val1, val3, vformat1) + vprefixe
                    Else
                        If Option3(1).Value Then  ' Rajouter à droite
                            vprefixe = vprefixe + Compteur(val1, val3, vformat1)
                        Else ' Remplacer le préfixe par le compteur
                            vprefixe = Compteur(val1, val3, vformat1)
                        End If
                    End If
                    val1 = val1 + val2
                End If
      
                If Check5.Value = 1 Then ' ajouter la taille
                    If Option3(3).Value Then  ' Rajouter à gauche
                        vprefixe = FileLen(vNomComplet) & vprefixe
                    Else
                        If Option3(4).Value Then  ' Rajouter à droite
                            vprefixe = vprefixe & FileLen(vNomComplet)
                        Else ' Remplacer le préfixe par la taille
                            vprefixe = FileLen(vNomComplet)
                        End If
                    End If
                End If
      
                If Check6.Value = 1 Then ' ajouter la date
                    If Option3(8).Value Then  ' Rajouter à gauche
                        vprefixe = FmtDate(FileDateTime(vNomComplet)) + vprefixe
                    Else
                        If Option3(7).Value Then  'Rajouter à droite
                            vprefixe = vprefixe + FmtDate(FileDateTime(vNomComplet))
                        Else 'Remplacer le préfixe par la date
                            vprefixe = FmtDate(FileDateTime(vNomComplet))
                        End If
                    End If
                End If
      
                If Check7.Value = 1 Then ' ajouter l'heure
                    If Option3(9).Value Then  ' Rajouter à gauche
                        vprefixe = Menage(FmtHeure(FileDateTime(vNomComplet))) + vprefixe
                    Else
                        If Option3(10).Value Then  'Rajouter à droite
                            vprefixe = vprefixe + Menage(FmtHeure(FileDateTime(vNomComplet)))
                        Else 'Remplacer le préfixe par la date
                            vprefixe = Menage(FmtHeure(FileDateTime(vNomComplet)))
                        End If
                    End If
                End If
      
                If FolderOk Then  ' Add folder's name
                    If Folder4 = 0 Then ' add to the left
                        vprefixe = Trim$(FolderPart(vNomComplet)) + vprefixe
                    Else
                        If Folder4 = 1 Then ' Add to the right
                            vprefixe = vprefixe + Trim$(FolderPart(vNomComplet))
                        Else ' Replace prefix
                            vprefixe = Trim$(FolderPart(vNomComplet))
                        End If
                    End If
                End If
      
                If Check1.Value = 1 Then ' ajouter les infos sur l'image
                    If IsFile Then
                        simginfo = ""
                        simginfo = ImgInfo(vNomComplet)
                        If Option3(17).Value Then  ' Rajouter à gauche
                            vprefixe = simginfo + vprefixe
                        Else
                            If Option3(16).Value Then  'Rajouter à droite
                                vprefixe = vprefixe + simginfo
                            Else 'Remplacer le préfixe par les infos
                                If Len(Trim$(simginfo)) <> 0 Then
                                    vprefixe = simginfo
                                End If
                            End If
                        End If
                    End If
                End If
                
                ' Tags EXIF des images
                If PicEXIF.UseEXIF Then
                    If IsFile Then
                        SonInfo = ""
                        SonInfo = PicEXIF.GetEXIFInfos(vNomComplet)
                        If PicEXIF.PlaceWhereToPut = 0 Then ' Rajouter à gauche
                            vprefixe = SonInfo + vprefixe
                        Else
                            If PicEXIF.PlaceWhereToPut = 1 Then ' Rajouter à droite
                                vprefixe = vprefixe + SonInfo
                            Else    ' Remplacer le préfixe par les infos
                                vprefixe = SonInfo
                            End If
                        End If
                    End If
                End If
      
                If UseMP3 Then
                    If IsFile Then
                        SonInfo = ""
                        SonInfo = MusMP3.GetMP3Infos(vNomComplet)
                        If MusMP3.PlaceWhereToPut = 0 Then ' Rajouter à gauche
                            vprefixe = SonInfo + vprefixe
                        Else
                            If MusMP3.PlaceWhereToPut = 1 Then ' Rajouter à droite
                                vprefixe = vprefixe + SonInfo
                            Else ' Remplacer le préfixe par les infos
                                If Len(Trim$(SonInfo)) <> 0 Then
                                    vprefixe = SonInfo
                                End If
                            End If
                        End If
                    End If
                End If ' Faut'il ajouter des infos de MP3 ?
      
                If UseWMA Then
                    If IsFile Then
                        SonInfo = ""
                        SonInfo = MusWMA.GetWMAInfos(vNomComplet)
                        If MusWMA.PlaceWhereToPut = 0 Then ' Rajouter à gauche
                            vprefixe = SonInfo + vprefixe
                        Else
                            If MusWMA.PlaceWhereToPut = 1 Then ' Rajouter à droite
                                vprefixe = vprefixe + SonInfo
                            Else ' Remplacer le préfixe par les infos
                                If Len(Trim$(SonInfo)) <> 0 Then
                                    vprefixe = SonInfo
                                End If
                            End If
                        End If
                    End If
                End If ' Faut'il ajouter des infos des WMA ?
      
                If UseVQF Then
                    If IsFile Then
                        SonInfo = ""
                        SonInfo = MusVQF.GetVQFInfos(vNomComplet)
                        If MusVQF.PlaceWhereToPut = 0 Then ' Rajouter à gauche
                            vprefixe = SonInfo + vprefixe
                        Else
                            If MusVQF.PlaceWhereToPut = 1 Then ' Rajouter à droite
                                vprefixe = vprefixe + SonInfo
                            Else ' Remplacer le préfixe par les infos
                                If Len(Trim$(SonInfo)) <> 0 Then
                                    vprefixe = SonInfo
                                End If
                            End If
                        End If
                    End If
                End If ' Faut'il ajouter des infos de VQF ?
      
                If UseOGG Then
                    If IsFile Then
                        SonInfo = ""
                        SonInfo = MusOgg.GetOggInfos(vNomComplet)
                        If MusOgg.PlaceWhereToPut = 0 Then ' Rajouter à gauche
                            vprefixe = SonInfo + vprefixe
                        Else
                            If MusOgg.PlaceWhereToPut = 1 Then ' Rajouter à droite
                                vprefixe = vprefixe + SonInfo
                            Else ' Remplacer le préfixe par les infos
                                If Len(Trim$(SonInfo)) <> 0 Then
                                    vprefixe = SonInfo
                                End If
                            End If
                        End If
                    End If
                End If ' Faut'il ajouter des infos de OGG ?
      
    
            Case 11 ' Remplacer avec le contenu du fichier
                If IsFile Then
                    Temoin11 = True
                    Open vnomfichier For Binary As #1
                    Get 1, , Buffer
                    Close #1
                    Buffer = Left$(Buffer, nboctets)
                    letexte = Replace(Buffer, Chr$(0), "")
                    Temoin11 = False
                    If Len(Trim$(letexte)) = 0 Then
                        MsgBox "Error, unable to rename file " + vnomfichier + " because file's content is empty !"
                        ExitLang
                        Exit Sub
                    End If
                    letexte = Replace(letexte, vbCrLf, "")
                    vprefixe = Trim$(Menage(letexte))
                    If Len(vprefixe) = 0 Then
                        MsgBox "Error, unable to rename file " + vnomfichier + " because file's content is empty after removing unavailable characters !"
                        ExitLang
                        Exit Sub
                    End If
                End If
    
            Case 12 ' Remplacer avec le nom interne d'une police Truetype
                If IsFile Then
                    vprefixe = Menage(GetFontName(vNomComplet))
                End If
    
            Case 13 ' Free form, mini language
                cmdprefix = Prefixe(vnomfichier)
                cmdextension = Suffixe(vnomfichier)
                ' Il faut gérer les recherches et remplacements
                LesRecherches1 cmdprefix, cmdextension, 0, 2
                
                For laboucle2 = 1 To CptFichier
                    vnom = ""
                    CreateTokenTabl vnomorig, chemin ' chargement des tokens pour le fichier courant
                    For laboucle = 1 To cmdencours
                        SonInfo = ""
                        Select Case commandes(laboucle, 1)
                            Case "-1" ' Commande de token
                                TokenVoulu = Val(Mid$(commandes(laboucle, 2), 3, Len(commandes(laboucle, 2)) - 3))
                                If Left$(UCase$(commandes(laboucle, 2)), 2) = "<P" Then   ' Token sur le prefixe
                                    If TokenVoulu <= VnbTokensPr Then
                                        If At(commandes(laboucle, 2), ",", 1) <> 0 Then ' On a mis un ou des paramètres
                                            If TokenVoulu = 0 Then  ' On demande le dernier tokken
                                                vtempo = TablTokensPr(VnbTokensPr)
                                            Else
                                                If Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", 2)) = "+" Or Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", 2)) = "-" Then
                                                    If Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", 2)) = "+" Then ' tout ce qu'il y a après le token n (y compris ce token)
                                                        vtempo = Mid$(cmdprefix, PTablTokensPr(TokenVoulu))
                                                    Else    ' -
                                                        vtempo = Left$(cmdprefix, (PTablTokensPr(TokenVoulu) + Len(TablTokensPr(TokenVoulu))) - 1)
                                                    End If
                                                Else
                                                    vtempo = TablTokensPr(TokenVoulu)
                                                End If
                                            End If
                                            For LaBoucle4 = 1 To CharOccurs(commandes(laboucle, 2), ",")
                                                vtempo = FmtToken(vtempo, Val(Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", LaBoucle4 + 1))))
                                            Next
                                            vnom = vnom + vtempo
                                        Else    ' Token sans paramètre
                                            If TokenVoulu = 0 Then  ' On demande le dernier tokken
                                                vnom = vnom + TablTokensPr(VnbTokensPr)
                                            Else
                                                vnom = vnom + TablTokensPr(TokenVoulu)
                                            End If
                                        End If
                                    End If
                                Else    ' Token sur l'extension ?
                                    If Left$(UCase$(commandes(laboucle, 2)), 2) = "<E" Then   ' Token sur l'extension
                                        If TokenVoulu <= VnbTokensEx Then
                                            If At(commandes(laboucle, 2), ",", 1) <> 0 Then ' On a mis un ou des paramètres
                                                If TokenVoulu = 0 Then  ' On demande le dernier tokken
                                                    vtempo = TablTokensEx(VnbTokensEx)
                                                Else
                                                    If Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", 2)) = "+" Or Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", 2)) = "-" Then
                                                        If Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", 2)) = "+" Then ' tout ce qu'il y a après le token n (y compris ce token)
                                                            vtempo = Mid$(cmdextension, PTablTokensEx(TokenVoulu))
                                                        Else    ' -
                                                            vtempo = Left$(cmdextension, (PTablTokensEx(TokenVoulu) + Len(TablTokensEx(TokenVoulu))) - 1)
                                                        End If
                                                    Else
                                                        vtempo = TablTokensEx(TokenVoulu)
                                                    End If
                                                End If
                                                For LaBoucle4 = 1 To CharOccurs(commandes(laboucle, 2), ",")
                                                    vtempo = FmtToken(vtempo, Val(Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", LaBoucle4 + 1))))
                                                Next
                                                vnom = vnom + vtempo
                                            Else    ' Token sans paramètre
                                                If TokenVoulu = 0 Then  ' On demande le dernier tokken
                                                    vnom = vnom + TablTokensEx(VnbTokensEx)
                                                Else
                                                    vnom = vnom + TablTokensEx(TokenVoulu)
                                                End If
                                            End If
                                        End If
                                    Else    ' Token pour le répertoire
                                        If TokenVoulu <= VnbTokensFo Then
                                            If At(commandes(laboucle, 2), ",", 1) <> 0 Then ' On a mis un ou des paramètres
                                                If TokenVoulu = 0 Then  ' On demande le dernier tokken
                                                    vtempo = TablTokensFo(VnbTokensFo)
                                                Else
                                                    If Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", 2)) = "+" Or Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", 2)) = "-" Then
                                                        If Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", 2)) = "+" Then ' tout ce qu'il y a après le token n (y compris ce token)
                                                            vtempo = Mid$(chemin, PTablTokensFo(TokenVoulu))
                                                        Else    ' -
                                                            vtempo = Left$(chemin, (PTablTokensFo(TokenVoulu) + Len(TablTokensFo(TokenVoulu))) - 1)
                                                        End If
                                                    Else
                                                        vtempo = TablTokensFo(TokenVoulu)
                                                    End If
                                                End If
                                                For LaBoucle4 = 1 To CharOccurs(commandes(laboucle, 2), ",")
                                                    vtempo = FmtToken(vtempo, Val(Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", LaBoucle4 + 1))))
                                                Next
                                                vnom = vnom + vtempo
                                            Else    ' Token sans paramètre
                                                If TokenVoulu = 0 Then  ' On demande le dernier tokken
                                                    vnom = vnom + TablTokensFo(VnbTokensFo)
                                                Else
                                                    vnom = vnom + TablTokensFo(TokenVoulu)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If  ' Test sur <P> ou <E> ou <F>
                            Case "0"  ' Litteral
                                vnom = vnom + commandes(laboucle, 2)
         
                            Case "1"  ' <curext>
                                vnom = vnom + cmdextension
         
                            Case "2" ' <curprefix>
                                vnom = vnom + cmdprefix
         
                            Case "3" ' <ddddd>
                                vnom = vnom + Compteur(val1, Val(commandes(laboucle, 2)), 0)
                                val1 = val1 + val2
         
                            Case "4" ' <EXCapital>
                                vnom = vnom + MyStrConv(cmdextension)
         
                            Case "5" ' <EXInvert>
                                vnom = vnom + StrReverse(cmdextension)
         
                            Case "6" ' <EXLower>
                                vnom = vnom + LCase$(cmdextension)
         
                            Case "7" ' <EXToggle>
                                vnom = vnom + ToggleCase(cmdextension)
         
                            Case "8" ' <EXUpper>
                                vnom = vnom + UCase$(cmdextension)
         
                            Case "9" ' <filecontent>
                                If IsFile Then
                                    If Val(commandes(laboucle, 2)) > 256 Then
                                        MsgBox "Numbers of characters to take from file is not valid. Change it !", vbOKOnly, "Error"
                                        ExitLang
                                    End If
                                    Open chemin & cmdprefix & "." & cmdextension For Input As #1
                                    letexte = Input(Val(commandes(laboucle, 2)), 1)
                                    Close #1
                                    If Len(Trim$(letexte)) = 0 Then
                                        MsgBox "Error, unable to rename file " + chemin & cmdprefix & "." & cmdextension + " because file's content is empty !"
                                        ExitLang
                                        Exit Sub
                                    End If
                                    vpos = InStr(Chr$(13) + Chr$(10), letexte)
                                    If vpos <> 0 Then
                                        letexte = Left$(letexte, vpos - 1)
                                    End If
                                    vnom = vnom + Trim$(Menage(letexte))
                                    If Len(vnom) = 0 Then
                                        MsgBox "Error, unable to rename file " + cmdprefix & "." & cmdextension + " because file's content is empty after removing unavailable characters !"
                                        ExitLang
                                        Exit Sub
                                    End If
                                End If
            
                            Case "10" ' <filedate>
                                vnom = vnom + FmtDate(FileDateTime(chemin & cmdprefix & "." & cmdextension))
         
                            Case "11" ' <filetime>
                                vnom = vnom + Menage(FmtHeure(FileDateTime(chemin & cmdprefix & "." & cmdextension)))
         
                            Case "12" ' <XXXXX>
                                vnom = vnom + Compteur(val1, Val(commandes(laboucle, 2)), 1)
                                val1 = val1 + val2
         
                            Case "13" ' <ooooo>
                                vnom = vnom + Compteur(val1, Val(commandes(laboucle, 2)), 2)
                                val1 = val1 + val2
         
                            Case "14" ' <PRCapital>
                                vnom = vnom + MyStrConv(cmdprefix)
         
                            Case "15" ' <PRInvert>
                                vnom = vnom + StrReverse(cmdprefix)
         
                            Case "16" ' <PRLower>
                                vnom = vnom + LCase$(cmdprefix)
         
                            Case "17" ' <PRToggle>
                                vnom = vnom + ToggleCase(cmdprefix)
         
                            Case "18" ' <PRUpper>
                                vnom = vnom + UCase$(cmdprefix)
         
                            Case "19" ' <systdate>
                                vnom = vnom + FmtDate(Now)
         
                            Case "20" ' <systtime>
                                vnom = vnom + Menage(FmtHeure(Time()))
         
                            Case "21" ' <ttfname>
                                If IsFile Then vnom = vnom + Menage(GetFontName(chemin & cmdprefix & "." & cmdextension))
            
                            Case "22" ' <html>
                                If IsFile Then vnom = vnom + GetHtmlName(chemin & cmdprefix & "." & cmdextension)
            
                            Case "23" ' <PRrtrim>
                                vnom = vnom + RTrim$(cmdprefix)
         
                            Case "24" ' <PRltrim>
                                vnom = vnom + LTrim$(cmdprefix)
         
                            Case "25" ' <PRtrim>
                                vnom = vnom + Trim$(cmdprefix)
         
                            Case "26" ' <EXrtrim>
                                vnom = vnom + RTrim$(cmdextension)
         
                            Case "27" ' <EXltrim>
                                vnom = vnom + LTrim$(cmdextension)
         
                            Case "28" ' <EXtrim>
                                vnom = vnom + Trim$(cmdextension)
         
                            Case "29" ' <ShortName>
                                vnom = vnom + ShortName(chemin & cmdprefix & "." & cmdextension)
         
                            Case "30" '<PRRemIntSp>
                                vnom = vnom + RInternalSpaces(cmdprefix)
         
                            Case "31" '<EXRemIntSp>
                                vnom = vnom + RInternalSpaces(cmdextension)
         
                            Case "32" '<Copy>
        
                            Case "33" ' <EXLeft>
                                If Val(commandes(laboucle, 2)) > 0 Then
                                    vnom = vnom + Left$(cmdextension, Val(commandes(laboucle, 2)))
                                Else
                                    If Len(cmdextension) + Val(commandes(laboucle, 2)) > 0 Then
                                        vnom = vnom + Left$(cmdextension, Len(cmdextension) + Val(commandes(laboucle, 2)))
                                    Else
                                        vnom = vnom + cmdextension
                                    End If
                                End If
         
                            Case "34" ' <EXRight>
                                If Val(commandes(laboucle, 2)) > 0 Then
                                    vnom = vnom + Right$(cmdextension, Val(commandes(laboucle, 2)))
                                Else
                                    If Len(cmdextension) + Val(commandes(laboucle, 2)) > 0 Then
                                        vnom = vnom + Right$(cmdextension, Len(cmdextension) + Val(commandes(laboucle, 2)))
                                    Else
                                        vnom = vnom + cmdextension
                                    End If
                                End If
         
                            Case "35" ' <PRLeft>
                                If Val(commandes(laboucle, 2)) > 0 Then
                                    vnom = vnom + Left$(cmdprefix, Val(commandes(laboucle, 2)))
                                Else
                                    If Len(cmdprefix) + Val(commandes(laboucle, 2)) > 0 Then
                                        vnom = vnom + Left$(cmdprefix, Len(cmdprefix) + Val(commandes(laboucle, 2)))
                                    Else
                                        vnom = vnom + cmdprefix
                                    End If
                                End If
         
                            Case "36" ' <PRRight>
                                If Val(commandes(laboucle, 2)) > 0 Then
                                    vnom = vnom + Right$(cmdprefix, Val(commandes(laboucle, 2)))
                                Else
                                    If Len(cmdprefix) + Val(commandes(laboucle, 2)) > 0 Then
                                        vnom = vnom + Right$(cmdprefix, Len(cmdprefix) + Val(commandes(laboucle, 2)))
                                    Else
                                        vnom = vnom + cmdprefix
                                    End If
                                End If
         
                            Case "37" ' <EXMiddle>
                                vnom = vnom + Mid$(cmdextension, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3)))
         
                            Case "38" ' <PRMiddle>
                                vnom = vnom + Mid$(cmdprefix, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3)))
         
                            Case "39" ' <EXPaddLeft>
                                If Len(cmdextension) <= Val(commandes(laboucle, 2)) Then
                                    vnom = vnom + String$(Val(commandes(laboucle, 2)) - Len(cmdextension), commandes(laboucle, 3)) + cmdextension
                                Else
                                    vnom = vnom + cmdextension
                                End If
         
                            Case "40" ' <EXPaddRight>
                                If Len(cmdextension) <= Val(commandes(laboucle, 2)) Then
                                    vnom = vnom + cmdextension + String$(Val(commandes(laboucle, 2)) - Len(cmdextension), commandes(laboucle, 3))
                                Else
                                    vnom = vnom + cmdextension
                                End If
         
                            Case "41" ' <PRPaddLeft>
                                If Len(cmdprefix) <= Val(commandes(laboucle, 2)) Then
                                    vnom = vnom + String$(Val(commandes(laboucle, 2)) - Len(cmdprefix), commandes(laboucle, 3)) + cmdprefix
                                Else
                                    vnom = vnom + cmdprefix
                                End If
         
                            Case "42" ' <PRPaddRight>
                                If Len(cmdprefix) <= Val(commandes(laboucle, 2)) Then
                                    vnom = vnom + cmdprefix + String$(Val(commandes(laboucle, 2)) - Len(cmdprefix), commandes(laboucle, 3))
                                Else
                                    vnom = vnom + cmdprefix
                                End If
         
                            Case "43" ' <EXToken>
                                ps.ParseDelimitedString cmdextension, commandes(laboucle, 3)
                                ps.MoveFirst
                                vnom = vnom + ps.TokenX(Val(commandes(laboucle, 2)) - 1)
         
                            Case "44" ' <PRToken>
                                ps.ParseDelimitedString cmdprefix, commandes(laboucle, 3)
                                ps.MoveFirst
                                vnom = vnom + ps.TokenX(Val(commandes(laboucle, 2)) - 1)
         
                            Case "45" ' <EXCapitalEX,0,0>
                                vnom = vnom + MyStrConv(Mid$(cmdextension, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "46" ' <EXInvertEX,0,0>
                                vnom = vnom + StrReverse(Mid$(cmdextension, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "47" ' <EXLowerEX,0,0>
                                vnom = vnom + LCase$(Mid$(cmdextension, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "48" ' <EXToggleEX,0,0>
                                vnom = vnom + ToggleCase(Mid$(cmdextension, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "49" ' <EXUpperEX,0,0>
                                vnom = vnom + UCase$(Mid$(cmdextension, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "50" ' <PRCapitalEX,0,0>
                                vnom = vnom + MyStrConv(Mid$(cmdprefix, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "51" ' <PRInvertEX,0,0>
                                vnom = vnom + StrReverse(Mid$(cmdprefix, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "52" ' <PRLowerEX,0,0>
                                vnom = vnom + LCase$(Mid$(cmdprefix, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "53" ' <PRToggleEX,0,0>
                                vnom = vnom + ToggleCase(Mid$(cmdprefix, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "54" ' <PRUpperEX,0,0>
                                vnom = vnom + UCase$(Mid$(cmdprefix, Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3))))
         
                            Case "55" ' <FileSize,0>
                                sataille = FileLen(chemin & cmdprefix & "." & cmdextension)
                                satailles = ""
                                Select Case Val(commandes(laboucle, 2))
                                    Case 0
                                        satailles = Trim$(Format$(sataille, "###############"))
                                    Case 1
                                        satailles = Trim$(Format$(sataille, "### ### ### ### ###"))
                                    Case Else
                                        satailles = Trim$(Format$(sataille, "### ### ### ### ###"))
                                End Select
                                vnom = vnom + satailles
         
                            Case "56" ' <FileAttr>
                                sesattrs = ""
                                sesattr = GetAttr(chemin & cmdprefix & "." & cmdextension)
                                ' chaine renvoyée au format RHSA
                                If sesattr And vbReadOnly Then sesattrs = sesattrs + "R"
                                If sesattr And vbHidden Then sesattrs = sesattrs + "H"
                                If sesattr And vbSystem Then sesattrs = sesattrs + "S"
                                If sesattr And vbArchive Then sesattrs = sesattrs + "A"
                                vnom = vnom + sesattrs
         
                            Case "57" ' <ImgInfo>
                                If IsFile Then
                                    simginfo = ImgInfo(chemin & cmdprefix & "." & cmdextension)
                                    vnom = vnom + simginfo
                                End If
            
                            Case "58" ' <CyclicSelection>
                                CompteurCyclic = CompteurCyclic + 1
                                If CompteurCyclic > VnbCyclic Then
                                    CompteurCyclic = 1
                                End If
                                If CompteurCyclic <= VnbCyclic Then
                                    vnom = vnom + LesCyclic(CompteurCyclic)
                                End If
          
                            Case "254" ' <PRBefore,0>
                                If InStr(1, cmdprefix, commandes(laboucle, 2)) - 1 > 0 Then
                                    vnom = vnom + Left$(cmdprefix, InStr(1, cmdprefix, commandes(laboucle, 2)) - 1)
                                End If
         
                            Case "60" ' <PRAfter,0>
                                vnom = vnom + Mid$(cmdprefix, InStr(1, cmdprefix, commandes(laboucle, 2)) + Len(commandes(laboucle, 2)))
         
                            Case "61" ' <PRBetween,0,0>
                                Pos1 = InStr(1, cmdprefix, commandes(laboucle, 2)) + Len(commandes(laboucle, 2))  ' Position de la première chaine
                                Pos2 = InStr(1, cmdprefix, commandes(laboucle, 3))  ' Position de la deuxième chaine
                                If InStr(1, cmdprefix, commandes(laboucle, 2)) > 0 Then
                                    If Pos2 - Pos1 > 0 Then
                                        vnom = vnom + Mid$(cmdprefix, Pos1, Pos2 - Pos1)
                                    End If
                                End If
         
                            Case "62" ' <EXBefore,0>
                                If InStr(1, cmdextension, commandes(laboucle, 2)) - 1 > 0 Then
                                    vnom = vnom + Left$(cmdextension, InStr(1, cmdextension, commandes(laboucle, 2)) - 1)
                                End If
         
                            Case "63" ' <EXAfter,0>
                                vnom = vnom + Mid$(cmdextension, InStr(1, cmdextension, commandes(laboucle, 2)) + Len(commandes(laboucle, 2)))
         
                            Case "64" ' <EXBetween,0,0>
                                Pos1 = InStr(1, cmdextension, commandes(laboucle, 2)) + Len(commandes(laboucle, 2))  ' Position de la première chaine
                                Pos2 = InStr(1, cmdextension, commandes(laboucle, 3))  ' Position de la deuxième chaine
                                If InStr(1, cmdextension, commandes(laboucle, 2)) > 0 Then
                                    If Pos2 - Pos1 > 0 Then
                                        vnom = vnom + Mid$(cmdextension, Pos1, Pos2 - Pos1)
                                    End If
                                End If
         
                            Case "65" ' <TextCounter>
                                vnom = vnom + Compteur(val1, Val(commandes(laboucle, 2)), 4)
                                val1 = val1 + val2
         
                            Case "66" ' <RandomPrefix>
                                LongRandom = 0
                                BoucleRandom = 0
                                NomRandom = ""
                                VraiRandom = True
                                While VraiRandom <> False
                                    While LongRandom = 0
                                        LongRandom = Int(Rnd * 24)
                                    Wend
                                    For BoucleRandom = 1 To LongRandom
                                        LetterRandom = 0
                                        While LetterRandom = 0
                                            LetterRandom = Int(Rnd * 26)
                                        Wend
                                        NomRandom = NomRandom + FBase26(LetterRandom)
                                    Next
                                    If LesOptions.UseLowerInLetterCounters = 1 Then ' passer en minuscules
                                        NomRandom = LCase$(NomRandom)
                                    End If
                                    VraiRandom = FileExists(chemin & vnom & NomRandom & "." & vsuffixe)
                                Wend
                                vnom = vnom & NomRandom
'                                vnom = vnom & GiveRandomName(chemin, vnom)
         
                            Case "67"  ' <EXCapitalFirst>
                                vnom = vnom + UCase$(Left$(cmdextension, 1)) + LCase$(Mid$(cmdextension, 2))
                            Case "68"  ' <PRCapitalFirst>
                                vnom = vnom + UCase$(Left$(cmdprefix, 1)) + LCase$(Mid$(cmdprefix, 2))
                            Case "69" ' <PRCowBoys>
                                vnom = vnom + CoWbOyS(cmdprefix)
                            Case "70" ' <EXCowBoys>
                                vnom = vnom + CoWbOyS(cmdextension)
                            Case "71" ' <PRRemMultSp>
                                vnom = vnom + RemoveMultipleSpacing(cmdprefix)
                            Case "72" ' <EXRemMultSp>
                                vnom = vnom + RemoveMultipleSpacing(cmdextension)
        
                            Case "73" ' <MP3Title,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Title, laboucle, vnom)
                                End If
                                
                            Case "74" ' <MP3Artist,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Artist, laboucle, vnom)
                                End If
                                
                            Case "75" ' <MP3Album,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Album, laboucle, vnom)
                                End If
                                
                            Case "76" ' <MP3Year,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Year, laboucle, vnom)
                                End If
                                
                            Case "77" ' <MP3Comment,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Comment, laboucle, vnom)
                                End If
                                
                            Case "78" ' <MP3Genre,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Genre, laboucle, vnom)
                                End If
                                
                            Case "79" ' <MP3Band,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Band, laboucle, vnom)
                                End If
                                
                            Case "80" ' <MP3BMP,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.BPM, laboucle, vnom)
                                End If
                                
                            Case "81" ' <MP3Composer,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Composer, laboucle, vnom)
                                End If
                                
                            Case "82" ' <MP3Conductor,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Conductor, laboucle, vnom)
                                End If
                                
                            Case "83" ' <MP3ContentGroup,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.ContentGroup, laboucle, vnom)
                                End If
            
                            Case "84" ' <MP3Copyright,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Copyright, laboucle, vnom)
                                End If
            
                            Case "85"  ' <PrIfEmpty,0>
                                If Trim$(Prefixe(vnom)) = "" Then
                                    If InStr(vnom, ".") = 0 Then  ' il n'y a pas de point
                                        vnom = commandes(laboucle, 2) + "." + Suffixe(vnom)
                                    Else
                                        vnom = commandes(laboucle, 2) + Suffixe(vnom)
                                    End If
                                End If
                                
                            Case "86"  ' <ExIfEmpty,0>
                                If Trim$(Suffixe(vnom)) = "" Then
                                    If InStr(vnom, ".") = 0 Then  ' il n'y a pas de point
                                        vnom = Prefixe(vnom) + "." + commandes(laboucle, 2)
                                    Else
                                        vnom = Prefixe(vnom) + commandes(laboucle, 2)
                                    End If
                                End If
                                
                            Case "87"   ' <PRLength>
                                vnom = vnom & Len(vprefixe)
                                
                            Case "88"   ' <EXLength>
                                vnom = vnom & Len(vsuffixe)
                                
                            Case "89"   ' <FileLength>
                                vnom = vnom & (Len(vprefixe) + Len(vsuffixe))
                                
                            Case "90"   ' <FullLength>
                                vnom = vnom & Len(vnomfichier)
                                
                            Case "91"   ' <PRSepWords>
                                vnom = vnom + ExtractWords(cmdprefix)
                                
                            Case "92"    ' <EXSepWords>
                                vnom = vnom + ExtractWords(cmdextension)
            
                            Case "93" ' <VQFArtist,,>
                                If IsFile Then
                                    SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusVQF.Author, laboucle, vnom)
                                End If
        
                            Case "94" ' <VQFBitrate,,>
                                If IsFile Then
                                    SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusVQF.BitRate, laboucle, vnom)
                                End If
        
                            Case "95" ' <VQFComment,,>
                                If IsFile Then
                                    SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusVQF.Comment, laboucle, vnom)
                                End If
        
                            Case "96" ' <VQFCopyright,,>
                                If IsFile Then
                                    SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusVQF.Copyright, laboucle, vnom)
                                End If
        
                            Case "97" ' <VQFFileSaveAs,,>
                                If IsFile Then
                                    SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusVQF.SaveAsFilename, laboucle, vnom)
                                End If
        
                            Case "98" ' <VQFMonoStereo,,>
                                If IsFile Then
                                    SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusVQF.Mono_Stereo, laboucle, vnom)
                                End If
        
                            Case "99" ' <VQFQuality,,>
                                If IsFile Then
                                    SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusVQF.Quality, laboucle, vnom)
                                End If
        
                            Case "100" ' <VQFSampleRate,,>
                                If IsFile Then
                                    SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusVQF.SampleRate, laboucle, vnom)
                                End If
        
                            Case "101" ' <VQFTitle>
                                If IsFile Then
                                    SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusVQF.Title, laboucle, vnom)
                                End If
                            Case "102" ' <MP3EncryptionMethod,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.EncryptionMethod, laboucle, vnom)
                                End If
        
                            Case "103" ' <MP3Date,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.mDate, laboucle, vnom)
                                End If

                            Case "104" ' <MP3EncodedBy,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.EncodedBy, laboucle, vnom)
                                End If
        
                            Case "105" ' <MP3SoftwareEncodingSettings,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.SoftwareEncodingSettings, laboucle, vnom)
                                End If
        
                            Case "106" ' <MP3FileOwner,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.FileOwner, laboucle, vnom)
                                End If
        
                            Case "107" ' <MP3FileType,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.FileType, laboucle, vnom)
                                End If
        
                            Case "108" ' <MP3GroupIdent,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.GroupIdent, laboucle, vnom)
                                End If
        
                            Case "109" ' <MP3InitialKey,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.InitialKey, laboucle, vnom)
                                End If
        
                            Case "110" ' <MP3InvolvedPeopleList,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.InvolvedPeopleList, laboucle, vnom)
                                End If
        
                            Case "111" ' <MP3Isrc,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.ISRC, laboucle, vnom)
                                End If
        
                            Case "112" ' <MP3Language,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Language, laboucle, vnom)
                                End If
        
                            Case "113" ' <MP3LinkedInformation,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.LinkedInformation, laboucle, vnom)
                                End If
        
                            Case "114" ' <MP3Lyricist,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Lyricist, laboucle, vnom)
                                End If
        
                            Case "115" ' <MP3MediaType,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.MediaType, laboucle, vnom)
                                End If
        
                            Case "116" ' <MP3MixArtist,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.MixArtist, laboucle, vnom)
                                End If
        
                            Case "117" ' <MP3NetRadioOwner,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.NetRadioOwner, laboucle, vnom)
                                End If
        
                            Case "118" ' <MP3NetRadioStation,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.NetRadioStation, laboucle, vnom)
                                End If
        
                            Case "119" ' <MP3OriginalAlbum,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.OriginalAlbum, laboucle, vnom)
                                End If
        
                            Case "120" ' <MP3OriginalArtist,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.OriginalArtist, laboucle, vnom)
                                End If
        
                            Case "121" ' <MP3OriginalFilename,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.OriginalFilename, laboucle, vnom)
                                End If
        
                            Case "122" ' <MP3OriginalLyricist,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.OriginalLyricist, laboucle, vnom)
                                End If
        
                            Case "123" ' <MP3OriginalYear,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.OriginalYear, laboucle, vnom)
                                End If
        
                            Case "124" ' <MP3PartOfASet,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.PartOfASet, laboucle, vnom)
                                End If
        
                            Case "125" ' <MP3PlayListDelay,,>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.PlayListDelay, laboucle, vnom)
                                End If
        
                            Case "126" ' <RandomNumber,0,100,4>
                                BoucleRandom = 0
                                NomRandom = ""
                                VraiRandom = True
                                While VraiRandom <> False
                                    LetterRandom = Val(commandes(laboucle, 2)) - 1
                                    While LetterRandom < Val(commandes(laboucle, 2)) Or LetterRandom > Val(commandes(laboucle, 3))
                                        LetterRandom = Int(Rnd * Val(commandes(laboucle, 3)))
                                    Wend
                                    NomRandom = Compteur(LetterRandom, Val(commandes(laboucle, 4)), 0)
                                    VraiRandom = FileExists(chemin & vnom & NomRandom & "." & vsuffixe)
                                Wend
                                vnom = vnom & NomRandom
'                                vnom = vnom & GiveRandomName(chemin, vnom)
        
                            Case "127" ' <SystDateEx,expr>
                                vnom = vnom + Format$(Now, commandes(laboucle, 2))
            
                            Case "128"  ' <PathPart,1>
                                vnom = vnom + Trim$(GetToken(ExtractPath(vNomComplet), "\", Val(commandes(laboucle, 2))))
        
                            Case "129" ' <FileDateEx,Date,Format>
                                Select Case Val(commandes(laboucle, 2))
                                    Case 1  ' Creation date
                                        vnom = vnom + LesDates(vNomComplet, 1, commandes(laboucle, 3))
                                    Case 2  ' Last Acces date
                                        vnom = vnom + LesDates(vNomComplet, 2, commandes(laboucle, 3))
                                    Case 3  ' Last modified date
                                        vnom = vnom + LesDates(vNomComplet, 3, commandes(laboucle, 3))
                                End Select
                            Case "130"  ' <PRRefomartNumber,4,0>
                                vnom = vnom + ReformatNumbers(cmdprefix, Val(commandes(laboucle, 2)), commandes(laboucle, 3))
            
                            Case "131"  ' <ExRefomartNumber,4,0>
                                vnom = vnom + ReformatNumbers(cmdextension, Val(commandes(laboucle, 2)), commandes(laboucle, 3))
            
                            Case "132"  ' <PRSepThousand, >
                                vnom = vnom + SeparateThousands(cmdprefix, commandes(laboucle, 2))
            
                            Case "133"  ' <EXSepThousand, >
                                vnom = vnom + SeparateThousands(cmdextension, commandes(laboucle, 2))
            
                            Case "134"  ' <MP3PopulariMeter>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.PopulariMeter, laboucle, vnom)
                                End If

                            Case "135"  ' <MP3Publisher>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Publisher, laboucle, vnom)
                                End If

                            Case "136"  ' <MP3RecordingDates>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.RecordingDates, laboucle, vnom)
                                End If

                            Case "137"  ' <MP3SongLength>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.SongLength, laboucle, vnom)
                                End If

                            Case "138"  ' <MP3SubTitle>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.SubTitle, laboucle, vnom)
                                End If

                            Case "139"  ' <MP3SynchronizedLyric>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.SynchronizedLyric, laboucle, vnom)
                                End If

                            Case "140"  ' <MP3TermsOfUse>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.TermsOfUse, laboucle, vnom)
                                End If

                            Case "141"  ' <MP3Time>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.Time, laboucle, vnom)
                                End If

                            Case "142"  ' <MP3TrackNumber>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.TrackNumber, laboucle, vnom)
                                End If

                            Case "143"  ' <MP3TotalTracks>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.TotalTracks, laboucle, vnom)
                                End If

                            Case "144"  ' <MP3UnsynchronizedLyric>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.UnsynchronizedLyric, laboucle, vnom)
                                End If

                            Case "145"  ' <MP3UserText>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.UserText, laboucle, vnom)
                                End If

                            Case "146"  ' <MP3wwwArtist>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.wwwArtist, laboucle, vnom)
                                End If

                            Case "147"  ' <MP3wwwAudioFile>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.wwwAudioFile, laboucle, vnom)
                                End If

                            Case "148"  ' <MP3wwwAudioSource>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.wwwAudioSource, laboucle, vnom)
                                End If

                            Case "149"  ' <MP3wwwCommercialInfo>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.wwwCommercialInfo, laboucle, vnom)
                                End If

                            Case "150"  ' <MP3wwwCopyright>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.wwwCopyright, laboucle, vnom)
                                End If

                            Case "151"  ' <MP3wwwPayment>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.wwwPayment, laboucle, vnom)
                                End If

                            Case "152"  ' <MP3wwwPublisher>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.wwwPublisher, laboucle, vnom)
                                End If

                            Case "153"  ' <MP3wwwRadioPage>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.wwwRadioPage, laboucle, vnom)
                                End If

                            Case "154"  ' <MP3wwwUserURL>
                                If IsFile Then
                                    SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                                    vnom = MP3Commands(MusMP3.wwwUserURL, laboucle, vnom)
                                End If
        
                            Case "155"  ' <CounterEx>
                                vnom = vnom + Compteur(val7, Val(commandes(laboucle, 5)), Val(commandes(laboucle, 2)))
                                If CompteurFichier = Val(commandes(laboucle, 6)) Then
                                    val7 = val7 + val8
                                    CompteurFichier = 0
                                End If
            
                            Case "156"  ' <PRPrevFileOldName>
                                vnom = vnom + PRPrevFileOldName
            
                            Case "157"  ' <PRPrevFileNewName>
                                vnom = vnom + PRPrevFileNewName
            
                            Case "158"  ' <EXPrevFileOldName>
                                vnom = vnom + EXPrevFileOldName
            
                            Case "159"  ' <EXPrevFileNewName>
                                vnom = vnom + EXPrevFileNewName
            
                            Case "160"  ' <PRModifyCounter,1,4,0>
                                vnom = vnom + ReformatNumbers(cmdprefix, Val(commandes(laboucle, 3)), commandes(laboucle, 4), Val(commandes(laboucle, 2)))
            
                            Case "161"  ' <EXModifyCounter,1,4,0>
                                vnom = vnom + ReformatNumbers(cmdextension, Val(commandes(laboucle, 3)), commandes(laboucle, 4), Val(commandes(laboucle, 2)))
            
                            Case "162" ' <OggNumberOfTags>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = vnom + MusOgg.NumberOfTags
                                End If
        
                            Case "163" ' <OggSerialNumber>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.SerialNumber, laboucle, vnom)
                                End If
        
                            Case "164" ' <OggEncoderVersion>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.EncoderVersion, laboucle, vnom)
                                End If
        
                            Case "165" ' <OggLowerBitrate>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.LowerBitrate, laboucle, vnom)
                                End If
        
                            Case "166" ' <OggUpperBitrate>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.UpperBitrate, laboucle, vnom)
                                End If
        
                            Case "167" ' <OggNominalBitrate>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.NominalBitrate, laboucle, vnom)
                                End If
        
                            Case "168" ' <OggAverageBitrate>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.AverageBitrate, laboucle, vnom)
                                End If
        
                            Case "169" ' <OggChannels>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Channels, laboucle, vnom)
                                End If
        
                            Case "170" ' <OggSampleRate>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.SampleRate, laboucle, vnom)
                                End If
        
                            Case "171" ' <OggVendor>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Vendor, laboucle, vnom)
                                End If
        
                            Case "172" ' <OggPlaytime>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Playtime, laboucle, vnom)
                                End If
        
                            Case "173" ' <OggLength>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Length, laboucle, vnom)
                                End If
        
                            Case "174" ' <OggISRC>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.ISRC, laboucle, vnom)
                                End If
        
                            Case "175" ' <OggDate>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.SaDate, laboucle, vnom)
                                End If
        
                            Case "176" ' <OggCopyRight>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Copyright, laboucle, vnom)
                                End If
        
                            Case "177" ' <OggLocation>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Location, laboucle, vnom)
                                End If
        
                            Case "178" ' <OggDescription>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Description, laboucle, vnom)
                                End If
        
                            Case "179" ' <OggOrganization>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Organization, laboucle, vnom)
                                End If
        
                            Case "180" ' <OggTotalTracks>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.TotalTracks, laboucle, vnom)
                                End If
        
                            Case "181" ' <OggTrackNumber>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.TrackNumber, laboucle, vnom)
                                End If
        
                            Case "182" ' <OggVersion>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Version, laboucle, vnom)
                                End If
        
                            Case "183" ' <OggComment>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Comment, laboucle, vnom)
                                End If
        
                            Case "184" ' <OggAlbum>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Album, laboucle, vnom)
                                End If
        
                            Case "185" ' <OggGenre>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Genre, laboucle, vnom)
                                End If
        
                            Case "186" ' <OggArtist>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Artist, laboucle, vnom)
                                End If
        
                            Case "187" ' <OggTitle>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Title, laboucle, vnom)
                                End If
            
                            Case "188" ' <OggComposer>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Composer, laboucle, vnom)
                                End If
        
                            Case "189" ' <OggConductor>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Conductor, laboucle, vnom)
                                End If
            
                            Case "190"  ' <OggEnsemble>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Ensemble, laboucle, vnom)
                                End If
        
                            Case "191"  ' <OggPerformer>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusOgg.Performer, laboucle, vnom)
                                End If
            
                            Case "192"  ' <OggGetUnknowTags,1,=>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = vnom + MusOgg.GetUnknowTags(Val(commandes(laboucle, 2)), commandes(laboucle, 3))
                                End If
        
                            Case "193"  ' <OggGetAllTags,1,=>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = vnom + MusOgg.GetAllTags(Val(commandes(laboucle, 2)), commandes(laboucle, 3))
                                End If
        
                            Case "194"  ' <OggTagByName,NomTag,Format,Separateur,[Literal],[Position]>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands2(MusOgg.GetTagByName(commandes(laboucle, 2), Val(commandes(laboucle, 3)), commandes(laboucle, 4)), laboucle, vnom)
                                End If
        
                            Case "195"  ' <OggTagByPosition,Position,Format,Separateur,[Literal],[Position]>
                                If IsFile Then
                                    SonInfo = MusOgg.GetOggInfos(vNomComplet, False)
                                    vnom = MP3Commands2(MusOgg.GetTagByPosition(Val(commandes(laboucle, 2)), Val(commandes(laboucle, 3)), commandes(laboucle, 4)), laboucle, vnom)
                                End If
        
                            Case "196"  ' <PRInsert,1,text>
                                If Val(commandes(laboucle, 2)) > 0 Then
                                    vnom = vnom + Left$(cmdprefix, Val(commandes(laboucle, 2)) - 1) + commandes(laboucle, 3) + Mid$(cmdprefix, Val(commandes(laboucle, 2)))
                                End If
        
                            Case "197"  ' <EXInsert,1,text>
                                If Val(commandes(laboucle, 2)) > 0 Then
                                    vnom = vnom + Left$(cmdextension, Val(commandes(laboucle, 2)) - 1) + commandes(laboucle, 3) + Mid$(cmdextension, Val(commandes(laboucle, 2)))
                                End If
         
                            Case "198"  ' <PRRemove,1,2> Paramètres, position de début et position de fin
                                If Val(commandes(laboucle, 2)) > 0 And Val(commandes(laboucle, 3)) > 0 Then
                                    vnom = vnom + Left$(cmdprefix, Val(commandes(laboucle, 2)) - 1) + Mid$(cmdprefix, Val(commandes(laboucle, 3)) + 1)
                                End If
            
                            Case "199"  ' <EXRemove,1,2> Paramètres, position de début et position de fin
                                If Val(commandes(laboucle, 2)) > 0 And Val(commandes(laboucle, 3)) > 0 Then
                                    vnom = vnom + Left$(cmdextension, Val(commandes(laboucle, 2)) - 1) + Mid$(cmdextension, Val(commandes(laboucle, 3)) + 1)
                                End If
         
                            Case "200"  ' <PRDeleteText,Text,1,-1,Yes> Paramètres, 1=Text, 2=Posdeb, 3=Count (-1=all), 4=Match Case
                                If Val(commandes(laboucle, 3)) > 0 Then
                                    If UCase$(commandes(laboucle, 5)) = "YES" Then  ' Match Case
                                        vnom = vnom + Replace(cmdprefix, commandes(laboucle, 2), "", Val(commandes(laboucle, 3)), Val(commandes(laboucle, 4)), vbBinaryCompare)
                                    Else    ' Don't match case
                                        vnom = vnom + Replace(cmdprefix, commandes(laboucle, 2), "", Val(commandes(laboucle, 3)), Val(commandes(laboucle, 4)), vbTextCompare)
                                    End If
                                End If
        
                            Case "201"  ' <EXDeleteText,Text,1,-1,Yes> Paramètres, 1=Text, 2=Posdeb, 3=Count (-1=all), 4=Match Case
                                If Val(commandes(laboucle, 3)) > 0 Then
                                    If UCase$(commandes(laboucle, 5)) = "YES" Then  ' Match Case
                                        vnom = vnom + Replace(cmdextension, commandes(laboucle, 2), "", Val(commandes(laboucle, 3)), Val(commandes(laboucle, 4)), vbBinaryCompare)
                                    Else    ' Don't match case
                                        vnom = vnom + Replace(cmdextension, commandes(laboucle, 2), "", Val(commandes(laboucle, 3)), Val(commandes(laboucle, 4)), vbTextCompare)
                                    End If
                                End If
        
                            Case "202"  ' <SelectedFilesCount>
                                vnom = vnom + Trim$(Str$(LVGetCountSelected(ListView1)))
                                
                            Case "203"  ' <NonSelectedFilesCount>
                                vnom = vnom + Trim$(Str$(ListView1.ListItems.Count - LVGetCountSelected(ListView1)))
                                
                            Case "204"  ' <TotalFilesCount>
                                vnom = vnom + Trim$(Str$(ListView1.ListItems.Count))
                                
                            Case "205"  ' <RomanCounter>
                                vnom = vnom + Compteur(val1, Val(commandes(laboucle, 2)), 5)
                                val1 = val1 + val2
                            
                            Case "206"  ' <AfmFontMetricsVersion>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.FontMetricsVersion, laboucle, vnom)
                                End If
                                
                            Case "207"  ' <AfmWeight>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.Weight, laboucle, vnom)
                                End If
                                
                            Case "208"  ' <AfmNotice>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.Notice, laboucle, vnom)
                                End If
                                
                            Case "209"  ' <AfmMetricsSets>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.MetricsSets, laboucle, vnom)
                                End If
                                
                            Case "210"  ' <AfmFullName>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.FullName, laboucle, vnom)
                                End If
                                
                            Case "211"  ' <AfmFontVersion>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.FontVersion, laboucle, vnom)
                                End If
                                
                            Case "212"  ' <AfmFontName>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.FontName, laboucle, vnom)
                                End If
                                
                            Case "213"  ' <AfmFamilyName>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.FamilyName, laboucle, vnom)
                                End If
                                
                            Case "214"  ' <AfmEncodingScheme>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.EncodingScheme, laboucle, vnom)
                                End If
                                
                            Case "215"  ' <AfmCharacterSet>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.CharacterSet, laboucle, vnom)
                                End If
                                
                            Case "216"  ' <AfmCopyright>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.Copyright, laboucle, vnom)
                                End If
                                
                            Case "217"  ' <AfmCreationDate>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.CreationDate, laboucle, vnom)
                                End If
                                
                            Case "218"  ' <AfmUniqueID>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.UniqueID, laboucle, vnom)
                                End If
                                
                            Case "219"  ' <AfmVMusage>
                                If IsFile Then
                                    AFM.GetAFMInfos vNomComplet
                                    vnom = MP3Commands(AFM.VMusage, laboucle, vnom)
                                End If
                                
                            Case "220"  ' <ExifAperture>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.Aperture, laboucle, vnom)
                                End If
                                
                            Case "221"  ' <ExifBrightness>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.Brightness, laboucle, vnom)
                                End If
                            
                            Case "222"  ' <ExifCompressedBitsPerPixel>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.CompressedBitsPerPixel, laboucle, vnom)
                                End If
                                
                            Case "223"  ' <ExifCopyright>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.Copyright, laboucle, vnom)
                                End If
                            
                            Case "224"  ' <ExifDateTime>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.DateTime, laboucle, vnom)
                                End If
                                
                            Case "225"  ' <ExifDateTimeDigitized>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.DateTimeDigitized, laboucle, vnom)
                                End If
                            
                            Case "226"  ' <ExifDateTimeOriginal>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.DateTimeOriginal, laboucle, vnom)
                                End If
                            
                            Case "227"  ' <ExifVersion>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.Version, laboucle, vnom)
                                End If
                                
                            Case "228"  ' <ExifExposureBias>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.ExposureBias, laboucle, vnom)
                                End If
                                
                            Case "229"  ' <ExifExposureProgram>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.ExposureProgram, laboucle, vnom)
                                End If
                            
                            Case "230"  ' <ExifExposureTime>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.ExposureTime, laboucle, vnom)
                                End If
                            
                            Case "231"  ' <ExifFirmwareVersion>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.FirmwareVersion, laboucle, vnom)
                                End If
                                
                            Case "232"  ' <ExifFlash>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.Flash, laboucle, vnom)
                                End If
                            
                            Case "233"  ' <ExifFNumber>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.FNumber, laboucle, vnom)
                                End If
                            
                            Case "234"  ' <ExifFocalLength>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.FocalLength, laboucle, vnom)
                                End If
                            
                            Case "235"  ' <ExifFocalPlaneResolutionUnit>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.FocalPlaneResolutionUnit, laboucle, vnom)
                                End If
                            
                            Case "236"  ' <ExifFocalPlaneXResolution>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.FocalPlaneXResolution, laboucle, vnom)
                                End If
                            
                            Case "237"  ' <ExifFocalPlaneYResolution>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.FocalPlaneYResolution, laboucle, vnom)
                                End If
                                
                            Case "238"  ' <ExifImageDescription>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.ImageDescription, laboucle, vnom)
                                End If
                            
                            Case "239"  ' <ExifImageHeight>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.ImageHeight, laboucle, vnom)
                                End If
                            
                            Case "240"  ' <ExifImageWidth>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.ImageWidth, laboucle, vnom)
                                End If
                            
                            Case "241"  ' <ExifISOSpeedRatings>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.ISOSpeedRatings, laboucle, vnom)
                                End If
                            
                            Case "242"  ' <ExifMake>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.Make, laboucle, vnom)
                                End If
                            
                            Case "243"  ' <ExifMaxAperture>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.MaxAperture, laboucle, vnom)
                                End If
                            
                            Case "244"  ' <ExifMeteringMode>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.MeteringMode, laboucle, vnom)
                                End If
                            
                            Case "245"  ' <ExifModel>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.Model, laboucle, vnom)
                                End If
                                
                            Case "246"  ' <ExifOrientation>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.Orientation, laboucle, vnom)
                                End If
                            
                            Case "247"  ' <ExifRelatedSoundFile>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.RelatedSoundFile, laboucle, vnom)
                                End If
                                
                            Case "248"  ' <ExifResolutionUnit>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.ResolutionUnit, laboucle, vnom)
                                End If
                            
                            Case "249"  ' <ExifShutterSpeed>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.ShutterSpeed, laboucle, vnom)
                                End If
                                
                            Case "250"  ' <ExifSubjectDistance>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.SubjectDistance, laboucle, vnom)
                                End If
                            
                            Case "251"  ' <ExifWhiteBalance>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.WhiteBalance, laboucle, vnom)
                                End If
                                
                            Case "252"  ' <ExifXResolution>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.XResolution, laboucle, vnom)
                                End If
                                
                            Case "253"  ' <ExifYResolution>
                                If IsFile Then
                                    SonInfo = PicEXIF.GetEXIFInfos(vNomComplet, False)
                                    vnom = MP3Commands(PicEXIF.YResolution, laboucle, vnom)
                                End If
                            
                            
                            Case "59"  ' <PRbeforeEx,Expr,Opt1,Opt2,Opt3....Opt10>
                                If InStr(1, cmdprefix, commandes(laboucle, 2)) - 1 > 0 Then
                                        vtempo = Left$(cmdprefix, InStr(1, cmdprefix, commandes(laboucle, 2)) - 1)
                                        For LaBoucle4 = 2 To 6
                                            vtempo = FmtToken(vtempo, Val(commandes(laboucle, LaBoucle4)))
                                        Next
                                        vnom = vnom + vtempo
                                End If

                            Case "255"  ' <WmaChannelMode>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.ChannelMode, laboucle, vnom)
                                End If
                            
                            Case "256"  ' <WmaSampleRate>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.SampleRate, laboucle, vnom)
                                End If
                            
                            Case "257"  ' <WmaDuration>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.Duration, laboucle, vnom)
                                End If
                            
                            Case "258"  ' <WmaBitRate>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.BitRate, laboucle, vnom)
                                End If
                            
                            Case "259"  ' <WmaTrack>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.Track, laboucle, vnom)
                                End If
                            
                            Case "260"  ' <WmaTitle>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.Title, laboucle, vnom)
                                End If
                            
                            Case "261"  ' <WmaArtist>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.Artist, laboucle, vnom)
                                End If
                            
                            Case "262"  ' <WmaAlbum>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.Album, laboucle, vnom)
                                End If
                            
                            Case "263"  ' <WmaGenre>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.Genre, laboucle, vnom)
                                End If
                            
                            Case "264"  ' <WmaComment>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.Comment, laboucle, vnom)
                                End If
                            
                            Case "265"  ' <WmaYear>
                                If IsFile Then
                                    SonInfo = MusWMA.GetWMAInfos(vNomComplet, False)
                                    vnom = MP3Commands(MusWMA.Year, laboucle, vnom)
                                End If
                        End Select ' Quelle commande ?
                    Next '    Boucle sur les commandes
                    If acpreview And vTemCopy Then
                        If TemoinTrop = False Then
                            preview.Command5.Visible = False
                            vMontre = True
                            If LesOptions.ShowWhenFileNameChange = 1 Then
                                If Trim$(vnomfichier) = Trim$(chemin & vnom) Then
                                    vMontre = False
                                End If
                            End If
                            If vMontre Then
                                Set itmX = preview.listPreview.ListItems.Add(, , vnomfichier)
                                itmX.SubItems(1) = chemin & vnom
                                preview.listsav.AddItem vnomdesti
                            End If
                        End If
                    Else
                        If vTemCopy Then
                            If DisplayRenMsg = 1 Then
                                état.Panels(1).Text = "Copying " + vnomfichier + " to " + vnom
                            End If
                            fileop.AddSourceFile chemin + cmdprefix + "." + cmdextension
                            fileop.AddDestFile chemin + vnom
                            fileop.ConfirmOperation = False
                            If Not fileop.CopyFiles Then
                                ExitLang
                                Exit Sub
                            End If
                            fileop.ClearSourceFiles
                            fileop.ClearDestFiles
                            vnom = ""
                        End If
                    End If
                Next ' Boucle sur les fichiers à copier en plusieurs exemplaires
                cmdprefix = Prefixe(vnom)
                cmdextension = Suffixe(vnom)
                
                LesRecherches1 cmdprefix, cmdextension, 1, 2
                vnom = cmdprefix & "." & cmdextension
    
            Case 14 ' Short name
                If Recursive Then
                    vnom = ShortName(vNomComplet)
                    vprefixe = Prefixe(vnom)
                    vsuffixe = Suffixe(vnom)
                    vnomdesti = Prefixe(vnom) + "." + Suffixe(vnom)
                    vnom = vnomdesti
                Else
                    vnom = ShortName(vnomorig)
                    vprefixe = Prefixe(vnom)
                    vsuffixe = Suffixe(vnom)
                    vnomdesti = vnom
                End If
    
            Case 15 ' Remove internal spaces
                vprefixe = RInternalSpaces(vprefixe)
    
            Case 17 ' Capitalize first word only
                vprefixe = UCase$(Left$(vprefixe, 1)) + LCase$(Mid$(vprefixe, 2))
                
            Case 18 ' Rename with a random prefix
                LongRandom = 0
                BoucleRandom = 0
                NomRandom = ""
                VraiRandom = True
                While VraiRandom <> False
                    While LongRandom = 0
                        LongRandom = Int(Rnd * 24)
                    Wend
                    For BoucleRandom = 1 To LongRandom
                        LetterRandom = 0
                        While LetterRandom = 0
                            LetterRandom = Int(Rnd * 26)
                        Wend
                        NomRandom = NomRandom + FBase26(LetterRandom)
                    Next
                    If LesOptions.UseLowerInLetterCounters = 1 Then ' passer en minuscules
                        NomRandom = LCase$(NomRandom)
                    End If
                    VraiRandom = FileExists(chemin & NomRandom & "." & vsuffixe)
                Wend
                vprefixe = NomRandom
'                vprefixe = GiveRandomName(chemin, vnom)
    
            Case 19 ' CoWbOyS
                vprefixe = CoWbOyS(vprefixe)
     
            Case 20 ' Remove Multiple Spacing
                vprefixe = RemoveMultipleSpacing(vprefixe)
                
            Case 21 ' Separate Words
                vprefixe = ExtractWords(vprefixe)
        End Select
   
        If RechGlob Then
            If LesOptions.SearchAndReplace = 1 Or LesOptions.SearchAndReplace = 2 Then
                rech3.SourceString = vprefixe + "." + vsuffixe
                vtmpGlob = rech3.BeginSearchAndReplace
                vprefixe = Prefixe(vtmpGlob)
                vsuffixe = Suffixe(vtmpGlob)
     
                rech3.SourceString = vprefixe + "." + vsuffixe
                vtmpGlob = rech3.BeginReplaceCharacters
                vprefixe = Prefixe(vtmpGlob)
                vsuffixe = Suffixe(vtmpGlob)
            End If
        End If
   
        If RechPref Then
            If LesOptions.SearchAndReplace = 1 Or LesOptions.SearchAndReplace = 2 Then
                rech1.SourceString = vprefixe
                vprefixe = rech1.BeginSearchAndReplace
                rech1.SourceString = vprefixe
                vprefixe = rech1.BeginReplaceCharacters
            End If
        End If
   
        ' Abréviations
        If OkUseAbbrev Then  ' il faut utiliser les abbréviations
            If LesOptions.SearchAndReplace = 0 Or LesOptions.SearchAndReplace = 2 Then
                For ij = 1 To CollAbrev.Count   ' Boucle sur toutes les abbréviations de la collection
                    If GetToken(CollAbrev.Item(ij), Chr$(254), 7) = "yes" Then ' on utilise des expressions régulières
                        'On Error Resume Next ********** Laisser tel quel
                        Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                            Case "YES"          ' on recherche partout (préfixe et suffixe)
                                str2 = vprefixe + "." + vsuffixe
                            Case "PREFIX"      ' Préfixe uniquement
                                str2 = vprefixe
                            Case "EXTENSION"    ' Extension uniquement
                                str2 = vsuffixe
                        End Select
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 4) = "yes" Then
                            vnbcount = 9999
                        Else
                            vnbcount = Val(GetToken(CollAbrev.Item(ij), Chr$(254), 4))
                        End If
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 5) = "yes" Then
                            str3 = 1
                        Else
                            str3 = 0
                        End If
                        Match = RegSub(str2, GetToken(CollAbrev.Item(ij), Chr$(254), 1), GetToken(CollAbrev.Item(ij), Chr$(254), 2), str1, str3, vnbcount, 0, 0)
                        If Match Then
                            Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                                Case "YES"          ' on recherche partout (préfixe et suffixe)
                                    vprefixe = Prefixe(str1)
                                    vsuffixe = Suffixe(str1)
                                Case "PREFIX"      ' Préfixe uniquement
                                    vprefixe = str1
                                Case "EXTENSION"    ' Extension uniquement
                                    vsuffixe = str1
                            End Select
                        End If
                    Else ' on n'utilise pas d'expression régulière, recherche "normale"
                        Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                            Case "YES"          ' on recherche partout (préfixe et suffixe)
                                str2 = vprefixe + "." + vsuffixe
                            Case "PREFIX"      ' Préfixe uniquement
                                str2 = vprefixe
                            Case "EXTENSION"    ' Extension uniquement
                                str2 = vsuffixe
                        End Select
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 4) = "yes" Then
                            vnbcount = 9999
                        Else
                            vnbcount = Val(GetToken(CollAbrev.Item(ij), Chr$(254), 4))
                        End If
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 5) = "yes" Then
                            str3 = vbBinaryCompare
                        Else
                            str3 = vbTextCompare
                        End If
                        str1 = Replace(str2, GetToken(CollAbrev.Item(ij), Chr$(254), 1), GetToken(CollAbrev.Item(ij), Chr$(254), 2), Val(GetToken(CollAbrev.Item(ij), Chr$(254), 3)), vnbcount, str3)
                        Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                            Case "YES"          ' on recherche partout (préfixe et suffixe)
                                vprefixe = Prefixe(str1)
                                vsuffixe = Suffixe(str1)
                            Case "PREFIX"      ' Préfixe uniquement
                                vprefixe = str1
                            Case "EXTENSION"    ' Extension uniquement
                                vsuffixe = str1
                        End Select
                    End If
                Next
            End If
        End If
   
        If RechSuff Then
            If LesOptions.SearchAndReplace = 0 Or LesOptions.SearchAndReplace = 2 Then
                rech2.SourceString = vsuffixe
                vsuffixe = rech2.BeginSearchAndReplace
                rech2.SourceString = vsuffixe
                vsuffixe = rech2.BeginReplaceCharacters
            End If
        End If

' *************************************************************************
' ***************** Action à effectuer sur le suffixe *********************
' *************************************************************************
        If Frame2.Visible Then   ' On essaye de gagner du temps, si les actions possibles sur le suffixe ne sont pas visibles, c'est pas la peine de faire des tests
            Select Case vaction2
                Case 11 ' Effacer le suffixe
                    vsuffixe = ""
                Case 1 ' Mettre en majuscules
                    vsuffixe = UCase$(vsuffixe)
                Case 2 ' Mettre en minuscules
                    vsuffixe = LCase$(vsuffixe)
                Case 3 ' Inverser majuscules/minucules
                    vsuffixe = ToggleCase(vsuffixe)
                Case 6 ' Inverser les lettres
                    vsuffixe = StrReverse(vsuffixe)
                Case 5 ' Capitalize all words
                    vsuffixe = MyStrConv(vsuffixe)
                Case 4 ' Conserver le suffixe
                    vsuffixe = vsuffixe
                Case 7 ' Remplacer par la date système
                    vsuffixe = FmtDate(Now)
                Case 8 ' Remplacer par l'heure système
                    vsuffixe = Menage(FmtHeure(Time()))
                Case 9 ' Remplacer par la date + l'heure système
                    vsuffixe = FmtDate(Now) + Menage(FmtHeure(Time()))
                Case 12 ' Remove internal spaces
                    vsuffixe = RInternalSpaces(vsuffixe)
                Case 13 ' CoWbOyS
                    vsuffixe = CoWbOyS(vsuffixe)
                Case 14 ' Remove Multiple Spacing
                    vsuffixe = RemoveMultipleSpacing(vsuffixe)
                Case 15 ' Separate Words
                    vsuffixe = ExtractWords(vsuffixe)
     
                Case 10 ' Modifier l'extension
                    If Option4(0).Value Then  ' remplacer par un texte fixe
                        vsuffixe = Text8.Text
                    Else
                        If Option4(1).Value Then  ' ajouter un texte fixe
                            If Option5(0).Value Then  ' au début
                                vsuffixe = Text15.Text + LTrim$(vsuffixe)
                            Else ' à la fin
                                vsuffixe = RTrim$(vsuffixe) + Text15.Text
                            End If
                        End If
                    End If
      
                    If Check11.Value = 1 Then ' ajouter un compteur
                        If Option3(26).Value Then   ' Rajouter à gauche
                            vsuffixe = Compteur(val4, val6, vformat2) + vsuffixe
                        Else
                            If Option3(25).Value Then  ' Rajouter à droite
                                vsuffixe = vsuffixe + Compteur(val4, val6, vformat2)
                            Else ' Remplacer le préfixe par le compteur
                                vsuffixe = Compteur(val4, val6, vformat2)
                            End If
                        End If
                        val4 = val4 + val5
                    End If
      
                    If Check12.Value = 1 Then ' ajouter la taille
                        If Option3(29).Value Then  ' Rajouter à gauche
                            vsuffixe = FileLen(vNomComplet) & vsuffixe
                        Else
                            If Option3(28).Value Then  ' Rajouter à droite
                                vsuffixe = vsuffixe & FileLen(vNomComplet)
                            Else ' Remplacer le préfixe par la taille
                                vsuffixe = FileLen(vNomComplet)
                            End If
                        End If
                    End If
      
                    If Check13.Value = 1 Then ' ajouter la date
                        If Option3(30).Value Then  ' Rajouter à gauche
                            vsuffixe = FmtDate(FileDateTime(vNomComplet)) + vsuffixe
                        Else
                            If Option3(31).Value Then  'Rajouter à droite
                                vsuffixe = vsuffixe + FmtDate(FileDateTime(vNomComplet))
                            Else 'Remplacer le préfixe par la date
                                vsuffixe = FmtDate(FileDateTime(vNomComplet))
                            End If
                        End If
                    End If
      
                    If Check4.Value = 1 Then ' ajouter l'heure
                        If Option3(14).Value Then  ' Rajouter à gauche
                            vsuffixe = Menage(FmtHeure(FileDateTime(vNomComplet))) + vsuffixe
                        Else
                            If Option3(13).Value Then  'Rajouter à droite
                                vsuffixe = vsuffixe + Menage(FmtHeure(FileDateTime(vNomComplet)))
                            Else 'Remplacer le préfixe par la date
                                vsuffixe = Menage(FmtHeure(FileDateTime(vNomComplet)))
                            End If
                        End If
                    End If
      
            End Select
        End If ' Si la frame pour l'extension est visible
    
        If RechSuff Then
            If LesOptions.SearchAndReplace = 1 Or LesOptions.SearchAndReplace = 2 Then
                rech2.SourceString = vsuffixe
                vsuffixe = rech2.BeginSearchAndReplace
                rech2.SourceString = vsuffixe
                vsuffixe = rech2.BeginReplaceCharacters
            End If
        End If
   
        ' Abréviations
        If OkUseAbbrev Then  ' il faut utiliser les abbréviations
            If LesOptions.SearchAndReplace = 1 Or LesOptions.SearchAndReplace = 2 Then
                For ij = 1 To CollAbrev.Count   ' Boucle sur toutes les abbréviations de la collection
                    If GetToken(CollAbrev.Item(ij), Chr$(254), 7) = "yes" Then ' on utilise des expressions régulières
                        Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                            Case "YES"          ' on recherche partout (préfixe et suffixe)
                                str2 = vprefixe + "." + vsuffixe
                            Case "PREFIX"      ' Préfixe uniquement
                                str2 = vprefixe
                            Case "EXTENSION"    ' Extension uniquement
                                str2 = vsuffixe
                        End Select
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 4) = "yes" Then
                            vnbcount = 9999
                        Else
                            vnbcount = Val(GetToken(CollAbrev.Item(ij), Chr$(254), 4))
                        End If
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 5) = "yes" Then
                            str3 = 1
                        Else
                            str3 = 0
                        End If
                        Match = RegSub(str2, GetToken(CollAbrev.Item(ij), Chr$(254), 1), GetToken(CollAbrev.Item(ij), Chr$(254), 2), str1, str3, vnbcount, 0, 0)
                        If Match Then
                            Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                                Case "YES"          ' on recherche partout (préfixe et suffixe)
                                    vprefixe = Prefixe(str1)
                                    vsuffixe = Suffixe(str1)
                                Case "PREFIX"      ' Préfixe uniquement
                                    vprefixe = str1
                                Case "EXTENSION"    ' Extension uniquement
                                    vsuffixe = str1
                            End Select
                        End If
                    Else ' on n'utilise pas d'expression régulière, recherche "normale"
                        Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                            Case "YES"          ' on recherche partout (préfixe et suffixe)
                                str2 = vprefixe + "." + vsuffixe
                            Case "PREFIX"      ' Préfixe uniquement
                                str2 = vprefixe
                            Case "EXTENSION"    ' Extension uniquement
                                str2 = vsuffixe
                        End Select
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 4) = "yes" Then
                            vnbcount = 9999
                        Else
                            vnbcount = Val(GetToken(CollAbrev.Item(ij), Chr$(254), 4))
                        End If
                        If GetToken(CollAbrev.Item(ij), Chr$(254), 5) = "yes" Then
                            str3 = vbBinaryCompare
                        Else
                            str3 = vbTextCompare
                        End If
                        str1 = Replace(str2, GetToken(CollAbrev.Item(ij), Chr$(254), 1), GetToken(CollAbrev.Item(ij), Chr$(254), 2), Val(GetToken(CollAbrev.Item(ij), Chr$(254), 3)), vnbcount, str3)
                        Select Case UCase$(GetToken(CollAbrev.Item(ij), Chr$(254), 6))
                            Case "YES"          ' on recherche partout (préfixe et suffixe)
                                vprefixe = Prefixe(str1)
                                vsuffixe = Suffixe(str1)
                            Case "PREFIX"      ' Préfixe uniquement
                                vprefixe = str1
                            Case "EXTENSION"    ' Extension uniquement
                                vsuffixe = str1
                        End Select
                    End If
                Next
            End If
        End If
   
        If vTemCopy = False Then
            If Frame2.Visible Then   ' On essaye de gagner du temps, si les actions possibles sur le suffixe ne sont pas visibles, c'est pas la peine de faire des tests
                vnomdesti = vprefixe & "." & vsuffixe
                vnom = vnomdesti
            Else  ' La frame de l'extension n'est pas visible
                vnomdesti = vnom
            End If ' Si la frame pour l'extension est visible

            PRPrevFileOldName = Prefixe(vnomorig)
            EXPrevFileOldName = Suffixe(vnomorig)

            fileop.AddSourceFile chemin + vnomorig
            If LesOptions.RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
                vnom = RemIllegals(vnom)
                vnomdesti = RemIllegals(vnomdesti)
            End If
            If LesOptions.RemoveStartingSpaces = 1 Then ' Il faut supprimer les espaces de début de fichier
                vnom = LTrim$(vnom)
                vnomdesti = LTrim$(vnomdesti)
            End If

            PRPrevFileNewName = Prefixe(vnom)
            EXPrevFileNewName = Suffixe(vnom)

            fileop.AddDestFile chemin + vnom
            If DisplayRenMsg = 1 Then
                état.Panels(1).Text = "Rename " + vnomorig + " to " + vnom
                état.Panels(2).Text = Trim$(Str$(vnb)) + "/" + Trim$(Str$(vnbtot))
            End If
    
            If acpreview = False Then  ' On n'est pas en preview
                If temoin1 = 1 Then ' Undofile
                    If Recursive = False Then
                        Print #1, "ren " + Chr$(34) + vnom + Chr$(34) + " " + Chr$(34) + vnomorig + Chr$(34)
                    End If
                End If
     
                If temoin3 = 1 Then ' Logfile
                    ' Format, Date, heure, ancien nom, nouveau nom
                    Print #3, Str$(Date) + vbTab + Str$(Time) + vbTab + chemin + vnomorig + vbTab + chemin + vnom
                End If
     
                If temoin2 = 1 Then ' Batch
                    If Recursive = False Then
                        vtmp1 = Replace(vnomorig, Chr$(0), "")
                        vtmp2 = Replace(vnom, Chr$(0), "")
                        Print #2, "ren " + Chr$(34) + vtmp1 + Chr$(34) + " " + Chr$(34) + vtmp2 + Chr$(34)
                    End If
                Else
                    ' Les ajouts dans les listes sont pour l'UNDO
                    If TemoinTrop = False Then
                        List2.AddItem chemin + vnomorig ' Nom d'origine.
                        List3.AddItem chemin + vnomdesti   ' Nom d'arrivée.
                    End If
                    If LesOptions.UseHistory Then
                        If TemoinTrop = False Then
                            lhistory.AddItem Trim$(Str$(Time())) + "|" + chemin + "|" + vnomorig + "|" + vnomdesti ' Historique
                        End If
                    End If
                    If Len(Trim$(LesOptions.prog1)) <> 0 Then ' Lancer un programme avant de renommer le fichier
                        prog11 = LesOptions.prog1
                        ExecCmd prog11, chemin + vnomorig
                    End If
                    If LesOptions.CopyRename Then   ' Renommer les fichier et copier
                        If Not fileop.RenameFiles Then
                        End If
                    Else ' On copie les fichiers, on ne les renomme pas
                        If Not fileop.CopyFiles Then
                        End If
                    End If
                End If
                If Len(Trim$(LesOptions.prog2)) <> 0 Then ' Lancer un programme après avoir renommé le fichier
                    prog22 = LesOptions.prog2
                    ExecCmd prog22, chemin + vnomdesti
                End If
                DT1.SetFileDateTime (chemin + vnomdesti)
                Attr1.ChangeAttr (chemin + vnomdesti)
            Else ' On est en preview *********************************************************************
                If Recursive = False Then
                    If TemoinTrop = False Then
                        vMontre = True
                        If LesOptions.ShowWhenFileNameChange = 1 Then
                            If Trim$(vnomorig) = Trim$(vnomdesti) Then
                                vMontre = False
                            End If
                        End If
                        If vMontre Then
                            Set itmX = preview.listPreview.ListItems.Add(, , vnomorig)
                            itmX.SubItems(1) = vnomdesti
                            preview.listsav.AddItem vnomdesti
                        End If
                    End If
                Else ' On n'est pas en mode récursif
                    If TemoinTrop = False Then
                        vMontre = True
                        If LesOptions.ShowWhenFileNameChange = 1 Then
                            If Trim$(vnomfichier) = Trim$(chemin + vnomdesti) Then
                                vMontre = False
                            End If
                        End If
                        If vMontre Then
                            Set itmX = preview.listPreview.ListItems.Add(, , vnomfichier)
                            itmX.SubItems(1) = chemin + vnomdesti
                            preview.listsav.AddItem vnomdesti
                        End If
                    End If
                End If
            End If ' Preview ou pas ? *********************************************************************************
    
            fileop.ClearSourceFiles
            fileop.ClearDestFiles
        End If
fin:
        i = LVGetItemSelected(ListView1, i)
    Wend

zsuite:
    If acpreview = False Then ' Si on n'est pas en preview, il faut rafraichir l'écran
        remplissage
    End If

    If DisplayRenMsg = 1 Then état.Panels(1).Text = "Ok"
    RENAME.MousePointer = 0
    Rem ** Mise à jour des paramètres
    With LesOptions
        .LastUseDate = Date
        .LastUseTime = Time
        .NumberOFiles = vnbtot
        .LastDirectory = Dir1Path
    End With

    If acpreview Then
        With preview
            .Command7.Visible = True
            .Command8.Visible = True
            .Command1.Visible = True
            .Command6.Visible = True
            .Command2.Visible = True
            .Command3.Visible = True
            .Command4.Visible = True
            .StatusBar1.SimpleText = Trim$(Str$(vnb)) + " file(s) to rename"
        End With
        If vTemCopy = False Then
            preview.Command5.Visible = True
        End If
        If Recursive Then preview.Command5.Visible = False
        If ShowPreviewList = 0 Then preview.listPreview.Visible = True
    Else ' On n'est pas en preview
        If Len(Trim$(LesOptions.prog3)) <> 0 Then ' Lancer un programme après avoir renommé tous les fichiers
            prog33 = LesOptions.prog3
            ExecCmd prog33, ""
        End If
    End If
    Close #1
    Close #2
    Close #3

    If List2.ListCount > 0 Then
        mundo.Enabled = True
    End If

    If DisplayRenMsg = 1 Then
        état.Panels(1).Text = ""
        état.Panels(2).Text = ""
    End If

    If LesOptions.ShutDown And acpreview = False Then Unload Me
    Exit Sub
' **************** Fin des haricots ************************************************************************************************************************************************
Erreur1:
    MsgBox "Error, unable to create the undo file " + LesOptions.UndoFile + ", verify it's name and path (did you only specify a name?)"
    Exit Sub

Erreur2:
    MsgBox "Error, unable to create the batch file " + LesOptions.batch + ", verify it's name path  (did you only specify a name?)"
    Exit Sub

Erreur3:
    MsgBox "Error, unable to create the log file " + LesOptions.LogFile + ", verify it's name and path  (did you only specify a name?)"
    Exit Sub

ErrGen:
    If Err.Number = 53 Or (Err.Number = 62 And Temoin11) Then ' File not found...
        Resume Next
    End If
    ErreurGrave "StartRename"
    Exit Sub
End Sub

Private Sub StepSelection()
 Dim vretour As String
 Dim i As Long
 Dim pas As Integer
 Dim vnb As Long, vnb2 As Long
 vretour = InputBox("Enter step for selection", "Step", "2")
 If vretour = "" Then Exit Sub
 RENAME.MousePointer = 11
 letat = True
 pas = Val(vretour)
 ListView1.Visible = False
 vnb2 = ListView1.ListItems.Count - 1
 For i = 0 To vnb2 Step pas
    LVSetItemSelected ListView1, i
    vnb = vnb + 1
 Next
 RENAME.MousePointer = 0
 état.Panels(4).Text = Trim$(Str$(vnb))
 ListView1.Visible = True
 letat = False
 ListView1.SetFocus
End Sub

Private Sub Unselect()
 Dim i As Long, vnb As Long
 letat = True
 RENAME.MousePointer = 11
 ListView1.Visible = False
 vnb = ListView1.ListItems.Count - 1
 For i = 0 To vnb
    LVSetItemNotSelected ListView1, i
 Next
 ListView1.Visible = True
 RENAME.MousePointer = 0
 état.Panels(4).Text = "0"
 letat = False
 ListView1.SetFocus
End Sub
Private Sub Acdsee_DblClick()
    Dim vret As Boolean
    Dim vtmp As String
    If Recursive = False Then
        vtmp = AddBackSlash(Dir1Path)
    Else
        vtmp = ""
    End If
    vret = FViewPict.ChargeImage(vtmp & ListView1.SelectedItem.Text, ListView1.SelectedItem.Text)
End Sub

Private Sub Acdsee_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 107    ' ZoomIn    (+)
            Acdsee.ZoomIn
        Case 109    ' Zoom Out  (-)
            Acdsee.ZoomOut
        Case 106    ' Stretch  (*)
            Acdsee.Stretch
    End Select
End Sub

Private Sub Acdsee_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu MPicture
    End If
End Sub

Private Sub Check1_Click()
 If Check1.Value = 1 Then
  Picture8.Visible = True
 Else
  Picture8.Visible = False
 End If
 ModifierLigne
End Sub

Private Sub Check11_Click()
 If Check11.Value = 1 Then
  onglcounter2.Visible = True
 Else
  onglcounter2.Visible = False
 End If
 ModifierLigne
End Sub

Private Sub Check12_Click()
If Check12.Value = 1 Then
 Picture13.Visible = True
Else
 Picture13.Visible = False
End If
ModifierLigne
End Sub

Private Sub Check13_Click()
If Check13.Value = 1 Then
 Picture14.Visible = True
Else
 Picture14.Visible = False
End If
ModifierLigne
End Sub

Private Sub Check3_Click()
 If Check3.Value = 1 Then
  onglcounter.Visible = True
  If LesOptions.CompleCounters = 0 Then
    If LesOptions.AskQuestion = 0 Then ' Variable cachée dans la base de registres pour éviter la question chiante
        MsgBox "Warning, you are using counters but with the current options, they will not be padded with zeros. Like this, you will have some names like file1.txt, file2.txt, file3.txt ... instead of file01.txt, file02.txt, file03.txt. If you want to change this, go in the options, tab 'Other' and verify that the option 'Complete counters with 0s' is checked", vbOKOnly, "WARNING"
    End If
  End If
 Else
  onglcounter.Visible = False
 End If
 ModifierLigne
End Sub

Private Sub Check4_Click()
 If Check4.Value = 1 Then
  Picture4.Visible = True
 Else
  Picture4.Visible = False
 End If
 ModifierLigne
End Sub
Private Sub Check5_Click()
 If Check5.Value = 1 Then
  Picture5.Visible = True
 Else
  Picture5.Visible = False
 End If
 ModifierLigne
End Sub

Private Sub Check6_Click()
 If Check6.Value = 1 Then
  Picture6.Visible = True
 Else
  Picture6.Visible = False
 End If
 ModifierLigne
End Sub

Private Sub Check7_Click()
 If Check7.Value = 1 Then
  Picture7.Visible = True
 Else
  Picture7.Visible = False
 End If
 ModifierLigne
End Sub
Private Sub cmdclear_Click()
txtlang.Text = ""
txtlang.SetFocus
End Sub

Private Sub cmdpictures_Click(Index As Integer)
    InsertTextInTextBoxFromMenu txtlang, cmdpictures(Index).Caption
End Sub

Private Sub Combo1_Click()
Dim tControle As Integer
Dim i As Integer

tControle = GetKeyState(VK_SHIFT)
If tControle = -127 Or tControle = -128 Or TemShift = True Then
 For i = 0 To Combo2.ListCount - 1
  If Trim$(Combo2.List(i)) = Trim$(Combo1.List(Combo1.ListIndex)) Then
    Combo2.ListIndex = i
   Exit For
  End If
 Next
 TemShift = False
End If
  
 For i = 1 To vnboptionp
  If Trim$(Combo1.List(Combo1.ListIndex)) = Trim$(optionp(i)) Then
   laidep.Caption = aidep(i)
   Exit For
  End If
 Next

If Trim$(Combo1.List(Combo1.ListIndex)) = "Rename from a list" Then
 état.Panels(1).Text = ""
 Command3.Visible = False
 Frame2.Visible = False
 laides.Visible = False
 panelcmd.Visible = False
 Frame1.height = Frame3.Top - Frame1.Top
 PanelList.Move FrameDroite.Left + PanelPrefix.Left, PanelPrefix.Top + FrameDroite.Top
 Set PanelList.Container = TabGen
 PanelList.Visible = True
 PanelList.height = Frame1.height - Combo1.height - Combo1.Top - 140
 m_oAutoPos.RefreshPositions
 Exit Sub
Else
 Frame1.height = 2600
 PanelList.Visible = False
 Frame2.Visible = True
 laides.Visible = True
End If

If Trim$(Combo1.List(Combo1.ListIndex)) = "Free form" Then
 état.Panels(1).Text = ""
 panelcmd.Move FrameDroite.Left + PanelPrefix.Left, PanelPrefix.Top + FrameDroite.Top
 panelcmd.height = Frame3.Top - panelcmd.Top
 TV1.height = panelcmd.height - TV1.Top
 TV1.width = (panelcmd.width - TV1.Left) '- 25
 panelcmd.Visible = True
 Frame2.Visible = False
 laides.Visible = False
 Combo6.Visible = False
 Command3.Visible = False
 Set panelcmd.Container = TabGen
Else
 m3contextuel.Visible = False
 Command3.Visible = True
 panelcmd.Visible = False
 Frame2.Visible = True
 laides.Visible = True
End If

If Trim$(Combo1.List(Combo1.ListIndex)) = "Replace long filename with short filename" Then
 Frame2.Visible = False
 laides.Visible = False
End If

If Trim$(Combo1.List(Combo1.ListIndex)) = "Replace with file's content" Then
 paneltext.Move FrameDroite.Left + PanelPrefix.Left, PanelPrefix.Top + FrameDroite.Top
 Set paneltext.Container = TabGen
 paneltext.Visible = True
Else
 paneltext.Visible = False
End If
  
 If Trim$(Combo1.List(Combo1.ListIndex)) = "Modify prefix" Then
  PanelPrefix.Visible = True
  If Option1(0).Value = True Then
   Option1_Click 0
  End If
  If Option1(1).Value = True Then
   Option1_Click 1
  End If
  paneltext.Visible = False
 Else
  PanelPrefix.Visible = False
  état.Panels(1).Text = ""
 End If
 ModifierLigne
End Sub

Private Sub Combo2_Click()
Dim tControle As Integer
Dim i As Integer

tControle = GetKeyState(VK_SHIFT)
If tControle = -127 Or tControle = -128 Or TemShift = True Then
 For i = 0 To Combo1.ListCount - 1
  If Trim$(Combo1.List(i)) = Trim$(Combo2.List(Combo2.ListIndex)) Then
    Combo1.ListIndex = i
   Exit For
  End If
 Next
 TemShift = False
End If
 
 For i = 1 To vnboptions
  If Trim$(Combo2.List(Combo2.ListIndex)) = Trim$(options(i)) Then
   laides.Caption = aides(i)
   Exit For
  End If
 Next
 
 If Trim$(Combo2.List(Combo2.ListIndex)) = "Modify extension" Then
  PanelExt.Visible = True
Rem ajouts **************************************
  Check11_Click
  Check4_Click
  Check12_Click
  Check13_Click
  If Option4(0).Value = True Then
   Option4_Click 0
  End If
  If Option4(1).Value = True Then
   Option4_Click 1
  End If
 Else
  PanelExt.Visible = False
  état.Panels(1).Text = ""
 End If
ModifierLigne
End Sub

Private Sub Combo3_Change()
 If Combo3.List(Combo3.ListIndex) = "Letters" Then
  If Text3.Text = "0" Then
   Text3.Text = "1"
  End If
 End If
End Sub

Private Sub Combo3_Click()
 If Combo3.List(Combo3.ListIndex) = "Letters" Then
  If Text3.Text = "0" Then
   Text3.Text = "1"
  End If
 End If
End Sub

Private Sub Combo4_Change()
 If Combo4.List(Combo4.ListIndex) = "Letters" Then
  If Text16.Text = "0" Then
   Text16.Text = "1"
  End If
 End If
End Sub

Private Sub Combo4_Click()
 If Combo4.List(Combo4.ListIndex) = "Letters" Then
  If Text16.Text = "0" Then
   Text16.Text = "1"
  End If
 End If
End Sub

Private Sub Combo5_Click()
 Dim vnb As Long
 Filtre = Trim$(Combo5.Text)
 LesOptions.LastFilter = Filtre
 If Right$(Filtre, 1) = ";" Then
  Filtre = Left$(Filtre, Len(Filtre) - 1)
 End If
 Combo5.Text = Filtre
 vnb = remplissage()
 état.Panels(3).Text = Trim$(Str$(vnb))
 état.Panels(4).Text = "0"
End Sub

Private Sub Combo5_GotFocus()
    SelAll Combo5
End Sub

Private Sub Combo5_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim vnb As Long
 If KeyCode = 13 Then
  Filtre = Trim$(Combo5.Text)
  If Right$(Filtre, 1) = ";" Then
   Filtre = Left$(Filtre, Len(Filtre) - 1)
  End If
  Combo5.Text = Filtre
  vnb = remplissage()
  état.Panels(3).Text = Trim$(Str$(vnb))
  état.Panels(4).Text = "0"
 End If
End Sub
Private Sub Combo6_DblClick()
    ' Le double clic est pris comme une validation
    txtlang.Text = Left$(txtlang.Text, LaPosSauve - 1) + Trim$(Mid$(Combo6.List(Combo6.ListIndex), LongSauve + 1)) + Trim$(Mid$(txtlang.Text, txtlang.SelStart))
    txtlang.SelStart = LaPosSauve + Len(Trim$(Mid$(Combo6.List(Combo6.ListIndex), LongSauve + 1))) - 1
    Combo6.Visible = False
    Combo6.Clear
End Sub
Private Sub Combo6_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then ' Touche Esc
    Combo6.Visible = False
    Combo6.Clear
    txtlang.SetFocus
End If

If KeyCode = 32 Or KeyCode = 13 Then ' Barre d'espace
    txtlang.Text = Left$(txtlang.Text, LaPosSauve - 1) + Trim$(Mid$(Combo6.List(Combo6.ListIndex), LongSauve + 1)) + Trim$(Mid$(txtlang.Text, txtlang.SelStart))
    txtlang.SelStart = LaPosSauve + Len(Trim$(Mid$(Combo6.List(Combo6.ListIndex), LongSauve + 1))) - 1
    Combo6.Visible = False
    Combo6.Clear
End If
End Sub

Private Sub Combo6_LostFocus()
    Combo6.Visible = False
End Sub

Private Sub Command1_Click()
    FPICT.Show 1
End Sub

Private Sub Command10_Click()
 Folder1 = 0
 Folder2 = 0
 Folder4 = 0
 Folder5 = ""
 Folder6 = ""
 FolderOk = False
 UseMP3 = False
 UseVQF = False
 UseOGG = False
 UseWMA = False
 MusMP3.Clear
 MusVQF.Clear
 Check7.Value = 0
 Check5.Value = 0
 Check6.Value = 0
 Option3(9).Value = False
 Option3(10).Value = False
 Option3(11).Value = False
 Option3(3).Value = False
 Option3(4).Value = False
 Option3(5).Value = False
 Option3(8).Value = False
 Option3(7).Value = False
 Option3(6).Value = False
 Option1(0).Value = False
 Text2.Text = ""
 Option1(1).Value = False
 Text14.Text = ""
 Option2(0).Value = True
 Option2(1).Value = False
 Check3.Value = 0
 Text3.Text = "1"
 Text4.Text = "1"
 Text5.Text = "4"
 Combo3.ListIndex = 0
 Option3(0).Value = False
 Option3(1).Value = False
 Option3(2).Value = False
 Folder1 = 0
 Folder2 = 0
 Folder3 = 0
 Folder4 = 0
 Folder5 = ""
 Folder6 = ""
 Text2.Visible = False
 Command8.Visible = False
 Option2(0).Visible = False
 Option2(1).Visible = False
 Text14.Visible = False
 rech1.ResetSearch
End Sub

Private Sub Command11_Click()
Check11.Value = 0
Text16.Text = "1"
Text17.Text = "1"
Text18.Text = "4"
Combo4.ListIndex = 0
Option3(26).Value = False
Option3(25).Value = False
Option3(24).Value = False
Option4(0).Value = False
Text8.Text = ""
Text15.Text = ""
Option5(0).Value = True
Option5(1).Value = False
Check4.Value = 0
Option3(14).Value = False
Option3(13).Value = False
Option3(12).Value = False
Check13.Value = 0
Option3(30).Value = False
Option3(31).Value = False
Option3(32).Value = False
Check12.Value = 0
Option3(29).Value = False
Option3(28).Value = False
Option3(27).Value = False
Text8.Visible = False
Text15.Visible = False
Option5(0).Visible = False
Option5(1).Visible = False
rech2.ResetSearch
Option4(1).Value = False
Option4(0).Value = False
End Sub

Private Sub Command12_Click()
Dim szFilename As String
Dim chemin As String
Dim ligne As String, Chaine1 As String, Chaine2 As String, vretour As Integer
Dim ff As Integer
If ListView2.ListItems.Count > 0 Then
    vretour = MsgBox("Current list contains items, do you want to delete them ?", vbOKCancel, "Open an existing list")
    If vretour = vbOK Then
        ListView2.ListItems.Clear
    End If
End If
szFilename = DialogFile(Me.hWnd, 1, "Open list file", "rename.list", "List" & Chr$(0) & "*.list" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "list")
If Trim$(szFilename) = "" Then
    Exit Sub
End If
ListView2.Visible = False
RENAME.MousePointer = 11
chemin = ExtractPath(szFilename)
ChDir chemin
TemMove = False
FolderTreeview1(0).Visible = False
FolderTreeview1(0).SelectedFolder = chemin
Dir1Path = chemin
FolderTreeview1(0).Visible = True
If LesOptions.ListDelimiter = 0 Then LesOptions.ListDelimiter = 9

ff = FreeFile
Open szFilename For Input As #ff
Line Input #ff, ligne
While Not EOF(ff)
    If Len(Trim$(ligne)) > 0 Then
        If LesOptions.RemoveGuill = 1 Then
            ligne = Replace(ligne, Chr$(34), "")
        End If
        Chaine1 = GetToken(ligne, Chr$(LesOptions.ListDelimiter), 1)
        Chaine2 = GetToken(ligne, Chr$(LesOptions.ListDelimiter), 2)
        Set itmX = ListView2.ListItems.Add(, , Chaine1)
        itmX.Text = Chaine1
        itmX.SubItems(1) = Chaine2
    End If
    Line Input #ff, ligne
Wend
Close #ff
If Len(Trim$(ligne)) > 0 Then
    If LesOptions.RemoveGuill = 1 Then ligne = Replace(ligne, Chr$(34), "")
    Chaine1 = GetToken(ligne, Chr$(LesOptions.ListDelimiter), 1)
    Chaine2 = GetToken(ligne, Chr$(LesOptions.ListDelimiter), 2)
    Set itmX = ListView2.ListItems.Add(, , Chaine1)
    itmX.Text = Chaine1
    itmX.SubItems(1) = Chaine2
End If

ListView2.Visible = True
RENAME.MousePointer = 0
Exit Sub

erreur:
    MsgBox "There was an error during process !"
    ListView2.Visible = True
    RENAME.MousePointer = 0
    Exit Sub
End Sub

Private Sub Command13_Click()
Dim i As Long
Dim R As Long
Dim vnb As Long
ListView2.Visible = False
RENAME.MousePointer = 11
vnb = ListView2.ListItems.Count
For i = vnb To 0 Step -1
  If LVIsSelected(ListView2, i) = True Then
   R = LVRemoveItem(ListView2, i)
 End If
Next
ListView2.Visible = True
RENAME.MousePointer = 0
End Sub

Private Sub Command14_Click()
    FCmd.Show 1
    ChargeVNBCommandes
End Sub

Private Sub Command15_Click()
    CopySelected True
End Sub

Private Sub Command16_Click()
Dim i As Long, vnb As Long
Dim sItem As String
RENAME.MousePointer = 11
ListView2.Visible = False
vnb = ListView1.ListItems.Count - 1
For i = 0 To vnb
 sItem = LVGetName(ListView1, i)
 Set itmX = ListView2.ListItems.Add(, , sItem)
 itmX.Text = sItem
 itmX.SubItems(1) = sItem
Next
ListView2.Visible = True
RENAME.MousePointer = 0
End Sub
Private Sub Command19_Click()
 ffolder.Show 1
End Sub

Private Sub Command2_Click()
    FLstMan.Show 0
End Sub

Private Sub Command3_Click()
 TemShift = True
 Combo1_Click
End Sub

Private Sub Command4_Click()
 TemShift = True
 Combo2_Click
End Sub

Private Sub Command5_Click()
 FMP3.Show 1
End Sub

Private Sub Command6_Click()
 OptionsCyclic = False
 Fcyclic.Show 1
End Sub

Private Sub Command7_Click()
Dim szFilename As String
Dim sItem1 As String, sItem2 As String
Dim i As Long
Dim vnb As Long
Dim ff As Integer
If ListView2.ListItems.Count <= 0 Then
    MsgBox "Sorry, there's nothing to save"
    Exit Sub
End If

szFilename = DialogFile(Me.hWnd, 2, "Save as", "rename.list", "Text" & Chr$(0) & "*.list" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "list")

RENAME.MousePointer = 11
If Trim$(szFilename) = "" Then
    RENAME.MousePointer = 0
    Exit Sub
End If

ff = FreeFile
Open szFilename For Output As #ff
If LesOptions.ListDelimiter = 0 Then LesOptions.ListDelimiter = 9
ListView2.Visible = False
vnb = ListView2.ListItems.Count - 1
For i = 0 To vnb
 sItem1 = LVGetName(ListView2, i)
 sItem2 = LVGetItemName(ListView2, i, 1)
 Print #ff, sItem1 & Chr$(LesOptions.ListDelimiter) & sItem2
Next
Close #ff
RENAME.MousePointer = 0
ListView2.Visible = True
Beep
End Sub

Private Sub Command8_Click()
 OptionsCyclic = True
 Fcyclic.Show 1
End Sub
Private Sub FolderTreeview1_FolderClick(Index As Integer, Folder As CCRPFolderTV6.Folder, Location As CCRPFolderTV6.ftvHitTestConstants)
 Dim fldCur As Folder
 Dim vnb As Long
 On Error Resume Next   ' fldCur's Folder object references may be Nothing
 Set fldCur = FolderTreeview1(Index).SelectedFolder
 If TemMove = False Then
  AjoutHistorique fldCur.FullPath
 Else
  TemMove = False
 End If
 Dir1Path = fldCur.FullPath
 ChDir Dir1Path
 If LesOptions.ShowPathInCaption = 1 Then
  Me.Caption = "THE Rename - " + Dir1Path
 End If
 vnb = 0
 état.Panels(4).Text = "0"
 vnb = remplissage()
 état.Panels(3).Text = Trim$(Str$(vnb))
 mundo.Enabled = False
 List2.Clear
 List3.Clear
End Sub

Private Sub FolderTreeview1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
 Dim tAlt As Integer
 tAlt = GetKeyState(VK_MENU)
 If (tAlt = -127 Or tAlt = -128) And KeyCode = 13 Then ' Afficher les propriétés
  mprop2_Click
 End If
 
 Select Case KeyCode
  Case 8
   TemMove = False
   mundo.Enabled = False
   List2.Clear
   List3.Clear
  Case 46
   mdelrep_Click
  Case 113
   mrendirect_Click
 End Select
End Sub

Private Sub FolderTreeview1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
  PopupMenu mcontext2
 End If
End Sub
Private Sub Form_Activate()
    If PbFtv1 = True Then ' On a demandé à ouvrir un fichier de settings sur la ligne de commande mais le FTV déconne
        PbFtv1 = False
        FolderTreeview1(0).SelectedFolder = Dir1Path
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
letat = False
If KeyCode = 27 Then   ' Esc
    Annuler = True
End If
End Sub
Private Sub Form_Load()
Dim leretour As Integer
Dim mvnb1 As Integer, mvnb2 As Integer, mvnb3 As Integer, vnb As Long, mvnb4 As Integer, mvnb5 As Integer
Dim tlibdate(3) As String
Dim i As Integer
Dim ctl As Control
On Error GoTo ErrGen
GlobalLoad = True
'AppPath = AddBackSlash(App.Path)
'OkUseAbbrev = False ' Par défaut on n'utilise pas les abréviations
'UseMP3 = False
'UseVQF = False
PbFtv1 = False
CurrentCommand = 0
'LoadTags ' Charge les tags pour les TIF
'UseCylcic = False
'VnbCyclic = 0
FavEncours = -1
acpreview = False
TemShift = False
aOuvrir = False
LeCancel = False
'OptionsCyclic = False
Piège1 = False

' MRU
Dim cR As New cRegistry
cR.ClassKey = HKEY_CURRENT_USER
cR.SectionKey = "Software\VB and VBA Program Settings\THERename"
m_cMRU.Load cR
m_cMRU.MaxFileCount = 5
pDisplayMRU False
' Fin MRU
'VnbHistory = 0
'RechPref = False
'RechSuff = False
'FolderOk = False
'Folder1 = 0
'Folder5 = "1"
'Folder2 = 0
'Folder3 = 0
'Folder6 = " "
'Folder4 = 0
tlibdate(1) = "Created"
tlibdate(2) = "Modified"
tlibdate(3) = "Access"
'VnbRep = 0
'TemMove = False
'recursive = False
m2recursive.Checked = False
'VancRep = ""
vnb = 0
ChargeVNBCommandes
lhistory.Clear
'LeCancel = False
'aOuvrir = False
'acpreview = False
vnboptionp = UBound(optionp)
vnboptions = UBound(options)
vnbcmd = UBound(hlplang)  ' Nombre de commandes du langage
'App.HelpFile = AppPath + "therename.hlp"
'TemDelete = False
EtchedLine RENAME, 0, Toolbar1.Top + Toolbar1.height + 50, RENAME.width

optionp(1) = "Upper Case"
aidep(1) = "File1.exe => FILE1.exe"
optionp(2) = "Lower Case"
aidep(2) = "FILE1.exe => file1.exe"
optionp(3) = "Toggle Case"
aidep(3) = "File1.exe => fILE1.exe"
optionp(4) = "Keep prefix"
aidep(4) = "File1.exe => File1.exe"
optionp(5) = "Capitalize all words"
aidep(5) = "file1 binary.exe => File1 Binary.exe"
optionp(6) = "Invert letters"
aidep(6) = "File1.exe => 1eliF.exe"
optionp(7) = "Replace with system date"
aidep(7) = "File1.exe => 19980515.exe"
optionp(8) = "Replace with system time"
aidep(8) = "File1.exe => 160258.exe"
optionp(9) = "Replace with date + time"
aidep(9) = "File1.exe => 19980515160258.exe"
optionp(10) = "Modify prefix"
aidep(10) = "File1.exe => change with your options"
optionp(11) = "Replace with file's content"
aidep(11) = "about.txt => THE Rename is a freeware.txt (take file's content to make name)"
optionp(12) = "Special truetypes, search internal name"
aidep(12) = "times.ttf => Times New Roman - Regular.ttf"
optionp(13) = "Free form"
aidep(13) = "Type commands yourself to rename files"
optionp(14) = "Replace long filename with short filename"
aidep(14) = "THERENAME.HLP => THEREN~1.HLP"
optionp(15) = "Remove internal spaces"
aidep(15) = "THE RENAME.HLP => THERENAME.HLP"
optionp(16) = "Rename from a list"
aidep(16) = "Program will uses a list to rename files"
optionp(17) = "Capitalize first word only"
aidep(17) = "file1 binary.exe => File1 binary.exe"
optionp(18) = "Rename prefix with a random name"
aidep(18) = "Prefix will be random"
optionp(19) = "CoWbOyS Option"
aidep(19) = "THE RENAME.HLP => ThErEnAmE.HLP"
optionp(20) = "Remove multiple spacing in name"
aidep(20) = "THE   RENAME.HLP => THE RENAME.HLP"
optionp(21) = "Separate Words"
aidep(21) = "TheRename.HLP => The Rename.HLP"

options(1) = "Upper Case"
aides(1) = "File1.exe => File1.EXE"
options(2) = "Lower Case"
aides(2) = "File1.EXE => File1.exe"
options(3) = "Toggle Case"
aides(3) = "File1.Exe => File1.eXE"
options(4) = "Keep extension"
aides(4) = "File1.Exe => File1.Exe"
options(5) = "Capitalize"
aides(5) = "File1.exe => File1.Exe"
options(6) = "Invert letters"
aides(6) = "File1.dat => File1.tad"
options(7) = "Replace with system date"
aides(7) = "File1.exe => File1.19980515.exe"
options(8) = "Replace with system time"
aides(8) = "File1.exe => File1.160258"
options(9) = "Replace with date + time"
aides(9) = "File1.exe => File1.19980515160258"
options(10) = "Modify extension"
aides(10) = "File1.exe => change with your options"
options(11) = "Delete extension"
aides(11) = "File1.exe => File1"
options(12) = "Remove internal spaces"
aides(12) = "THE RENAME.H L P => THE RENAME.HLP"
options(13) = "CoWbOyS Option"
aides(13) = "THE RENAME.HELP => THE RENAME.HeLp"
options(14) = "Remove multiple spacing in name"
aides(14) = "THE RENAME.Help   File => THE RENAME.Help File"
options(15) = "Separate Words"
aides(15) = "THE RENAME.HelpFile => THE RENAME.Help File"

Combo3.ListIndex = 0
Combo4.ListIndex = 0

langage(1) = "<curext>"
hlplang(1) = 352
langage(2) = "<curprefix>"
hlplang(2) = 317
langage(3) = "<ddddd>"
hlplang(3) = 463
langage(4) = "<EXCapital>"
hlplang(4) = 356
langage(5) = "<EXInvert>"
hlplang(5) = 361
langage(6) = "<EXLower>"
hlplang(6) = 365
langage(7) = "<EXToggle>"
hlplang(7) = 378
langage(8) = "<EXUpper>"
hlplang(8) = 382
langage(9) = "<FileContent,10>"
hlplang(9) = 386
langage(10) = "<FileDate>"
hlplang(10) = 387
langage(11) = "<FileTime>"
hlplang(11) = 391
langage(12) = "<xxxxx>"
hlplang(12) = 467
langage(13) = "<ooooo>"
hlplang(13) = 464
langage(14) = "<PRCapital>"
hlplang(14) = 322
langage(15) = "<PRInvert>"
hlplang(15) = 327
langage(16) = "<PRLower>"
hlplang(16) = 331
langage(17) = "<PRToggle>"
hlplang(17) = 344
langage(18) = "<PRUpper>"
hlplang(18) = 348
langage(19) = "<systdate>"
hlplang(19) = 468
langage(20) = "<systtime>"
hlplang(20) = 470
langage(21) = "<ttfname>"
hlplang(21) = 395
langage(22) = "<html>"
hlplang(22) = 393
langage(23) = "<PRrtrim>"
hlplang(23) = 341
langage(24) = "<PRltrim>"
hlplang(24) = 333
langage(25) = "<PRtrim>"
hlplang(25) = 347
langage(26) = "<EXrtrim>"
hlplang(26) = 375
langage(27) = "<EXltrim>"
hlplang(27) = 367
langage(28) = "<EXtrim>"
hlplang(28) = 381
langage(29) = "<ShortName>"
hlplang(29) = 351
langage(30) = "<PRRemIntSp>"
hlplang(30) = 338
langage(31) = "<EXRemIntSp>"
hlplang(31) = 372
langage(32) = "<copyfile,0>"
hlplang(32) = 384
langage(33) = "<EXLeft,0>"
hlplang(33) = 363
langage(34) = "<EXRight,0>"
hlplang(34) = 374
langage(35) = "<PRLeft,0>"
hlplang(35) = 329
langage(36) = "<PRRight,0>"
hlplang(36) = 340
langage(37) = "<EXMiddle,0,0>"
hlplang(37) = 368
langage(38) = "<PRMiddle,0,0>"
hlplang(38) = 334
langage(39) = "<EXPaddLeft,0,0>"
hlplang(39) = 369
langage(40) = "<EXPaddRight,0,0>"
hlplang(40) = 370
langage(41) = "<PRPaddLeft,0,0>"
hlplang(41) = 335
langage(42) = "<PRPaddRight,0,0>"
hlplang(42) = 336
langage(43) = "<EXToken,0,delim>"
hlplang(43) = 380
langage(44) = "<PRToken,0,delim>"
hlplang(44) = 346
langage(45) = "<EXCapitalEX,0,0>"
hlplang(45) = 357
langage(46) = "<EXInvertEX,0,0>"
hlplang(46) = 362
langage(47) = "<EXLowerEX,0,0>"
hlplang(47) = 366
langage(48) = "<EXToggleEX,0,0>"
hlplang(48) = 379
langage(49) = "<EXUpperEX,0,0>"
hlplang(49) = 383
langage(50) = "<PRCapitalEX,0,0>"
hlplang(50) = 323
langage(51) = "<PRInvertEX,0,0>"
hlplang(51) = 328
langage(52) = "<PRLowerEX,0,0>"
hlplang(52) = 332
langage(53) = "<PRToggleEX,0,0>"
hlplang(53) = 345
langage(54) = "<PRUpperEX,0,0>"
hlplang(54) = 349
langage(55) = "<FileSize,0>"
hlplang(55) = 390
langage(56) = "<FileAttr>"
hlplang(56) = 385
langage(57) = "<ImgInfo>"
hlplang(57) = 394
langage(58) = "<CyclicSelection>"
hlplang(58) = 396
langage(59) = "<PRbeforeEx>"
hlplang(59) = 538
langage(60) = "<PRAfter,0>"
hlplang(60) = 318
langage(61) = "<PRBetween,0,0>"
hlplang(61) = 320
langage(62) = "<EXBefore,0>"
hlplang(62) = 354
langage(63) = "<EXAfter,0>"
hlplang(63) = 353
langage(64) = "<EXBetween,0,0>"
hlplang(64) = 355
langage(65) = "<TextCounter>"
hlplang(65) = 466
langage(66) = "<RandomPrefix>"
hlplang(66) = 350
langage(67) = "<EXCapitalFirst>"
hlplang(67) = 358
langage(68) = "<PRCapitalFirst>"
hlplang(68) = 324
langage(69) = "<PRCowboys>"
hlplang(69) = 325
langage(70) = "<EXCowboys>"
hlplang(70) = 359
langage(71) = "<PRRemMultSp>"
hlplang(71) = 339
langage(72) = "<EXRemMultSp>"
hlplang(72) = 373
langage(73) = "<MP3Title>"
hlplang(73) = 439
langage(74) = "<MP3Artist>"
hlplang(74) = 398
langage(75) = "<MP3Album>"
hlplang(75) = 306
langage(76) = "<MP3Year>"
hlplang(76) = 453
langage(77) = "<MP3Comment>"
hlplang(77) = 401
langage(78) = "<MP3Genre>"
hlplang(78) = 411
langage(79) = "<MP3Band>"
hlplang(79) = 399
langage(80) = "<MP3Bpm>"
hlplang(80) = 400
langage(81) = "<MP3Composer>"
hlplang(81) = 402
langage(82) = "<MP3Conductor>"
hlplang(82) = 403
langage(83) = "<MP3ContentGroup>"
hlplang(83) = 404
langage(84) = "<MP3Copyright>"
hlplang(84) = 405
langage(85) = "<PRIfEmpty,0>"
hlplang(85) = 326
langage(86) = "<ExIfEmpty,0>"
hlplang(86) = 360
langage(87) = "<PRLength>"
hlplang(87) = 330
langage(88) = "<EXLength>"
hlplang(88) = 364
langage(89) = "<FileLength>"
hlplang(89) = 389
langage(90) = "<FullLength>"
hlplang(90) = 392
langage(91) = "<PRSepWords>"
hlplang(91) = 343
langage(92) = "<EXSepWords>"
hlplang(92) = 377
langage(93) = "<VQFArtist>"
hlplang(93) = 454
langage(94) = "<VQFBitrate>"
hlplang(94) = 455
langage(95) = "<VQFComment>"
hlplang(95) = 456
langage(96) = "<VQFCopyright>"
hlplang(96) = 457
langage(97) = "<VQFFileSaveAs>"
hlplang(97) = 458
langage(98) = "<VQFMonoStereo>"
hlplang(98) = 459
langage(99) = "<VQFQuality>"
hlplang(99) = 460
langage(100) = "<VQFSampleRate>"
hlplang(100) = 461
langage(101) = "<VQFTitle>"
hlplang(101) = 462
langage(102) = "<MP3EncryptionMethod>"
hlplang(102) = 408
langage(103) = "<MP3Date>"
hlplang(103) = 406
langage(104) = "<MP3EncodedBy>"
hlplang(104) = 407
langage(105) = "<MP3SoftwareEncodingSettings>"
hlplang(105) = 433
langage(106) = "<MP3FileOwner>"
hlplang(106) = 409
langage(107) = "<MP3FileType>"
hlplang(107) = 410
langage(108) = "<MP3GroupIdent>"
hlplang(108) = 412
langage(109) = "<MP3InitialKey>"
hlplang(109) = 413
langage(110) = "<MP3InvolvedPeopleList>"
hlplang(110) = 414
langage(111) = "<MP3Isrc>"
hlplang(111) = 415
langage(112) = "<MP3Language>"
hlplang(112) = 416
langage(113) = "<MP3LinkedInformation>"
hlplang(113) = 417
langage(114) = "<MP3Lyricist>"
hlplang(114) = 418
langage(115) = "<MP3MediaType>"
hlplang(115) = 419
langage(116) = "<MP3MixArtist>"
hlplang(116) = 420
langage(117) = "<MP3NetRadioOwner>"
hlplang(117) = 421
langage(118) = "<MP3NetRadioStation>"
hlplang(118) = 422
langage(119) = "<MP3OriginalAlbum>"
hlplang(119) = 423
langage(120) = "<MP3OriginalArtist>"
hlplang(120) = 424
langage(121) = "<MP3OriginalFilename>"
hlplang(121) = 425
langage(122) = "<MP3OriginalLyricist>"
hlplang(122) = 426
langage(123) = "<MP3OriginalYear>"
hlplang(123) = 427
langage(124) = "<MP3PartOfASet>"
hlplang(124) = 428
langage(125) = "<MP3PlayListDelay>"
hlplang(125) = 429
langage(126) = "<RandomNumber,0,100,4>"
hlplang(126) = 465
langage(127) = "<SystDateEx,expr>"
hlplang(127) = 469
langage(128) = "<PathPart,1>"
hlplang(128) = 397
langage(129) = "<FileDateEx,1,dd>"
hlplang(129) = 388
langage(130) = "<PRRefomartNumber,4,0>"
hlplang(130) = 337
langage(131) = "<ExRefomartNumber,4,0>"
hlplang(131) = 371
langage(132) = "<PRSepThousands,\w>"
hlplang(132) = 342
langage(133) = "<ExSepThousands,\w>"
hlplang(133) = 376
langage(134) = "<MP3PopulariMeter>"
hlplang(134) = 430
langage(135) = "<MP3Publisher>"
hlplang(135) = 431
langage(136) = "<MP3RecordingDates>"
hlplang(136) = 432
langage(137) = "<MP3SongLength>"
hlplang(137) = 434
langage(138) = "<MP3SubTitle>"
hlplang(138) = 435
langage(139) = "<MP3SynchronizedLyric>"
hlplang(139) = 436
langage(140) = "<MP3TermsOfUse>"
hlplang(140) = 437
langage(141) = "<MP3Time>"
hlplang(141) = 438
langage(142) = "<MP3TrackNumber>"
hlplang(142) = 441
langage(143) = "<MP3TotalTracks>"
hlplang(143) = 440
langage(144) = "<MP3UnsynchronizedLyric>"
hlplang(144) = 442
langage(145) = "<MP3UserText>"
hlplang(145) = 443
langage(146) = "<MP3wwwArtist>"
hlplang(146) = 444
langage(147) = "<MP3wwwAudioFile>"
hlplang(147) = 445
langage(148) = "<MP3wwwAudioSource>"
hlplang(148) = 446
langage(149) = "<MP3wwwCommercialInfo>"
hlplang(149) = 447
langage(150) = "<MP3wwwCopyright>"
hlplang(150) = 448
langage(151) = "<MP3wwwPayment>"
hlplang(151) = 449
langage(152) = "<MP3wwwPublisher>"
hlplang(152) = 450
langage(153) = "<MP3wwwRadioPage>"
hlplang(153) = 451
langage(154) = "<MP3wwwUserURL>"
hlplang(154) = 452
langage(155) = "<CounterEx,0,1,1,4,2>"
hlplang(155) = 475
langage(156) = "<PRPrevFileOldName>"
hlplang(156) = 481
langage(157) = "<PRPrevFileNewName>"
hlplang(157) = 482
langage(158) = "<EXPrevFileOldName>"
hlplang(158) = 483
langage(159) = "<EXPrevFileNewName>"
hlplang(159) = 484
langage(160) = "<PRModifyCounter,1,4,0>"
hlplang(160) = 485
langage(161) = "<EXModifyCounter,1,4,0>"
hlplang(161) = 486
langage(162) = "<OggNumberOfTags>"
hlplang(162) = 512
langage(163) = "<OggSerialNumber>"
hlplang(163) = 517
langage(164) = "<OggEncoderVersion>"
hlplang(164) = 502
langage(165) = "<OggLowerBitrate>"
hlplang(165) = 510
langage(166) = "<OggUpperBitrate>"
hlplang(166) = 523
langage(167) = "<OggNominalBitrate>"
hlplang(167) = 511
langage(168) = "<OggAverageBitrate>"
hlplang(168) = 494
langage(169) = "<OggChannels>"
hlplang(169) = 495
langage(170) = "<OggSampleRate>"
hlplang(170) = 516
langage(171) = "<OggVendor>"
hlplang(171) = 524
langage(172) = "<OggPlaytime>"
hlplang(172) = 515
langage(173) = "<OggLength>"
hlplang(173) = 508
langage(174) = "<OggISRC>"
hlplang(174) = 507
langage(175) = "<OggDate>"
hlplang(175) = 500
langage(176) = "<OggCopyRight>"
hlplang(176) = 499
langage(177) = "<OggLocation>"
hlplang(177) = 509
langage(178) = "<OggDescription>"
hlplang(178) = 501
langage(179) = "<OggOrganization>"
hlplang(179) = 513
langage(180) = "<OggTotalTracks>"
hlplang(180) = 521
langage(181) = "<OggTrackNumber>"
hlplang(181) = 522
langage(182) = "<OggVersion>"
hlplang(182) = 525
langage(183) = "<OggComment>"
hlplang(183) = 496
langage(184) = "<OggAlbum>"
hlplang(184) = 491
langage(185) = "<OggGenre>"
hlplang(185) = 504
langage(186) = "<OggArtist>"
hlplang(186) = 493
langage(187) = "<OggTitle>"
hlplang(187) = 520
langage(188) = "<OggComposer>"
hlplang(188) = 497
langage(189) = "<OggConductor>"
hlplang(189) = 498
langage(190) = "<OggEnsemble>"
hlplang(190) = 503
langage(191) = "<OggPerformer>"
hlplang(191) = 514
langage(192) = "<OggGetUnknowTags,1,=>"
hlplang(192) = 506
langage(193) = "<OggGetAllTags,1,=>"
hlplang(193) = 505
langage(194) = "<OggTagByName,Artist,1,=>"
hlplang(194) = 518
langage(195) = "<OggTagByPosition,1,1,=>"
hlplang(195) = 519
langage(196) = "<PRInsert,1,text>"
hlplang(196) = 526
langage(197) = "<EXInsert,1,text>"
hlplang(197) = 527
langage(198) = "<PRRemove,1,2>"
hlplang(198) = 528
langage(199) = "<EXRemove,1,2>"
hlplang(199) = 529
langage(200) = "<PRDeleteText,Text,1,-1,Yes>"
hlplang(200) = 530
langage(201) = "<EXDeleteText,Text,1,-1,Yes>"
hlplang(201) = 531
langage(202) = "<SelectedFilesCount>"
hlplang(202) = 533
langage(203) = "<NonSelectedFilesCount>"
hlplang(203) = 534
langage(204) = "<TotalFilesCount>"
hlplang(204) = 535
langage(205) = "<RomanCounter>"
hlplang(205) = 536
langage(206) = "<AfmFontMetricsVersion>"
hlplang(206) = 4
langage(207) = "<AfmWeight>"
hlplang(207) = 5
langage(208) = "<AfmNotice>"
hlplang(208) = 6
langage(209) = "<AfmMetricsSets>"
hlplang(209) = 7
langage(210) = "<AfmFullName>"
hlplang(210) = 12
langage(211) = "<AfmFontVersion>"
hlplang(211) = 27
langage(212) = "<AfmFontName>"
hlplang(212) = 28
langage(213) = "<AfmFamillyName>"
hlplang(213) = 29
langage(214) = "<AfmEncodingScheme>"
hlplang(214) = 30
langage(215) = "<AfmCharacterSet>"
hlplang(215) = 31
langage(216) = "<AfmCopyright>"
hlplang(216) = 70
langage(217) = "<AfmCreationDate>"
hlplang(217) = 71
langage(218) = "<AfmUniqueID>"
hlplang(218) = 72
langage(219) = "<AfmVMusage>"
hlplang(219) = 73
langage(220) = "<ExifAperture>"
hlplang(220) = 81
langage(221) = "<ExifBrightness>"
hlplang(221) = 82
langage(222) = "<ExifCompressedBitsPerPixel>"
hlplang(222) = 83
langage(223) = "<ExifCopyright>"
hlplang(223) = 84
langage(224) = "<ExifDateTime>"
hlplang(224) = 85
langage(225) = "<ExifDateTimeDigitized>"
hlplang(225) = 86
langage(226) = "<ExifDateTimeOriginal>"
hlplang(226) = 87
langage(227) = "<ExifVersion>"
hlplang(227) = 88
langage(228) = "<ExifExposureBias>"
hlplang(228) = 89
langage(229) = "<ExifExposureProgram>"
hlplang(229) = 90
langage(230) = "<ExifExposureTime>"
hlplang(230) = 91
langage(231) = "<ExifFirmwareVersion>"
hlplang(231) = 92
langage(232) = "<ExifFlash>"
hlplang(232) = 93
langage(233) = "<ExifFNumber>"
hlplang(233) = 94
langage(234) = "<ExifFocalLength>"
hlplang(234) = 95
langage(235) = "<ExifFocalPlaneResolutionUnit>"
hlplang(235) = 96
langage(236) = "<ExifFocalPlaneXResolution>"
hlplang(236) = 97
langage(237) = "<ExifFocalPlaneYResolution>"
hlplang(237) = 98
langage(238) = "<ExifImageDescription>"
hlplang(238) = 99
langage(239) = "<ExifImageHeight>"
hlplang(239) = 100
langage(240) = "<ExifImageWidth>"
hlplang(240) = 101
langage(241) = "<ExifISOSpeedRatings>"
hlplang(241) = 102
langage(242) = "<ExifMake>"
hlplang(242) = 112
langage(243) = "<ExifMaxAperture>"
hlplang(243) = 145
langage(244) = "<ExifMeteringMode>"
hlplang(244) = 155
langage(245) = "<ExifModel>"
hlplang(245) = 160
langage(246) = "<ExifOrientation>"
hlplang(246) = 163
langage(247) = "<ExifRelatedSoundFile>"
hlplang(247) = 164
langage(248) = "<ExifResolutionUnit>"
hlplang(248) = 242
langage(249) = "<ExifShutterSpeed>"
hlplang(249) = 243
langage(250) = "<ExifSubjectDistance>"
hlplang(250) = 252
langage(251) = "<ExifWhiteBalance>"
hlplang(251) = 283
langage(252) = "<ExifXResolution>"
hlplang(252) = 321
langage(253) = "<ExifYResolution>"
hlplang(253) = 478
langage(254) = "<PRBefore,0>"
hlplang(254) = 319
langage(255) = "<WmaChannelMode>"
hlplang(255) = 539
langage(256) = "<WmaSampleRate>"
hlplang(256) = 540
langage(257) = "<WmaDuration>"
hlplang(257) = 541
langage(258) = "<WmaBitRate>"
hlplang(258) = 542
langage(259) = "<WmaTrack>"
hlplang(259) = 543
langage(260) = "<WmaTitle>"
hlplang(260) = 544
langage(261) = "<WmaArtist>"
hlplang(261) = 545
langage(262) = "<WmaAlbum>"
hlplang(262) = 546
langage(263) = "<WmaGenre>"
hlplang(263) = 547
langage(264) = "<WmaComment>"
hlplang(264) = 548
langage(265) = "<WmaYear>"
hlplang(265) = 549

LngCmd(1, 1) = 0
LngCmd(2, 1) = 0
LngCmd(3, 1) = 0
LngCmd(4, 1) = 0
LngCmd(5, 1) = 0
LngCmd(6, 1) = 0
LngCmd(7, 1) = 0
LngCmd(8, 1) = 0
LngCmd(9, 1) = 12
LngCmd(10, 1) = 0
LngCmd(11, 1) = 0
LngCmd(12, 1) = 0
LngCmd(13, 1) = 0
LngCmd(14, 1) = 0
LngCmd(15, 1) = 0
LngCmd(16, 1) = 0
LngCmd(17, 1) = 0
LngCmd(18, 1) = 0
LngCmd(19, 1) = 0
LngCmd(20, 1) = 0
LngCmd(21, 1) = 0
LngCmd(22, 1) = 0
LngCmd(23, 1) = 0
LngCmd(24, 1) = 0
LngCmd(25, 1) = 0
LngCmd(26, 1) = 0
LngCmd(27, 1) = 0
LngCmd(28, 1) = 0
LngCmd(29, 1) = 0
LngCmd(30, 1) = 0
LngCmd(31, 1) = 0
LngCmd(1, 2) = 0
LngCmd(2, 2) = 0
LngCmd(3, 2) = 0
LngCmd(4, 2) = 0
LngCmd(5, 2) = 0
LngCmd(6, 2) = 0
LngCmd(7, 2) = 0
LngCmd(8, 2) = 0
LngCmd(9, 2) = 1
LngCmd(10, 2) = 0
LngCmd(11, 2) = 0
LngCmd(12, 2) = 0
LngCmd(13, 2) = 0
LngCmd(14, 2) = 0
LngCmd(15, 2) = 0
LngCmd(16, 2) = 0
LngCmd(17, 2) = 0
LngCmd(18, 2) = 0
LngCmd(19, 2) = 0
LngCmd(20, 2) = 0
LngCmd(21, 2) = 0
LngCmd(22, 2) = 0
LngCmd(23, 2) = 0
LngCmd(24, 2) = 0
LngCmd(25, 2) = 0
LngCmd(26, 2) = 0
LngCmd(27, 2) = 0
LngCmd(28, 2) = 0
LngCmd(29, 2) = 0
LngCmd(30, 2) = 0
LngCmd(31, 2) = 0
LngCmd(32, 1) = 9
LngCmd(32, 2) = 1
LngCmd(33, 1) = 7
LngCmd(33, 2) = 1
LngCmd(34, 1) = 8
LngCmd(34, 2) = 1
LngCmd(35, 1) = 7
LngCmd(35, 2) = 1
LngCmd(36, 1) = 8
LngCmd(36, 2) = 1
LngCmd(37, 1) = 9
LngCmd(37, 2) = 2
LngCmd(38, 1) = 9
LngCmd(38, 2) = 2
LngCmd(39, 1) = 11
LngCmd(39, 2) = 2
LngCmd(40, 1) = 12
LngCmd(40, 2) = 2
LngCmd(41, 1) = 11
LngCmd(41, 2) = 2
LngCmd(42, 1) = 12
LngCmd(42, 2) = 2
LngCmd(43, 1) = 8
LngCmd(43, 2) = 2
LngCmd(44, 1) = 8
LngCmd(44, 2) = 2
LngCmd(45, 1) = 12
LngCmd(45, 2) = 2
LngCmd(46, 1) = 11
LngCmd(46, 2) = 2
LngCmd(47, 1) = 10
LngCmd(47, 2) = 2
LngCmd(48, 1) = 11
LngCmd(48, 2) = 2
LngCmd(49, 1) = 10
LngCmd(49, 2) = 2
LngCmd(50, 1) = 12
LngCmd(50, 2) = 2
LngCmd(51, 1) = 11
LngCmd(51, 2) = 2
LngCmd(52, 1) = 10
LngCmd(52, 2) = 2
LngCmd(53, 1) = 11
LngCmd(53, 2) = 2
LngCmd(54, 1) = 10
LngCmd(54, 2) = 2
LngCmd(55, 1) = 9
LngCmd(55, 2) = 1
LngCmd(56, 1) = 0
LngCmd(56, 2) = 0
LngCmd(57, 1) = 0
LngCmd(57, 2) = 0
LngCmd(58, 1) = 0
LngCmd(58, 2) = 0
LngCmd(59, 1) = 11
LngCmd(59, 2) = -10
LngCmd(60, 1) = 8
LngCmd(60, 2) = 1
LngCmd(61, 1) = 10
LngCmd(61, 2) = 2
LngCmd(62, 1) = 9
LngCmd(62, 2) = 1
LngCmd(63, 1) = 8
LngCmd(63, 2) = 1
LngCmd(64, 1) = 10
LngCmd(64, 2) = 2
LngCmd(65, 1) = 0
LngCmd(65, 2) = 0
LngCmd(66, 1) = 0
LngCmd(66, 2) = 0
LngCmd(67, 1) = 0
LngCmd(67, 2) = 0
LngCmd(68, 1) = 0
LngCmd(68, 2) = 0
LngCmd(69, 1) = 0
LngCmd(69, 2) = 0
LngCmd(70, 1) = 0
LngCmd(70, 2) = 0
LngCmd(71, 1) = 0
LngCmd(71, 2) = 0
LngCmd(72, 1) = 0
LngCmd(72, 2) = 0
LngCmd(73, 1) = 9
LngCmd(73, 2) = -2
LngCmd(74, 1) = 10
LngCmd(74, 2) = -2
LngCmd(75, 1) = 9
LngCmd(75, 2) = -2
LngCmd(76, 1) = 8
LngCmd(76, 2) = -2
LngCmd(77, 1) = 11
LngCmd(77, 2) = -2
LngCmd(78, 1) = 9
LngCmd(78, 2) = -2
LngCmd(79, 1) = 8
LngCmd(79, 2) = -2
LngCmd(80, 1) = 7
LngCmd(80, 2) = -2
LngCmd(81, 1) = 12
LngCmd(81, 2) = -2
LngCmd(82, 1) = 13
LngCmd(82, 2) = -2
LngCmd(83, 1) = 16
LngCmd(83, 2) = -2
LngCmd(84, 1) = 13
LngCmd(84, 2) = -2
LngCmd(85, 1) = 10
LngCmd(85, 2) = 1
LngCmd(86, 1) = 10
LngCmd(86, 2) = 1
LngCmd(87, 1) = 0
LngCmd(87, 2) = 0
LngCmd(88, 1) = 0
LngCmd(88, 2) = 0
LngCmd(89, 1) = 0
LngCmd(89, 2) = 0
LngCmd(90, 1) = 0
LngCmd(90, 2) = 0
LngCmd(91, 1) = 0
LngCmd(91, 2) = 0
LngCmd(92, 1) = 0
LngCmd(92, 2) = 0
LngCmd(93, 1) = 10
LngCmd(93, 2) = -2
LngCmd(94, 1) = 11
LngCmd(94, 2) = -2
LngCmd(95, 1) = 11
LngCmd(95, 2) = -2
LngCmd(96, 1) = 13
LngCmd(96, 2) = -2
LngCmd(97, 1) = 14
LngCmd(97, 2) = -2
LngCmd(98, 1) = 14
LngCmd(98, 2) = -2
LngCmd(99, 1) = 11
LngCmd(99, 2) = -2
LngCmd(100, 1) = 14
LngCmd(100, 2) = -2
LngCmd(101, 1) = 9
LngCmd(101, 2) = -2
LngCmd(102, 1) = 20
LngCmd(102, 2) = -2
LngCmd(103, 1) = 8
LngCmd(103, 2) = -2
LngCmd(104, 1) = 13
LngCmd(104, 2) = -2
LngCmd(105, 1) = 28
LngCmd(105, 2) = -2
LngCmd(106, 1) = 13
LngCmd(106, 2) = -2
LngCmd(107, 1) = 12
LngCmd(107, 2) = -2
LngCmd(108, 1) = 14
LngCmd(108, 2) = -2
LngCmd(109, 1) = 14
LngCmd(109, 2) = -2
LngCmd(110, 1) = 22
LngCmd(110, 2) = -2
LngCmd(111, 1) = 8
LngCmd(111, 2) = -2
LngCmd(112, 1) = 12
LngCmd(112, 2) = -2
LngCmd(113, 1) = 21
LngCmd(113, 2) = -2
LngCmd(114, 1) = 12
LngCmd(114, 2) = -2
LngCmd(115, 1) = 13
LngCmd(115, 2) = -2
LngCmd(116, 1) = 13
LngCmd(116, 2) = -2
LngCmd(117, 1) = 17
LngCmd(117, 2) = -2
LngCmd(118, 1) = 19
LngCmd(118, 2) = -2
LngCmd(119, 1) = 17
LngCmd(119, 2) = -2
LngCmd(120, 1) = 18
LngCmd(120, 2) = -2
LngCmd(121, 1) = 20
LngCmd(121, 2) = -2
LngCmd(122, 1) = 20
LngCmd(122, 2) = -2
LngCmd(123, 1) = 16
LngCmd(123, 2) = -2
LngCmd(124, 1) = 14
LngCmd(124, 2) = -2
LngCmd(125, 1) = 14
LngCmd(125, 2) = -2
LngCmd(126, 1) = 13
LngCmd(126, 2) = 3
LngCmd(127, 1) = 11
LngCmd(127, 2) = 1
LngCmd(128, 1) = 9
LngCmd(128, 2) = 1
LngCmd(129, 1) = 11
LngCmd(129, 2) = 2
LngCmd(130, 1) = 17
LngCmd(130, 2) = 2
LngCmd(131, 1) = 17
LngCmd(131, 2) = 2
LngCmd(132, 1) = 15
LngCmd(132, 2) = 1
LngCmd(133, 1) = 15
LngCmd(133, 2) = 1
LngCmd(134, 1) = 17
LngCmd(134, 2) = -2
LngCmd(135, 1) = 13
LngCmd(135, 2) = -2
LngCmd(136, 1) = 18
LngCmd(136, 2) = -2
LngCmd(137, 1) = 14
LngCmd(137, 2) = -2
LngCmd(138, 1) = 12
LngCmd(138, 2) = -2
LngCmd(139, 1) = 21
LngCmd(139, 2) = -2
LngCmd(140, 1) = 14
LngCmd(140, 2) = -2
LngCmd(141, 1) = 8
LngCmd(141, 2) = -2
LngCmd(142, 1) = 15
LngCmd(142, 2) = -2
LngCmd(143, 1) = 15
LngCmd(143, 2) = -2
LngCmd(144, 1) = 23
LngCmd(144, 2) = -2
LngCmd(145, 1) = 12
LngCmd(145, 2) = -2
LngCmd(146, 1) = 13
LngCmd(146, 2) = -2
LngCmd(147, 1) = 16
LngCmd(147, 2) = -2
LngCmd(148, 1) = 18
LngCmd(148, 2) = -2
LngCmd(149, 1) = 21
LngCmd(149, 2) = -2
LngCmd(150, 1) = 16
LngCmd(150, 2) = -2
LngCmd(151, 1) = 14
LngCmd(151, 2) = -2
LngCmd(152, 1) = 16
LngCmd(152, 2) = -2
LngCmd(153, 1) = 16
LngCmd(153, 2) = -2
LngCmd(154, 1) = 14
LngCmd(154, 2) = -2
LngCmd(155, 1) = 10
LngCmd(155, 2) = 5
LngCmd(156, 1) = 0
LngCmd(156, 2) = 0
LngCmd(157, 1) = 0
LngCmd(157, 2) = 0
LngCmd(158, 1) = 0
LngCmd(158, 2) = 0
LngCmd(159, 1) = 0
LngCmd(159, 2) = 0
LngCmd(160, 1) = 16
LngCmd(160, 2) = 3
LngCmd(161, 1) = 16
LngCmd(161, 2) = 3
LngCmd(162, 1) = 0
LngCmd(162, 2) = 0
LngCmd(163, 1) = 16
LngCmd(163, 2) = -2
LngCmd(164, 1) = 18
LngCmd(164, 2) = -2
LngCmd(165, 1) = 16
LngCmd(165, 2) = -2
LngCmd(166, 1) = 16
LngCmd(166, 2) = -2
LngCmd(167, 1) = 18
LngCmd(167, 2) = -2
LngCmd(168, 1) = 18
LngCmd(168, 2) = -2
LngCmd(169, 1) = 12
LngCmd(169, 2) = -2
LngCmd(170, 1) = 14
LngCmd(170, 2) = -2
LngCmd(171, 1) = 10
LngCmd(171, 2) = -2
LngCmd(172, 1) = 12
LngCmd(172, 2) = -2
LngCmd(173, 1) = 10
LngCmd(173, 2) = -2
LngCmd(174, 1) = 8
LngCmd(174, 2) = -2
LngCmd(175, 1) = 8
LngCmd(175, 2) = -2
LngCmd(176, 1) = 13
LngCmd(176, 2) = -2
LngCmd(177, 1) = 12
LngCmd(177, 2) = -2
LngCmd(178, 1) = 15
LngCmd(178, 2) = -2
LngCmd(179, 1) = 16
LngCmd(179, 2) = -2
LngCmd(180, 1) = 15
LngCmd(180, 2) = -2
LngCmd(181, 1) = 15
LngCmd(181, 2) = -2
LngCmd(182, 1) = 11
LngCmd(182, 2) = -2
LngCmd(183, 1) = 11
LngCmd(183, 2) = -2
LngCmd(184, 1) = 9
LngCmd(184, 2) = -2
LngCmd(185, 1) = 9
LngCmd(185, 2) = -2
LngCmd(186, 1) = 10
LngCmd(186, 2) = -2
LngCmd(187, 1) = 9
LngCmd(187, 2) = -2
LngCmd(188, 1) = 12
LngCmd(188, 2) = -2
LngCmd(189, 1) = 13
LngCmd(189, 2) = -2
LngCmd(190, 1) = 12
LngCmd(190, 2) = -2
LngCmd(191, 1) = 13
LngCmd(191, 2) = -2
LngCmd(192, 1) = 17
LngCmd(192, 2) = 2
LngCmd(193, 1) = 14
LngCmd(193, 2) = 2
LngCmd(194, 1) = 13
LngCmd(194, 2) = -5
LngCmd(195, 1) = 17
LngCmd(195, 2) = -5
LngCmd(196, 1) = 9
LngCmd(196, 2) = 2
LngCmd(197, 1) = 9
LngCmd(197, 2) = 2
LngCmd(198, 1) = 9
LngCmd(198, 2) = 2
LngCmd(199, 1) = 9
LngCmd(199, 2) = 2
LngCmd(200, 1) = 13
LngCmd(200, 2) = 4
LngCmd(201, 1) = 13
LngCmd(201, 2) = 4
LngCmd(202, 1) = 0
LngCmd(202, 2) = 0
LngCmd(203, 1) = 0
LngCmd(203, 2) = 0
LngCmd(204, 1) = 0
LngCmd(204, 2) = 0
LngCmd(205, 1) = 0
LngCmd(205, 2) = 0
LngCmd(206, 1) = 22
LngCmd(206, 2) = -2
LngCmd(207, 1) = 10
LngCmd(207, 2) = -2
LngCmd(208, 1) = 10
LngCmd(208, 2) = -2
LngCmd(209, 1) = 15
LngCmd(209, 2) = -2
LngCmd(210, 1) = 12
LngCmd(210, 2) = -2
LngCmd(211, 1) = 15
LngCmd(211, 2) = -2
LngCmd(212, 1) = 12
LngCmd(212, 2) = -2
LngCmd(213, 1) = 14
LngCmd(213, 2) = -2
LngCmd(214, 1) = 18
LngCmd(214, 2) = -2
LngCmd(215, 1) = 16
LngCmd(215, 2) = -2
LngCmd(216, 1) = 13
LngCmd(216, 2) = -2
LngCmd(217, 1) = 16
LngCmd(217, 2) = -2
LngCmd(218, 1) = 12
LngCmd(218, 2) = -2
LngCmd(219, 1) = 11
LngCmd(219, 2) = -2
LngCmd(220, 1) = 13 ' <ExifAperture>
LngCmd(220, 2) = -2
LngCmd(221, 1) = 15 ' <ExifBrightness>
LngCmd(221, 2) = -2
LngCmd(222, 1) = 27 ' <ExifCompressedBitsPerPixel>
LngCmd(222, 2) = -2
LngCmd(223, 1) = 14 ' <ExifCopyright>
LngCmd(223, 2) = -2
LngCmd(224, 1) = 13 ' <ExifDateTime>
LngCmd(224, 2) = -2
LngCmd(225, 1) = 22 ' <ExifDateTimeDigitized>
LngCmd(225, 2) = -2
LngCmd(226, 1) = 21 ' <ExifDateTimeOriginal>
LngCmd(226, 2) = -2
LngCmd(227, 1) = 12 ' <ExifVersion>
LngCmd(227, 2) = -2
LngCmd(228, 1) = 17 ' <ExifExposureBias>
LngCmd(228, 2) = -2
LngCmd(229, 1) = 20 ' <ExifExposureProgram>
LngCmd(229, 2) = -2
LngCmd(230, 1) = 17 ' <ExifExposureTime>
LngCmd(230, 2) = -2
LngCmd(231, 1) = 20 ' <ExifFirmwareVersion>
LngCmd(231, 2) = -2
LngCmd(232, 1) = 10 ' <ExifFlash>
LngCmd(232, 2) = -2
LngCmd(233, 1) = 12 ' <ExifFNumber>
LngCmd(233, 2) = -2
LngCmd(234, 1) = 16 ' <ExifFocalLength>
LngCmd(234, 2) = -2
LngCmd(235, 1) = 29 ' <ExifFocalPlaneResolutionUnit>
LngCmd(235, 2) = -2
LngCmd(236, 1) = 26 ' <ExifFocalPlaneXResolution>
LngCmd(236, 2) = -2
LngCmd(237, 1) = 26 ' <ExifFocalPlaneYResolution>
LngCmd(237, 2) = -2
LngCmd(238, 1) = 21 ' <ExifImageDescription>
LngCmd(238, 2) = -2
LngCmd(239, 1) = 16 ' <ExifImageHeight>
LngCmd(239, 2) = -2
LngCmd(240, 1) = 15 ' <ExifImageWidth>
LngCmd(240, 2) = -2
LngCmd(241, 1) = 20 ' <ExifISOSpeedRatings>
LngCmd(241, 2) = -2
LngCmd(242, 1) = 9  ' <ExifMake>
LngCmd(242, 2) = -2
LngCmd(243, 1) = 16 ' <ExifMaxAperture>
LngCmd(243, 2) = -2
LngCmd(244, 1) = 17 ' <ExifMeteringMode>
LngCmd(244, 2) = -2
LngCmd(245, 1) = 10 ' <ExifModel>
LngCmd(245, 2) = -2
LngCmd(246, 1) = 16 ' <ExifOrientation>
LngCmd(246, 2) = -2
LngCmd(247, 1) = 21 ' <ExifRelatedSoundFile>
LngCmd(247, 2) = -2
LngCmd(248, 1) = 19 ' <ExifResolutionUnit>
LngCmd(248, 2) = -2
LngCmd(249, 1) = 17 ' <ExifShutterSpeed>
LngCmd(249, 2) = -2
LngCmd(250, 1) = 20 ' <ExifSubjectDistance>
LngCmd(250, 2) = -2
LngCmd(251, 1) = 17 ' <ExifWhiteBalance>
LngCmd(251, 2) = -2
LngCmd(252, 1) = 16 ' <ExifXResolution>
LngCmd(252, 2) = -2
LngCmd(253, 1) = 16 ' <ExifYResolution>
LngCmd(253, 2) = -2
LngCmd(254, 1) = 9 ' <PRbefore>
LngCmd(254, 2) = 1
LngCmd(255, 1) = 15 ' <WmaChannelMode>
LngCmd(255, 2) = -2
LngCmd(256, 1) = 14 ' <WmaSampleRate>
LngCmd(256, 2) = -2
LngCmd(257, 1) = 12 ' <WmaDuration>
LngCmd(257, 2) = -2
LngCmd(258, 1) = 11 ' <WmaBitRate>
LngCmd(258, 2) = -2
LngCmd(259, 1) = 9 ' <WmaTrack>
LngCmd(259, 2) = -2
LngCmd(260, 1) = 9 ' <WmaTitle>
LngCmd(260, 2) = -2
LngCmd(261, 1) = 10 ' <WmaArtist>
LngCmd(261, 2) = -2
LngCmd(262, 1) = 9 ' <WmaAlbum>
LngCmd(262, 2) = -2
LngCmd(263, 1) = 9 ' <WmaGenre>
LngCmd(263, 2) = -2
LngCmd(264, 1) = 11 ' <WmaComment>
LngCmd(264, 2) = -2
LngCmd(265, 1) = 8 ' <WmaYear>
LngCmd(265, 2) = -2

' Pour le retaillage
With m_oAutoPos
    .AddAssignment Me.Frame3, Me.FrameDroite, tCONTAINER_RELATIVE_POS_BOTTOM
    .AddAssignment Me.Command13, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
    .AddAssignment Me.Command12, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
    .AddAssignment Me.Command7, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
    .AddAssignment Me.Command15, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
    .AddAssignment Me.Command16, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
    .AddAssignment Me.Command2, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
    .AddAssignment Me.ListView2, Me.PanelList, tCONTAINER_HEIGHT_DELTA_BOTTOM
End With

' Chargement des règles
'LesRegles.LoadRulesFromFile AppPath & "Rules.ini"

' Chargement des commandes dans les menus contextuels
For i = 1 To vnbcmd
    listcmd.AddItem langage(i)
    listcmd.ItemData(listcmd.NewIndex) = hlplang(i)    ' Ajout du topic du fichier d'aide a appeler
Next
Dim nodX As Node
With TV1
    .Nodes.Add , , "Prefix", "Prefix"
    .Nodes.Add , , "Extension", "Extension"
    .Nodes.Add , , "General", "General"
    .Nodes.Add , , "Music", "Music"
    .Nodes.Add , , "Pictures", "Pictures"
End With
Set nodX = TV1.Nodes.Add("Music", tvwChild, "MusMP3", "MP3")
Set nodX = TV1.Nodes.Add("Music", tvwChild, "MusOGG", "OGG (Vorbis)")
Set nodX = TV1.Nodes.Add("Music", tvwChild, "MusVQF", "VQF")
Set nodX = TV1.Nodes.Add("Music", tvwChild, "MusWMA", "WMA")

LoadAvailableDrives
mvnb1 = -1
mvnb2 = -1
mvnb3 = -1
mvnb4 = -1
mvnb5 = -1

For i = 0 To vnbcmd - 1
    If UCase$(Left$(listcmd.List(i), 3)) = "<PR" Then
        mvnb1 = mvnb1 + 1
        If mvnb1 <> 0 Then Load m3cmdprefix(mvnb1)
        m3cmdprefix(mvnb1).Caption = listcmd.List(i)
        Set nodX = TV1.Nodes.Add("Prefix", tvwChild, Trim$(Str$(i)) + "|" + Trim$(Str$(listcmd.ItemData(i))), listcmd.List(i))
    Else
        If UCase$(Left$(listcmd.List(i), 5)) = "<EXIF" And UCase$(Left$(listcmd.List(i), 10)) <> "<EXIFEMPTY" Then
            mvnb5 = mvnb5 + 1
            If mvnb5 <> 0 Then Load cmdpictures(mvnb5)
            cmdpictures(mvnb5).Caption = listcmd.List(i)
            Set nodX = TV1.Nodes.Add("Pictures", tvwChild, Trim$(Str$(i)) + "|" + Trim$(Str$(listcmd.ItemData(i))), listcmd.List(i))
        Else
            If UCase$(Left$(listcmd.List(i), 3)) = "<EX" Then
                mvnb2 = mvnb2 + 1
                If mvnb2 <> 0 Then Load m3cmdextension(mvnb2)
                m3cmdextension(mvnb2).Caption = listcmd.List(i)
                Set nodX = TV1.Nodes.Add("Extension", tvwChild, Trim$(Str$(i)) + "|" + Trim$(Str$(listcmd.ItemData(i))), listcmd.List(i))
            Else
                If UCase$(Left$(listcmd.List(i), 4)) = "<VQF" Or UCase$(Left$(listcmd.List(i), 4)) = "<MP3" Or UCase$(Left$(listcmd.List(i), 4)) = "<OGG" Or UCase$(Left$(listcmd.List(i), 4)) = "<WMA" Then
                    mvnb4 = mvnb4 + 1
                    If mvnb4 <> 0 Then Load mimusic(mvnb4)
                    mimusic(mvnb4).Caption = listcmd.List(i)
                    If UCase$(Left$(listcmd.List(i), 4)) = "<VQF" Then
                        Set nodX = TV1.Nodes.Add("MusVQF", tvwChild, Trim$(Str$(i)) + "|" + Trim$(Str$(listcmd.ItemData(i))), listcmd.List(i))
                    Else
                        If UCase$(Left$(listcmd.List(i), 4)) = "<MP3" Then
                            Set nodX = TV1.Nodes.Add("MusMP3", tvwChild, Trim$(Str$(i)) + "|" + Trim$(Str$(listcmd.ItemData(i))), listcmd.List(i))
                        Else
                            If UCase$(Left$(listcmd.List(i), 4)) = "<WMA" Then
                                Set nodX = TV1.Nodes.Add("MusWMA", tvwChild, Trim$(Str$(i)) + "|" + Trim$(Str$(listcmd.ItemData(i))), listcmd.List(i))
                            Else
                                Set nodX = TV1.Nodes.Add("MusOGG", tvwChild, Trim$(Str$(i)) + "|" + Trim$(Str$(listcmd.ItemData(i))), listcmd.List(i))
                            End If
                        End If
                    End If
                Else
                    mvnb3 = mvnb3 + 1
                    If mvnb3 <> 0 Then Load mlang(mvnb3)
                    mlang(mvnb3).Caption = listcmd.List(i)
                    Set nodX = TV1.Nodes.Add("General", tvwChild, Trim$(Str$(i)) + "|" + Trim$(Str$(listcmd.ItemData(i))), listcmd.List(i))
                End If
            End If
        End If
    End If
Next

For i = 1 To vnboptionp
    Combo1.AddItem (optionp(i))
Next
For i = 1 To vnboptions
    Combo2.AddItem (options(i))
Next
AncTitre = Me.Caption
'rafraichir = True

Rem *** Lecture des paramètres pas défaut.
leretour = LoadPref()
LoadMenuPicture ' Chargement du menu des images

' Lecture des paramètres cachés ************************************************************************************
RefreshRate = GetSetting("THERename", "HiddenParam", "RefreshRate", 20)
DisplayRenMsg = GetSetting("THERename", "HiddenParam", "DisplayRenMsg", 1)
ShowPreviewList = GetSetting("THERename", "HiddenParam", "ShowPreviewList", 1)
' *******************************************************************************************************************
Text11.Text = LesOptions.PicturesFormat

    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
        'Case "CommandButton", "TextBox", "ComboBox", "ImageCombo", "HScrollBar", "LabelText"
        'Case "CommandButton", "TextBox", "ImageCombo", "HScrollBar", "LabelText"
        Case "CommandButton", "ImageCombo", "HScrollBar", "LabelText"
            ReDim Preserve K(i)
            Set K(i) = New cControlFlater
            K(i).Attach ctl
            i = i + 1
        End Select
    Next

FolderTreeview1(0).VirtualFolders = LesOptions.IncVirtualFolders
FolderTreeview1(0).HiddenFolders = LesOptions.IncHiddenFolders
If LesOptions.ToolbarButtons = 1 Then
    RENAME.Toolbar1.Style = tbrFlat
Else
    RENAME.Toolbar1.Style = tbrStandard
End If
ListView1.ColumnHeaders(3).Text = tlibdate(LesOptions.Dateformat + 1)
Combo1.ListIndex = LesOptions.DefOption1
Combo2.ListIndex = LesOptions.DefOption2
Combo1.Text = Combo1.List(LesOptions.DefOption1)
Combo2.Text = Combo2.List(LesOptions.DefOption2)
If LesOptions.ColumnsWiths = True Then
    ListView1.ColumnHeaders(1).width = LesOptions.WCol1
    ListView1.ColumnHeaders(2).width = LesOptions.WCol2
    ListView1.ColumnHeaders(3).width = LesOptions.WCol3
    ListView1.ColumnHeaders(4).width = LesOptions.WCol4
    ListView1.ColumnHeaders(5).width = LesOptions.WCol5
End If
If LesOptions.UseDefaultCyclicFile = 1 Then    ' Il faut utiliser un fichier de sélections cycliques par défaut
    OpenCyclic LesOptions.DefaultCyclicFile
End If
If LesOptions.UseDefaultAbbrevFile = 1 Then    ' Il faut utiliser un fichier d'abbréviations par défaut
    OpenAbbrev LesOptions.DefaultAbbrevFile
End If
mviewmp3tab.Checked = LesOptions.ShowMP3Tab
mviewpicturetab.Checked = LesOptions.ShowMusicTab
mviewtexttab.Checked = LesOptions.ShowTextTab
TabGen.TabVisible(2) = LesOptions.ShowMP3Tab
TabGen.TabVisible(3) = LesOptions.ShowMusicTab
TabGen.TabVisible(4) = LesOptions.ShowTextTab
If LesOptions.UseHistory = True Then ' Menu history
    mhistory.Enabled = True
Else
    mhistory.Enabled = False
End If
ListView1.FullRowSelect = LesOptions.FullRow
ListView1.GridLines = LesOptions.GridLines
If LesOptions.Center0rSave = 1 Then ' centrer la fenêtre
    Me.Move (Screen.width - Me.width) / 2, (Screen.height - Me.height) / 2
Else ' Rappeler sa position précédente
    RENAME.Left = LesOptions.lLeft
    RENAME.Top = LesOptions.lTOp
End If
RENAME.WindowState = LesOptions.wWindowState

If LesOptions.RememberWSize = 1 Then
    If LesOptions.wHeight <> -1 And LesOptions.wWidth <> -1 Then
        RENAME.height = LesOptions.wHeight
        RENAME.width = LesOptions.wWidth
    Else
        RENAME.WindowState = vbMaximized
    End If
End If
GlobalLoad = False
LesOptions.NumberOfRuns = LesOptions.NumberOfRuns + 1
NomSettings = ""
txtlang.Text = LesOptions.DefaultCommand
If LesOptions.RememberLastCommand = 1 Then txtlang.Text = LesOptions.LastCommand
If Trim$(LesOptions.LogFile) <> "" Then
    mviewlog.Enabled = True
Else
    mviewlog.Enabled = False
End If

If Len(Trim$(Command$)) > 0 Then
 If InStr(UCase$(Command$), ".REN") > 0 Then  ' On a spécifié un fichier .ren sur la ligne de commande
  aOuvrir = True
  PbFtv1 = True
  NomSettings = Replace(Trim$(Command$), Chr$(34), "")
  mopenset_Click
 Else ' On a demandé à ouvrir THE Rename depuis l'explorateur
  If UCase$(Command$) = "/CLEAN" Then
    DeleteSetting "THERename"
    MsgBox "All settings have been deleted !"
    End
  End If
  If UCase$(Command$) = "/WSIZE" Then
    RENAME.WindowState = vbNormal
    RENAME.height = 7740
    RENAME.width = 11085
  End If
  FolderTreeview1(0).SelectedFolder = Replace(Trim$(Command$), Chr$(34), "")
  Dir1Path = Replace(Trim$(Command$), Chr$(34), "")
  PbFtv1 = True
 End If
Else ' Rien sur la ligne de commande
    If LesOptions.RememberLastFolder = 1 Then
        If Trim$(LesOptions.LastFolder) <> "" Then
            FolderTreeview1(0).SelectedFolder = LesOptions.LastFolder
            Dir1Path = LesOptions.LastFolder
            PbFtv1 = True
        End If
    Else
        If Trim$(LesOptions.StartupDir) <> "" Then
            FolderTreeview1(0).SelectedFolder = LesOptions.StartupDir
            PbFtv1 = True
            Dir1Path = LesOptions.StartupDir
        End If
    End If
    
    If LesOptions.ShowPathInCaption = 1 Then
        Me.Caption = "THE Rename - " + Dir1Path
    End If
  'End If
End If

If LesOptions.UseAutoSave = 1 Then
    aOuvrir = True
    NomSettings = AppPath + "autosave.ren"
    mopenset_Click
End If

Exit Sub

ErrGen:
ErreurGrave "Form_Load"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim cR As New cRegistry, resultat As Integer
 If LesOptions.RememberLastCommand = 1 Then
    LesOptions.LastCommand = txtlang.Text
 End If
 If LesOptions.RememberLastFolder = 1 Then
    LesOptions.LastFolder = FolderTreeview1(0).SelectedFolder
  End If
  LesOptions.wWidth = RENAME.width
  LesOptions.wHeight = RENAME.height

cR.ClassKey = HKEY_CURRENT_USER
cR.SectionKey = "Software\VB and VBA Program Settings\THERename"
m_cMRU.Save cR
Unload preview
LesOptions.lLeft = RENAME.Left
LesOptions.lTOp = RENAME.Top
LesOptions.PicturesFormat = Text11.Text
resultat = SaveSettings()
If LesOptions.UseAutoSave = 1 Then
    NomSettings = AppPath + "autosave.ren"
    msave_Click ' et on lance la sauvegarde dans le répertoire du programme
End If
End Sub

Private Sub Form_Resize()
 If GlobalLoad Then Exit Sub
 On Error Resume Next
 If RENAME.width < 11085 Then
    RENAME.width = 11085
 End If
 If RENAME.height < 7740 Then
    RENAME.height = 7740
 End If
 ListView1.height = Me.ScaleHeight - état.height - ListView1.Top
 TabGen.Left = Me.ScaleWidth - TabGen.width - 10
 TabGen.height = ListView1.height
 FolderTreeview1(0).height = TabGen.height - TabGen.TabHeight - 150
 LvMP3.height = FolderTreeview1(0).height
 Acdsee.height = Acdsee.width
 If Acdsee.height > FolderTreeview1(0).height Then
    Acdsee.height = FolderTreeview1(0).height
 End If
 
 ListView1.width = TabGen.Left - ListView1.Left - 50
 Combo1_Click
 FrameDroite.height = TabGen.height - 500
 Text1.height = LvMP3.height
 Text1.width = LvMP3.width
 
 m_oAutoPos.RefreshPositions
 If PanelList.Visible = True Then
     Frame1.height = Frame3.Top - Frame1.Top
    PanelList.height = Frame1.height - Combo1.height - Combo1.Top - 140
 End If
 Text1.height = LvMP3.height
 Text1.width = LvMP3.width
 m_oAutoPos.RefreshPositions
 panelcmd.height = Frame3.Top - panelcmd.Top
 TV1.height = panelcmd.height - TV1.Top '- 100
 TV1.width = (panelcmd.width - TV1.Left) '- 25
 EtchedLine RENAME, 0, Toolbar1.Top + Toolbar1.height + 50, RENAME.width
 LesOptions.wHeight = RENAME.height
 LesOptions.wWidth = RENAME.width
End Sub

Private Sub Form_Terminate()
 If acpreview = True Then
  Unload preview
  acpreview = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim resultat As Integer
 If acpreview = True Then
  Unload preview
  acpreview = False
  End If
 LesOptions.lLeft = RENAME.Left
 LesOptions.lTOp = RENAME.Top
 LesOptions.wWidth = RENAME.width
 LesOptions.wHeight = RENAME.height
 
 If LesOptions.RememberLastCommand = 1 Then
   LesOptions.LastCommand = txtlang.Text
 End If
 If LesOptions.RememberLastFolder = 1 Then
    LesOptions.LastFolder = FolderTreeview1(0).SelectedFolder
  End If
 
 resultat = SaveSettings()
End Sub
Private Sub HTMLReport_Click()
Dim i As Long, sItem As String, vnb As Long, taille As Long, vnb2 As Long, vtempo As String
Dim vrep As String, szFilename As String, largeur As String, LeLien As String, VnbTagMP3 As Integer
Dim VnbTagVQF As Integer, Boucle1 As Integer, ChMP3 As String, ChVQF As String
Dim MP3Tags(57) As String, VQFTags(9) As String, SonInfo As String, vNomComplet As String
Dim ff As Integer

largeur = "20%"
If LesOptions.IncludePictInfo = 0 Then
    largeur = "35%"
End If

taille = 0
vnb = 0
On Error GoTo ErrorHandler
szFilename = ""
szFilename = DialogFile(Me.hWnd, 2, "Generate HTML report", "Report.html", "HTML Files (*.html)" & Chr$(0) & "*.html" & Chr$(0) & "All files (*.*)" & Chr$(0) & "*.*", Dir1Path, "htm")
If Trim$(szFilename) = "" Then
    Exit Sub
End If

Me.MousePointer = 11
If Recursive = False Then
    vrep = AddBackSlash(Dir1Path)
Else
    vrep = ""
End If
ff = FreeFile
Open szFilename For Output As #ff
Print #ff, "<!DOCTYPE HTML PUBLIC " + Chr$(34) + "-//W3C//DTD HTML 4.0 Transitional//EN" + Chr$(34) + ">"
Print #ff, "<HTML>" + vbCrLf + "<HEAD>" + vbCrLf + "<TITLE>THE Rename</TITLE>" + vbCrLf + "</HEAD>"
Print #ff, "<BODY BGCOLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + " TEXT=" + Chr$(34) + "#000000" + Chr$(34) + ">"
Print #ff, "<H1 ALIGN=" + Chr$(34) + "center" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#0000FF" + Chr$(34) + " SIZE=" + Chr$(34) + "+4" + Chr$(34) + ">THE Rename</FONT></H1>"
Print #ff, "<DIV ALIGN=" + Chr$(34) + "CENTER" + Chr$(34) + "><CENTER>"
Print #ff, "<P><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + " COLOR=" + Chr$(34) + "#0000FF" + Chr$(34) + ">" + "Content of directory " + Dir1Path + "   on " + Format$(Date, "Long Date") + " at " + Format$(Time, "Long Time") + "</FONT></P>"
Print #ff, "<TABLE BORDER=" + Chr$(34) + "3" + Chr$(34) + " CELLSPACING=" + Chr$(34) + "0" + Chr$(34) + " CELLPADDING=" + Chr$(34) + "2" + Chr$(34) + " WIDTH=" + Chr$(34) + "100%" + Chr$(34) + ">"
Print #ff, "<TR>"
Print #ff, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Name</FONT></TH>"
If LesOptions.HtmlIncFolder = 1 Then
    Print #ff, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Folder</FONT></TH>"
End If
If LesOptions.HtmlIncSize = 1 Then
    Print #ff, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Size</FONT></TH>"
End If
If LesOptions.HtmlIncDate = 1 Then
    Print #ff, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Date</FONT></TH>"
End If
If LesOptions.IncludePictInfo = 1 Then
    Print #ff, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Pict Info</FONT></TH>"
End If
If LesOptions.HtmlIncAttr = 1 Then
    Print #ff, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Attrib</FONT></TH>"
End If
If LesOptions.HtmlIncMusic = 1 Then ' Il faut inclure les infos sur les MP3 et sur les VQF
    VnbTagMP3 = NbMP3Tags(ChMP3)
    VnbTagVQF = NbVQFTags(ChVQF)
    For Boucle1 = 1 To VnbTagMP3
        Print #ff, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">" + MP3Caption(Boucle1) + "</FONT></TH>"
    Next
    For Boucle1 = 1 To VnbTagVQF
        Print #ff, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">" + VQFCaption(Boucle1) + "</FONT></TH>"
    Next
End If
Print #ff, "</TR>"
vnb2 = ListView1.ListItems.Count - 1
For i = 0 To vnb2
    sItem = LVGetName(ListView1, i)
    If LVGetItemName(ListView1, i, 4) = "File" Then
        vtempo = ""
        vtempo = ImgInfo(vrep + sItem)
        If vtempo = "" Then
            vtempo = "&nbsp"
        End If
    End If
    Print #ff, "<TR>"
    If LesOptions.IncludeLinks = 1 Then ' Il faut inclure des liens vers les images
        If vtempo <> "&nbsp" Then   ' Seulement si on arrive à lire ses propriétés
            LeLien = "<A HREF=" + Chr$(34) + Prefixe(sItem) & "." & Suffixe(sItem) + Chr$(34) + ">"
            Print #ff, "<TD><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp " + LeLien + Prefixe(sItem) & "." & Suffixe(sItem) + "</A></FONT></TD>"
        Else    ' On n'a pas réussi à lire les propriétés de l'image
            Print #ff, "<TD><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp " + Prefixe(sItem) & "." & Suffixe(sItem) + "</FONT></TD>"
        End If
    Else    ' Il ne faut pas inclure de liens vers les images.
        Print #ff, "<TD><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp " + Prefixe(sItem) & "." & Suffixe(sItem) + "</FONT></TD>"
    End If
    If LesOptions.HtmlIncFolder = 1 Then
        Print #ff, "<TD><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp " + ExtractPath(vrep + sItem) + "</FONT></TD>"
    End If
    If LesOptions.HtmlIncSize = 1 Then
        Print #ff, "<TD ALIGN=right><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + Format$(LVGetItemName(ListView1, i, 1), "### ### ### ###") + "&nbsp</FONT></TD>"
    End If
    If LesOptions.HtmlIncDate = 1 Then
        Print #ff, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + LVGetItemName(ListView1, i, 2) + "</FONT></TD>"
    End If
    If LesOptions.IncludePictInfo = 1 Then
        Print #ff, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + vtempo + "</FONT></TD>"
    End If
    vtempo = LVGetItemName(ListView1, i, 3)
    If Trim$(vtempo) = "" Then
        vtempo = "&nbsp;"
    End If
    If LesOptions.HtmlIncAttr = 1 Then
        Print #ff, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + vtempo + "</FONT></TD>"
    End If
    
    If LesOptions.HtmlIncMusic = 1 Then ' Il faut inclure les infos sur les MP3 et sur les VQF
        For Boucle1 = 1 To 57
            MP3Tags(Boucle1) = ""
        Next
        For Boucle1 = 1 To 9
            VQFTags(Boucle1) = ""
        Next
        vNomComplet = vrep + Prefixe(sItem) & "." & Suffixe(sItem)
        If UCase$(Suffixe(sItem)) = "VQF" Then   ' Chargement des infos sur le VQF
            SonInfo = ""
            SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
            VQFTags(1) = Blanc(MusVQF.Author)
            VQFTags(2) = Blanc(MusVQF.BitRate)
            VQFTags(3) = Blanc(MusVQF.Comment)
            VQFTags(4) = Blanc(MusVQF.Copyright)
            VQFTags(5) = Blanc(MusVQF.SaveAsFilename)
            VQFTags(6) = Blanc(MusVQF.Mono_Stereo)
            VQFTags(7) = Blanc(MusVQF.Quality)
            VQFTags(8) = Blanc(MusVQF.SampleRate)
            VQFTags(9) = Blanc(MusVQF.Title)
            For Boucle1 = 1 To VnbTagVQF
                Print #ff, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + VQFTags(Val(GetToken(ChVQF, "|", Boucle1))) + "</FONT></TD>"
            Next
            For Boucle1 = 1 To VnbTagMP3
                Print #ff, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp</FONT></TD>"
            Next
        End If

        If UCase$(Suffixe(sItem)) = "MP3" Then   ' Chargement des infos sur le MP3
            SonInfo = ""
            SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
            MP3Tags(1) = Blanc(MusMP3.Album)
            MP3Tags(2) = Blanc(MusMP3.Artist)
            MP3Tags(3) = Blanc(MusMP3.Band)
            MP3Tags(4) = Blanc(MusMP3.BPM)
            MP3Tags(5) = Blanc(MusMP3.Comment)
            MP3Tags(6) = Blanc(MusMP3.Composer)
            MP3Tags(7) = Blanc(MusMP3.Conductor)
            MP3Tags(8) = Blanc(MusMP3.ContentGroup)
            MP3Tags(9) = Blanc(MusMP3.Copyright)
            MP3Tags(10) = Blanc(MusMP3.mDate)
            MP3Tags(11) = Blanc(MusMP3.EncodedBy)
            MP3Tags(12) = Blanc(MusMP3.EncryptionMethod)
            MP3Tags(13) = Blanc(MusMP3.FileOwner)
            MP3Tags(14) = Blanc(MusMP3.FileType)
            MP3Tags(15) = Blanc(MusMP3.Genre)
            MP3Tags(16) = Blanc(MusMP3.GroupIdent)
            MP3Tags(17) = Blanc(MusMP3.InitialKey)
            MP3Tags(18) = Blanc(MusMP3.InvolvedPeopleList)
            MP3Tags(19) = Blanc(MusMP3.ISRC)
            MP3Tags(20) = Blanc(MusMP3.Language)
            MP3Tags(21) = Blanc(MusMP3.LinkedInformation)
            MP3Tags(22) = Blanc(MusMP3.Lyricist)
            MP3Tags(23) = Blanc(MusMP3.MediaType)
            MP3Tags(24) = Blanc(MusMP3.MixArtist)
            MP3Tags(25) = Blanc(MusMP3.NetRadioOwner)
            MP3Tags(26) = Blanc(MusMP3.NetRadioStation)
            MP3Tags(27) = Blanc(MusMP3.OriginalAlbum)
            MP3Tags(28) = Blanc(MusMP3.OriginalArtist)
            MP3Tags(29) = Blanc(MusMP3.OriginalFilename)
            MP3Tags(30) = Blanc(MusMP3.OriginalLyricist)
            MP3Tags(31) = Blanc(MusMP3.OriginalYear)
            MP3Tags(32) = Blanc(MusMP3.PartOfASet)
            MP3Tags(33) = Blanc(MusMP3.PlayListDelay)
            MP3Tags(34) = Blanc(MusMP3.PopulariMeter)
            MP3Tags(35) = Blanc(MusMP3.Publisher)
            MP3Tags(36) = Blanc(MusMP3.RecordingDates)
            MP3Tags(37) = Blanc(MusMP3.SoftwareEncodingSettings)
            MP3Tags(38) = Blanc(MusMP3.SongLength)
            MP3Tags(39) = Blanc(MusMP3.SubTitle)
            MP3Tags(40) = Blanc(MusMP3.SynchronizedLyric)
            MP3Tags(41) = Blanc(MusMP3.TermsOfUse)
            MP3Tags(42) = Blanc(MusMP3.Time)
            MP3Tags(43) = Blanc(MusMP3.Title)
            MP3Tags(44) = Blanc(MusMP3.TotalTracks)
            MP3Tags(45) = Blanc(MusMP3.TrackNumber)
            MP3Tags(46) = Blanc(MusMP3.UnsynchronizedLyric)
            MP3Tags(47) = Blanc(MusMP3.UserText)
            MP3Tags(48) = Blanc(MusMP3.wwwArtist)
            MP3Tags(49) = Blanc(MusMP3.wwwAudioFile)
            MP3Tags(50) = Blanc(MusMP3.wwwAudioSource)
            MP3Tags(51) = Blanc(MusMP3.wwwCommercialInfo)
            MP3Tags(52) = Blanc(MusMP3.wwwCopyright)
            MP3Tags(53) = Blanc(MusMP3.wwwPayment)
            MP3Tags(54) = Blanc(MusMP3.wwwPublisher)
            MP3Tags(55) = Blanc(MusMP3.wwwRadioPage)
            MP3Tags(56) = Blanc(MusMP3.wwwUserURL)
            MP3Tags(57) = Blanc(MusMP3.Year)
            For Boucle1 = 1 To VnbTagMP3
                Print #ff, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + MP3Tags(Val(GetToken(ChMP3, "|", Boucle1))) + "</FONT></TD>"
            Next
            For Boucle1 = 1 To VnbTagVQF
                Print #ff, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp</FONT></TD>"
            Next
        End If
        If UCase$(Suffixe(sItem)) <> "MP3" And UCase$(Suffixe(sItem)) <> "VQF" Then
            For Boucle1 = 1 To VnbTagMP3 + VnbTagVQF
                Print #ff, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp;</FONT></TD>"
            Next
        End If
    End If
    Print #ff, "</TR>"
    vnb = vnb + 1
    taille = taille + Val(LVGetItemName(ListView1, i, 1))
Next
Print #ff, "</TABLE></CENTER></DIV>"
Print #ff, "<BR><BR><TABLE BORDER=" + Chr$(34) + "0" + Chr$(34) + " CELLSPACING=" + Chr$(34) + "0" + Chr$(34) + " WIDTH=" + Chr$(34) + "100%" + Chr$(34) + "><TR>"
Print #ff, "<TD WIDTH=" + Chr$(34) + "15%" + Chr$(34) + ">Number of files</TD>"
Print #ff, "<TD WIDTH=" + Chr$(34) + "85%" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#0000FF" + Chr$(34) + ">" + Trim$(Str$(vnb)) + "</FONT></TD>"
Print #ff, "</TR>"
Print #ff, "<TR>"
Print #ff, "<TD WIDTH=" + Chr$(34) + "15%" + Chr$(34) + ">Total Size</TD>"
Print #ff, "<TD WIDTH=" + Chr$(34) + "85%" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#0000FF" + Chr$(34) + ">" + Format$(taille, "### ### ### ###") + "</FONT></TD>"
Print #ff, "</TR></TABLE><BR><BR><BR>"
Print #ff, "</BODY></HTML>"

Close #ff
Me.MousePointer = 0
RefreshF5
Exit Sub

ErrorHandler:
 RENAME.MousePointer = 0
 MsgBox "There was a problem while generating the report...!!!"
 Exit Sub
End Sub
Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim vnb As Long
Piège1 = False
If Not Cancel Then
    On Error GoTo ErrorNewname
    If VancFichier <> AddBackSlash(Dir1Path) & NewString Then
        If LesOptions.RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
            NewString = RemIllegals(NewString)
        End If
        If LesOptions.RemoveStartingSpaces = 1 Then ' Il faut supprimer les espaces en début de fichier
            NewString = LTrim$(NewString)
        End If
        Name VancFichier As AddBackSlash(Dir1Path) + NewString
    End If
    vnb = remplissage()
    Exit Sub
End If

ErrorNewname:
    If Err.Number = 58 Then
        MsgBox "Error, a file already exist with this name", vbInformation, "Error"
    End If
    vnb = remplissage()
    Exit Sub
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
 On Error GoTo ErrGen
 Piège1 = True
 VancFichier = AddBackSlash(Dir1Path) + ListView1.SelectedItem.Text
Exit Sub
ErrGen:
ErreurGrave "ListView1_BeforeLabelEdit"
End Sub
Private Sub ListView1_Click()
On Error Resume Next
Dim vnb As Long
Dim Pref As String
Dim vtmp As String
If letat = False Then
 vnb = LVGetCountSelected(ListView1)
 état.Panels(4).Text = Trim$(Str$(vnb))
End If
Pref = ""
If ListView1.ListItems.Count > 0 Then
    Pref = UCase$(Suffixe(ListView1.SelectedItem.Text))
End If

If TabGen.TabVisible(4) = True Then
    If InStr(LesOptions.TextToView, Pref) <> 0 Then
        If Recursive = False Then
            vtmp = AddBackSlash(Dir1Path)
        Else
            vtmp = ""
        End If
        LoadText vtmp & ListView1.SelectedItem.Text
    End If
End If

If mviewmp3tab.Checked Then
    Select Case Pref
        Case "MP3"
            LoadLvMP3
        Case "OGG"
            LoadLvOgg
        Case "VQF"
            loadLvVQF
        Case "AFM"
            LoadAFM
'        Case "WAV"
'            LoadWav
        Case "WMA"
            LoadWMA
        Case Else
            LvMP3.ListItems.Clear
    End Select
End If


If mviewpicturetab.Checked Then
    Dim lWidth As Long
    Dim lHeight As Long
    If Pref = "JPG" Or Pref = "BMP" Or Pref = "GIF" Or Pref = "JPEG" Or Pref = "DIB" Or Pref = "WMF" Or Pref = "EMF" Or Pref = "ICO" Or Pref = "CUR" Then
        Me.MousePointer = vbHourglass
        If Recursive = False Then
            vtmp = AddBackSlash(Dir1Path)
        Else
            vtmp = ""
        End If
        Set Acdsee.Picture = LoadPicture(vtmp & ListView1.SelectedItem.Text)
        Info_JPG vtmp & ListView1.SelectedItem.Text, lWidth, lHeight
        Select Case LesOptions.PicturesPreview
            Case 0  ' Real Size
                Acdsee.ZoomReal
            Case 1  ' Stretch
                Acdsee.Stretch
            Case 2  ' Best Fit
                Acdsee.BestFit
        End Select
        Me.MousePointer = vbDefault
        If Pref = "JPG" Or Pref = "JPEG" Or Pref = "JPE" Or Pref = "TIF" Or Pref = "TIFF" And mviewmp3tab.Checked Then
            Dim col As ExifTags
            Dim tg As ExifTag
            Dim exobj As ExifPage
            ' create ExifPage object
            Set exobj = New ExifPage
            ' extract exif info in the file named and it will return ExifTags collection object
            Set col = exobj.ExtractExifInfo(vtmp & ListView1.SelectedItem.Text)
            ' get all the ExifTag objects for information
            For Each tg In col
                Set itmX = LvMP3.ListItems.Add(, , tg.Name)
                itmX.SubItems(1) = tg.Value
            Next
            ResizeLvMp3
        End If
    Else
        Acdsee.Clear
    End If
End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrGen
Dim currSortKey As Integer
Me.MousePointer = 11
   sOrder = Not sOrder
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.SortOrder = Abs(Not ListView1.SortOrder = 1)
   Select Case ColumnHeader.Index - 1
      Case 0:
               ListView1.Sorted = True
               If LesOptions.UseNaturalSort = 1 Then
                    SendMessageLong ListView1.hWnd, LVM_SORTITEMS, ListView1.hWnd, AddressOf CompareNatural
                End If
      Case 3:
               ListView1.Sorted = True
      Case 2:
               ListView1.Sorted = True
               SendMessageLong ListView1.hWnd, LVM_SORTITEMS, ListView1.hWnd, AddressOf CompareDates
      Case 1:
               ListView1.Sorted = True
               SendMessageLong ListView1.hWnd, LVM_SORTITEMS, ListView1.hWnd, AddressOf CompareValues
   End Select
Me.MousePointer = vbDefault
ListView1.SortKey = ColumnHeader.Index - 1
currSortKey = ListView1.SortKey
Exit Sub

ErrGen:
If Err.Number = 13 Then
    MsgBox "Warning, may it is not possible to sort files according to the date and time format you have selected, Sorry", vbOKOnly, "Error"
    Resume Next
End If
ErreurGrave "ListView1_ColumnClick"
End Sub

Private Sub ListView1_DblClick()
On Error GoTo ErrGen
 Select Case LesOptions.ActDblClick
  Case 0
   mdelete_Click
  Case 1
   mopen_Click
  Case 2
   mpropertyes_Click
  Case 3
   mprint_Click
  Case 4
   mexec_Click
  Case 5
   mexplorer_Click
 End Select
Exit Sub
ErrGen:
ErreurGrave "ListView1_DblClick"
End Sub
Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error GoTo ErrGen
 Select Case KeyAscii
  Case 13  ' Entrée
   ListView1_DblClick
 End Select
Exit Sub
ErrGen:
ErreurGrave "ListView1_KeyPress"
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
 Dim tAlt As Integer
 On Error GoTo ErrGen
 tAlt = GetKeyState(VK_MENU)
 If (tAlt = -127 Or tAlt = -128) And KeyCode = 13 Then ' Afficher les propriétés
    mpropertyes_Click
 End If
 If Shift = 1 And (KeyCode = 40 Or KeyCode = 38 Or KeyCode = 35 Or KeyCode = 36 Or KeyCode = 34 Or KeyCode = 33) Then
    ListView1_Click
 End If

 If Shift = 0 And (KeyCode = 40 Or KeyCode = 38 Or KeyCode = 35 Or KeyCode = 36 Or KeyCode = 34 Or KeyCode = 33) Then
    ListView1_Click
 End If

 If (tAlt = -127 Or tAlt = -128) And KeyCode = 38 Then ' Déplacer vers le haut
    MoveFilesUp ListView1, RENAME
 End If

 If (tAlt = -127 Or tAlt = -128) And KeyCode = 40 Then ' Déplacer vers le bas
    MoveFilesDown ListView1, RENAME
 End If

 Select Case KeyCode
  Case 13 'Entrée
    If Piège1 = True Then
        Piège1 = False
    End If
  Case 27 ' Esc
     Piège1 = False
  Case 111 ' / invert
    InvertSelection
    ListView1.SetFocus
  Case 106  ' * select all
    SelectAll
    ListView1.SetFocus
  Case 109 ' - unselect
    Unselect
    ListView1.SetFocus
  Case 107  ' + step
    StepSelection
    ListView1.SetFocus
  Case 46 ' Supress
   If Shift = 1 Then
    TemDelete = True
   End If
   mdelete_Click
  Case 113 ' F2
   If Shift = 0 Then
    RenameManually
   End If
 End Select
 état.Panels(4).Text = LVGetCountSelected(RENAME.ListView1)
Exit Sub
ErrGen:
ErreurGrave "ListView1_KeyUp"
End Sub
Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim vnb As Long
vnb = LVGetCountSelected(RENAME.ListView1)
état.Panels(4).Text = vnb
On Error GoTo ErrGen

If Button = 2 Then
    If Recursive = True Then
        madd.Visible = True
    Else
        madd.Visible = False
    End If
    If vnb = 2 Then
        mnuswap.Visible = True
    Else
        mnuswap.Visible = False
    End If
    If vnb > 0 Then
        mregrename.Enabled = True
    Else
        mregrename.Enabled = False
    End If
    PopupMenu mcontextuel
End If
Exit Sub

ErrGen:
ErreurGrave "ListView1_MouseUp"
End Sub
Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrGen
    Dim vnb As Long, i As Integer, vnb1 As Integer, vnb2 As Long
    Dim fichier As String, vrai As Boolean
    Dim VancRep As String
    ' Variables utilisées dans le cas ou les fichiers droppés ne sont pas du même répertoire
    Dim clsFind As New clsFindFile, chaine As String
    Dim strFile As String, attributs As Long

    RENAME.MousePointer = 11
    Unselect
    List1.Clear
    vnb = 0
    If Data.GetFormat(vbCFFiles) Then
        For i = 1 To Data.Files.Count
            List1.AddItem Data.Files(i)
        Next
        vrai = True
        VancRep = ExtractPath(Trim$(List1.List(0)))
        For i = 0 To List1.ListCount - 1
            If ExtractPath(Trim$(List1.List(i))) <> VancRep Then
                vrai = False  ' Il faudra se mettre en mode récursif
                Exit For
            End If
        Next
        ChDir ExtractPath(List1.List(0))
        FolderTreeview1(0).SelectedFolder = ExtractPath(List1.List(0))
        Dir1Path = ExtractPath(List1.List(0))

        If vrai Then ' On est sur le même répertoire
            vnb1 = CharOccurs(List1.List(0), "\")
            vnb1 = At(List1.List(0), "\", vnb1)
            vnb = remplissage
            Unselect
            If vnb = List1.ListCount Then 'On gagne du temps si tous les fichiers on été sélectionnés
                SelectAll
            Else ' Seuls certains fichiers du répertoire ont été déposés, il faut les retrouver.
                For i = 0 To List1.ListCount - 1
                    vnb1 = CharOccurs(List1.List(0), "\")
                    vnb1 = At(List1.List(0), "\", vnb1)
                    fichier = Mid$(List1.List(i), vnb1 + 1)
                    vnb2 = LVSearch(ListView1, fichier + Chr$(0))
                    If vnb2 <> -1 Then
                        ListView1.ListItems(vnb2 + 1).Selected = True
                    End If
                Next
            End If
        Else ' On passe en mode récursif
            If Toolbar1.Buttons(13).Value <> 1 Then
                Toolbar1.Buttons(13).Value = tbrPressed
                m2recursive.Checked = True
                MsgBox "Warning, you have dropped files from different directories so I'm going to use the recursive mode"
            End If
            Recursive = True
            ListView1.ListItems.Clear
            clsFind.Dateformat = "short Date"
            For i = 0 To List1.ListCount - 1
                strFile = clsFind.Find(List1.List(i), False)
                If Len(strFile) > 0 Then
                    If (clsFind.FileAttributes And vbDirectory) = 0 Then
                        attributs = clsFind.FileAttributes
                        chaine = ""
                        If attributs And FILE_ATTRIBUTE_READONLY Then chaine = "R"
                        If attributs And FILE_ATTRIBUTE_HIDDEN Then chaine = chaine + "H"
                        If attributs And FILE_ATTRIBUTE_SYSTEM Then chaine = chaine + "S"
                        If attributs And FILE_ATTRIBUTE_ARCHIVE Then chaine = chaine + "A"
                        Set itmX = ListView1.ListItems.Add(, , Trim$(List1.List(i)))
                        With itmX
                            .Text = Trim$(List1.List(i))
                            .SubItems(1) = clsFind.FileSize
                            .SubItems(2) = clsFind.GetCreationDate
                            .SubItems(3) = chaine
                        End With
                    End If
                End If
            Next
            SelectAll  ' Sélection de tous les fichiers
        End If
    Else
        MsgBox "Error, you can just drop files on THE Rename !"
    End If ' Sont ce des données acceptables par ce superbe programme ?
    état.Panels(3).Text = Trim$(Str$(ListView1.ListItems.Count))
    état.Panels(4).Text = Trim$(Str$(ListView1.ListItems.Count))
    If ListView1.ListItems.Count > 0 Then
        ListView1.ListItems(1).EnsureVisible
    End If
    RENAME.MousePointer = 0
    Exit Sub
ErrGen:
    ErreurGrave "ListView1_OLEDragDrop"
End Sub
Private Sub ListView2_AfterLabelEdit(Cancel As Integer, NewString As String)
 CharInterdits NewString
End Sub

Private Sub ListView2_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = 113 Then
  ListView2.StartLabelEdit
 End If
End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
  PopupMenu m4contextuel
 End If
End Sub

Private Sub ListView2_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    CopySelected False
End Sub

Private Sub LvMP3_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub M1Invert_Click()
 InvertSelection
End Sub

Private Sub m1selectAll_Click()
 SelectAll
End Sub

Private Sub M1Step_Click()
 StepSelection
End Sub

Private Sub M1Unselect_Click()
 Unselect
End Sub

Private Sub m2manually_Click()
 RenameManually
End Sub

Private Sub m2preview_Click()
 PreviewRename
End Sub

Private Sub m2recursive_Click()
 Dim vnb As Long
 If Toolbar1.Buttons(13).Value <> 1 Then
  Toolbar1.Buttons(13).Value = tbrPressed
  Recursive = True
  m2recursive.Checked = True
 Else
  Toolbar1.Buttons(13).Value = tbrUnpressed
  Recursive = False
  m2recursive.Checked = False
 End If
 vnb = remplissage()
 état.Panels(3).Text = Trim$(Str$(vnb))
 état.Panels(4).Text = "0"
End Sub

Private Sub M2Start_Click()
 StartRename
End Sub

Private Sub m3cmdextension_Click(Index As Integer)
 InsertTextInTextBoxFromMenu txtlang, m3cmdextension(Index).Caption
End Sub

Private Sub m3cmdprefix_Click(Index As Integer)
    InsertTextInTextBoxFromMenu txtlang, m3cmdprefix(Index).Caption
End Sub

Private Sub mabrev_Click()
 Fabbrev.Show 1
End Sub

Private Sub madd_Click()
 Dim szFilename As String
 Dim attributs As Long
 Dim chaine As String
 szFilename = DialogFile(Me.hWnd, 1, "Add File(s)", "*.*", "All files" & Chr$(0) & "*.*", Dir1Path, "*.*")
 If Trim$(szFilename) = "" Then
  Exit Sub
 End If
 Set itmX = ListView1.ListItems.Add(, , szFilename)
 itmX.Text = szFilename
 itmX.SubItems(1) = FileLen(szFilename)
 itmX.SubItems(2) = Format$(FileDateTime(szFilename), "short date")
 attributs = GetAttr(szFilename)
 chaine = ""
 If attributs And FILE_ATTRIBUTE_READONLY Then chaine = "R"
 If attributs And FILE_ATTRIBUTE_HIDDEN Then chaine = chaine + "H"
 If attributs And FILE_ATTRIBUTE_SYSTEM Then chaine = chaine + "S"
 If attributs And FILE_ATTRIBUTE_ARCHIVE Then chaine = chaine + "A"
 itmX.SubItems(3) = chaine
 état.Panels(3).Text = Trim$(Str$(ListView1.ListItems.Count + 1))
End Sub

Private Sub maddbag_Click()
Dim vret As Integer
vret = Bag(2)
End Sub
Private Sub madddirectory_Click()
Dim repvoulu As String
Dim i As Integer

repvoulu = Trim$(Dir1Path)
For i = 1 To 20
 If fav(i) = repvoulu Then
  Beep
  MsgBox "This directory is already in your favorites"
  Exit Sub
 End If
Next

For i = 20 To 2 Step -1
 fav(i) = fav(i - 1)
Next
fav(1) = repvoulu

For i = 0 To 19
 RENAME.menufav(i).Caption = "&" + Chr$(65 + i) + " " + fav(i + 1)
 RENAME.mnufav(i).Caption = "&" + Chr$(65 + i) + " " + fav(i + 1)
Next
End Sub

Private Sub maddfavorites_Click()
 madddirectory_Click
End Sub

Private Sub mapropos_Click()
 About2.Show 1
End Sub
Private Sub mattrib_Click()
 AttrEncours = 1
 attributs.Show 1
End Sub
Private Sub mchangetab_Click()
    Dim vnum As Integer
    Dim vrai As Boolean
    vrai = False
    vnum = TabGen.Tab + 1
    
    While Not vrai
        If vnum >= TabGen.Tabs Then
            vnum = 0
        End If
        If TabGen.TabVisible(vnum) Then
            TabGen.Tab = vnum
            vrai = True
        Else
            vnum = vnum + 1
        End If
    Wend
End Sub

Private Sub mchg1_Click()
' Change files date and time now
Dim i As Long, vnb As Long
Dim sItem As String, chemin As String
i = 0
chemin = AddBackSlash(Trim$(Dir1Path))
If Recursive = True Then
 chemin = ""
End If

DTEnCours = 4
fDT.Show 1

i = LVGetItemSelected(ListView1, -1)
While i <> -1
 sItem = LVGetName(ListView1, i)
 DT4.SetFileDateTime (Trim$(chemin + sItem))
 i = LVGetItemSelected(ListView1, i)
Wend
vnb = remplissage()
état.Panels(3).Text = Trim$(Str$(vnb))
état.Panels(4).Text = "0"
End Sub

Private Sub mchg2_Click()
' Changes files attributes now
Dim i As Long, vnb As Long
Dim sItem As String, chemin As String
i = 0
chemin = AddBackSlash(Trim$(Dir1Path))
If Recursive = True Then
 chemin = ""
End If

AttrEncours = 4
attributs.Show 1

i = LVGetItemSelected(ListView1, -1)
While i <> -1
 sItem = LVGetName(ListView1, i)
 Attr4.ChangeAttr (chemin + sItem)
 i = LVGetItemSelected(ListView1, i)
Wend
vnb = remplissage()
état.Panels(3).Text = Trim$(Str$(vnb))
état.Panels(4).Text = "0"
End Sub

Private Sub mchgattdir_Click()
' Changes files attributes now
Dim i As Long, vnb As Long
Dim chemin As String
chemin = AddBackSlash(Trim$(Dir1Path))
AttrEncours = 4
attributs.Show 1
Attr4.ChangeAttr (chemin)
vnb = remplissage()
état.Panels(3).Text = Trim$(Str$(vnb))
état.Panels(4).Text = "0"
End Sub
Private Sub mclearbag_Click()
 ListView3.ListItems.Clear
End Sub

Private Sub mcopy_Click()
Dim i As Long, chaine As String, vtmp As String, vnb As Long
On Error Resume Next    ' Pour éviter les débordements de chaine
FCopyName.Show 1
If LOk = False Then ' L'utilisateur a abandonné
    Exit Sub
End If
RENAME.MousePointer = 11
chaine = ""
If LOption2 = 0 Then ' Copy All
    vnb = ListView1.ListItems.Count - 1
    For i = 0 To vnb
        Select Case LOption1
            Case 0  ' Prefix only
                vtmp = Prefixe(LVGetName(ListView1, i))
            Case 1  ' Extension only
                vtmp = Suffixe(LVGetName(ListView1, i))
            Case 2  ' Prefix + extension
                vtmp = Prefixe(LVGetName(ListView1, i)) & "." & Suffixe(LVGetName(ListView1, i))
            Case 3  ' Path only
                If Recursive = True Then
                    vtmp = ExtractPath(LVGetName(ListView1, i))
                Else
                    vtmp = Dir1Path
                End If
            Case 4  ' Path + Full Path name
                If Recursive = True Then
                    vtmp = LVGetName(ListView1, i)
                Else
                    vtmp = AddBackSlash(Dir1Path) + LVGetName(ListView1, i)
                End If
        End Select
        chaine = chaine + vtmp + vbCrLf
    Next
Else    ' Copy selected
i = LVGetItemSelected(ListView1, -1)
    While i <> -1
        Select Case LOption1
            Case 0  ' Prefix only
                vtmp = Prefixe(LVGetName(ListView1, i))
            Case 1  ' Extension only
                vtmp = Suffixe(LVGetName(ListView1, i))
            Case 2  ' Prefix + extension
                vtmp = Prefixe(LVGetName(ListView1, i)) & "." & Suffixe(LVGetName(ListView1, i))
            Case 3  ' Path only
                If Recursive = True Then
                    vtmp = ExtractPath(LVGetName(ListView1, i))
                Else
                    vtmp = Dir1Path
                End If
            Case 4  ' Path + Full Path name
                If Recursive = True Then
                    vtmp = LVGetName(ListView1, i)
                Else
                    vtmp = AddBackSlash(Dir1Path) + LVGetName(ListView1, i)
                End If
        End Select
        chaine = chaine + vtmp + vbCrLf
        i = LVGetItemSelected(ListView1, i)
    Wend
End If
Clipboard.Clear
Clipboard.SetText chaine
chaine = ""
RENAME.MousePointer = 0
End Sub

Private Sub mcopy2_Click()
Clipboard.Clear
Clipboard.SetText Dir1Path
Beep
End Sub

Private Sub mcopybag_Click()
Dim vret As Integer
vret = Bag(1)
End Sub

Private Sub mcreatefold_Click()
Dim i As Long
Dim vnb As Long
Dim vnbren As Long
On Error Resume Next
fcreatfold.Show 1   ' Affichage de la fenêtre
If LOk = False Then ' L'utilisateur a abandonné
    Exit Sub
End If
RENAME.MousePointer = 11

If LesOptions.Misc5 = 0 Then ' All
    vnb = ListView1.ListItems.Count
    For i = 0 To vnb - 1
        If LVGetItemName(ListView1, i, 4) = "File" Then
            MoveCopyFile i, vnb, i + 1
        End If
    Next
Else    ' Selected
    vnb = LVGetCountSelected(ListView1)
    i = LVGetItemSelected(ListView1, -1)
    While i <> -1
        If LVGetItemName(ListView1, i, 4) = "File" Then
            vnbren = vnbren + 1
            MoveCopyFile i, vnb, vnbren
        End If
        i = LVGetItemSelected(ListView1, i)
    Wend
End If
If LesOptions.Misc4 = 2 Then RefreshF5
état.Panels(1).Text = ""
état.Panels(2).Text = ""
RENAME.MousePointer = 0
End Sub

Private Sub mcreatfoldman_Click()
    FFoldMan.Show vbModal
End Sub

Private Sub mcutadditive_Click()
Dim vret As Integer
vret = Bag(4)
End Sub

Private Sub mcutbag_Click()
Dim vret As Integer
vret = Bag(3)
End Sub

Private Sub mdatetime_Click()
DTEnCours = 1
fDT.Show 1
End Sub

Private Sub mdelete_Click()
Dim fileop As New CSHFileOp, chemin As String, vnb As Long
Dim i As Long
Dim sItem As String
With fileop
    .ParentWnd = hWnd
    .ClearSourceFiles
    .ClearDestFiles
End With
chemin = AddBackSlash(Trim$(Dir1Path))
If Recursive = True Then
 chemin = ""
End If

i = LVGetItemSelected(ListView1, -1)
While i <> -1
 sItem = LVGetName(ListView1, i)
 fileop.AddSourceFile chemin + sItem
 i = LVGetItemSelected(ListView1, i)
Wend

If TemDelete = True Then
 fileop.AllowUndo = False
 TemDelete = False
End If
If Not fileop.DeleteFiles Then
End If
'File1.Refresh
vnb = remplissage()
état.Panels(3).Text = Trim$(Str$(vnb))
état.Panels(4).Text = "0"
End Sub

Private Sub mdelrep_Click()
Dim fileop As New CSHFileOp
With fileop
    .ParentWnd = hWnd
    .ClearSourceFiles
    .ClearDestFiles
    .AddSourceFile Dir1Path
    .DeleteFiles
End With
remplissage
End Sub

Private Sub mdgroupe_Click()
 fmdgroupe.Show 1
 RefreshF5
End Sub

Private Sub mdisconnect_Click()
 Dim R As Long
 R = WNetDisconnectDialog(RENAME.hWnd, RESOURCETYPE_DISK)
End Sub

Private Sub mdosprompthere_Click()
 Dim i As Long
 On Error Resume Next
 ChDrive Left$(Dir1Path, 3)
 ChDir Dir1Path
 i = Shell("command.com", 1)
End Sub

Private Sub mend_Click(Index As Integer)
 Dim resultat As Integer
 LesOptions.lLeft = RENAME.Left
 LesOptions.lTOp = RENAME.Top
  If LesOptions.RememberLastCommand = 1 Then
    LesOptions.LastCommand = txtlang.Text
 End If
 If LesOptions.RememberLastFolder = 1 Then
    LesOptions.LastFolder = FolderTreeview1(0).SelectedFolder
  End If
 resultat = SaveSettings()
 If LesOptions.UseAutoSave = 1 Then
    NomSettings = AppPath + "autosave.ren"
    msave_Click ' et on lance la sauvegarde dans la répertoire du programme
 End If
 End
End Sub

Private Sub menufav_Click(Index As Integer)
Dim repertoire As String, unite As String
On Error GoTo ErrorHandler
If Len(Trim$(menufav(Index).Caption)) > 2 Then ' l'option contient un répertoire
  unite = Mid$(Trim$(menufav(Index).Caption), 4, 3)
  repertoire = Trim$(Mid$(Trim$(menufav(Index).Caption), 7))
  ChDir unite & repertoire
  TemMove = False
  FolderTreeview1(0).Visible = False
  FolderTreeview1(0).SelectedFolder = unite & repertoire
  Dir1Path = unite & repertoire
  FolderTreeview1(0).Visible = True
  mundo.Enabled = False
  List2.Clear
  List3.Clear
End If
Exit Sub

ErrorHandler:   ' Error-handling routine.
Select Case Err.Number  ' Evaluate error number.
 Case 76
  MsgBox "Error - This directory doesn't not exist any more !"
 Case 68
  MsgBox "Error - This drive is unavailable !"
End Select
Exit Sub
End Sub

Private Sub mexec_Click()
Dim titre As String
Dim Message As String
SHExecuteDlg hWnd, 0, 0, titre, Message, 0
End Sub

Private Sub mexecmd_Click()
    ExeCmd = True
    FExecCmd.Show 1
    ExeCmd = False
End Sub

Private Sub mexplorer_Click()
FileExecutor Me.hWnd, Dir1Path, "Explore"
End Sub

Private Sub mexport_Click()
    FExport.Show 1
End Sub

Private Sub mfilefind_Click()
 Dim VK_ACTION As Long
 VK_ACTION = &H46
 Call keybd_event(VK_LWIN, 0, 0, 0)
 Call keybd_event(VK_ACTION, 0, 0, 0)
 Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub mfirst_Click()
  ListView1.ListItems(1).Selected = True
  ListView1.ListItems(1).EnsureVisible
End Sub

Private Sub mglob_Click()
    RechEnCours = 3
    If RechPref = True Or RechSuff = True Then
        If MsgBox("Warning, search and replace in the prefix or in the extension is already active. You can't use global search and replace in this case. Would you like to deactivate search and replace in the prefix and extension ?", vbYesNo, "Warning") = vbNo Then
            Exit Sub
        End If
    End If
    RechPref = False
    RechSuff = False
    advsearch.Show 1
End Sub

Private Sub mgo1_Click()
    GoFTV 1
End Sub

Private Sub mgolast_Click()
    GoFTV 4
End Sub

Private Sub mgonext_Click()
    GoFTV 2
End Sub

Private Sub mgoprev_Click()
    GoFTV 3
End Sub

Private Sub mgoto_Click()
 fgoto.Show 1
End Sub

Private Sub mgreat_Click()
    Dim i As Long, max As Currency
    Dim nomfic As String, test1 As String
    Dim vprefixe As String, vsuffixe As String, vnomorig As String
    Dim sItem As String
    Dim vretour As Integer
    Dim clsFind As New clsFindFile
    Dim strFile As String

    max = -1

    vretour = MsgBox("Would you like to specify an expression to filter files ?", vbYesNo, "Find biggest counter")
    If vretour = vbYes Then
        FFilter.Show 1
        If FilterOk = False Then
            Exit Sub
        End If
    End If

    ' Filtre mais pas par expression régulière
    If vretour = vbOK And FilterRegular = False Then
        strFile = clsFind.Find(FilterExpr, False)
        Do While Len(strFile)
            état.Panels(1).Text = "Search with " & strFile
            test1 = Menage2(strFile)
            If Len(test1) > 0 Then
                If Val(test1) > max Then
                    max = Val(test1)
                    nomfic = strFile
                End If
            End If
            strFile = clsFind.FindNext()
        Loop
    End If

    For i = 0 To ListView1.ListItems.Count - 1
        sItem = LVGetName(ListView1, i)
        vprefixe = Prefixe(sItem) ' Le préfixe uniquement
        vsuffixe = Suffixe(sItem) ' Le suffixe uniquement
        vnomorig = vprefixe + "." + vsuffixe
        If vretour <> vbOK Then
            test1 = Menage2(vnomorig)
            If Len(test1) > 0 Then
                If Val(test1) > max Then
                    max = Val(test1)
                    nomfic = sItem
                End If
            End If
        Else
            If vretour = vbOK And FilterRegular = True Then
                If vnomorig Like FilterExpr Then
                    état.Panels(1).Text = "Search with " & vnomorig
                    test1 = Menage2(vnomorig)
                    If Len(test1) > 0 Then
                        If Val(test1) > max Then
                            max = Val(test1)
                            nomfic = sItem
                        End If
                    End If
                End If
            End If
        End If
    Next

    If max <> -1 Then
        MsgBox "Biggest counter is " + Trim$(Str$(max)) + " on file " + nomfic, , "Biggest counter"
    Else
        MsgBox "No counter found", , "Biggest counter"
    End If
    ModifierLigne
End Sub

Private Sub mhistory_Click()
 fHistory.Show 1
End Sub

Private Sub mimusic_Click(Index As Integer)
  InsertTextInTextBoxFromMenu txtlang, mimusic(Index).Caption
End Sub

Private Sub mindex_Click()
Dim X As Long
X = WinHelp(Me.hWnd, App.HelpFile, HELP_CONTEXT, 1)
End Sub

Private Sub minfos_Click()
 Infos.Show 1
End Sub
Private Sub mlang_Click(Index As Integer)
 InsertTextInTextBoxFromMenu txtlang, mlang(Index).Caption
End Sub

Private Sub mlast_Click()
  ListView1.ListItems(ListView1.ListItems.Count).Selected = True
  ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
End Sub

Private Sub mmakedir_Click()
    Dim rep As String
    Dim chem As String
    Dim vnb As Integer
    rep = ""
    chem = AddBackSlash(Dir1Path)
    rep = InputBox("Enter directory name", "Make directory", chem)
    If rep = "" Then
        Exit Sub
    End If
    On Error Resume Next
    vnb = CreateNestedFoldersByPath(rep)
    RefreshF5
End Sub
Private Sub mmap_Click()
 Dim R As Long
 R = WNetConnectionDialog(RENAME.hWnd, RESOURCETYPE_DISK)
End Sub

Private Sub mmapped_drives_Click()
 FMappedDrives.Show 1
End Sub

Private Sub mmiddle_Click()
 ListView1.ListItems(Int(ListView1.ListItems.Count / 2)).Selected = True
 ListView1.ListItems(Int(ListView1.ListItems.Count / 2)).EnsureVisible
End Sub

Private Sub mnewformlist_Click()
    Dim szFilename As String, chemin As String
    Dim ligne As String
    Dim i As Integer
    Dim ff As Integer
    i = 0
    If ListView2.ListItems.Count = 0 Then
        MsgBox "error, this can only be use on a list containing files"
        Exit Sub
    End If
    szFilename = DialogFile(Me.hWnd, 1, "Open a list file containing only new names", "rename.list", "List" & Chr$(0) & "*.list" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "list")
    If Trim$(szFilename) = "" Then
        Exit Sub
    End If
    ListView2.Visible = False
    RENAME.MousePointer = 11
    chemin = ExtractPath(szFilename)
    ChDir chemin
    TemMove = False
    FolderTreeview1(0).Visible = False
    FolderTreeview1(0).SelectedFolder = chemin
    Dir1Path = chemin
    FolderTreeview1(0).Visible = True
    If LesOptions.ListDelimiter = 0 Then LesOptions.ListDelimiter = 9
    ff = FreeFile
    Open szFilename For Input As #ff
    Line Input #ff, ligne
    While Not EOF(ff)
        If Len(Trim$(ligne)) > 0 Then
            i = i + 1
            If i <= ListView2.ListItems.Count Then
                If LesOptions.RemoveGuill = 1 Then
                    ligne = Replace(ligne, Chr$(34), "")
                End If
                ListView2.ListItems.Item(i).Text = ligne
            End If
        End If
        Line Input #ff, ligne
    Wend
    Close #ff
    If Len(Trim$(ligne)) > 0 Then
        i = i + 1
        If i <= ListView2.ListItems.Count Then
            ListView2.ListItems.Item(i).Text = ligne
        End If
    End If

    ListView2.Visible = True
    RENAME.MousePointer = 0
    Exit Sub
erreur:
    MsgBox "There was an error during process !"
    ListView2.Visible = True
    RENAME.MousePointer = 0
    Exit Sub

End Sub

Private Sub mnudrives_Click(Index As Integer)
  FolderTreeview1(0).Visible = False
  FolderTreeview1(0).SelectedFolder = mnudrives(Index).Caption
  Dir1Path = mnudrives(Index).Caption
  FolderTreeview1(0).Visible = True
  mundo.Enabled = False
  List2.Clear
  List3.Clear
End Sub

Private Sub mnufav_Click(Index As Integer)
 menufav_Click (Index)
End Sub

Private Sub mnuFile_Click(Index As Integer)
aOuvrir = True
NomSettings = m_cMRU.File(CLng(mnuFile(Index).Tag))
mopenset_Click
pDisplayMRU True
End Sub

Private Sub mnuhistory_Click(Index As Integer)
 ChangeRepHistorique Index
End Sub
Private Sub mnuswap_Click()
' Inverse le nom de 2 fichiers sélectionnés (et seulement de 2)
Dim fileop As New CSHFileOp
Dim i As Long
Dim vnom1 As String
Dim vnom2 As String
Dim chemin1 As String
Dim chemin2 As String
chemin1 = AddBackSlash(Trim$(Dir1Path))
chemin2 = AddBackSlash(Trim$(Dir1Path))

i = LVGetItemSelected(ListView1, -1)
vnom1 = LVGetName(ListView1, i)
If Recursive = True Then
  chemin1 = AddBackSlash(ExtractPath(vnom1)) ' Si on est en récursif, il faut récupérer le chemin du fichier
End If
i = LVGetItemSelected(ListView1, i)
vnom2 = LVGetName(ListView1, i)
If Recursive = True Then
 chemin2 = AddBackSlash(ExtractPath(vnom2)) ' Si on est en récursif, il faut récupérer le chemin du fichier
End If
With fileop
    .ParentWnd = hWnd
    .ConfirmOperation = False
    .ClearSourceFiles
    .ClearDestFiles
    ' on renomme le premier fichier avec un nom bidon
    .AddSourceFile chemin1 + Prefixe(vnom1) & "." & Suffixe(vnom1)
    .AddDestFile chemin1 + "$hthouzard$"
    .RenameFiles
    ' on renomme le deuxième fichier
    .ClearSourceFiles
    .ClearDestFiles
    .AddSourceFile chemin2 + Prefixe(vnom2) & "." & Suffixe(vnom2)
    .AddDestFile chemin2 + Prefixe(vnom1) & "." & Suffixe(vnom1)
    .RenameFiles
    ' on re renomme le premier
    .ClearSourceFiles
    .ClearDestFiles
    .AddSourceFile chemin1 + "$hthouzard$"
    .AddDestFile chemin1 + Prefixe(vnom2) & "." & Suffixe(vnom2)
    .RenameFiles
End With
' et on raffaichit
SendKeys "{F5}"
End Sub
Private Sub mopen_Click()
Dim chemin As String
chemin = AddBackSlash(Trim$(Dir1Path))
FileExecutor Me.hWnd, chemin + ListView1.ListItems(ListView1.SelectedItem.Index), "Open"
End Sub

Private Sub mopenset_Click()
Dim Version As String, szFilename As String, ChemSave As String
On Error GoTo errloadset

ChemSave = Dir1Path
If aOuvrir = False Then
 If Len(Trim$(LesOptions.SettingsDirectory)) > 0 Then
  szFilename = DialogFile(Me.hWnd, 1, "Open settings", "settings.ren", "Rename" & Chr$(0) & "*.ren" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "ren")
 Else
  szFilename = DialogFile(Me.hWnd, 1, "Open settings", "settings.ren", "Rename" & Chr$(0) & "*.ren" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "ren")
 End If
 If Trim$(szFilename) = "" Then
  Exit Sub
 End If
Else
 szFilename = NomSettings
 aOuvrir = False
End If

NomSettings = szFilename  ' Nom du fichier de settings en cours
msave.Enabled = True
DTEnCours = 1

Version = LoadSet("General", szFilename, "Version")
If Version <> "2.0" Then
    GoTo SuiteLoad
End If

Toolbar1.Buttons(13).Value = Val(LoadSet("General", szFilename, "Recursive"))
txtlang.Text = LoadSet("General", szFilename, "txtlang")
Combo1.ListIndex = Val(LoadSet("General", szFilename, "Combo1"))
cmdtxt1.Text = LoadSet("General", szFilename, "CmdTxt1")
cmdtxt2.Text = LoadSet("General", szFilename, "CmdTxt2")
Text9.Text = LoadSet("General", szFilename, "Txt9")
Check7.Value = Val(LoadSet("General", szFilename, "Check7"))
Folder1 = Val(LoadSet("General", szFilename, "Folder1"))
Folder2 = Val(LoadSet("General", szFilename, "Folder2"))
Folder3 = Val(LoadSet("General", szFilename, "Folder3"))
Folder4 = Val(LoadSet("General", szFilename, "Folder4"))
Folder5 = LoadSet("General", szFilename, "Folder5")
Folder6 = LoadSet("General", szFilename, "Folder6")
FolderOk = Val(LoadSet("General", szFilename, "FolderOk"))
Option3(9).Value = Val(LoadSet("General", szFilename, "Option3_9"))
Option3(10).Value = Val(LoadSet("General", szFilename, "Option3_10"))
Option3(11).Value = Val(LoadSet("General", szFilename, "Option3_11"))
Option1(0).Value = Val(LoadSet("General", szFilename, "Option1_0"))
Text2.Text = LoadSet("General", szFilename, "Text2")
Option1(1).Value = Val(LoadSet("General", szFilename, "Option1_1"))
Text14.Text = LoadSet("General", szFilename, "Text14")
Option2(0).Value = Val(LoadSet("General", szFilename, "Option2_0"))
Option2(1).Value = Val(LoadSet("General", szFilename, "Option2_1"))
Check5.Value = Val(LoadSet("General", szFilename, "Check5"))
Option3(3).Value = Val(LoadSet("General", szFilename, "Option3_3"))
Option3(4).Value = Val(LoadSet("General", szFilename, "Option3_4"))
Option3(5).Value = Val(LoadSet("General", szFilename, "Option3_5"))
Check6.Value = Val(LoadSet("General", szFilename, "Check6"))
Option3(8).Value = Val(LoadSet("General", szFilename, "Option3_8"))
Option3(7).Value = Val(LoadSet("General", szFilename, "Option3_7"))
Option3(6).Value = Val(LoadSet("General", szFilename, "Option3_6"))
Check3.Value = Val(LoadSet("General", szFilename, "Check3"))
Text3.Text = LoadSet("General", szFilename, "Text3")
Text4.Text = LoadSet("General", szFilename, "Text4")
Text5.Text = LoadSet("General", szFilename, "Text5")
Combo3.ListIndex = Val(LoadSet("General", szFilename, "Combo3"))
Option3(0).Value = Val(LoadSet("General", szFilename, "Option3_0"))
Option3(1).Value = Val(LoadSet("General", szFilename, "Option3_1"))
Option3(2).Value = Val(LoadSet("General", szFilename, "Option3_2"))
Check1.Value = Val(LoadSet("General", szFilename, "Check1"))
Option3(17).Value = Val(LoadSet("General", szFilename, "Option3_17"))
Option3(16).Value = Val(LoadSet("General", szFilename, "Option3_16"))
Option3(15).Value = Val(LoadSet("General", szFilename, "Option3_15"))
Combo2.ListIndex = Val(LoadSet("General", szFilename, "Combo2"))
Check11.Value = Val(LoadSet("General", szFilename, "Check11"))
Text16.Text = LoadSet("General", szFilename, "Text16")
Text17.Text = LoadSet("General", szFilename, "Text17")
Text18.Text = LoadSet("General", szFilename, "Text18")
Combo4.ListIndex = Val(LoadSet("General", szFilename, "Combo4"))
Option3(26).Value = Val(LoadSet("General", szFilename, "Option3_26"))
Option3(25).Value = Val(LoadSet("General", szFilename, "Option3_25"))
Option3(24).Value = Val(LoadSet("General", szFilename, "Option3_24"))
Check4.Value = Val(LoadSet("General", szFilename, "Check4"))
Option3(14).Value = Val(LoadSet("General", szFilename, "Option3_14"))
Option3(13).Value = Val(LoadSet("General", szFilename, "Option3_13"))
Option3(12).Value = Val(LoadSet("General", szFilename, "Option3_12"))
Option4(0).Value = Val(LoadSet("General", szFilename, "Option4_0"))
Text8.Text = LoadSet("General", szFilename, "Text8")
Option4(1).Value = Val(LoadSet("General", szFilename, "Option4_1"))
Text15.Text = LoadSet("General", szFilename, "Text15")
Option5(0).Value = Val(LoadSet("General", szFilename, "Option5_0"))
Option5(1).Value = Val(LoadSet("General", szFilename, "Option5_1"))
Check12.Value = Val(LoadSet("General", szFilename, "Check12"))
Option3(29).Value = Val(LoadSet("General", szFilename, "Option3_29"))
Option3(28).Value = Val(LoadSet("General", szFilename, "Option3_28"))
Option3(27).Value = Val(LoadSet("General", szFilename, "Option3_27"))
Check13.Value = Val(LoadSet("General", szFilename, "Check13"))
Option3(30).Value = Val(LoadSet("General", szFilename, "Option3_30"))
Option3(31).Value = Val(LoadSet("General", szFilename, "Option3_31"))
Option3(32).Value = Val(LoadSet("General", szFilename, "Option3_32"))

Attr1.Archive = Val(LoadSet("Attributs", szFilename, "Archive"))
Attr1.System = Val(LoadSet("Attributs", szFilename, "System"))
Attr1.ReadOnly = Val(LoadSet("Attributs", szFilename, "ReadOnly"))
Attr1.Hidden = Val(LoadSet("Attributs", szFilename, "Hidden"))
Attr1.SETArchive = Val(LoadSet("Attributs", szFilename, "SetArchive"))
Attr1.SETSystem = Val(LoadSet("Attributs", szFilename, "SetSystem"))
Attr1.SETReadOnly = Val(LoadSet("Attributs", szFilename, "SetReadOnly"))
Attr1.SETHidden = Val(LoadSet("Attributs", szFilename, "SetHidden"))
Attr1.AtrOk = Val(LoadSet("Attributs", szFilename, "UseAttr"))

DT1.DTOk = Val(LoadSet("Date and Time", szFilename, "UseDateTime"))
DT1.CreateDOption = Val(LoadSet("Date and Time", szFilename, "CreateDOption"))
DT1.CreateDFixed = IIf(Trim$(LoadSet("Date and Time", szFilename, "CreateDFixed")) <> "", LoadSet("Date and Time", szFilename, "CreateDFixed"), 0)
DT1.CreateDInc1 = Val(LoadSet("Date and Time", szFilename, "CreateDInc1"))
DT1.CreateDInc2 = Val(LoadSet("Date and Time", szFilename, "CreateDInc2"))
DT1.CreateDInc3 = Val(LoadSet("Date and Time", szFilename, "CreateDInc3"))
DT1.CreateTOption = Val(LoadSet("Date and Time", szFilename, "CreateTOption"))
DT1.CreateTFixed = LoadSet("Date and Time", szFilename, "CreateTFixed")
DT1.AccessDOption = Val(LoadSet("Date and Time", szFilename, "AccessDOption"))
DT1.AccessDFixed = IIf(Trim$(LoadSet("Date and Time", szFilename, "AccessDFixed")) <> "", LoadSet("Date and Time", szFilename, "AccessDFixed"), 0)
DT1.AccessDInc1 = Val(LoadSet("Date and Time", szFilename, "AccessDInc1"))
DT1.AccessDInc2 = Val(LoadSet("Date and Time", szFilename, "AccessDInc2"))
DT1.AccessDInc3 = Val(LoadSet("Date and Time", szFilename, "AccessDInc3"))
DT1.AccessTOption = Val(LoadSet("Date and Time", szFilename, "AccessTOption"))
DT1.AccessTFixed = LoadSet("Date and Time", szFilename, "AccessTFixed")
DT1.WriteDOption = Val(LoadSet("Date and Time", szFilename, "WriteDOption"))
DT1.WriteDFixed = IIf(Trim$(LoadSet("Date and Time", szFilename, "WriteDFixed")) <> "", LoadSet("Date and Time", szFilename, "WriteDFixed"), 0)
DT1.WriteDInc1 = Val(LoadSet("Date and Time", szFilename, "WriteDInc1"))
DT1.WriteDInc2 = Val(LoadSet("Date and Time", szFilename, "WriteDInc2"))
DT1.WriteDInc3 = Val(LoadSet("Date and Time", szFilename, "WriteDInc3"))
DT1.WriteTOption = Val(LoadSet("Date and Time", szFilename, "WriteTOption"))
DT1.WriteTFixed = LoadSet("Date and Time", szFilename, "WriteTFixed")

If Trim$(LoadSet("Folder", szFilename, "CurrentFolder")) <> "<DoNotRestore>" Then
    Dir1Path = Trim$(LoadSet("Folder", szFilename, "CurrentFolder"))
End If

rech1.SearchString = LoadSet("Search and Replace in Prefix", szFilename, "SearchString")
rech1.ReplaceString = LoadSet("Search and Replace in Prefix", szFilename, "ReplaceString")
rech1.ReplaceAll = Val(LoadSet("Search and Replace in Prefix", szFilename, "ReplaceAll"))
rech1.MatchCase = Val(LoadSet("Search and Replace in Prefix", szFilename, "MatchCase"))
rech1.SearchFromleft = Val(LoadSet("Search and Replace in Prefix", szFilename, "SearchFromLeft"))
rech1.SearchFrom = Val(LoadSet("Search and Replace in Prefix", szFilename, "SearchFrom"))
rech1.SearchTo = Val(LoadSet("Search and Replace in Prefix", szFilename, "SearchTo"))
rech1.SearchCharacters = LoadSet("Search and Replace in Prefix", szFilename, "SearchCharacters")
rech1.ReplaceCharacters = LoadSet("Search and Replace in Prefix", szFilename, "ReplaceCharacters")
rech1.UseRegExp = Val(LoadSet("Search and Replace in Prefix", szFilename, "UseRegExp"))

rech2.SearchString = LoadSet("Search and Replace in Extension", szFilename, "SearchString")
rech2.ReplaceString = LoadSet("Search and Replace in Extension", szFilename, "ReplaceString")
rech2.ReplaceAll = Val(LoadSet("Search and Replace in Extension", szFilename, "ReplaceAll"))
rech2.MatchCase = Val(LoadSet("Search and Replace in Extension", szFilename, "MatchCase"))
rech2.SearchFromleft = Val(LoadSet("Search and Replace in Extension", szFilename, "SearchFromLeft"))
rech2.SearchFrom = Val(LoadSet("Search and Replace in Extension", szFilename, "SearchFrom"))
rech2.SearchTo = Val(LoadSet("Search and Replace in Extension", szFilename, "SearchTo"))
rech2.SearchCharacters = LoadSet("Search and Replace in Extension", szFilename, "SearchCharacters")
rech2.ReplaceCharacters = LoadSet("Search and Replace in Extension", szFilename, "ReplaceCharacters")
rech2.UseRegExp = Val(LoadSet("Search and Replace in Extension", szFilename, "UseRegExp"))

rech3.SearchString = LoadSet("Global Search and Replace", szFilename, "SearchString")
rech3.ReplaceString = LoadSet("Global Search and Replace", szFilename, "ReplaceString")
rech3.ReplaceAll = Val(LoadSet("Global Search and Replace", szFilename, "ReplaceAll"))
rech3.MatchCase = Val(LoadSet("Global Search and Replace", szFilename, "MatchCase"))
rech3.SearchFromleft = Val(LoadSet("Global Search and Replace", szFilename, "SearchFromLeft"))
rech3.SearchFrom = Val(LoadSet("Global Search and Replace", szFilename, "SearchFrom"))
rech3.SearchTo = Val(LoadSet("Global Search and Replace", szFilename, "SearchTo"))
rech3.SearchCharacters = LoadSet("Global Search and Replace", szFilename, "SearchCharacters")
rech3.ReplaceCharacters = LoadSet("Global Search and Replace", szFilename, "ReplaceCharacters")
rech3.UseRegExp = Val(LoadSet("Global Search and Replace", szFilename, "UseRegExp"))

RechPref = Val(LoadSet("Search and Replace", szFilename, "UseSearchPref"))
RechSuff = Val(LoadSet("Search and Replace", szFilename, "UseSearchSuff"))
RechGlob = Val(LoadSet("Search and Replace", szFilename, "UseGlobalSearch"))

' Infos sur les MP3
UseMP3 = Val(LoadSet("MP3", szFilename, "Use"))
MusMP3.Rule = LoadSet("MP3", szFilename, "Rule")
MusMP3.PlaceWhereToPut = Val(LoadSet("MP3", szFilename, "PlaceWhereToPut"))
MusMP3.DefaultArtistToUse = LoadSet("MP3", szFilename, "DefaultArtistToUse")
MusMP3.DefaultYearToUse = LoadSet("MP3", szFilename, "DefaultYearToUse")
MusMP3.DefaultGenreToUse = LoadSet("MP3", szFilename, "DefaultGenreToUse")
MusMP3.DefaultAlbumToUse = LoadSet("MP3", szFilename, "DefaultAlbumToUse")
MusMP3.DefaultTitleToUse = LoadSet("MP3", szFilename, "DefaultTitleToUse")

' Infos sur les Ogg
UseOGG = Val(LoadSet("OGG", szFilename, "Use"))
MusOgg.Rule = LoadSet("OGG", szFilename, "Rule")
MusOgg.PlaceWhereToPut = Val(LoadSet("OGG", szFilename, "PlaceWhereToPut"))
MusOgg.DefaultAlbumToUse = LoadSet("OGG", szFilename, "DefaultAlbumToUse")
MusOgg.DefaultArtistToUse = LoadSet("OGG", szFilename, "DefaultArtistToUse")
MusOgg.DefaultGenreToUse = LoadSet("OGG", szFilename, "DefaultGenreToUse")
MusOgg.DefaultTitleToUse = LoadSet("OGG", szFilename, "DefaultTitleToUse")

' Infos sur les VQF
UseVQF = Val(LoadSet("VQF", szFilename, "Use"))
MusVQF.Rule = LoadSet("VQF", szFilename, "Rule")
MusVQF.PlaceWhereToPut = Val(LoadSet("VQF", szFilename, "PlaceWhereToPut"))
MusVQF.DefaultArtistToUse = LoadSet("VQF", szFilename, "DefaultArtistToUse")
MusVQF.DefaultTitle = LoadSet("VQF", szFilename, "DefaultTitle")

' Infos sur les tags EXIF
PicEXIF.UseEXIF = Val(LoadSet("EXIF", szFilename, "Use"))
PicEXIF.Rule = LoadSet("EXIF", szFilename, "Rule")
PicEXIF.PlaceWhereToPut = Val(LoadSet("EXIF", szFilename, "PlaceWhereToPut"))

If Toolbar1.Buttons(13).Value = tbrPressed Then
    Recursive = True
    remplissage
End If
    
' ********************************************************************************
' MRU
SuiteLoad:
m_cMRU.AddFile szFilename
pDisplayMRU True

On Error GoTo Erreur2
If Trim$(LoadSet("Folder", szFilename, "CurrentFolder")) <> "<DoNotRestore>" Then
    If Mid$(Dir1Path, 2, 1) = ":" Then
        ChDrive Left$(Dir1Path, 1)
    End If
    ChDir Dir1Path
    FolderTreeview1(0).SelectedFolder = Dir1Path
End If
état.Panels(1).Text = "Settings opened"
Exit Sub

errloadset:
 MsgBox "There was an error during the process. May be the settings file is corrupted or can't be found. Settings files older than this version can't be opened ... sorry !"
 Exit Sub

Erreur2:
 MsgBox "Error. The program was unable to connect to" + vbCrLf + Dir1Path
 Dir1Path = ChemSave
 Exit Sub
End Sub

Private Sub moptions_Click()
 OptionVis = True
 doptions.Show 1
 OptionVis = False
 If Trim$(LesOptions.LogFile) <> "" Then
    mviewlog.Enabled = True
 Else
    mviewlog.Enabled = False
 End If
 Text11.Text = LesOptions.PicturesFormat
End Sub
Private Sub morganyze_Click()
 ffavoris.Show 1
End Sub

Private Sub mpastebag_Click()
Dim vret As Integer
vret = Bag(5)
End Sub

Private Sub mpastekeep_Click()
Dim vret As Integer
vret = Bag(6)
End Sub

Private Sub mpastenewnames_Click()
Dim NewNames As String
Dim i As Integer
Dim j As Integer
Dim vnb As Integer
Dim lachaine As String
vnb = ListView2.ListItems.Count
j = 0

 If Clipboard.GetFormat(vbCFText) Then
    NewNames = Clipboard.GetText(vbCFText)
    If NewNames <> "" Then
        While i <= vnb And j <= Len(NewNames)
            j = j + 1
            If Mid$(NewNames, j, 1) <> Chr$(10) Then
                If Mid$(NewNames, j, 1) <> Chr$(13) Then
                    lachaine = lachaine + Mid$(NewNames, j, 1)
                End If
            Else
                i = i + 1
                If i <= vnb Then
                    ListView2.ListItems.Item(i).Text = lachaine
                    lachaine = ""
                End If
            End If
        Wend
    End If
 Else
    MsgBox "I can't paste names from the clipboard as it is not in text format, sorry !", vbOKOnly, "Sorry..."
 End If
End Sub

Private Sub MPict_Click(Index As Integer)
Select Case Index
    Case 0  ' Zoom In
        Acdsee.ZoomIn
    Case 1  ' Zoom Out
        Acdsee.ZoomOut
    Case 2  ' Real Size
        Acdsee.ZoomReal
    Case 3  ' Stretch
        Acdsee.Stretch
    Case 4  ' Best Fit
        Acdsee.BestFit
End Select
End Sub

Private Sub mprevioustab_Click()
    Dim vnum As Integer
    Dim vrai As Boolean
    vrai = False
    vnum = TabGen.Tab - 1
    If vnum < 0 Then
        vnum = TabGen.Tabs - 1
    End If
    
    While Not vrai
        If vnum < 0 Then
            vnum = TabGen.Tabs - 1
        End If
        If TabGen.TabVisible(vnum) Then
            TabGen.Tab = vnum
            vrai = True
        Else
            vnum = vnum - 1
        End If
    Wend
End Sub

Private Sub mprint_Click()
Dim chemin As String
chemin = AddBackSlash(Trim$(Dir1Path))
FileExecutor Me.hWnd, chemin + ListView1.ListItems(ListView1.SelectedItem.Index), "Print"
End Sub
Private Sub mprintdir_Click()
Dim i As Long
Dim sItem As String, vnb As Long, taille As Long
Dim vtempo As String
Dim vrep As String
Dim vnb2 As Long
taille = 0
Dim TailleMax As Integer
vnb = 0
On Error GoTo ErrorHandler
RENAME.MousePointer = 11
TailleMax = 0
If Recursive = False Then
    vrep = AddBackSlash(Dir1Path)
Else
    vrep = ""
End If
Printer.Print
Printer.Print "Content of directory " + Dir1Path + " on " + Format$(Date, "Long Date") + " at " + Format$(Time, "Long Time")
Printer.Print " "
Printer.Print " "
vnb2 = ListView1.ListItems.Count
For i = 0 To vnb2
    sItem = LVGetName(ListView1, i)
    If Len(Trim$(sItem)) > TailleMax Then
        TailleMax = Len(sItem)
    End If
Next

If LesOptions.DirectoryReport <> 1 Then
    Printer.Print "Filename" + Chr$(9) + Chr$(9) + "Size" + Chr$(9) + "Date" + Chr$(9) + Chr$(9) + "Attributes" + Chr$(9) + Chr$(9) + "Pict Info"
    Printer.Print " "
End If
For i = 0 To ListView1.ListItems.Count - 1
 sItem = LVGetName(ListView1, i)
 If LesOptions.DirectoryReport = 1 Then
  Printer.Print sItem
 Else
  If LesOptions.IncludePictInfo = 1 Then
    If LVGetItemName(ListView1, i, 4) = "File" Then
        vtempo = ""
        vtempo = ImgInfo(vrep + sItem)
    End If
  Else
    vtempo = ""
  End If
  Printer.Print Trim$(sItem) + String$(TailleMax - Len(Trim$(sItem)), Chr$(32)) + Chr$(9) + Format$(LVGetItemName(ListView1, i, 1), "### ### ### ###") + Chr$(9) + LVGetItemName(ListView1, i, 2) + Chr$(9) + LVGetItemName(ListView1, i, 3) + Chr$(9) + Chr$(9) + vtempo
 End If
 vnb = vnb + 1
 taille = taille + Val(LVGetItemName(ListView1, i, 1))
Next
Printer.Print " "
Printer.Print " "
Printer.Print "Total of " + Trim$(Str$(vnb)) + " file(s), " + Format$(taille, "### ### ### ###") + " bytes"
Printer.EndDoc

RENAME.MousePointer = 0
Exit Sub

ErrorHandler:
 RENAME.MousePointer = 0
 MsgBox "There was a problem printing to your printer."
 Exit Sub
End Sub

Private Sub mprop2_Click()
ShowProperties Dir1Path, Me
End Sub

Private Sub mpropertyes_Click()
Dim fName As String
Dim chemin As String
chemin = AddBackSlash(Trim$(Dir1Path))
If Recursive = True Then
 chemin = ""
End If
fName = chemin + ListView1.ListItems(ListView1.SelectedItem.Index)
ShowProperties fName, Me
End Sub

Private Sub Mrefresh_Click()
  RefreshF5
End Sub

Private Sub mregrenam2_Click()
    mregrename_Click
End Sub

Private Sub mregrename_Click()
    Dim i As Long, vnb As Long
    On Error Resume Next
    FRegRename.Show 1   ' Affichage de la fenêtre
    
    If LOk = False Then ' L'utilisateur a abandonné
        Exit Sub
    End If
    
    RENAME.MousePointer = 11

    If LOption1 = 1 Then ' All
        vnb = ListView1.ListItems.Count - 1
        For i = 0 To vnb
            SRegRename i
        Next
    Else    ' Selected
        i = LVGetItemSelected(ListView1, -1)
        While i <> -1
            SRegRename i
            i = LVGetItemSelected(ListView1, i)
        Wend
    End If
    RefreshF5
    If List2.ListCount > 0 Then
        mundo.Enabled = True
    End If
    RENAME.MousePointer = 0
End Sub

Private Sub mremdisp_Click()
' Remove from display
Dim vnb As Long
Dim vnb2 As Long
Dim R As Boolean, i As Long
vnb = Val(état.Panels(3).Text)
vnb2 = ListView1.ListItems.Count
For i = vnb2 To 0 Step -1
  If LVIsSelected(ListView1, i) = True Then
   R = LVRemoveItem(ListView1, i)
  vnb = vnb - 1
 End If
Next
état.Panels(3).Text = Trim$(Str$(vnb))
état.Panels(4).Text = "0"
End Sub

Private Sub mrendirect_Click()
Dim NewName As String
Dim fileop As New CSHFileOp
With fileop
    .ParentWnd = hWnd
    .ClearSourceFiles
    .ClearDestFiles
    .ConfirmOperation = False
End With

NewName = InputBox("Select a new name for " + Dir1Path, "Rename directory", Dir1Path)
If Trim$(NewName) = "" Then
 Exit Sub
End If
fileop.AddSourceFile Dir1Path
If LesOptions.RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
 NewName = RemIllegals(NewName, True)
End If
If LesOptions.RemoveStartingSpaces = 1 Then ' Il faut supprimer les espaces en début de fichier
 NewName = LTrim$(NewName)
End If

fileop.AddDestFile NewName
fileop.RenameFiles
remplissage
End Sub
Private Sub mrnclipboard_Click()
Dim i As Long, chaine As String, vnb As Long
Dim NewName As String, OldName As String
On Error Resume Next
chaine = ""
RENAME.MousePointer = 11
vnb = ListView2.ListItems.Count - 1
For i = 0 To vnb
  NewName = LVGetName(ListView2, i)
  OldName = LVGetItemName(ListView2, i, 1)
  chaine = chaine + NewName + vbTab + OldName + vbCrLf
Next
Clipboard.Clear
Clipboard.SetText chaine
chaine = ""
RENAME.MousePointer = 0
End Sub
Private Sub mrnremoveall_Click()
 ListView2.ListItems.Clear
End Sub

Private Sub mrules_Click()
    FRules.Show vbModal
End Sub

Private Sub msave_Click()
Dim szFilename As String
Dim RepCourant As String
On Error GoTo errsaveset

If Len(Trim$(NomSettings)) = 0 Then
 If Len(Trim$(LesOptions.SettingsDirectory)) > 0 Then
  szFilename = DialogFile(Me.hWnd, 2, "Save settings as", "settings.ren", "Rename" & Chr$(0) & "*.ren" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "ren")
 Else
  szFilename = DialogFile(Me.hWnd, 2, "Save settings as", "settings.ren", "Rename" & Chr$(0) & "*.ren" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "ren")
 End If

 If Trim$(szFilename) = "" Then
  LeCancel = True
  Exit Sub
 End If
Else
    LeCancel = False
    szFilename = NomSettings
End If
If InStr(szFilename, "autosave.ren") = 0 Then
    If MsgBox("Would you like to restore the current directory when you will open this setting file later ?", vbYesNo, "Current folder") = vbNo Then
        RepCourant = "<DoNotRestore>"
    Else
        RepCourant = Dir1Path
    End If
Else
    RepCourant = "<DoNotRestore>"
End If

' Préférences "Générales"
 SavSet "General", szFilename, "Version", "2.0"
 SavSet "General", szFilename, "Recursive", Toolbar1.Buttons(13).Value
 SavSet "General", szFilename, "txtlang", txtlang.Text
 SavSet "General", szFilename, "Combo1", Combo1.ListIndex
 SavSet "General", szFilename, "CmdTxt1", cmdtxt1.Text
 SavSet "General", szFilename, "CmdTxt2", cmdtxt2.Text
 SavSet "General", szFilename, "Txt9", Text9.Text
 SavSet "General", szFilename, "Check7", Check7.Value
 SavSet "General", szFilename, "Folder1", Folder1
 SavSet "General", szFilename, "Folder2", Folder2
 SavSet "General", szFilename, "Folder3", Folder3
 SavSet "General", szFilename, "Folder4", Folder4
 SavSet "General", szFilename, "Folder5", Folder5
 SavSet "General", szFilename, "Folder6", Folder6
 SavSet "General", szFilename, "FolderOk", FolderOk
 SavSet "General", szFilename, "Option3_9", Option3(9).Value
 SavSet "General", szFilename, "Option3_10", Option3(10).Value
 SavSet "General", szFilename, "Option3_11", Option3(11).Value
 SavSet "General", szFilename, "Option1_0", Option1(0).Value
 SavSet "General", szFilename, "Text2", Text2.Text
 SavSet "General", szFilename, "Option1_1", Option1(1).Value
 SavSet "General", szFilename, "Text14", Text14.Text
 SavSet "General", szFilename, "Option2_0", Option2(0).Value
 SavSet "General", szFilename, "Option2_1", Option2(1).Value
 SavSet "General", szFilename, "Check5", Check5.Value
 SavSet "General", szFilename, "Option3_3", Option3(3).Value
 SavSet "General", szFilename, "Option3_4", Option3(4).Value
 SavSet "General", szFilename, "Option3_5", Option3(5).Value
 SavSet "General", szFilename, "Check6", Check6.Value
 SavSet "General", szFilename, "Option3_8", Option3(8).Value
 SavSet "General", szFilename, "Option3_7", Option3(7).Value
 SavSet "General", szFilename, "Option3_6", Option3(6).Value
 SavSet "General", szFilename, "Check3", Check3.Value
 SavSet "General", szFilename, "Text3", Text3.Text
 SavSet "General", szFilename, "Text4", Text4.Text
 SavSet "General", szFilename, "Text5", Text5.Text
 SavSet "General", szFilename, "Combo3", Combo3.ListIndex
 SavSet "General", szFilename, "Option3_0", Option3(0).Value
 SavSet "General", szFilename, "Option3_1", Option3(1).Value
 SavSet "General", szFilename, "Option3_2", Option3(2).Value
 SavSet "General", szFilename, "Check1", Check1.Value
 SavSet "General", szFilename, "Option3_17", Option3(17).Value
 SavSet "General", szFilename, "Option3_16", Option3(16).Value
 SavSet "General", szFilename, "Option3_15", Option3(15).Value

' Extension ***********************
 SavSet "General", szFilename, "Combo2", Combo2.ListIndex
 SavSet "General", szFilename, "Check11", Check11.Value
 SavSet "General", szFilename, "Text16", Text16.Text
 SavSet "General", szFilename, "Text17", Text17.Text
 SavSet "General", szFilename, "Text18", Text18.Text
 SavSet "General", szFilename, "Combo4", Combo4.ListIndex
 SavSet "General", szFilename, "Option3_26", Option3(26).Value
 SavSet "General", szFilename, "Option3_25", Option3(25).Value
 SavSet "General", szFilename, "Option3_24", Option3(24).Value
 SavSet "General", szFilename, "Check4", Check4.Value
 SavSet "General", szFilename, "Option3_14", Option3(14).Value
 SavSet "General", szFilename, "Option3_13", Option3(13).Value
 SavSet "General", szFilename, "Option3_12", Option3(12).Value
 SavSet "General", szFilename, "Option4_0", Option4(0).Value
 SavSet "General", szFilename, "Text8", Text8.Text
 SavSet "General", szFilename, "Option4_1", Option4(1).Value
 SavSet "General", szFilename, "Text15", Text15.Text
 SavSet "General", szFilename, "Option5_0", Option5(0).Value
 SavSet "General", szFilename, "Option5_1", Option5(1).Value
 SavSet "General", szFilename, "Check12", Check12.Value
 SavSet "General", szFilename, "Option3_29", Option3(29).Value
 SavSet "General", szFilename, "Option3_28", Option3(28).Value
 SavSet "General", szFilename, "Option3_27", Option3(27).Value
 SavSet "General", szFilename, "Check13", Check13.Value
 SavSet "General", szFilename, "Option3_30", Option3(30).Value
 SavSet "General", szFilename, "Option3_31", Option3(31).Value
 SavSet "General", szFilename, "Option3_32", Option3(32).Value

' Attributs
 SavSet "Attributs", szFilename, "Archive", Attr1.Archive
 SavSet "Attributs", szFilename, "System", Attr1.System
 SavSet "Attributs", szFilename, "ReadOnly", Attr1.ReadOnly
 SavSet "Attributs", szFilename, "Hidden", Attr1.Hidden
 SavSet "Attributs", szFilename, "SetArchive", Attr1.SETArchive
 SavSet "Attributs", szFilename, "SetSystem", Attr1.SETSystem
 SavSet "Attributs", szFilename, "SetReadOnly", Attr1.SETReadOnly
 SavSet "Attributs", szFilename, "SetHidden", Attr1.SETHidden
 SavSet "Attributs", szFilename, "UseAttr", Attr1.AtrOk

' Date and Time
 SavSet "Date and Time", szFilename, "UseDateTime", DT1.DTOk
 SavSet "Date and Time", szFilename, "CreateDOption", DT1.CreateDOption
 SavSet "Date and Time", szFilename, "CreateDFixed", DT1.CreateDFixed
 SavSet "Date and Time", szFilename, "CreateDInc1", DT1.CreateDInc1
 SavSet "Date and Time", szFilename, "CreateDInc2", DT1.CreateDInc2
 SavSet "Date and Time", szFilename, "CreateDInc3", DT1.CreateDInc3
 SavSet "Date and Time", szFilename, "CreateTOption", DT1.CreateTOption
 SavSet "Date and Time", szFilename, "CreateTFixed", DT1.CreateTFixed
 SavSet "Date and Time", szFilename, "AccessDOption", DT1.AccessDOption
 SavSet "Date and Time", szFilename, "AccessDFixed", DT1.AccessDFixed
 SavSet "Date and Time", szFilename, "AccessDInc1", DT1.AccessDInc1
 SavSet "Date and Time", szFilename, "AccessDInc2", DT1.AccessDInc2
 SavSet "Date and Time", szFilename, "AccessDInc3", DT1.AccessDInc3
 SavSet "Date and Time", szFilename, "AccessTOption", DT1.AccessTOption
 SavSet "Date and Time", szFilename, "AccessTFixed", DT1.AccessTFixed
 SavSet "Date and Time", szFilename, "WriteDOption", DT1.WriteDOption
 SavSet "Date and Time", szFilename, "WriteDFixed", DT1.WriteDFixed
 SavSet "Date and Time", szFilename, "WriteDInc1", DT1.WriteDInc1
 SavSet "Date and Time", szFilename, "WriteDInc2", DT1.WriteDInc2
 SavSet "Date and Time", szFilename, "WriteDInc3", DT1.WriteDInc3
 SavSet "Date and Time", szFilename, "WriteTOption", DT1.WriteTOption
 SavSet "Date and Time", szFilename, "WriteTFixed", DT1.WriteTFixed

' Lecteur et répertoire
 SavSet "Folder", szFilename, "CurrentFolder", RepCourant

' Recherche et remplacement
 ' Recherche globale
 SavSet "Global Search and Replace", szFilename, "SearchString", rech3.SearchString
 SavSet "Global Search and Replace", szFilename, "ReplaceString", rech3.ReplaceString
 SavSet "Global Search and Replace", szFilename, "ReplaceAll", rech3.ReplaceAll
 SavSet "Global Search and Replace", szFilename, "MatchCase", rech3.MatchCase
 SavSet "Global Search and Replace", szFilename, "SearchFromLeft", rech3.SearchFromleft
 SavSet "Global Search and Replace", szFilename, "SearchFrom", rech3.SearchFrom
 SavSet "Global Search and Replace", szFilename, "SearchTo", rech3.SearchTo
 SavSet "Global Search and Replace", szFilename, "SearchCharacters", rech3.SearchCharacters
 SavSet "Global Search and Replace", szFilename, "ReplaceCharacters", rech3.ReplaceCharacters
 SavSet "Global Search and Replace", szFilename, "UseRegExp", rech3.UseRegExp
 
' Dans le préfixe
 SavSet "Search and Replace in Prefix", szFilename, "SearchString", rech1.SearchString
 SavSet "Search and Replace in Prefix", szFilename, "ReplaceString", rech1.ReplaceString
 SavSet "Search and Replace in Prefix", szFilename, "ReplaceAll", rech1.ReplaceAll
 SavSet "Search and Replace in Prefix", szFilename, "MatchCase", rech1.MatchCase
 SavSet "Search and Replace in Prefix", szFilename, "SearchFromLeft", rech1.SearchFromleft
 SavSet "Search and Replace in Prefix", szFilename, "SearchFrom", rech1.SearchFrom
 SavSet "Search and Replace in Prefix", szFilename, "SearchTo", rech1.SearchTo
 SavSet "Search and Replace in Prefix", szFilename, "SearchCharacters", rech1.SearchCharacters
 SavSet "Search and Replace in Prefix", szFilename, "ReplaceCharacters", rech1.ReplaceCharacters
 SavSet "Search and Replace in Prefix", szFilename, "UseRegExp", rech1.UseRegExp

' Dans l'extension
 SavSet "Search and Replace in Extension", szFilename, "SearchString", rech2.SearchString
 SavSet "Search and Replace in Extension", szFilename, "ReplaceString", rech2.ReplaceString
 SavSet "Search and Replace in Extension", szFilename, "ReplaceAll", rech2.ReplaceAll
 SavSet "Search and Replace in Extension", szFilename, "MatchCase", rech2.MatchCase
 SavSet "Search and Replace in Extension", szFilename, "SearchFromLeft", rech2.SearchFromleft
 SavSet "Search and Replace in Extension", szFilename, "SearchFrom", rech2.SearchFrom
 SavSet "Search and Replace in Extension", szFilename, "SearchTo", rech2.SearchTo
 SavSet "Search and Replace in Extension", szFilename, "SearchCharacters", rech2.SearchCharacters
 SavSet "Search and Replace in Extension", szFilename, "ReplaceCharacters", rech2.ReplaceCharacters
 SavSet "Search and Replace in Extension", szFilename, "UseRegExp", rech2.UseRegExp

' Général pour la recherche et le remplacement
 SavSet "Search and Replace", szFilename, "UseSearchPref", RechPref
 SavSet "Search and Replace", szFilename, "UseSearchSuff", RechSuff
 SavSet "Search and Replace", szFilename, "UseGlobalSearch", RechGlob

' Infos sur les MP3
 SavSet "MP3", szFilename, "Use", UseMP3
 SavSet "MP3", szFilename, "Rule", MusMP3.Rule
 SavSet "MP3", szFilename, "PlaceWhereToPut", MusMP3.PlaceWhereToPut
 SavSet "MP3", szFilename, "DefaultArtistToUse", MusMP3.DefaultArtistToUse
 SavSet "MP3", szFilename, "DefaultYearToUse", MusMP3.DefaultYearToUse
 SavSet "MP3", szFilename, "DefaultGenreToUse", MusMP3.DefaultGenreToUse
 SavSet "MP3", szFilename, "DefaultAlbumToUse", MusMP3.DefaultAlbumToUse
 SavSet "MP3", szFilename, "DefaultTitleToUse", MusMP3.DefaultTitleToUse

' Infos sur les OGG
 SavSet "OGG", szFilename, "Use", UseOGG
 SavSet "OGG", szFilename, "Rule", MusOgg.Rule
 SavSet "OGG", szFilename, "PlaceWhereToPut", MusOgg.PlaceWhereToPut
 SavSet "OGG", szFilename, "DefaultAlbumToUse", MusOgg.DefaultAlbumToUse
 SavSet "OGG", szFilename, "DefaultArtistToUse", MusOgg.DefaultArtistToUse
 SavSet "OGG", szFilename, "DefaultGenreToUse", MusOgg.DefaultGenreToUse
 SavSet "OGG", szFilename, "DefaultTitleToUse", MusOgg.DefaultTitleToUse
 
' Infos sur les VQF
 SavSet "VQF", szFilename, "Use", UseVQF
 SavSet "VQF", szFilename, "PlaceWhereToPut", MusVQF.PlaceWhereToPut
 SavSet "VQF", szFilename, "DefaultArtistToUse", MusVQF.DefaultArtistToUse
 SavSet "VQF", szFilename, "DefaultTitle", MusVQF.DefaultTitle
 SavSet "VQF", szFilename, "Rule", MusVQF.Rule
 
' Infos sur les tags EXIF
 SavSet "EXIF", szFilename, "Use", PicEXIF.UseEXIF
 SavSet "EXIF", szFilename, "Rule", PicEXIF.Rule
 SavSet "EXIF", szFilename, "PlaceWhereToPut", PicEXIF.PlaceWhereToPut

RefreshF5
état.Panels(1).Text = "Settings saved"
Exit Sub

errsaveset:
 MsgBox "An error was detecting while saving. Sorry."
 
End Sub

Private Sub msaveas_Click()
Dim nomsave As String
nomsave = NomSettings
NomSettings = ""
msave_Click
If LeCancel = True Then
 NomSettings = nomsave
Else
 msave.Enabled = True
End If
End Sub

Private Sub msaveonlynewnames_Click()
Dim szFilename As String, sItem1 As String
Dim i As Long, vnb As Long
Dim ff As Integer
szFilename = DialogFile(Me.hWnd, 2, "Save as", "rename.list", "Text" & Chr$(0) & "*.list" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "list")

RENAME.MousePointer = 11
If Trim$(szFilename) = "" Then
 RENAME.MousePointer = 0
 Exit Sub
End If
ff = FreeFile
Open szFilename For Output As #ff
ListView2.Visible = False
vnb = ListView2.ListItems.Count - 1
For i = 0 To vnb
 sItem1 = LVGetName(ListView2, i)
 Print #ff, sItem1
Next
Close #ff
RENAME.MousePointer = 0
ListView2.Visible = True
Beep
End Sub

Private Sub msearch_Click()
 Dim vretour As String
 Dim vnb2 As Long
 vretour = ""
 vretour = InputBox("Enter file to search", "Search", "")
 If vretour = "" Then
  Exit Sub
 End If
 vnb2 = LVSearch(ListView1, vretour)
 If vnb2 <> -1 Then
  ListView1.ListItems(vnb2 + 1).Selected = True
  ListView1.ListItems(vnb2 + 1).EnsureVisible
 Else
  MsgBox "File not found"
 End If
End Sub

Private Sub msearchext_Click()
    If RechGlob = True Then
        If MsgBox("Warning, you are already using global search and replace. You can't use both, search and replace in the extension and global search and replace. Do you want to deactivate global search and replace ?", vbYesNo, "Warning !") = vbNo Then
            Exit Sub
        End If
    End If
    RechGlob = False
    RechEnCours = 2
    advsearch.Show 1
End Sub

Private Sub msearchpref_Click()
    If RechGlob = True Then
        If MsgBox("Warning, you are already using global search and replace. You can't use both, search and replace in prefix and global search and replace. Do you want to deactivate global search and replace ?", vbYesNo, "Warning !") = vbNo Then
            Exit Sub
        End If
    End If
    RechGlob = False
    RechEnCours = 1
    advsearch.Show 1
End Sub
Private Sub msetvolabel_Click()
 fsetvol.Show 1
End Sub

Private Sub mshformat_Click()
 fformat.Show 1
End Sub
Private Sub mstartup_Click()
 LesOptions.StartupDir = Dir1Path
End Sub

Private Sub mundo_Click()
 Dim vnb As Integer, vnom1 As String, vnom2 As String, i As Integer, vnb2 As Long
 Dim fileop As New CSHFileOp, ff As Integer
 With fileop
    .ParentWnd = hWnd
    .ClearSourceFiles
    .ClearDestFiles
    .AllowUndo = False
    .ConfirmOperation = False
 End With
 vnb = List2.ListCount
 RENAME.MousePointer = 11
 
 For i = 0 To vnb - 1
    vnom1 = Trim$(List2.List(i)) ' Nom d'origine
    vnom2 = Trim$(List3.List(i)) ' Nouveau nom
    état.Panels(1).Text = "(UNDO) Rename " + vnom2 + " => " + vnom1
    état.Panels(2).Text = Trim$(Str$(i + 1)) + "/" + Trim$(Str$(vnb))
    With fileop
        .AddSourceFile vnom2
        .AddDestFile vnom1
        .RenameFiles
        .ClearSourceFiles
        .ClearDestFiles
    End With
 Next
 
On Error GoTo ErrUndo
If Len(LesOptions.UndoFile) > 0 Then
    ff = FreeFile
    Open LesOptions.UndoFile For Input As #ff
    Close #ff
    Dim vretour As Integer
    vretour = MsgBox("An undo file =>" + LesOptions.UndoFile + "<= have been created, would you like to delete it ?", vbOKCancel, "Delete the undo file ?")
    If vretour = vbOK Then
        Kill LesOptions.UndoFile
    End If
End If

ErrUndo:
 état.Panels(1).Text = "(UNDO) Ok !"
 état.Panels(2).Text = ""
 vnb2 = remplissage()
 état.Panels(3).Text = Trim$(Str$(vnb2))
 état.Panels(4).Text = "0"
 RENAME.MousePointer = 0
 List2.Clear
 List3.Clear
 mundo.Enabled = False
End Sub
Private Sub mviewbag_Click()
 fBag.Show 1
End Sub

Private Sub mviewlog_Click()
    FViewLog.Show 1
End Sub

Private Sub mviewmp3tab_Click()
    MTetat1
End Sub

Private Sub mviewpicturetab_Click()
    MTetat2
End Sub

Private Sub mviewtexttab_Click()
    MTetat3
End Sub
Private Sub Option1_Click(Index As Integer)
 If Index = 0 Then
  Text2.Visible = True
  Command8.Top = 495
  Command8.Visible = True
  Text14.Visible = False
  Option2(0).Visible = False
  Option2(1).Visible = False
 Else
  Text2.Visible = False
  Text14.Visible = True
  Command8.Top = 850
  Command8.Visible = True
  Option2(0).Visible = True
  Option2(1).Visible = True
 End If
End Sub
Private Sub Option4_Click(Index As Integer)
 If Index = 0 Then
  Text8.Visible = True
  Text15.Visible = False
  Option5(0).Visible = False
  Option5(1).Visible = False
  Text8.SetFocus
 Else
  Text8.Visible = False
  Text15.Visible = True
  Option5(0).Visible = True
  Option5(1).Visible = True
  Text15.SetFocus
 End If
End Sub
Function remplissage() As Long
On Error GoTo ErrGen
 Dim extrait As String, i As Integer, vnb As Integer
 Dim nbfichiers As Long, clsFind As New clsFindFile
 Dim strFile As String, chemin As String, attributs As Long
 Dim chaine As String, afficher As Boolean
 Dim colonne As ColumnHeader
 Dim j As Integer
 Set colonne = ListView1.ColumnHeaders.Item(1)
 RENAME.MousePointer = 11
 If Recursive = True Then  ' Mode récursif ***************************************************************************
    remplissage = srecursive()
    If LesOptions.AutoArrange = True Then
        For j = 1 To 5
            Set colonne = ListView1.ColumnHeaders.Item(j)
            AutoSizeColumnHeader ListView1, colonne, True
        Next
    End If
    RENAME.MousePointer = 0
    If ListView1.ListItems.Count > 0 Then
        mregrenam2.Enabled = True
    Else
        mregrenam2.Enabled = False
    End If
    If LesOptions.SelectAllFiles = 1 Then
        SelectAll
    End If
    Exit Function
 End If
 
 If Rafraichir = False Then
  Rafraichir = True
  Exit Function
 End If
 clsFind.Dateformat = "short Date"
 ListView1.ListItems.Clear
 nbfichiers = 0
 chemin = AddBackSlash(Trim$(Dir1Path))
 Filtre = Trim$(Filtre)
 ' Suppression des caractères en trop
 If Right$(Filtre, 1) = ";" Then
  Filtre = Left$(Filtre, Len(Filtre) - 1)
 End If
 If Left$(Filtre, 1) = ";" Then
  Filtre = Mid$(Filtre, 2)
 End If
 ' normalement la condition de filtre est bonne
 vnb = CharOccurs(Filtre, ";")
 vnb = vnb + 1
 
 For i = 1 To vnb
    extrait = GetToken(Filtre, ";", i)
    strFile = clsFind.Find(chemin & extrait, True)
        Do While Len(strFile)
            afficher = True
            Select Case LesOptions.FilesToInclude
                Case 0 ' Files Only
                    If (clsFind.FileAttributes And vbDirectory) = 0 Then ' Si ce n'est pas un répertoire
                        If InStr(Suffixe(extrait), "*") = 0 Then  ' Test de correspondance sur l'intégralité du masque de sélection
                            If UCase$(Suffixe(strFile)) <> UCase$(Suffixe(extrait)) Then
                                afficher = False
                            End If
                        End If
                        attributs = clsFind.FileAttributes
                        chaine = ""
                        If (attributs And FILE_ATTRIBUTE_READONLY) And LesOptions.ReadOnly = False Then afficher = False
                        If (attributs And FILE_ATTRIBUTE_HIDDEN) And LesOptions.Hidden = False Then afficher = False
                        If (attributs And FILE_ATTRIBUTE_SYSTEM) And LesOptions.System = False Then afficher = False
                        If afficher = True Then
                            If attributs And FILE_ATTRIBUTE_READONLY Then chaine = "R"
                            If attributs And FILE_ATTRIBUTE_HIDDEN Then chaine = chaine + "H"
                            If attributs And FILE_ATTRIBUTE_SYSTEM Then chaine = chaine + "S"
                            If attributs And FILE_ATTRIBUTE_ARCHIVE Then chaine = chaine + "A"
                            If chaine = "" Then
                                chaine = " "
                            End If
                            nbfichiers = nbfichiers + 1
                            Set itmX = ListView1.ListItems.Add(, , strFile)
                            itmX.Text = strFile
                            itmX.SubItems(1) = clsFind.FileSize
                            Select Case LesOptions.Dateformat
                                Case 0
                                    itmX.SubItems(2) = clsFind.GetCreationDate
                                Case 1
                                    itmX.SubItems(2) = clsFind.GetLastWriteDate
                                Case 2
                                    itmX.SubItems(2) = clsFind.GetLastAccessDate
                            End Select
                            itmX.SubItems(3) = chaine
                            itmX.SubItems(4) = "File"  ' Type fichier
                        End If
                    End If
                
                Case 1 ' Files and Folders
                    If InStr(Suffixe(extrait), "*") = 0 Then  ' Test de correspondance sur l'intégralité du masque de sélection
                        If UCase$(Suffixe(strFile)) <> UCase$(Suffixe(extrait)) Then
                            afficher = False
                        End If
                    End If
                    attributs = clsFind.FileAttributes
                    chaine = ""
                    If Trim$(strFile) = "." Or Trim$(strFile) = ".." Then afficher = False
                    If (attributs And FILE_ATTRIBUTE_READONLY) And LesOptions.ReadOnly = False Then afficher = False
                    If (attributs And FILE_ATTRIBUTE_HIDDEN) And LesOptions.Hidden = False Then afficher = False
                    If (attributs And FILE_ATTRIBUTE_SYSTEM) And LesOptions.System = False Then afficher = False
                    If afficher = True Then
                        If attributs And FILE_ATTRIBUTE_READONLY Then chaine = "R"
                        If attributs And FILE_ATTRIBUTE_HIDDEN Then chaine = chaine + "H"
                        If attributs And FILE_ATTRIBUTE_SYSTEM Then chaine = chaine + "S"
                        If attributs And FILE_ATTRIBUTE_ARCHIVE Then chaine = chaine + "A"
                        If chaine = "" Then
                            chaine = " "
                        End If
                        nbfichiers = nbfichiers + 1
                        If (clsFind.FileAttributes And vbDirectory) <> 0 Then ' Si c'est un répertoire
                            Set itmX = ListView1.ListItems.Add(, , strFile)
                            itmX.Bold = True
                            itmX.SubItems(4) = "Dir"  ' Type répertoire
                        Else ' Ce n'est pas un répertoire
                            Set itmX = ListView1.ListItems.Add(, , strFile)
                            itmX.SubItems(4) = "File"  ' Type fichier
                        End If
                        itmX.Text = strFile
                        itmX.SubItems(1) = clsFind.FileSize
                        Select Case LesOptions.Dateformat
                            Case 0
                                itmX.SubItems(2) = clsFind.GetCreationDate
                            Case 1
                                itmX.SubItems(2) = clsFind.GetLastWriteDate
                            Case 2
                                itmX.SubItems(2) = clsFind.GetLastAccessDate
                        End Select
                        itmX.SubItems(3) = chaine
                    End If
                
                Case 2 ' Folders only
                    If (clsFind.FileAttributes And vbDirectory) <> 0 Then ' Si ce n'est pas un répertoire
                        If InStr(Suffixe(extrait), "*") = 0 Then  ' Test de correspondance sur l'intégralité du masque de sélection
                            If UCase$(Suffixe(strFile)) <> UCase$(Suffixe(extrait)) Then
                                afficher = False
                            End If
                        End If
                        attributs = clsFind.FileAttributes
                        chaine = ""
                        If Trim$(strFile) = "." Or Trim$(strFile) = ".." Then afficher = False
                        If (attributs And FILE_ATTRIBUTE_READONLY) And LesOptions.ReadOnly = False Then afficher = False
                        If (attributs And FILE_ATTRIBUTE_HIDDEN) And LesOptions.Hidden = False Then afficher = False
                        If (attributs And FILE_ATTRIBUTE_SYSTEM) And LesOptions.System = False Then afficher = False
                        If afficher = True Then
                            If attributs And FILE_ATTRIBUTE_READONLY Then chaine = "R"
                            If attributs And FILE_ATTRIBUTE_HIDDEN Then chaine = chaine + "H"
                            If attributs And FILE_ATTRIBUTE_SYSTEM Then chaine = chaine + "S"
                            If attributs And FILE_ATTRIBUTE_ARCHIVE Then chaine = chaine + "A"
                            If chaine = "" Then
                                chaine = " "
                            End If
                            nbfichiers = nbfichiers + 1
                            Set itmX = ListView1.ListItems.Add(, , strFile)
                            itmX.Bold = True
                            itmX.Text = strFile
                            itmX.SubItems(1) = clsFind.FileSize
                            Select Case LesOptions.Dateformat
                                Case 0
                                    itmX.SubItems(2) = clsFind.GetCreationDate
                                Case 1
                                    itmX.SubItems(2) = clsFind.GetLastWriteDate
                                Case 2
                                    itmX.SubItems(2) = clsFind.GetLastAccessDate
                            End Select
                            itmX.SubItems(3) = chaine
                            itmX.SubItems(4) = "Dir"  ' Type répertoire
                        End If
                    End If
               
            End Select
   
            strFile = clsFind.FindNext()
        Loop
 Next
RENAME.MousePointer = 0
If LesOptions.AutoArrange = True Then
    For j = 1 To 5
        Set colonne = ListView1.ColumnHeaders.Item(j)
        AutoSizeColumnHeader ListView1, colonne, True
    Next
End If
If ListView1.ListItems.Count > 0 Then
   mregrenam2.Enabled = True
Else
   mregrenam2.Enabled = False
End If

If nbfichiers > 0 Then
 ListView1.ListItems(1).EnsureVisible
End If
If LesOptions.SelectAllFiles = 1 Then
    SelectAll
End If
remplissage = nbfichiers
Exit Function

ErrGen:
ErreurGrave "Remplissage"
End Function

Private Sub ModifierLigne()
    Dim chnmodifier As String
    Dim vrai As Boolean
    vrai = False
    If Trim$(Combo1.List(Combo1.ListIndex)) <> "Modify prefix" Or Trim$(Combo2.List(Combo2.ListIndex)) <> "Modify extension" Then
        état.Panels(1).Text = ""
        Exit Sub
    End If
 
    chnmodifier = "Prefix: "
    If Trim$(Combo1.List(Combo1.ListIndex)) = "Modify prefix" Then
        ' 1-Préfixe
        If Check3.Value = 1 Then
            chnmodifier = chnmodifier + "Counter/"
            vrai = True
        End If
        If Check5.Value = 1 Then
            chnmodifier = chnmodifier + "Size/"
            vrai = True
        End If
        If Check6.Value = 1 Then
            chnmodifier = chnmodifier + "Date/"
            vrai = True
        End If
        If Check7.Value = 1 Then
            chnmodifier = chnmodifier + "Time/"
            vrai = True
        End If
        If Check1.Value = 1 Then
            chnmodifier = chnmodifier + "Picture/"
            vrai = True
        End If
        If Len(Trim$(Text2.Text)) > 0 And Option1(0).Value = True Then
            chnmodifier = chnmodifier + "Repl txt/"
            vrai = True
        End If
        If Len(Trim$(Text14.Text)) > 0 And Option1(1).Value = True Then
            chnmodifier = chnmodifier + "Add txt/"
            vrai = True
        End If
    End If
    If Not vrai Then chnmodifier = chnmodifier + "<nothing>"
    vrai = False
 
    ' 2-Extension
    If Right$(chnmodifier, 1) = "/" Then chnmodifier = Left$(chnmodifier, Len(chnmodifier) - 1)

    chnmodifier = chnmodifier + " | Extension: "
    If Trim$(Combo2.List(Combo2.ListIndex)) = "Modify extension" Then
        If Check11.Value = 1 Then
            chnmodifier = chnmodifier + "Counter/"
            vrai = True
        End If
        If Check12.Value = 1 Then
            chnmodifier = chnmodifier + "Size/"
            vrai = True
        End If
        If Check13.Value = 1 Then
            chnmodifier = chnmodifier + "Date/"
            vrai = True
        End If
        If Check4.Value = 1 Then
            chnmodifier = chnmodifier + "Time/"
            vrai = True
        End If
 
        If Len(Trim$(Text8.Text)) > 0 And Option4(0).Value = True Then
            chnmodifier = chnmodifier + "Repl txt/"
            vrai = True
        End If
        If Len(Trim$(Text15.Text)) > 0 And Option4(1).Value = True Then
            chnmodifier = chnmodifier + "Add txt/"
            vrai = True
        End If
    End If
    If Not vrai Then chnmodifier = chnmodifier + "<nothing>"
    If Right$(chnmodifier, 1) = "/" Then
        chnmodifier = Left$(chnmodifier, Len(chnmodifier) - 1)
        vrai = True
    End If
    état.Panels(1).Text = chnmodifier
End Sub
Private Sub Text1_DblClick()
Dim chemin As String
chemin = AddBackSlash(Trim$(Dir1Path))
FileExecutor Me.hWnd, chemin + ListView1.ListItems(ListView1.SelectedItem.Index), "Open"
Text1.Text = ""
LoadText chemin + ListView1.ListItems(ListView1.SelectedItem.Index)
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrGen
Dim vtmp As String
If Recursive = False Then
    vtmp = AddBackSlash(Dir1Path)
Else
    vtmp = ""
End If
LoadText vtmp & ListView1.SelectedItem.Text
Exit Sub

ErrGen:
ErreurGrave "Text1_OLEDragDrop"

End Sub
Private Sub Text11_Validate(Cancel As Boolean)
LesOptions.PicturesFormat = Text11.Text
End Sub

Private Sub Text14_Change()
 CharInterdits Text14.Text
 ModifierLigne
End Sub

Private Sub Text14_GotFocus()
    SelAll Text14
End Sub

Private Sub Text15_Change()
 CharInterdits Text15.Text
 ModifierLigne
End Sub

Private Sub Text15_GotFocus()
SelAll Text15
End Sub
Private Sub Text2_Change()
 CharInterdits Text2.Text
 ModifierLigne
End Sub

Private Sub Text2_GotFocus()
SelAll Text2
End Sub
Private Sub Text8_Change()
 CharInterdits Text8.Text
 ModifierLigne
End Sub

Private Sub Text8_GotFocus()
SelAll Text8
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Dim vnb As Long
 Select Case Button.Index
  Case 2 ' select all
   SelectAll
  Case 3 ' unselect
   Unselect
  Case 4 ' Invert
   InvertSelection
  Case 5 ' step
   StepSelection
  Case 7 ' start
   StartRename
  Case 8 ' preview
   PreviewRename
  Case 9 ' Rename manually
   RenameManually
  Case 11 ' Drop files
   DropFiles
  Case 13 ' Recursive
   If Button.Value = 0 Then
    Recursive = False
    m2recursive.Checked = False
   Else
    Recursive = True
    m2recursive.Checked = True
   End If
   vnb = remplissage()
   état.Panels(3).Text = Trim$(Str$(vnb))
   état.Panels(4).Text = "0"
   
  Case 15 ' Up
   MoveUp
  Case 16 ' Root
   MoveRoot
  Case 18
   If VnbHistory > 0 Then
    PopupMenu mgenhistory, , Button.Left, Toolbar1.Top + Toolbar1.height
   End If
  Case 20 ' Add to your favorites
   madddirectory_Click
  Case 21 'Organize your favorites
   morganyze_Click
  Case 22 ' First favorite
   NavFav 1
  Case 23 ' Previous favorite
   NavFav 2
  Case 24 ' Next favorite
   NavFav 3
  Case 25 ' Last favorite
   NavFav 4
  Case 27 ' Help Pointer
   Me.WhatsThisMode
 End Select
End Sub

Private Sub Toolbar1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim fileop As New CSHFileOp
    Dim chemin As String, i As Long, j As Integer, sItem As String, vtempo As String
    With fileop
        .ParentWnd = hWnd
        .ConfirmOperation = LesOptions.ConfirmOperation
        .RenameOnCollision = LesOptions.RenameOnCollision
        .SilentMode = LesOptions.SilentMode
        .AllowUndo = LesOptions.AllowUndo
        .ConfirmMakeDir = LesOptions.ConfirmMakeDir
    End With
  
    If VnbRep > 0 Then
        If Data.GetFormat(vbCFFiles) Then ' Ce sont des fichiers qui proviennent de l'explorateur ou d'une autre fenêtre que celle de THE Rename
            For i = 1 To Data.Files.Count
                For j = 1 To VnbRep
                    With fileop
                        .ClearSourceFiles
                        .ClearDestFiles
                        .AddSourceFile Data.Files(i)
                        .AddDestFile LesRepertoires(j)
                    End With
                    état.Panels(1).Text = "Copy " + Data.Files(i) + " to " + LesRepertoires(j)
                    vtempo = AddBackSlash(LesRepertoires(j))
                    DoEvents
                    If fileop.CopyFiles Then
                        DT3.SetFileDateTime (vtempo + Prefixe(Data.Files(i)) & "." & Suffixe(Data.Files(i)))
                        Attr3.ChangeAttr (vtempo + Prefixe(Data.Files(i)) & "." & Suffixe(Data.Files(i)))
                    End If
                Next
            Next
        Else
            chemin = AddBackSlash(Dir1Path)
            If Recursive = True Then
                chemin = ""
            End If
            If Screen.ActiveControl.Name = "ListView1" Then
                i = LVGetItemSelected(ListView1, -1)
                While i <> -1
                    sItem = LVGetName(ListView1, i)
                    For j = 1 To VnbRep
                        With fileop
                            .ClearSourceFiles
                            .ClearDestFiles
                            .AddSourceFile chemin & sItem
                            .AddDestFile LesRepertoires(j)
                        End With
                        état.Panels(1).Text = "Copy " + sItem + " to " + LesRepertoires(j)
                        DoEvents
                        vtempo = AddBackSlash(LesRepertoires(j))
                        If fileop.CopyFiles Then
                            DT3.SetFileDateTime (vtempo & sItem)
                            Attr3.ChangeAttr (vtempo & sItem)
                        End If
                    Next
                    i = LVGetItemSelected(ListView1, i)
                Wend
            End If ' Si on est sur le listview
            état.Panels(1).Text = "Ok"
        End If
    End If
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            txtlang.Text = ""
            txtlang.SetFocus
        Case 2
            FCmd.Show 1
            ChargeVNBCommandes
    End Select
End Sub
Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Index
        Case 1  ' Parameters
            FParamCmd.Show 1
    End Select
End Sub

Private Sub TV1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub TV1_Click()
    If Combo6.Visible = True Then
        Combo6.Visible = False
    End If
End Sub

Private Sub TV1_DblClick()
    If Left$(TV1.SelectedItem.Text, 1) = "<" Then
        InsertTextInTextBoxFromText txtlang, TV1.SelectedItem.Text
    End If
End Sub

Private Sub TV1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim X As Long
Dim ID As String
ID = TV1.SelectedItem.Key
If KeyCode = 112 Then ' F1
    If InStr(1, ID, "|") <> 0 And Val(GetToken(ID, "|", 2)) <> 0 Then
        X = WinHelp(Me.hWnd, App.HelpFile, HELP_CONTEXT, Val(GetToken(ID, "|", 2)))
    Else
        If InStr(1, ID, "|") <> 0 And Val(GetToken(ID, "|", 2)) = 0 Then
            MsgBox "Sorry but there is, for this moment, no help for this command...", vbOKOnly, "Free Form Help"
        End If
    End If
End If

End Sub

Private Sub txtlang_Click()
    If Combo6.Visible = True Then
        Combo6.Visible = False
    End If
End Sub
Private Sub txtlang_KeyUp(KeyCode As Integer, Shift As Integer)
' A optimiser
    On Error Resume Next
    Dim tControle As Integer
    Dim phrase As String ' Contient la commande en cours (complète)
    Dim PosDeb As Integer
    Dim PosFin As Integer
    Dim PosSauve As Integer
    Dim extrait As String
    Dim longueur As Integer
    Dim i As Integer
    Dim sValue As String
    Dim chemin As String
    Dim ARemplacer As String
    chemin = AppPath
    chemin = chemin + "commands.ini"
    tControle = GetKeyState(VK_CONTROL)
    If tControle = -127 Or tControle = -128 Then ' La touche controle a été appuyée
        Select Case KeyCode
            Case 8  ' BackSpace
            Case 39 ' Flèche droite
            Case 37 ' Flèche gauche
            
            Case 36 ' Touche Home, première commande
                If MaxCommand = 0 Then
                    Exit Sub
                End If
                CurrentCommand = 1
                With SIni
                    .Path = chemin
                    .Section = "Commands"
                    .Key = "Command" & Trim$(Str$(CurrentCommand))
                    sValue = .Value
                End With
                If Trim$(sValue) <> "" Then
                    txtlang.Text = sValue
                End If
            Case 35 ' Touche End, dernière commande
                If MaxCommand = 0 Then
                    Exit Sub
                End If
                CurrentCommand = MaxCommand
                With SIni
                    .Path = chemin
                    .Section = "Commands"
                    .Key = "Command" & Trim$(Str$(CurrentCommand))
                    sValue = .Value
                End With
                If Trim$(sValue) <> "" Then
                    txtlang.Text = sValue
                End If
            Case 38 ' Flèche vers le haut, commande précédente
                If MaxCommand = 0 Then
                    Exit Sub
                End If
                CurrentCommand = CurrentCommand - 1
                If CurrentCommand < 0 Then ' MaxCommand Then
                    CurrentCommand = MaxCommand
                End If
                With SIni
                    .Path = chemin
                    .Section = "Commands"
                    .Key = "Command" & Trim$(Str$(CurrentCommand))
                    sValue = .Value
                End With
                If Trim$(sValue) <> "" Then
                    txtlang.Text = sValue
                End If
            Case 40 ' Flèche vers le bas, commande suivante
                If MaxCommand = 0 Then
                    Exit Sub
                End If
                CurrentCommand = CurrentCommand + 1
                If CurrentCommand > MaxCommand Then
                    CurrentCommand = 1
                End If
                With SIni
                    .Path = chemin
                    .Section = "Commands"
                    .Key = "Command" & Trim$(Str$(CurrentCommand))
                    sValue = .Value
                End With
                If Trim$(sValue) <> "" Then
                    txtlang.Text = sValue
                End If
            Case 32 ' Barre d'espace
                txtlang.Text = Left$(txtlang.Text, txtlang.SelStart - 1) + Mid$(txtlang.Text, txtlang.SelStart)
                phrase = txtlang.Text
                ' 1) Il faut déterminer sur quel mot on est positionné
                PosDeb = txtlang.SelStart
                PosFin = InStrRev(Left$(phrase, txtlang.SelStart), "<")
                If PosFin = 0 Then
                    Exit Sub
                End If
                Combo6.Clear
                ' on extrait ce qui a été tapé
                extrait = RTrim$(UCase$(Mid$(phrase, PosFin, txtlang.SelStart - PosFin + 1)))
                longueur = Len(extrait)
                For i = 0 To vnbcmd - 1 ' Boucle sur les commandes du langage
                    If UCase$(Left$(listcmd.List(i), longueur)) = extrait Then    ' La commande ressemble à ce qui a été tapé
                        Combo6.AddItem listcmd.List(i)  ' on l'ajoute dans le combo puisqu'il correspond
                    End If
                Next
                If Combo6.ListCount = 0 Then ' La recherche n'a rien donnée
                    Exit Sub    ' On peut s'en aller
                End If
                PosSauve = txtlang.SelStart
                LongSauve = longueur
                If Combo6.ListCount > 1 Then ' Il y a eu plusieurs occurences trouvées, il va falloir l'afficher
                    Combo6.Top = txtlang.Top + txtlang.height + 10
                    LaPosSauve = txtlang.SelStart
                    ARemplacer = extrait
                    Combo6.Visible = True
                    Combo6.ListIndex = 0
                    Combo6.SetFocus
                Else ' Il n'y a eu qu'une seule occurence de trouvée, on la met directement dans le texte
                    txtlang.Text = Left$(phrase, txtlang.SelStart - 1) + Trim$(Mid$(Combo6.List(0), longueur + 1)) + Trim$(Mid$(phrase, txtlang.SelStart))
                    txtlang.SelStart = PosSauve + Len(Trim$(Mid$(Combo6.List(0), longueur + 1))) - 1
                    Combo6.Clear ' et à la fin on efface son contenu (par propreté)
                End If
        End Select
    End If
End Sub
Private Sub txtlang_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And Shift = 1 Then
    ChargeYourCmd
    PopupMenu m3contextuel
    txtlang.Enabled = True
    txtlang.SetFocus
End If
End Sub
Private Function srecursive() As Long
On Error GoTo ErrGen
Dim colFiles As New Collection, colDirs As New Collection
Dim intDirsFound As Integer, vntItem As Variant, nbfichiers As Long
Dim szFilename As String, i As Integer, vnb As Integer, extrait As String
Recursive = True
m2recursive.Checked = True
ListView1.ListItems.Clear
nbfichiers = 0
szFilename = AddBackSlash(Trim$(Dir1Path))
Filtre = Trim$(Filtre)
If Right$(Filtre, 1) = ";" Then
 Filtre = Left$(Filtre, Len(Filtre) - 1)
End If
If Left$(Filtre, 1) = ";" Then
 Filtre = Mid$(Filtre, 2)
End If
vnb = CharOccurs(Filtre, ";")
vnb = vnb + 1
 
 For i = 1 To vnb
  extrait = GetToken(Filtre, ";", i)
  Set colDirs = Nothing
  Set colFiles = Nothing
  Set colDirs = Nothing
  intDirsFound = 0
  colDirs.Add szFilename
  If extrait = "*.*" Then
   intDirsFound = FindAllFiles(szFilename, "*.*", colFiles, colDirs, , True)
  Else
   intDirsFound = FindAllFiles(szFilename, "*.*", , colDirs, True) ' Recherche des répertoires
   For Each vntItem In colDirs
    FindAllFiles CStr(vntItem), extrait, colFiles  ' Recherche des fichiers
   Next
  End If
 Next

If ListView1.ListItems.Count > 0 Then
 ListView1.ListItems(1).EnsureVisible
End If
srecursive = ListView1.ListItems.Count
Exit Function
ErrGen:
ErreurGrave "srecursive"
End Function

Private Sub ChangeRepHistorique(Index As Integer)
 If Trim$(RENAME.mnuhistory(Index).Caption) <> "" Then
  TemMove = True
  FolderTreeview1(0).SelectedFolder = RENAME.mnuhistory(Index).Caption
 End If
End Sub

Private Sub LoadAvailableDrives()
  Dim R As Long, lpBuffer  As String * 256, longueur As Long, i As Integer
  Dim n&
  On Error GoTo ErrGen
  i = 0
  lpBuffer = Space$(256)
  longueur = Len(lpBuffer)
  R = GetLogicalDriveStrings(longueur, lpBuffer)
  If R Then
    Do
     n = InStr(lpBuffer, Chr$(0))
     If n > 1 Then
      If i <> 0 Then Load mnudrives(i)
      mnudrives(i).Caption = Left$(lpBuffer, n - 1)
      i = i + 1
      lpBuffer = Mid$(lpBuffer, n + 1)
     End If
    Loop Until n <= 1
   End If
    
Exit Sub
ErrGen:
ErreurGrave "LoadAvailableDrives"
End Sub

Private Function Bag(action As Integer) As Integer
 Dim sItem As String, actions As String, i As Long, vnb2 As Long, chemin As String, qfaire As String
 Dim fileop As New CSHFileOp
 Dim chemin2 As String, chemin3 As String
 Dim vnb As Long
 With fileop
    .ParentWnd = hWnd
    .ConfirmOperation = LesOptions.ConfirmOperation
    .RenameOnCollision = LesOptions.RenameOnCollision
    .SilentMode = LesOptions.SilentMode
    .AllowUndo = LesOptions.AllowUndo
    .ConfirmMakeDir = LesOptions.ConfirmMakeDir
 End With
 If Piège1 = True Then ' Le truc pour éviter le piège des Ctrl C, Ctrl X, Ctrl V (et autres) quand on est en édition de nom de fichier (à la main) sur le listview principal
    GoTo suite
 End If
 RENAME.MousePointer = 11
 chemin = AddBackSlash(Dir1Path)
 chemin2 = chemin
 If Recursive = True Then
  chemin = ""
 Else
 End If
 
 Select Case action
  Case 1
   actions = "Copy"
  Case 2
   actions = "Add to bin"
  Case 3
   actions = "Cut to bin"
  Case 4
   actions = "Cut additive"
  Case 5
   actions = "Paste"
  Case 6
   actions = "Paste and keep"
 End Select
 
If Screen.ActiveControl.Name = "ListView1" Or Screen.ActiveControl.Name = "FolderTreeview1" Then
 If action = 1 Or action = 3 Then
   ListView3.ListItems.Clear
 End If
End If

Select Case Screen.ActiveControl.Name
 Case "ListView1"
  If action > 0 And action < 5 Then ' Copy ou Cut peut importe que se soit additive ou pas, le ménage est déjà fait
   i = LVGetItemSelected(ListView1, -1)
   While i <> -1
    sItem = LVGetName(ListView1, i)
    Set itmX = ListView3.ListItems.Add(, , actions)
    With itmX
        .Text = actions
        .SubItems(1) = chemin + LVGetName(ListView1, i)
        .SubItems(2) = LVGetItemName(ListView1, i, 1)
        .SubItems(3) = LVGetItemName(ListView1, i, 2)
        .SubItems(4) = LVGetItemName(ListView1, i, 3)
    End With
    i = LVGetItemSelected(ListView1, i)
   Wend
  Else ' coller
   vnb2 = ListView3.ListItems.Count - 1
   For i = 0 To vnb2
    qfaire = Trim$(LVGetName(ListView3, i))
    With fileop
        .ClearSourceFiles
        .ClearDestFiles
        .AddSourceFile LVGetItemName(ListView3, i, 1)
        .AddDestFile chemin2
    End With
    état.Panels(2).Text = Trim$(Str$(i + 1)) + "/" + Trim$(Str$(ListView3.ListItems.Count))
    If qfaire = "Cut to bin" Or qfaire = "Cut additive" Then
     état.Panels(1).Text = "Move " + LVGetItemName(ListView3, i, 1) + " to " + chemin2
     If fileop.MoveFiles Then
      chemin3 = AddBackSlash(chemin2)
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView3, i, 1)) + "." + Suffixe(LVGetItemName(ListView3, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
     End If
    Else
     état.Panels(1).Text = "Copy " + LVGetItemName(ListView3, i, 1) + " to " + chemin2
     If fileop.CopyFiles Then
      chemin3 = AddBackSlash(chemin2)
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView3, i, 1)) + "." + Suffixe(LVGetItemName(ListView3, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
     End If
    End If
   Next
   If action = 5 Then
    ListView3.ListItems.Clear
   End If
   vnb = remplissage()
   With état
        .Panels(1).Text = ""
        .Panels(3).Text = Trim$(Str$(vnb))
        .Panels(4).Text = "0"
    End With
  End If
 
 Case "FolderTreeview1"
  If action > 0 And action < 5 Then ' Copy ou Cut peut importe que se soit additive ou pas, le ménage est déjà fait
    Set itmX = ListView3.ListItems.Add(, , actions)
    itmX.Text = actions
    itmX.SubItems(1) = Dir1Path
  Else ' Coller
   For i = 0 To ListView3.ListItems.Count - 1
    qfaire = Trim$(LVGetName(ListView3, i))
    With fileop
        .ClearSourceFiles
        .ClearDestFiles
        .AddSourceFile LVGetItemName(ListView3, i, 1)
        .AddDestFile chemin2
    End With
    état.Panels(2).Text = Trim$(Str$(i + 1)) + "/" + Trim$(Str$(ListView3.ListItems.Count))
    If qfaire = "Cut to bin" Or qfaire = "Cut additive" Then
     état.Panels(1).Text = "Move " + LVGetItemName(ListView3, i, 1) + " to " + chemin2
     If fileop.MoveFiles Then
      chemin3 = AddBackSlash(chemin2)
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView3, i, 1)) + "." + Suffixe(LVGetItemName(ListView3, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
     End If
    Else
     état.Panels(1).Text = "Copy " + LVGetItemName(ListView3, i, 1) + " to " + chemin2
     If fileop.CopyFiles Then
      chemin3 = AddBackSlash(chemin2)
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView3, i, 1)) + "." + Suffixe(LVGetItemName(ListView3, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
     End If
    End If
   Next
   If action = 5 Then
    ListView3.ListItems.Clear
   End If
   vnb = remplissage()
   With état
        .Panels(1).Text = ""
        .Panels(3).Text = Trim$(Str$(vnb))
        .Panels(4).Text = "0"
    End With
  End If

 Case Else ' On doit être sur un controle autre que le listview ou le foldertreeview
suite:
  Select Case action
   Case 1 ' Copier
    mcopybag.Enabled = False
    SendKeys "^C"
    mcopybag.Enabled = True
   Case 3 ' Couper
    mcutbag.Enabled = False
    SendKeys "^X"
    mcutbag.Enabled = True
   Case 5 ' Coller
    mpastebag.Enabled = False
    SendKeys "^V"
    mpastebag.Enabled = True
  End Select
End Select
RENAME.MousePointer = 0
End Function

Private Sub pDisplayMRU(coche As Boolean)
Dim iFile As Long
    For iFile = 1 To m_cMRU.FileCount
        If (m_cMRU.FileExists(iFile)) Then
            If coche Then
             If iFile = 1 Then mnuFile(iFile + 1).Checked = True
            End If
            mnuFile(iFile + 1).Visible = True
            mnuFile(iFile + 1).Caption = m_cMRU.MenuCaption(iFile)
            mnuFile(iFile + 1).Tag = CStr(iFile)
        End If
    Next
    mnuFile(6).Visible = (m_cMRU.FileCount > 0)
End Sub

Sub NavFav(action As Integer)
' Permet de naviguer de favoris en favoris
Dim i As Integer
On Error GoTo ErrorHandler
Select Case action
 Case 1 ' Premier
  If Len(Trim$(fav(1))) > 0 Then
   FavEncours = 1
  Else
   Exit Sub
  End If
 
 Case 2 ' Précédent
  If FavEncours = -1 Then
   FavEncours = 1
  Else
   If FavEncours - 1 > 0 Then
    FavEncours = FavEncours - 1
   End If
  End If
 
 Case 3 ' Suivant
  If FavEncours + 1 > 20 Then
   FavEncours = 0
  End If
  FavEncours = FavEncours + 1
  If Len(Trim$(fav(FavEncours))) <= 0 Then
   FavEncours = 1
  End If
 
 Case 4 ' Dernier
  FavEncours = -1
  For i = 20 To 1 Step -1
   If Len(Trim$(fav(i))) > 0 Then
    FavEncours = i
    i = 1
   End If
  Next
End Select

If FavEncours <> -1 Then
 If Len(Trim$(fav(FavEncours))) > 0 Then
  ChDir fav(FavEncours)
  TemMove = False
  FolderTreeview1(0).Visible = False
  FolderTreeview1(0).SelectedFolder = fav(FavEncours)
  Dir1Path = fav(FavEncours)
  FolderTreeview1(0).Visible = True
  mundo.Enabled = False
  List2.Clear
  List3.Clear
 End If
End If
Exit Sub

ErrorHandler:   ' Error-handling routine.
Select Case Err.Number  ' Evaluate error number.
 Case 76
  MsgBox "Error - This directory doesn't not exist any more."
 Case 68
  MsgBox "Error - This drive is unavailable."
End Select
Exit Sub

End Sub
Private Sub SavSet(LSection As String, chemin As String, cle As String, valeur As Variant)
Dim ValSave As Variant

ValSave = valeur
If TypeName(valeur) = "String" Then
    ValSave = Replace(valeur, Chr$(32), Chr$(255))
End If
If TypeName(valeur) = "Boolean" Then
    If valeur = True Then
        ValSave = "-1"
    Else
        ValSave = "0"
    End If
End If

With SIni
 .Path = chemin
 .Section = LSection
 .Key = cle
 .Value = ValSave
 If Not (.Success) Then
  MsgBox "Failed to save value " & valeur & " in key " & cle, vbInformation
 End If
End With
End Sub

Private Function LoadSet(LSection As String, ByVal chemin As String, cle As String) As Variant
Dim sValue As Variant
With SIni
    .Path = chemin
    .Section = LSection
    .Key = cle
    sValue = .Value
End With
sValue = Replace(sValue, Chr$(255), " ")
LoadSet = sValue
End Function

Public Sub RefreshF5()
  Dim vnb As Long
  FolderTreeview1(0).Refresh
  vnb = remplissage()
  état.Panels(3).Text = Trim$(Str$(vnb))
  état.Panels(4).Text = "0"
End Sub

Private Sub ChargeVNBCommandes()
    Dim sValue As String
    Dim chemin As String
    chemin = AppPath
    chemin = chemin + "commands.ini"
    sValue = ""
    With SIni
     .Path = chemin
     .Section = "General"
     .Key = "NumberOfCommands"
     sValue = .Value
    End With
    MaxCommand = Val(sValue)
End Sub
Private Sub yourcmd_Click(Index As Integer)
    txtlang.Text = yourcmd(Index).Caption
End Sub
Private Sub ChargeYourCmd()
    Dim i As Integer
    Dim sValue As String
    Dim vnbcmd As Integer
    Dim chemin As String
    chemin = AppPath
    chemin = chemin + "commands.ini"
        If yourcmd.Count > 1 Then
            For i = 0 To yourcmd.Count - 1
                If i <> 0 Then
                    Unload yourcmd(i)
                End If
            Next
        End If
        yourcmd(0).Caption = ""
        txtlang.Enabled = False
        sValue = ""
        With SIni
            .Path = chemin
            .Section = "General"
            .Key = "NumberOfCommands"
            sValue = .Value
        End With
        vnbcmd = Val(sValue) - 1
        For i = 0 To vnbcmd
         With SIni
            .Path = chemin
            .Section = "Commands"
            .Key = "Command" & Trim$(Str$(i + 1))
            sValue = .Value
            End With
            If i <> 0 Then
                Load yourcmd(i)
            End If
            yourcmd(i).Caption = sValue
        Next
End Sub
Private Sub CreateTokenTabl(Filename As String, chemin As String)
' Remplit les tableaux des tokens avec les tokens du préfixe ET du suffixe
    Dim zPrefixe As String, zSuffixe As String, zchemin As String
    Dim i As Integer, j As Integer
    Dim ps As New clsParseString
    VnbTokensFo = 0
    If Filename = "" Then
        VnbTokensPr = 0
        VnbTokensEx = 0
    End If
    zchemin = Replace(chemin, ":", "")
    zPrefixe = Prefixe(Filename)
    zSuffixe = Suffixe(Filename)
    ' 1) réinitialisation des tableaux
    For i = 1 To 100
        TablTokensPr(i) = ""    ' Tokens du préfix
        TablTokensEx(i) = ""    ' Tokens de l'extension
        TablTokensFo(i) = ""    ' Tokens du répertoire
    Next
    ' 2) Les Tokens du prefix
    ps.ParseDelimitedString zPrefixe, LesOptions.CharTokens
    j = ps.Count
    If j > 100 Then
        j = 100
    End If
    ps.MoveFirst
    For i = 1 To j
        TablTokensPr(i) = ps.Token
        PTablTokensPr(i) = ps.Pos
        ps.MoveNext
    Next
    VnbTokensPr = j
    ' 3) Les Tokens de l'extension
    ps.ParseDelimitedString zSuffixe, LesOptions.CharTokens
    j = ps.Count
    If j > 100 Then
        j = 100
    End If
    ps.MoveFirst
    For i = 1 To j
        TablTokensEx(i) = ps.Token
        PTablTokensEx(i) = ps.Pos
        ps.MoveNext
    Next
    VnbTokensEx = j
    ' 4) Les Tokens du répertoire
    ps.ParseDelimitedString zchemin, "\"
    j = ps.Count
    If j > 100 Then
        j = 100
    End If
    ps.MoveFirst
    For i = 1 To j
        TablTokensFo(i) = ps.Token
        PTablTokensFo(i) = ps.Pos
        ps.MoveNext
    Next
    VnbTokensFo = j
End Sub
Private Sub MoveCopyFile(i As Long, vnbit As Long, vnbren As Long)
Dim vtmp As String, vtmp2 As String
Dim fileop As New CSHFileOp
Dim chemin As String
Dim ChemDest As String
Dim NomFichier As String
Dim j As Integer
On Error Resume Next

If Recursive = False Then
    chemin = AddBackSlash(Dir1Path)
End If
fileop.ClearSourceFiles
fileop.ClearDestFiles
NomFichier = LVGetName(ListView1, i)
Select Case LesOptions.Misc3
    Case 0  ' Prefix
        vtmp = Prefixe(NomFichier)
    Case 1  ' Extension
        vtmp = Suffixe(NomFichier)
End Select

If LesOptions.Misc9 = 1 Then ' On utilise qu'une partie du nom
    If LesOptions.Misc10 = 0 Then  ' From ... to ...
        vtmp = Mid$(vtmp, Val(LesOptions.Misc11), (Val(LesOptions.Misc12) - Val(LesOptions.Misc11)) + 1)
    Else    ' Last ... characters
        vtmp = Right$(vtmp, Val(LesOptions.Misc13))
    End If
End If

If LesOptions.Misc6 = 1 Then    ' Stop to numeric character
    vtmp2 = ""
    For j = 1 To Len(vtmp)
        If Not IsNumeric(Mid$(vtmp, j, 1)) Then
            vtmp2 = vtmp2 & Mid$(vtmp, j, 1)
        Else
            j = Len(vtmp)
        End If
    Next
    vtmp = Trim$(vtmp2)
End If
    
If LesOptions.Misc7 = 1 Then ' Replace _ with space
    vtmp = Replace(vtmp, "_", " ")
    vtmp = Trim$(vtmp)
End If
    
If LesOptions.Misc8 = 1 Then ' Capitalize all words
    vtmp = MyStrConv(vtmp)
End If

' Création du répertoire
If Recursive = True Then
    chemin = AddBackSlash(ExtractPath(NomFichier))
End If
ChemDest = chemin + vtmp
MkDir (ChemDest)
Select Case LesOptions.Misc4
    Case 0  ' Just create
    Case 1  ' Copy
        If Recursive = False Then
            fileop.AddSourceFile chemin + NomFichier
        Else
            fileop.AddSourceFile NomFichier
        End If
        fileop.AddDestFile ChemDest + "\" + NomFichier
        état.Panels(1).Text = "Copying " + chemin + NomFichier + " to " + ChemDest + "\" + NomFichier
        état.Panels(2).Text = Trim$(Str$(vnbren)) + "/" + Trim$(Str$(vnbit))
        fileop.CopyFiles
    Case 2  ' Move
        If Recursive = False Then
            fileop.AddSourceFile chemin + NomFichier
        Else
            fileop.AddSourceFile NomFichier
        End If
        fileop.AddDestFile ChemDest + "\" + NomFichier
        état.Panels(1).Text = "Moving " + chemin + NomFichier + " to " + ChemDest + "\" + NomFichier
        état.Panels(2).Text = Trim$(Str$(vnbren)) + "/" + Trim$(Str$(vnbit))
        fileop.MoveFiles
End Select
End Sub
Private Sub SRegRename(i As Long)
Dim vtmp As String
Dim fileop As New CSHFileOp
Dim chemin As String
Dim NomFichier As String
Dim ReturnString As String
On Error Resume Next

vtmp = LVGetName(ListView1, i)
If Recursive = False Then
    chemin = Dir1Path
Else
    chemin = ExtractPath(vtmp)
End If

chemin = AddBackSlash(chemin)

NomFichier = Prefixe(vtmp) & "." & Suffixe(vtmp)
If RegSub(NomFichier, LChaine1 + Chr$(0), LChaine2 + Chr$(0), ReturnString, LOption2, LOption3, 0, 0) Then
    With fileop
        .ParentWnd = hWnd
        .ClearSourceFiles
        .ClearDestFiles
        .ConfirmOperation = False
        .AddSourceFile chemin + NomFichier
        .AddDestFile chemin + ReturnString
    End With
    If LesOptions.CopyRename = True Then
        fileop.RenameFiles
    Else
        fileop.CopyFiles
    End If
    If LesOptions.UseHistory = True Then
        lhistory.AddItem Trim$(Str$(Time())) + "|" + chemin + "|" + NomFichier + "|" + ReturnString ' Historique
    End If
    List2.AddItem chemin + NomFichier       ' Nom d'origine.
    List3.AddItem chemin + ReturnString     ' Nom d'arrivée.
End If
End Sub

Private Function MP3Commands2(Donnee As String, laboucle As Integer, vnom As String) As String
    ' Premier cas, commande du style <MP3Album>
    If commandes(laboucle, 5) = "" And commandes(laboucle, 6) = "" Then
        MP3Commands2 = vnom + Donnee
    Else
        ' Commande du style <MP3Album,Literal>
        If commandes(laboucle, 5) <> "" And commandes(laboucle, 6) = "" Then
            If Trim$(Donnee) = "" Then
                MP3Commands2 = vnom + commandes(laboucle, 5)
            Else
                MP3Commands2 = vnom + Donnee
            End If
        Else    ' Commande du style <MP3Album,Literal,Position>
            If Trim$(Donnee) <> "" Then
                If commandes(laboucle, 6) = "0" Or UCase$(commandes(laboucle, 6)) = "L" Then
                    MP3Commands2 = vnom + commandes(laboucle, 5) + Donnee
                Else    ' Ajouter à droite
                    MP3Commands2 = vnom + Donnee + commandes(laboucle, 5)
                End If
            Else    ' le tag ne contient rien
                MP3Commands2 = vnom
            End If
        End If
    End If
End Function

Private Function MP3Commands(Donnee As String, laboucle As Integer, vnom As String) As String
    ' Premier cas, commande du style <MP3Album>
    If commandes(laboucle, 2) = "" And commandes(laboucle, 3) = "" Then
        MP3Commands = vnom + Donnee
    Else
        ' Commande du style <MP3Album,Literal>
        If commandes(laboucle, 2) <> "" And commandes(laboucle, 3) = "" Then
            If Trim$(Donnee) = "" Then
                MP3Commands = vnom + commandes(laboucle, 2)
            Else
                MP3Commands = vnom + Donnee
            End If
        Else    ' Commande du style <MP3Album,Literal,Position>
            If Trim$(Donnee) <> "" Then
                If commandes(laboucle, 3) = "0" Or UCase$(commandes(laboucle, 3)) = "L" Then
                    MP3Commands = vnom + commandes(laboucle, 2) + Donnee
                Else    ' Ajouter à droite
                    MP3Commands = vnom + Donnee + commandes(laboucle, 2)
                End If
            Else    ' le tag ne contient rien
                MP3Commands = vnom
            End If
        End If
    End If
End Function
' Renvoie le nombre de tags MP3 et VQF qu'il faut afficher
Private Function NbMP3Tags(ByRef lachaine As String) As Integer
Dim chemin As String
Dim vnbcmd As Integer
Dim i As Integer
Dim vnb As Integer
Dim sValue As String
lachaine = ""
chemin = AppPath + "Music.ini"
With SIni
    .Path = chemin
    .Section = "MP3Tags"
    .Key = "NumberOfTags"
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "MP3Tags"
        .Key = "Tag" & Trim$(Str$(i))
        sValue = .Value
    End With
    If Val(GetToken(sValue, "|", 1)) = 1 Then
        vnb = vnb + 1
        lachaine = lachaine + GetToken(sValue, "|", 3) + "|"
    End If
Next
NbMP3Tags = vnb
End Function

Private Function NbVQFTags(lachaine As String) As Integer
Dim chemin As String, vnbcmd As Integer, i As Integer, vnb As Integer, sValue As String
lachaine = ""
chemin = AppPath + "Music.ini"
With SIni
    .Section = "VQFTags"
    .Key = "NumberOfTags"
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "VQFTags"
        .Key = "Tag" & Trim$(Str$(i))
        sValue = .Value
    End With
    If Val(GetToken(sValue, "|", 1)) = 1 Then
        vnb = vnb + 1
        lachaine = lachaine + GetToken(sValue, "|", 3) + "|"
    End If
Next
NbVQFTags = vnb
End Function
' Fonction qui renvoie le titre à utiliser
Private Function MP3Caption(lindex As Integer) As String
Dim chemin As String
Dim vnbcmd As Integer
Dim i As Integer
Dim vnb As Integer
Dim sValue As String
chemin = AppPath + "Music.ini"
With SIni
    .Path = chemin
    .Section = "MP3Tags"
    .Key = "NumberOfTags"
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "MP3Tags"
        .Key = "Tag" & Trim$(Str$(i))
        sValue = .Value
    End With
    If Val(GetToken(sValue, "|", 1)) = 1 Then
        vnb = vnb + 1
        If vnb = lindex Then
            MP3Caption = GetToken(sValue, "|", 2)
            Exit Function
        End If
    End If
Next
MP3Caption = ""
End Function

Private Function VQFCaption(lindex As Integer) As String
Dim chemin As String
Dim vnbcmd As Integer
Dim i As Integer
Dim vnb As Integer
Dim sValue As String
chemin = AppPath + "Music.ini"
With SIni
    .Path = chemin
    .Section = "VQFTags"
    .Key = "NumberOfTags"
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "VQFTags"
        .Key = "Tag" & Trim$(Str$(i))
        sValue = .Value
    End With
    If Val(GetToken(sValue, "|", 1)) = 1 Then
        vnb = vnb + 1
        If vnb = lindex Then
            VQFCaption = GetToken(sValue, "|", 2)
            Exit Function
        End If
    End If
Next
VQFCaption = ""
End Function

Private Function Blanc(lachaine As String) As String
If Trim$(lachaine) = "" Then
    Blanc = "&nbsp;"
Else
    Blanc = lachaine
End If
End Function
' Permet de se déplacer dans l'arborescence du FolderTreeview
Private Sub GoFTV(Ou As Integer)
  Dim fldCur As CCRPFolderTV6.Folder
  Dim vtmp As String
  On Error Resume Next
  Set fldCur = FolderTreeview1(0).SelectedFolder
  Select Case Ou
    Case 1  ' First
        vtmp = fldCur.FirstSibling
    Case 2  ' Next
        vtmp = fldCur.NextSibling
    Case 3  ' Previous
        vtmp = fldCur.PrevSibling
    Case 4  ' Last
        vtmp = fldCur.LastSibling
  End Select
  FolderTreeview1(0).Visible = False
  FolderTreeview1(0).SelectedFolder = vtmp
  Dir1Path = vtmp
  FolderTreeview1(0).Visible = True
End Sub
' Affiche les infos sur les fichiers Afm
Private Sub LoadAFM()
    Dim vNomFic As String
    vNomFic = AddBackSlash(Dir1Path) + ListView1.SelectedItem.Text
    AFM.GetAFMInfos vNomFic
    AFM.FillTagsList LvMP3
End Sub
' Charge les infos sur le fichier vqf sélectionné
Private Sub loadLvVQF()
    Dim vNomFic As String
    Dim SonInfo As String
    vNomFic = AddBackSlash(Dir1Path) + ListView1.SelectedItem.Text
    SonInfo = MusVQF.GetVQFInfos(vNomFic, False)
    MusVQF.FillTagsList LvMP3
End Sub
' Charge les infos sur le fichier WMA sélectionné
Private Sub LoadWMA()
    Dim vNomFic As String
    Dim SonInfo As String
    vNomFic = AddBackSlash(Dir1Path) + ListView1.SelectedItem.Text
    SonInfo = MusWMA.GetWMAInfos(vNomFic, False)
    MusWMA.FillTagsList LvMP3
End Sub

' Charge les infos sur le fichier Ogg sélectionné
Private Sub LoadLvOgg()
    Dim vNomFic As String
    Dim SonInfo As String
    vNomFic = AddBackSlash(Dir1Path) + ListView1.SelectedItem.Text
    SonInfo = MusOgg.GetOggInfos(vNomFic, False)
    MusOgg.FillTagsList LvMP3
End Sub
' Charge les infos sur le fichier MP3 sélectionné
Private Sub LoadLvMP3()
Dim vNomFic As String
Dim SonInfo As String
vNomFic = AddBackSlash(Dir1Path) + ListView1.SelectedItem.Text
LvMP3.ListItems.Clear
SonInfo = MusMP3.GetMP3Infos(vNomFic, False)

AddOneTag MusMP3.Album, "Album"
AddOneTag MusMP3.Artist, "Artist"
AddOneTag MusMP3.Band, "Band"
AddOneTag MusMP3.BPM, "BPM"
AddOneTag MusMP3.Comment, "Comment"
AddOneTag MusMP3.Composer, "Composer"
AddOneTag MusMP3.Conductor, "Conductor"
AddOneTag MusMP3.ContentGroup, "ContentGroup"
AddOneTag MusMP3.Copyright, "Copyright"
AddOneTag MusMP3.EncodedBy, "EncodedBy"
AddOneTag MusMP3.EncryptionMethod, "EncryptionMethod"
AddOneTag MusMP3.FileOwner, "FileOwner"
AddOneTag MusMP3.FileType, "FileType"
AddOneTag MusMP3.Genre, "Genre"
AddOneTag MusMP3.GroupIdent, "GroupIdent"
AddOneTag MusMP3.InitialKey, "InitialKey"
AddOneTag MusMP3.InvolvedPeopleList, "InvolvedPeopleList"
AddOneTag MusMP3.ISRC, "ISRC"
AddOneTag MusMP3.Language, "Language"
AddOneTag MusMP3.LinkedInformation, "LinkedInformation"
AddOneTag MusMP3.Lyricist, "Lyricist"
AddOneTag MusMP3.mDate, "Date"
AddOneTag MusMP3.MediaType, "MediaType"
AddOneTag MusMP3.MixArtist, "MixArtist"
AddOneTag MusMP3.NetRadioOwner, "NetRadioOwner"
AddOneTag MusMP3.OriginalAlbum, "OriginalAlbum"
AddOneTag MusMP3.OriginalArtist, "OriginalArtist"
AddOneTag MusMP3.OriginalFilename, "OriginalFilename"
AddOneTag MusMP3.OriginalLyricist, "OriginalLyricist"
AddOneTag MusMP3.OriginalYear, "OriginalYear"
AddOneTag MusMP3.PartOfASet, "PartOfASet"
AddOneTag MusMP3.PlayListDelay, "PlayListDelay"
AddOneTag MusMP3.PopulariMeter, "PopulariMeter"
AddOneTag MusMP3.Publisher, "Publisher"
AddOneTag MusMP3.RecordingDates, "RecordingDates"
AddOneTag MusMP3.SoftwareEncodingSettings, "SoftwareEncodingSettings"
AddOneTag MusMP3.SongLength, "SongLength"
AddOneTag MusMP3.SubTitle, "SubTitle"
AddOneTag MusMP3.SynchronizedLyric, "SynchronizedLyric"
AddOneTag MusMP3.TermsOfUse, "TermsOfUse"
AddOneTag MusMP3.Time, "Time"
AddOneTag MusMP3.Title, "Title"
AddOneTag MusMP3.TotalTracks, "TotalTracks"
AddOneTag MusMP3.TrackNumber, "TrackNumber"
AddOneTag MusMP3.UnsynchronizedLyric, "UnsynchronizedLyric"
AddOneTag MusMP3.UserText, "UserText"
AddOneTag MusMP3.wwwArtist, "wwwArtist"
AddOneTag MusMP3.wwwAudioFile, "wwwAudioFile"
AddOneTag MusMP3.wwwAudioSource, "wwwAudioSource"
AddOneTag MusMP3.wwwCommercialInfo, "wwwCommercialInfo"
AddOneTag MusMP3.wwwCopyright, "wwwCopyright"
AddOneTag MusMP3.wwwPayment, "wwwPayment"
AddOneTag MusMP3.wwwPublisher, "wwwPublisher"
AddOneTag MusMP3.wwwRadioPage, "wwwRadioPage"
AddOneTag MusMP3.wwwUserURL, "wwwUserURL"
AddOneTag MusMP3.Year, "Year"

ResizeLvMp3
End Sub

Private Sub AddOneTag(Donnee As String, texte As String)
Dim Aff As Boolean
Aff = True
If LesOptions.RemoveEmptyTags = 1 And Trim$(Donnee) = "" Then
    Aff = False
End If
If Aff Then
    Set itmX = LvMP3.ListItems.Add(, , texte)
    itmX.SubItems(1) = Donnee
End If
End Sub
' Gestion de l'onglet TAGS
Private Sub MTetat1()
    If mviewmp3tab.Checked = True Then
        mviewmp3tab.Checked = False
    Else
        mviewmp3tab.Checked = True
    End If
    LesOptions.ShowMP3Tab = mviewmp3tab.Checked
    TabGen.TabVisible(2) = mviewmp3tab.Checked
End Sub
' Gestion de l'onglet Pictures
Private Sub MTetat2()
    If mviewpicturetab.Checked = True Then
        mviewpicturetab.Checked = False
    Else
        mviewpicturetab.Checked = True
    End If
    LesOptions.ShowMusicTab = mviewpicturetab.Checked
    TabGen.TabVisible(3) = mviewpicturetab.Checked
End Sub
' Gestion de l'onglet Text
Private Sub MTetat3()
    If mviewtexttab.Checked = True Then
        mviewtexttab.Checked = False
    Else
        mviewtexttab.Checked = True
    End If
    LesOptions.ShowTextTab = mviewtexttab.Checked
    TabGen.TabVisible(4) = mviewtexttab.Checked
End Sub

' Copie les fichiers sélectionnés dans la liste des fichiers vers la liste de l'option "Rename from a list"
Private Sub CopySelected(Cache As Boolean)
Dim i As Long, sItem As String
On Error GoTo ErrGen
RENAME.MousePointer = 11
If Cache Then
    ListView2.Visible = False
End If
i = LVGetItemSelected(ListView1, -1)
While i <> -1
 sItem = LVGetName(ListView1, i)
 Set itmX = ListView2.ListItems.Add(, , sItem)
 itmX.Text = sItem
 itmX.SubItems(1) = sItem
 i = LVGetItemSelected(ListView1, i)
Wend
If Cache Then
    ListView2.Visible = True
End If
RENAME.MousePointer = 0
Exit Sub

ErrGen:
ErreurGrave "CopySelected"
End Sub
' Charge un fichier texte dans textbox1
Private Sub LoadText(fichier As String)
    On Error GoTo ErrGen
    Dim ff As Integer
    Dim Buffer As String
    Dim longueur As Long
    ff = FreeFile
    longueur = FileLen(fichier)
    Open fichier For Binary As #ff Len = longueur
    Buffer = Input(longueur, #ff)
    Text1.Text = Buffer
    Close #ff
    Exit Sub
ErrGen:
    If Err.Number = 6 Then
        Text1.Text = "This file is too large, sorry I can't view it..."
    End If
End Sub

Private Sub LesRecherches1(Chaine1 As String, Chaine2 As String, test1 As Integer, test2 As Integer)
    Dim vtmpGlob As String
    If RechGlob = True Then
            If LesOptions.SearchAndReplace = test1 Or LesOptions.SearchAndReplace = test2 Then
                rech3.SourceString = Chaine1 + "." + Chaine2
                vtmpGlob = rech3.BeginSearchAndReplace
                Chaine1 = Prefixe(vtmpGlob)
                Chaine2 = Suffixe(vtmpGlob)
                rech3.SourceString = Chaine1 + "." + Chaine2
                vtmpGlob = rech3.BeginReplaceCharacters
                Chaine1 = Prefixe(vtmpGlob)
                Chaine2 = Suffixe(vtmpGlob)
            End If
    Else
        If RechPref = True Then
            If LesOptions.SearchAndReplace = test1 Or LesOptions.SearchAndReplace = test2 Then
                rech1.SourceString = Chaine1
                Chaine1 = rech1.BeginSearchAndReplace
                rech1.SourceString = Chaine1
                Chaine1 = rech1.BeginReplaceCharacters
            End If
        End If
     
        If RechSuff = True Then
            If LesOptions.SearchAndReplace = test1 Or LesOptions.SearchAndReplace = test2 Then
                rech2.SourceString = Chaine2
                Chaine2 = rech2.BeginSearchAndReplace
                rech2.SourceString = Chaine2
                Chaine2 = rech2.BeginReplaceCharacters
            End If
        End If
    End If
End Sub

Private Sub LoadMenuPicture()
    mpict(0).Caption = "Zoom In"
    Load mpict(1)
    mpict(1).Caption = "Zoom Out"
    Load mpict(2)
    mpict(2).Caption = "Real Size"
    Load mpict(3)
    mpict(3).Caption = "Stretch"
    Load mpict(4)
    mpict(4).Caption = "Best Fit"
End Sub

Private Function FmtToken(Token As String, Format As Integer) As String
Select Case Format
        Case 1   ' Capitalize First word only
            FmtToken = UCase$(Left$(Token, 1)) + LCase$(Mid$(Token, 2))
        Case 2   ' Capitalize all words
            FmtToken = MyStrConv(Token)
        Case 3   ' CowBoys
            FmtToken = CoWbOyS(Token)
        Case 4   ' Invert
            FmtToken = StrReverse(Token)
        Case 5   ' Lower
            FmtToken = LCase$(Token)
        Case 6   ' Ltrim
            FmtToken = LTrim$(Token)
        Case 7   ' Rtrim
            FmtToken = RTrim$(Token)
        Case 8   ' Trim
            FmtToken = Trim$(Token)
        Case 9   ' Toggle
            FmtToken = ToggleCase(Token)
        Case 10  ' Upper
            FmtToken = UCase$(Token)
        Case 11  ' lTrim
            FmtToken = LTrim$(Token)
        Case 12  ' RTrim
            FmtToken = RTrim$(Token)
        Case 13  ' Trim
            FmtToken = Trim$(Token)
        Case Else
            FmtToken = Token
    End Select

End Function

Private Sub ClearCommands()
Dim i As Integer
For i = 1 To 300
    commandes(i, 1) = ""
    commandes(i, 2) = ""
    commandes(i, 3) = ""
Next
End Sub

Private Function GiveRandomName(chemin As String, vnom As String) As String
'Dim LongRandom As Integer, LetterRandom As Long, VraiRandom As Boolean, NomRandom As String, BoucleRandom As Integer
'VraiRandom = True
'While VraiRandom <> False
'    While LongRandom = 0
'        LongRandom = Int(Rnd * 24)
'    Wend
'    For BoucleRandom = 1 To LongRandom
'        LetterRandom = 0
'        While LetterRandom = 0
'            LetterRandom = Int(Rnd * 26)
'        Wend
'        NomRandom = NomRandom + FBase26(LetterRandom)
'    Next
'    If LesOptions.UseLowerInLetterCounters = 1 Then ' passer en minuscules
'        NomRandom = LCase$(NomRandom)
'    End If
'    VraiRandom = FileExists(chemin & vnom & NomRandom & "." & vsuffixe)
'Wend
'GiveRandomName = NomRandom
End Function
