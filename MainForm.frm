VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "THE Rename"
   ClientHeight    =   9000
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11655
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHolder 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11655
      TabIndex        =   101
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
   End
   Begin VB.PictureBox FrameGauche 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8115
      Left            =   0
      ScaleHeight     =   8115
      ScaleWidth      =   5940
      TabIndex        =   1
      Top             =   615
      Visible         =   0   'False
      Width           =   5940
      Begin VB.PictureBox Picture10 
         AutoRedraw      =   -1  'True
         Height          =   135
         Left            =   5640
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   100
         Top             =   120
         Width           =   135
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sample "
         Height          =   700
         Left            =   30
         TabIndex        =   94
         Top             =   5280
         WhatsThisHelpID =   178
         Width           =   4335
         Begin VB.Label laides 
            BackStyle       =   0  'Transparent
            Height          =   210
            Left            =   120
            TabIndex        =   96
            ToolTipText     =   "Sample of what extension will be"
            Top             =   405
            WhatsThisHelpID =   178
            Width           =   4095
         End
         Begin VB.Label laidep 
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   120
            TabIndex        =   95
            ToolTipText     =   "Sample of what prefix will be"
            Top             =   195
            WhatsThisHelpID =   178
            Width           =   4095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Extension "
         Height          =   2380
         Left            =   30
         TabIndex        =   52
         Top             =   2880
         Width           =   5895
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   93
            ToolTipText     =   "Scroll to select action to perform on extension"
            Top             =   240
            WhatsThisHelpID =   176
            Width           =   5175
         End
         Begin VB.PictureBox PanelExt 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1650
            Left            =   120
            ScaleHeight     =   1650
            ScaleWidth      =   5715
            TabIndex        =   54
            Top             =   650
            Width           =   5715
            Begin VB.CommandButton Command11 
               Caption         =   "Reset"
               Height          =   300
               HelpContextID   =   110
               Left            =   4800
               TabIndex        =   56
               ToolTipText     =   "Reset all selections to their default's value"
               Top             =   1350
               WhatsThisHelpID =   181
               Width           =   870
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Search && Replace..."
               Height          =   300
               HelpContextID   =   110
               Left            =   3060
               TabIndex        =   55
               ToolTipText     =   "Click on this button to use search and replace"
               Top             =   1350
               WhatsThisHelpID =   180
               Width           =   1635
            End
            Begin TabDlg.SSTab SSTab2 
               Height          =   1300
               Left            =   0
               TabIndex        =   57
               Top             =   0
               Width           =   5700
               _ExtentX        =   10054
               _ExtentY        =   2302
               _Version        =   393216
               Style           =   1
               Tabs            =   5
               TabsPerRow      =   5
               TabHeight       =   520
               TabCaption(0)   =   "Counter"
               TabPicture(0)   =   "MainForm.frx":0000
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Check11"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "onglcounter2"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).ControlCount=   2
               TabCaption(1)   =   "Size"
               TabPicture(1)   =   "MainForm.frx":001C
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Check12"
               Tab(1).Control(1)=   "Picture13"
               Tab(1).ControlCount=   2
               TabCaption(2)   =   "Date"
               TabPicture(2)   =   "MainForm.frx":0038
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Picture14"
               Tab(2).Control(1)=   "Check13"
               Tab(2).ControlCount=   2
               TabCaption(3)   =   "Time"
               TabPicture(3)   =   "MainForm.frx":0054
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Picture4"
               Tab(3).Control(1)=   "Check4"
               Tab(3).ControlCount=   2
               TabCaption(4)   =   "Text "
               TabPicture(4)   =   "MainForm.frx":0070
               Tab(4).ControlEnabled=   0   'False
               Tab(4).ControlCount=   0
               Begin VB.PictureBox onglcounter2 
                  AutoRedraw      =   -1  'True
                  BorderStyle     =   0  'None
                  Height          =   580
                  Left            =   120
                  ScaleHeight     =   585
                  ScaleWidth      =   5475
                  TabIndex        =   81
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   5475
                  Begin VB.ComboBox Combo4 
                     Height          =   315
                     ItemData        =   "MainForm.frx":008C
                     Left            =   4095
                     List            =   "MainForm.frx":009F
                     Style           =   2  'Dropdown List
                     TabIndex        =   89
                     ToolTipText     =   "Choose file's counter representation (Dec/Hex/Oct)"
                     Top             =   35
                     WhatsThisHelpID =   187
                     Width           =   1300
                  End
                  Begin VB.TextBox Text17 
                     Height          =   285
                     Left            =   1980
                     TabIndex        =   88
                     Text            =   "1"
                     ToolTipText     =   "Enter increment value (1 for example)"
                     Top             =   35
                     WhatsThisHelpID =   185
                     Width           =   495
                  End
                  Begin VB.TextBox Text18 
                     Height          =   285
                     Left            =   3285
                     TabIndex        =   87
                     Text            =   "4"
                     ToolTipText     =   "Enter number of digits for the counter"
                     Top             =   35
                     WhatsThisHelpID =   186
                     Width           =   495
                  End
                  Begin VB.PictureBox Picture11 
                     BorderStyle     =   0  'None
                     Height          =   255
                     Left            =   165
                     ScaleHeight     =   255
                     ScaleWidth      =   5055
                     TabIndex        =   83
                     Top             =   345
                     WhatsThisHelpID =   188
                     Width           =   5055
                     Begin VB.OptionButton Option3 
                        Caption         =   "Replace extension"
                        Height          =   255
                        Index           =   24
                        Left            =   3375
                        TabIndex        =   86
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
                        TabIndex        =   85
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
                        TabIndex        =   84
                        ToolTipText     =   "Counter will be added to the left of extension"
                        Top             =   0
                        Value           =   -1  'True
                        WhatsThisHelpID =   188
                        Width           =   1365
                     End
                  End
                  Begin VB.TextBox Text16 
                     Height          =   285
                     Left            =   795
                     TabIndex        =   82
                     Text            =   "1"
                     ToolTipText     =   "Enter begin's value"
                     Top             =   35
                     WhatsThisHelpID =   184
                     Width           =   495
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     Caption         =   "Begin"
                     Height          =   195
                     Left            =   180
                     TabIndex        =   92
                     Top             =   60
                     WhatsThisHelpID =   184
                     Width           =   405
                  End
                  Begin VB.Label Label13 
                     AutoSize        =   -1  'True
                     Caption         =   "Step"
                     Height          =   195
                     Left            =   1455
                     TabIndex        =   91
                     Top             =   60
                     WhatsThisHelpID =   185
                     Width           =   330
                  End
                  Begin VB.Label Label14 
                     AutoSize        =   -1  'True
                     Caption         =   "Digits"
                     Height          =   195
                     Left            =   2730
                     TabIndex        =   90
                     Top             =   60
                     WhatsThisHelpID =   186
                     Width           =   390
                  End
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Left            =   -73200
                  TabIndex        =   80
                  ToolTipText     =   "Current extension will be replace with this text"
                  Top             =   480
                  Visible         =   0   'False
                  WhatsThisHelpID =   197
                  Width           =   2775
               End
               Begin VB.TextBox Text15 
                  Height          =   285
                  Left            =   -73200
                  TabIndex        =   79
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
                  TabIndex        =   78
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
                  TabIndex        =   77
                  ToolTipText     =   "Click to add or remove a text to the extension"
                  Top             =   870
                  WhatsThisHelpID =   199
                  Width           =   945
               End
               Begin VB.PictureBox Picture9 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -71070
                  ScaleHeight     =   240
                  ScaleWidth      =   1335
                  TabIndex        =   74
                  Top             =   900
                  Width           =   1335
                  Begin VB.OptionButton Option5 
                     Caption         =   "Begin"
                     Height          =   195
                     Index           =   0
                     Left            =   0
                     TabIndex        =   76
                     ToolTipText     =   "Text will be added to the left of extension"
                     Top             =   0
                     Value           =   -1  'True
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   855
                  End
                  Begin VB.OptionButton Option5 
                     Caption         =   "End"
                     Height          =   195
                     Index           =   1
                     Left            =   720
                     TabIndex        =   75
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
                  TabIndex        =   73
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
                  TabIndex        =   69
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   190
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   29
                     Left            =   15
                     TabIndex        =   72
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
                     TabIndex        =   71
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
                     TabIndex        =   70
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
                  TabIndex        =   65
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   192
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace extension"
                     Height          =   255
                     Index           =   32
                     Left            =   3375
                     TabIndex        =   68
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
                     TabIndex        =   67
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
                     TabIndex        =   66
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
                  TabIndex        =   64
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
                  TabIndex        =   60
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   194
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   14
                     Left            =   15
                     TabIndex        =   63
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
                     TabIndex        =   62
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
                     TabIndex        =   61
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
                  TabIndex        =   59
                  ToolTipText     =   "Click to add file's time"
                  Top             =   360
                  WhatsThisHelpID =   193
                  Width           =   1275
               End
               Begin VB.CheckBox Check12 
                  Caption         =   "Add file's size"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   58
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
            Left            =   5400
            TabIndex        =   53
            ToolTipText     =   "Use same option for prefix and extension"
            Top             =   240
            WhatsThisHelpID =   177
            Width           =   255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Prefix "
         Height          =   2600
         Left            =   30
         TabIndex        =   2
         Top             =   240
         Width           =   5895
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "MainForm.frx":00D1
            Left            =   120
            List            =   "MainForm.frx":00D3
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   51
            ToolTipText     =   "Scroll to select action to perform on prefix"
            Top             =   240
            WhatsThisHelpID =   175
            Width           =   5175
         End
         Begin VB.PictureBox PanelPrefix 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   90
            ScaleHeight     =   1815
            ScaleWidth      =   5715
            TabIndex        =   4
            Top             =   630
            Width           =   5715
            Begin VB.CommandButton Command10 
               Caption         =   "Reset"
               Height          =   300
               HelpContextID   =   110
               Left            =   4800
               TabIndex        =   6
               ToolTipText     =   "Reset all selections to their default's value"
               Top             =   1400
               WhatsThisHelpID =   181
               Width           =   870
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Search && Replace..."
               Height          =   300
               HelpContextID   =   110
               Left            =   3060
               TabIndex        =   5
               ToolTipText     =   "Click on this button to use search and replace"
               Top             =   1400
               WhatsThisHelpID =   180
               Width           =   1635
            End
            Begin TabDlg.SSTab SSTab1 
               Height          =   1335
               Left            =   0
               TabIndex        =   7
               Top             =   0
               Width           =   5700
               _ExtentX        =   10054
               _ExtentY        =   2355
               _Version        =   393216
               Style           =   1
               Tabs            =   8
               Tab             =   5
               TabsPerRow      =   8
               TabHeight       =   520
               TabCaption(0)   =   "Counter"
               TabPicture(0)   =   "MainForm.frx":00D5
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "Check3"
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "Size"
               TabPicture(1)   =   "MainForm.frx":00F1
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Check5"
               Tab(1).Control(1)=   "Picture5"
               Tab(1).ControlCount=   2
               TabCaption(2)   =   "Date"
               TabPicture(2)   =   "MainForm.frx":010D
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Picture6"
               Tab(2).Control(1)=   "Check6"
               Tab(2).ControlCount=   2
               TabCaption(3)   =   "Time"
               TabPicture(3)   =   "MainForm.frx":0129
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Check7"
               Tab(3).Control(1)=   "Picture7"
               Tab(3).ControlCount=   2
               TabCaption(4)   =   "Folder's name"
               TabPicture(4)   =   "MainForm.frx":0145
               Tab(4).ControlEnabled=   0   'False
               Tab(4).ControlCount=   0
               TabCaption(5)   =   "Text "
               TabPicture(5)   =   "MainForm.frx":0161
               Tab(5).ControlEnabled=   -1  'True
               Tab(5).Control(0)=   "Command8"
               Tab(5).Control(0).Enabled=   0   'False
               Tab(5).Control(1)=   "Option1(0)"
               Tab(5).Control(1).Enabled=   0   'False
               Tab(5).Control(2)=   "Text2"
               Tab(5).Control(2).Enabled=   0   'False
               Tab(5).Control(3)=   "Picture2"
               Tab(5).Control(3).Enabled=   0   'False
               Tab(5).Control(4)=   "Option1(1)"
               Tab(5).Control(4).Enabled=   0   'False
               Tab(5).Control(5)=   "Text14"
               Tab(5).Control(5).Enabled=   0   'False
               Tab(5).ControlCount=   6
               TabCaption(6)   =   "Pictures"
               TabPicture(6)   =   "MainForm.frx":017D
               Tab(6).ControlEnabled=   0   'False
               Tab(6).Control(0)=   "Picture8"
               Tab(6).Control(1)=   "Check1"
               Tab(6).ControlCount=   2
               TabCaption(7)   =   "Sounds"
               TabPicture(7)   =   "MainForm.frx":0199
               Tab(7).ControlEnabled=   0   'False
               Tab(7).Control(0)=   "Command5"
               Tab(7).ControlCount=   1
               Begin VB.PictureBox onglcounter 
                  AutoRedraw      =   -1  'True
                  BorderStyle     =   0  'None
                  Height          =   580
                  Left            =   -74880
                  ScaleHeight     =   585
                  ScaleWidth      =   5475
                  TabIndex        =   39
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   5475
                  Begin VB.ComboBox Combo3 
                     Height          =   315
                     ItemData        =   "MainForm.frx":01B5
                     Left            =   4095
                     List            =   "MainForm.frx":01C8
                     Style           =   2  'Dropdown List
                     TabIndex        =   47
                     ToolTipText     =   "Choose file's counter representation (Dec/Hex/Oct)"
                     Top             =   35
                     WhatsThisHelpID =   187
                     Width           =   1300
                  End
                  Begin VB.PictureBox Picture3 
                     BorderStyle     =   0  'None
                     Height          =   255
                     Left            =   210
                     ScaleHeight     =   255
                     ScaleWidth      =   4815
                     TabIndex        =   43
                     Top             =   360
                     WhatsThisHelpID =   188
                     Width           =   4815
                     Begin VB.OptionButton Option3 
                        Caption         =   "Add to the left"
                        Height          =   255
                        Index           =   0
                        Left            =   15
                        TabIndex        =   46
                        ToolTipText     =   "Counter will be added to the left of prefix"
                        Top             =   0
                        Value           =   -1  'True
                        WhatsThisHelpID =   188
                        Width           =   1455
                     End
                     Begin VB.OptionButton Option3 
                        Caption         =   "Add to the right"
                        Height          =   255
                        Index           =   1
                        Left            =   1665
                        TabIndex        =   45
                        ToolTipText     =   "Counter will be added to the right of prefix"
                        Top             =   0
                        WhatsThisHelpID =   188
                        Width           =   1395
                     End
                     Begin VB.OptionButton Option3 
                        Caption         =   "Replace prefix"
                        Height          =   255
                        Index           =   2
                        Left            =   3375
                        TabIndex        =   44
                        ToolTipText     =   "Counter will replace file's prefixe"
                        Top             =   0
                        WhatsThisHelpID =   188
                        Width           =   2295
                     End
                  End
                  Begin VB.TextBox Text5 
                     Height          =   285
                     Left            =   3285
                     TabIndex        =   42
                     Text            =   "4"
                     ToolTipText     =   "Enter number of digits for the counter"
                     Top             =   35
                     WhatsThisHelpID =   186
                     Width           =   495
                  End
                  Begin VB.TextBox Text4 
                     Height          =   285
                     Left            =   1980
                     TabIndex        =   41
                     Text            =   "1"
                     ToolTipText     =   "Enter increment value (1 for example)"
                     Top             =   35
                     WhatsThisHelpID =   185
                     Width           =   495
                  End
                  Begin VB.TextBox Text3 
                     Height          =   285
                     Left            =   795
                     TabIndex        =   40
                     Text            =   "1"
                     ToolTipText     =   "Enter begin's value"
                     Top             =   35
                     WhatsThisHelpID =   184
                     Width           =   495
                  End
                  Begin VB.Label Label6 
                     AutoSize        =   -1  'True
                     Caption         =   "Digits"
                     Height          =   195
                     Left            =   2730
                     TabIndex        =   50
                     Top             =   60
                     WhatsThisHelpID =   186
                     Width           =   390
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     Caption         =   "Step"
                     Height          =   195
                     Left            =   1455
                     TabIndex        =   49
                     Top             =   60
                     WhatsThisHelpID =   185
                     Width           =   330
                  End
                  Begin VB.Label Label4 
                     AutoSize        =   -1  'True
                     Caption         =   "Begin"
                     Height          =   195
                     Left            =   225
                     TabIndex        =   48
                     Top             =   60
                     WhatsThisHelpID =   184
                     Width           =   405
                  End
               End
               Begin VB.CommandButton Command19 
                  Caption         =   "Options..."
                  Height          =   375
                  Left            =   -72720
                  TabIndex        =   38
                  Top             =   600
                  WhatsThisHelpID =   195
                  Width           =   1455
               End
               Begin VB.TextBox Text14 
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   37
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
                  Left            =   120
                  TabIndex        =   36
                  ToolTipText     =   "Click to add or remove a text to the prefix"
                  Top             =   870
                  WhatsThisHelpID =   199
                  Width           =   960
               End
               Begin VB.PictureBox Picture2 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   3960
                  ScaleHeight     =   255
                  ScaleWidth      =   1335
                  TabIndex        =   33
                  Top             =   900
                  Width           =   1335
                  Begin VB.OptionButton Option2 
                     Caption         =   "Begin"
                     Height          =   195
                     Index           =   0
                     Left            =   0
                     TabIndex        =   35
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
                     Left            =   720
                     TabIndex        =   34
                     ToolTipText     =   "Text will be added at the end of prefix"
                     Top             =   0
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   615
                  End
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   32
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
                  Left            =   120
                  TabIndex        =   31
                  ToolTipText     =   "Click to replace or not file's prefix with a text"
                  Top             =   525
                  WhatsThisHelpID =   196
                  Width           =   1635
               End
               Begin VB.CheckBox Check5 
                  Caption         =   "Add file's size"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   30
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
                  TabIndex        =   26
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   190
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   4
                     Left            =   1650
                     TabIndex        =   29
                     ToolTipText     =   "File's size will be added to the right"
                     Top             =   0
                     WhatsThisHelpID =   190
                     Width           =   1440
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   3
                     Left            =   0
                     TabIndex        =   28
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
                     TabIndex        =   27
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
                  TabIndex        =   22
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   192
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   8
                     Left            =   15
                     TabIndex        =   25
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
                     TabIndex        =   24
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
                     TabIndex        =   23
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
                  TabIndex        =   21
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
                  TabIndex        =   17
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   194
                  Width           =   4695
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     Index           =   11
                     Left            =   3375
                     TabIndex        =   20
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
                     TabIndex        =   19
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
                     TabIndex        =   18
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
                  TabIndex        =   16
                  ToolTipText     =   "Click to add file's time"
                  Top             =   360
                  WhatsThisHelpID =   193
                  Width           =   1320
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "Add a counter"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   15
                  ToolTipText     =   "Click to add a counter"
                  Top             =   360
                  WhatsThisHelpID =   183
                  Width           =   1365
               End
               Begin VB.PictureBox Picture8 
                  BorderStyle     =   0  'None
                  Height          =   300
                  Left            =   -74640
                  ScaleHeight     =   300
                  ScaleWidth      =   4815
                  TabIndex        =   11
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   203
                  Width           =   4815
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     HelpContextID   =   162
                     Index           =   15
                     Left            =   3375
                     TabIndex        =   14
                     ToolTipText     =   "Replace prefix"
                     Top             =   0
                     WhatsThisHelpID =   203
                     Width           =   2295
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     HelpContextID   =   162
                     Index           =   16
                     Left            =   1650
                     TabIndex        =   13
                     ToolTipText     =   "Add to the right"
                     Top             =   0
                     WhatsThisHelpID =   203
                     Width           =   1395
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     HelpContextID   =   162
                     Index           =   17
                     Left            =   15
                     TabIndex        =   12
                     ToolTipText     =   "Add to the left"
                     Top             =   0
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
                  TabIndex        =   10
                  ToolTipText     =   "See the options to modify it's format"
                  Top             =   360
                  WhatsThisHelpID =   202
                  Width           =   2445
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "Options..."
                  Height          =   375
                  Left            =   -72720
                  TabIndex        =   9
                  Top             =   600
                  WhatsThisHelpID =   204
                  Width           =   1455
               End
               Begin VB.CommandButton Command8 
                  Height          =   300
                  HelpContextID   =   168
                  Left            =   5325
                  Picture         =   "MainForm.frx":01FA
                  Style           =   1  'Graphical
                  TabIndex        =   8
                  ToolTipText     =   "Use cyclic selection"
                  Top             =   495
                  Visible         =   0   'False
                  WhatsThisHelpID =   198
                  Width           =   300
               End
            End
         End
         Begin VB.CommandButton Command3 
            Caption         =   "="
            Height          =   315
            Left            =   5400
            TabIndex        =   3
            ToolTipText     =   "Use same option for prefix and extension"
            Top             =   240
            WhatsThisHelpID =   177
            Width           =   255
         End
      End
      Begin VB.Label etat1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Undo file not active"
         Height          =   195
         Left            =   4410
         TabIndex        =   99
         ToolTipText     =   "See options"
         Top             =   5370
         WhatsThisHelpID =   179
         Width           =   1410
      End
      Begin VB.Label etat2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Log file not active"
         Height          =   195
         Left            =   4410
         TabIndex        =   98
         ToolTipText     =   "See options"
         Top             =   5565
         WhatsThisHelpID =   179
         Width           =   1410
      End
      Begin VB.Label etat3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Batch file not active"
         Height          =   195
         Left            =   4410
         TabIndex        =   97
         ToolTipText     =   "See options"
         Top             =   5775
         WhatsThisHelpID =   179
         Width           =   1410
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   6840
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
            Picture         =   "MainForm.frx":02E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0838
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":12E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1834
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1D88
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":22DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2830
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":32D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":382C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":3D80
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":42D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4828
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":52D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":5824
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":5D78
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":62CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6820
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar état 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8730
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14367
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
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "06/03/01"
            Object.ToolTipText     =   "system's date"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   979
            MinWidth        =   970
            TextSave        =   "07:54"
            Object.ToolTipText     =   "system's time"
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
      Begin VB.Menu msep11 
         Caption         =   "-"
      End
      Begin VB.Menu mabrev 
         Caption         =   "A&bbreviations..."
      End
      Begin VB.Menu msep50 
         Caption         =   "-"
      End
      Begin VB.Menu mundo 
         Caption         =   "&Undo last rename"
         Enabled         =   0   'False
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
         Shortcut        =   ^A
      End
      Begin VB.Menu msep58 
         Caption         =   "-"
      End
      Begin VB.Menu m1selectAll 
         Caption         =   "Select &All"
         HelpContextID   =   149
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
   End
   Begin VB.Menu mView 
      Caption         =   "&View"
      HelpContextID   =   157
      Begin VB.Menu Mrefresh 
         Caption         =   "Refres&h (F5)"
         HelpContextID   =   158
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
         Shortcut        =   ^H
      End
      Begin VB.Menu mchangetab 
         Caption         =   "&Change tab"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mrun 
      Caption         =   "&Run"
      HelpContextID   =   148
      Begin VB.Menu M2Start 
         Caption         =   "&Start"
         HelpContextID   =   153
      End
      Begin VB.Menu m2preview 
         Caption         =   "&Preview"
         HelpContextID   =   154
      End
      Begin VB.Menu m2manually 
         Caption         =   "&Rename manually"
         HelpContextID   =   155
      End
      Begin VB.Menu msep59 
         Caption         =   "-"
      End
      Begin VB.Menu m2recursive 
         Caption         =   "Recursive &mode"
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
         Shortcut        =   ^R
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
      Begin VB.Menu msepchg2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuswap 
         Caption         =   "Swap filenames"
         Visible         =   0   'False
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
         Caption         =   "Add..."
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
            Caption         =   "Ma&ke directory..."
         End
         Begin VB.Menu mdgroupe 
            Caption         =   "Make a &group of directories..."
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
         Caption         =   "Pefix"
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
         Caption         =   "Copy to clibpoard"
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
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
'    Dim frm1 As Form1
'    Dim frm3 As Form3
    Dim frm4 As Form4
'    Set frm1 = New Form1
'    frm1.Show
'    Set frm3 = New Form3
'    frm3.Show
    Set frm4 = New Form4
    frm4.Show
    
    Load frmDock
    frmDock.ShowWindow True
    Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload frmDock
End Sub

Private Sub Picture10_Click()
 If FrameGauche.width > 5000 Then
    Frame1.Visible = False
    Frame2.Visible = False
    Picture10.left = 10
    FrameGauche.width = 135
 Else
    FrameGauche.width = 5940
    Frame1.Visible = True
    Frame2.Visible = True
    Picture10.left = 5640
 End If
End Sub
