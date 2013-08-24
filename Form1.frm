VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   8820
   ScaleWidth      =   9465
   Begin VB.PictureBox FrameDroite 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   8295
      Left            =   0
      ScaleHeight     =   8295
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   0
      Width           =   8475
      Begin VB.Frame Frame1 
         Caption         =   "Prefix "
         Height          =   2600
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Width           =   5895
         Begin VB.CommandButton Command3 
            Caption         =   "="
            Height          =   315
            Left            =   5400
            TabIndex        =   95
            ToolTipText     =   "Use same option for prefix and extension"
            Top             =   240
            WhatsThisHelpID =   177
            Width           =   255
         End
         Begin VB.PictureBox PanelPrefix 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1815
            Left            =   90
            ScaleHeight     =   1815
            ScaleWidth      =   5715
            TabIndex        =   48
            Top             =   630
            Width           =   5715
            Begin VB.CommandButton Command1 
               Caption         =   "Search && Replace..."
               Height          =   300
               HelpContextID   =   110
               Left            =   3060
               TabIndex        =   50
               ToolTipText     =   "Click on this button to use search and replace"
               Top             =   1400
               WhatsThisHelpID =   180
               Width           =   1635
            End
            Begin VB.CommandButton Command10 
               Caption         =   "Reset"
               Height          =   300
               HelpContextID   =   110
               Left            =   4800
               TabIndex        =   49
               ToolTipText     =   "Reset all selections to their default's value"
               Top             =   1400
               WhatsThisHelpID =   181
               Width           =   870
            End
            Begin TabDlg.SSTab SSTab1 
               Height          =   1335
               Left            =   0
               TabIndex        =   51
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
               TabPicture(0)   =   "Form1.frx":0000
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "Check3"
               Tab(0).Control(1)=   "onglcounter"
               Tab(0).ControlCount=   2
               TabCaption(1)   =   "Size"
               TabPicture(1)   =   "Form1.frx":001C
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Picture5"
               Tab(1).Control(1)=   "Check5"
               Tab(1).ControlCount=   2
               TabCaption(2)   =   "Date"
               TabPicture(2)   =   "Form1.frx":0038
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Check6"
               Tab(2).Control(1)=   "Picture6"
               Tab(2).ControlCount=   2
               TabCaption(3)   =   "Time"
               TabPicture(3)   =   "Form1.frx":0054
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Picture7"
               Tab(3).Control(1)=   "Check7"
               Tab(3).ControlCount=   2
               TabCaption(4)   =   "Folder's name"
               TabPicture(4)   =   "Form1.frx":0070
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "Command19"
               Tab(4).ControlCount=   1
               TabCaption(5)   =   "Text "
               TabPicture(5)   =   "Form1.frx":008C
               Tab(5).ControlEnabled=   -1  'True
               Tab(5).Control(0)=   "Text14"
               Tab(5).Control(0).Enabled=   0   'False
               Tab(5).Control(1)=   "Option1(1)"
               Tab(5).Control(1).Enabled=   0   'False
               Tab(5).Control(2)=   "Picture1"
               Tab(5).Control(2).Enabled=   0   'False
               Tab(5).Control(3)=   "Text2"
               Tab(5).Control(3).Enabled=   0   'False
               Tab(5).Control(4)=   "Option1(0)"
               Tab(5).Control(4).Enabled=   0   'False
               Tab(5).Control(5)=   "Command8"
               Tab(5).Control(5).Enabled=   0   'False
               Tab(5).ControlCount=   6
               TabCaption(6)   =   "Pictures"
               TabPicture(6)   =   "Form1.frx":00A8
               Tab(6).ControlEnabled=   0   'False
               Tab(6).Control(0)=   "Check1"
               Tab(6).Control(1)=   "Picture8"
               Tab(6).ControlCount=   2
               TabCaption(7)   =   "Sounds"
               TabPicture(7)   =   "Form1.frx":00C4
               Tab(7).ControlEnabled=   0   'False
               Tab(7).Control(0)=   "Command5"
               Tab(7).ControlCount=   1
               Begin VB.CommandButton Command8 
                  Height          =   300
                  HelpContextID   =   168
                  Left            =   5325
                  Picture         =   "Form1.frx":00E0
                  Style           =   1  'Graphical
                  TabIndex        =   94
                  ToolTipText     =   "Use cyclic selection"
                  Top             =   495
                  Visible         =   0   'False
                  WhatsThisHelpID =   198
                  Width           =   300
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "Options..."
                  Height          =   375
                  Left            =   -72720
                  TabIndex        =   93
                  Top             =   600
                  WhatsThisHelpID =   204
                  Width           =   1455
               End
               Begin VB.CheckBox Check1 
                  Caption         =   "Add picture's width and height"
                  Height          =   255
                  HelpContextID   =   162
                  Left            =   -74880
                  TabIndex        =   92
                  ToolTipText     =   "See the options to modify it's format"
                  Top             =   360
                  WhatsThisHelpID =   202
                  Width           =   2445
               End
               Begin VB.PictureBox Picture8 
                  BorderStyle     =   0  'None
                  Height          =   300
                  Left            =   -74640
                  ScaleHeight     =   300
                  ScaleWidth      =   4815
                  TabIndex        =   88
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   203
                  Width           =   4815
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     HelpContextID   =   162
                     Index           =   17
                     Left            =   15
                     TabIndex        =   91
                     ToolTipText     =   "Add to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   203
                     Width           =   1365
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     HelpContextID   =   162
                     Index           =   16
                     Left            =   1650
                     TabIndex        =   90
                     ToolTipText     =   "Add to the right"
                     Top             =   0
                     WhatsThisHelpID =   203
                     Width           =   1395
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     HelpContextID   =   162
                     Index           =   15
                     Left            =   3375
                     TabIndex        =   89
                     ToolTipText     =   "Replace prefix"
                     Top             =   0
                     WhatsThisHelpID =   203
                     Width           =   2295
                  End
               End
               Begin VB.CheckBox Check3 
                  Caption         =   "Add a counter"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   87
                  ToolTipText     =   "Click to add a counter"
                  Top             =   360
                  WhatsThisHelpID =   183
                  Width           =   1365
               End
               Begin VB.CheckBox Check7 
                  Caption         =   "Add file's time"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   86
                  ToolTipText     =   "Click to add file's time"
                  Top             =   360
                  WhatsThisHelpID =   193
                  Width           =   1320
               End
               Begin VB.PictureBox Picture7 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   -74640
                  ScaleHeight     =   285
                  ScaleWidth      =   4695
                  TabIndex        =   82
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   194
                  Width           =   4695
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   9
                     Left            =   15
                     TabIndex        =   85
                     ToolTipText     =   "File's time will be added to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   194
                     Width           =   1365
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   10
                     Left            =   1650
                     TabIndex        =   84
                     ToolTipText     =   "File's time will be added to the right"
                     Top             =   0
                     WhatsThisHelpID =   194
                     Width           =   1440
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     Index           =   11
                     Left            =   3375
                     TabIndex        =   83
                     ToolTipText     =   "File's time will replace prefix"
                     Top             =   0
                     WhatsThisHelpID =   194
                     Width           =   2295
                  End
               End
               Begin VB.CheckBox Check6 
                  Caption         =   "Add file's date"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   81
                  ToolTipText     =   "Click to add file's date"
                  Top             =   360
                  WhatsThisHelpID =   191
                  Width           =   1365
               End
               Begin VB.PictureBox Picture6 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -74640
                  ScaleHeight     =   240
                  ScaleWidth      =   5055
                  TabIndex        =   77
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   192
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     Index           =   6
                     Left            =   3375
                     TabIndex        =   80
                     ToolTipText     =   "File's date will replace prefix"
                     Top             =   0
                     WhatsThisHelpID =   192
                     Width           =   1575
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   7
                     Left            =   1650
                     TabIndex        =   79
                     ToolTipText     =   "File's date will be added to the right"
                     Top             =   0
                     WhatsThisHelpID =   192
                     Width           =   1410
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   8
                     Left            =   15
                     TabIndex        =   78
                     ToolTipText     =   "File's date will be added to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   192
                     Width           =   1365
                  End
               End
               Begin VB.PictureBox Picture5 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -74640
                  ScaleHeight     =   240
                  ScaleWidth      =   5055
                  TabIndex        =   73
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   190
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace prefix"
                     Height          =   255
                     Index           =   5
                     Left            =   3375
                     TabIndex        =   76
                     ToolTipText     =   "File's size will replace prefix"
                     Top             =   0
                     WhatsThisHelpID =   190
                     Width           =   1470
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   3
                     Left            =   0
                     TabIndex        =   75
                     ToolTipText     =   "File's size will be added to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   190
                     Width           =   1320
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   4
                     Left            =   1650
                     TabIndex        =   74
                     ToolTipText     =   "File's size will be added to the right"
                     Top             =   0
                     WhatsThisHelpID =   190
                     Width           =   1440
                  End
               End
               Begin VB.CheckBox Check5 
                  Caption         =   "Add file's size"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   72
                  ToolTipText     =   "Click to add file's size"
                  Top             =   360
                  WhatsThisHelpID =   189
                  Width           =   1320
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Replace with text"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   71
                  ToolTipText     =   "Click to replace or not file's prefix with a text"
                  Top             =   525
                  WhatsThisHelpID =   196
                  Width           =   1635
               End
               Begin VB.TextBox Text2 
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   70
                  ToolTipText     =   "Current prefix will be replace with this text"
                  Top             =   480
                  Visible         =   0   'False
                  WhatsThisHelpID =   197
                  Width           =   2775
               End
               Begin VB.PictureBox Picture1 
                  AutoRedraw      =   -1  'True
                  AutoSize        =   -1  'True
                  BorderStyle     =   0  'None
                  Height          =   255
                  Left            =   3960
                  ScaleHeight     =   255
                  ScaleWidth      =   1335
                  TabIndex        =   67
                  Top             =   900
                  Width           =   1335
                  Begin VB.OptionButton Option2 
                     Caption         =   "End"
                     Height          =   195
                     Index           =   1
                     Left            =   720
                     TabIndex        =   69
                     ToolTipText     =   "Text will be added at the end of prefix"
                     Top             =   0
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   615
                  End
                  Begin VB.OptionButton Option2 
                     Caption         =   "Begin"
                     Height          =   195
                     Index           =   0
                     Left            =   0
                     TabIndex        =   68
                     ToolTipText     =   "Text will be added to the left of prefix"
                     Top             =   0
                     Value           =   -1  'True
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   735
                  End
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Add text"
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   66
                  ToolTipText     =   "Click to add or remove a text to the prefix"
                  Top             =   870
                  WhatsThisHelpID =   199
                  Width           =   960
               End
               Begin VB.TextBox Text14 
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   65
                  ToolTipText     =   "Add text to the prefix"
                  Top             =   825
                  Visible         =   0   'False
                  WhatsThisHelpID =   200
                  Width           =   2055
               End
               Begin VB.CommandButton Command19 
                  Caption         =   "Options..."
                  Height          =   375
                  Left            =   -72720
                  TabIndex        =   64
                  Top             =   600
                  WhatsThisHelpID =   195
                  Width           =   1455
               End
               Begin VB.PictureBox onglcounter 
                  AutoRedraw      =   -1  'True
                  BorderStyle     =   0  'None
                  Height          =   580
                  Left            =   -74880
                  ScaleHeight     =   585
                  ScaleWidth      =   5475
                  TabIndex        =   52
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   5475
                  Begin VB.TextBox Text3 
                     Height          =   285
                     Left            =   795
                     TabIndex        =   60
                     Text            =   "1"
                     ToolTipText     =   "Enter begin's value"
                     Top             =   35
                     WhatsThisHelpID =   184
                     Width           =   495
                  End
                  Begin VB.TextBox Text4 
                     Height          =   285
                     Left            =   1980
                     TabIndex        =   59
                     Text            =   "1"
                     ToolTipText     =   "Enter increment value (1 for example)"
                     Top             =   35
                     WhatsThisHelpID =   185
                     Width           =   495
                  End
                  Begin VB.TextBox Text5 
                     Height          =   285
                     Left            =   3285
                     TabIndex        =   58
                     Text            =   "4"
                     ToolTipText     =   "Enter number of digits for the counter"
                     Top             =   35
                     WhatsThisHelpID =   186
                     Width           =   495
                  End
                  Begin VB.PictureBox Picture2 
                     BorderStyle     =   0  'None
                     Height          =   255
                     Left            =   210
                     ScaleHeight     =   255
                     ScaleWidth      =   4815
                     TabIndex        =   54
                     Top             =   360
                     WhatsThisHelpID =   188
                     Width           =   4815
                     Begin VB.OptionButton Option3 
                        Caption         =   "Replace prefix"
                        Height          =   255
                        Index           =   2
                        Left            =   3375
                        TabIndex        =   57
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
                        TabIndex        =   56
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
                        TabIndex        =   55
                        ToolTipText     =   "Counter will be added to the left of prefix"
                        Top             =   0
                        Value           =   -1  'True
                        WhatsThisHelpID =   188
                        Width           =   1455
                     End
                  End
                  Begin VB.ComboBox Combo3 
                     Height          =   315
                     ItemData        =   "Form1.frx":01CA
                     Left            =   4095
                     List            =   "Form1.frx":01DD
                     Style           =   2  'Dropdown List
                     TabIndex        =   53
                     ToolTipText     =   "Choose file's counter representation (Dec/Hex/Oct)"
                     Top             =   35
                     WhatsThisHelpID =   187
                     Width           =   1300
                  End
                  Begin VB.Label Label4 
                     AutoSize        =   -1  'True
                     Caption         =   "Begin"
                     Height          =   195
                     Left            =   225
                     TabIndex        =   63
                     Top             =   60
                     WhatsThisHelpID =   184
                     Width           =   405
                  End
                  Begin VB.Label Label5 
                     AutoSize        =   -1  'True
                     Caption         =   "Step"
                     Height          =   195
                     Left            =   1455
                     TabIndex        =   62
                     Top             =   60
                     WhatsThisHelpID =   185
                     Width           =   330
                  End
                  Begin VB.Label Label6 
                     AutoSize        =   -1  'True
                     Caption         =   "Digits"
                     Height          =   195
                     Left            =   2730
                     TabIndex        =   61
                     Top             =   60
                     WhatsThisHelpID =   186
                     Width           =   390
                  End
               End
            End
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form1.frx":020F
            Left            =   120
            List            =   "Form1.frx":0211
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   47
            ToolTipText     =   "Scroll to select action to perform on prefix"
            Top             =   240
            WhatsThisHelpID =   175
            Width           =   5175
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Extension "
         Height          =   2380
         Left            =   0
         TabIndex        =   4
         Top             =   2640
         Width           =   5895
         Begin VB.CommandButton Command4 
            Caption         =   "="
            Height          =   315
            Left            =   5400
            TabIndex        =   45
            ToolTipText     =   "Use same option for prefix and extension"
            Top             =   240
            WhatsThisHelpID =   177
            Width           =   255
         End
         Begin VB.PictureBox PanelExt 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1650
            Left            =   120
            ScaleHeight     =   1650
            ScaleWidth      =   5715
            TabIndex        =   6
            Top             =   650
            Width           =   5715
            Begin VB.CommandButton Command2 
               Caption         =   "Search && Replace..."
               Height          =   300
               HelpContextID   =   110
               Left            =   3060
               TabIndex        =   8
               ToolTipText     =   "Click on this button to use search and replace"
               Top             =   1350
               WhatsThisHelpID =   180
               Width           =   1635
            End
            Begin VB.CommandButton Command11 
               Caption         =   "Reset"
               Height          =   300
               HelpContextID   =   110
               Left            =   4800
               TabIndex        =   7
               ToolTipText     =   "Reset all selections to their default's value"
               Top             =   1350
               WhatsThisHelpID =   181
               Width           =   870
            End
            Begin TabDlg.SSTab SSTab2 
               Height          =   1300
               Left            =   0
               TabIndex        =   9
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
               TabPicture(0)   =   "Form1.frx":0213
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "onglcounter2"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Check11"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).ControlCount=   2
               TabCaption(1)   =   "Size"
               TabPicture(1)   =   "Form1.frx":022F
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Picture13"
               Tab(1).Control(1)=   "Check12"
               Tab(1).ControlCount=   2
               TabCaption(2)   =   "Date"
               TabPicture(2)   =   "Form1.frx":024B
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "Check13"
               Tab(2).Control(1)=   "Picture14"
               Tab(2).ControlCount=   2
               TabCaption(3)   =   "Time"
               TabPicture(3)   =   "Form1.frx":0267
               Tab(3).ControlEnabled=   0   'False
               Tab(3).Control(0)=   "Check4"
               Tab(3).Control(1)=   "Picture4"
               Tab(3).ControlCount=   2
               TabCaption(4)   =   "Text "
               TabPicture(4)   =   "Form1.frx":0283
               Tab(4).ControlEnabled=   0   'False
               Tab(4).Control(0)=   "Picture3"
               Tab(4).Control(1)=   "Option4(1)"
               Tab(4).Control(2)=   "Option4(0)"
               Tab(4).Control(3)=   "Text15"
               Tab(4).Control(4)=   "Text8"
               Tab(4).ControlCount=   5
               Begin VB.CheckBox Check12 
                  Caption         =   "Add file's size"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   44
                  ToolTipText     =   "Click to add file's size"
                  Top             =   360
                  WhatsThisHelpID =   189
                  Width           =   1320
               End
               Begin VB.CheckBox Check4 
                  Caption         =   "Add file's time"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   43
                  ToolTipText     =   "Click to add file's time"
                  Top             =   360
                  WhatsThisHelpID =   193
                  Width           =   1275
               End
               Begin VB.PictureBox Picture4 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -74640
                  ScaleHeight     =   240
                  ScaleWidth      =   5055
                  TabIndex        =   39
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   194
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace extension"
                     Height          =   255
                     Index           =   12
                     Left            =   3375
                     TabIndex        =   42
                     ToolTipText     =   "Click to replace extension will file's time"
                     Top             =   0
                     WhatsThisHelpID =   194
                     Width           =   2295
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   13
                     Left            =   1650
                     TabIndex        =   41
                     ToolTipText     =   "File's time will be added to to the right"
                     Top             =   0
                     WhatsThisHelpID =   194
                     Width           =   1440
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   14
                     Left            =   15
                     TabIndex        =   40
                     ToolTipText     =   "File's time will be added to to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   194
                     Width           =   1365
                  End
               End
               Begin VB.CheckBox Check13 
                  Caption         =   "Add file's date"
                  Height          =   255
                  Left            =   -74880
                  TabIndex        =   38
                  ToolTipText     =   "Click to add file's date"
                  Top             =   360
                  WhatsThisHelpID =   191
                  Width           =   1320
               End
               Begin VB.PictureBox Picture14 
                  BorderStyle     =   0  'None
                  Height          =   285
                  Left            =   -74640
                  ScaleHeight     =   285
                  ScaleWidth      =   5055
                  TabIndex        =   34
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   192
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   30
                     Left            =   15
                     TabIndex        =   37
                     ToolTipText     =   "Click to add file's date to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   192
                     Width           =   1365
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   31
                     Left            =   1650
                     TabIndex        =   36
                     ToolTipText     =   "Click to add file's date to the right"
                     Top             =   0
                     WhatsThisHelpID =   192
                     Width           =   1440
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace extension"
                     Height          =   255
                     Index           =   32
                     Left            =   3375
                     TabIndex        =   35
                     ToolTipText     =   "Click to replace file's extension with file's date"
                     Top             =   0
                     WhatsThisHelpID =   192
                     Width           =   2295
                  End
               End
               Begin VB.PictureBox Picture13 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -74640
                  ScaleHeight     =   240
                  ScaleWidth      =   5055
                  TabIndex        =   30
                  Top             =   720
                  Visible         =   0   'False
                  WhatsThisHelpID =   190
                  Width           =   5055
                  Begin VB.OptionButton Option3 
                     Caption         =   "Replace extension"
                     Height          =   255
                     Index           =   27
                     Left            =   3375
                     TabIndex        =   33
                     ToolTipText     =   "Click to replace file's extension with file's size"
                     Top             =   0
                     WhatsThisHelpID =   190
                     Width           =   2295
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the right"
                     Height          =   255
                     Index           =   28
                     Left            =   1650
                     TabIndex        =   32
                     ToolTipText     =   "Click to add file's size to the right"
                     Top             =   0
                     WhatsThisHelpID =   190
                     Width           =   1395
                  End
                  Begin VB.OptionButton Option3 
                     Caption         =   "Add to the left"
                     Height          =   255
                     Index           =   29
                     Left            =   15
                     TabIndex        =   31
                     ToolTipText     =   "Click to add file's size to the left"
                     Top             =   0
                     Value           =   -1  'True
                     WhatsThisHelpID =   190
                     Width           =   1365
                  End
               End
               Begin VB.CheckBox Check11 
                  Caption         =   "Add a counter"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   29
                  ToolTipText     =   "Click to add a counter"
                  Top             =   360
                  WhatsThisHelpID =   183
                  Width           =   1320
               End
               Begin VB.PictureBox Picture3 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   -71070
                  ScaleHeight     =   240
                  ScaleWidth      =   1335
                  TabIndex        =   26
                  Top             =   900
                  Width           =   1335
                  Begin VB.OptionButton Option5 
                     Caption         =   "End"
                     Height          =   195
                     Index           =   1
                     Left            =   720
                     TabIndex        =   28
                     ToolTipText     =   "Text will be added at the end of extension"
                     Top             =   0
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   615
                  End
                  Begin VB.OptionButton Option5 
                     Caption         =   "Begin"
                     Height          =   195
                     Index           =   0
                     Left            =   0
                     TabIndex        =   27
                     ToolTipText     =   "Text will be added to the left of extension"
                     Top             =   0
                     Value           =   -1  'True
                     Visible         =   0   'False
                     WhatsThisHelpID =   201
                     Width           =   855
                  End
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "Add text"
                  Height          =   195
                  Index           =   1
                  Left            =   -74880
                  TabIndex        =   25
                  ToolTipText     =   "Click to add or remove a text to the extension"
                  Top             =   870
                  WhatsThisHelpID =   199
                  Width           =   945
               End
               Begin VB.OptionButton Option4 
                  Caption         =   "Replace with text"
                  Height          =   195
                  Index           =   0
                  Left            =   -74880
                  TabIndex        =   24
                  ToolTipText     =   "Click to replace or not file's extension with a text"
                  Top             =   525
                  WhatsThisHelpID =   196
                  Width           =   1635
               End
               Begin VB.TextBox Text15 
                  Height          =   285
                  Left            =   -73200
                  TabIndex        =   23
                  ToolTipText     =   "Add a text to the extension"
                  Top             =   840
                  Visible         =   0   'False
                  WhatsThisHelpID =   200
                  Width           =   2055
               End
               Begin VB.TextBox Text8 
                  Height          =   285
                  Left            =   -73200
                  TabIndex        =   22
                  ToolTipText     =   "Current extension will be replace with this text"
                  Top             =   480
                  Visible         =   0   'False
                  WhatsThisHelpID =   197
                  Width           =   2775
               End
               Begin VB.PictureBox onglcounter2 
                  AutoRedraw      =   -1  'True
                  BorderStyle     =   0  'None
                  Height          =   580
                  Left            =   120
                  ScaleHeight     =   585
                  ScaleWidth      =   5475
                  TabIndex        =   10
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   5475
                  Begin VB.TextBox Text16 
                     Height          =   285
                     Left            =   795
                     TabIndex        =   18
                     Text            =   "1"
                     ToolTipText     =   "Enter begin's value"
                     Top             =   35
                     WhatsThisHelpID =   184
                     Width           =   495
                  End
                  Begin VB.PictureBox Picture11 
                     BorderStyle     =   0  'None
                     Height          =   255
                     Left            =   165
                     ScaleHeight     =   255
                     ScaleWidth      =   5055
                     TabIndex        =   14
                     Top             =   345
                     WhatsThisHelpID =   188
                     Width           =   5055
                     Begin VB.OptionButton Option3 
                        Caption         =   "Add to the left"
                        Height          =   255
                        Index           =   26
                        Left            =   15
                        TabIndex        =   17
                        ToolTipText     =   "Counter will be added to the left of extension"
                        Top             =   0
                        Value           =   -1  'True
                        WhatsThisHelpID =   188
                        Width           =   1365
                     End
                     Begin VB.OptionButton Option3 
                        Caption         =   "Add to the right"
                        Height          =   255
                        Index           =   25
                        Left            =   1665
                        TabIndex        =   16
                        ToolTipText     =   "Counter will be added to the right of extension"
                        Top             =   0
                        WhatsThisHelpID =   188
                        Width           =   1440
                     End
                     Begin VB.OptionButton Option3 
                        Caption         =   "Replace extension"
                        Height          =   255
                        Index           =   24
                        Left            =   3375
                        TabIndex        =   15
                        ToolTipText     =   "Counter will replace file's prefixe"
                        Top             =   0
                        WhatsThisHelpID =   188
                        Width           =   2295
                     End
                  End
                  Begin VB.TextBox Text18 
                     Height          =   285
                     Left            =   3285
                     TabIndex        =   13
                     Text            =   "4"
                     ToolTipText     =   "Enter number of digits for the counter"
                     Top             =   35
                     WhatsThisHelpID =   186
                     Width           =   495
                  End
                  Begin VB.TextBox Text17 
                     Height          =   285
                     Left            =   1980
                     TabIndex        =   12
                     Text            =   "1"
                     ToolTipText     =   "Enter increment value (1 for example)"
                     Top             =   35
                     WhatsThisHelpID =   185
                     Width           =   495
                  End
                  Begin VB.ComboBox Combo4 
                     Height          =   315
                     ItemData        =   "Form1.frx":029F
                     Left            =   4095
                     List            =   "Form1.frx":02B2
                     Style           =   2  'Dropdown List
                     TabIndex        =   11
                     ToolTipText     =   "Choose file's counter representation (Dec/Hex/Oct)"
                     Top             =   35
                     WhatsThisHelpID =   187
                     Width           =   1300
                  End
                  Begin VB.Label Label14 
                     AutoSize        =   -1  'True
                     Caption         =   "Digits"
                     Height          =   195
                     Left            =   2730
                     TabIndex        =   21
                     Top             =   60
                     WhatsThisHelpID =   186
                     Width           =   390
                  End
                  Begin VB.Label Label13 
                     AutoSize        =   -1  'True
                     Caption         =   "Step"
                     Height          =   195
                     Left            =   1455
                     TabIndex        =   20
                     Top             =   60
                     WhatsThisHelpID =   185
                     Width           =   330
                  End
                  Begin VB.Label Label12 
                     AutoSize        =   -1  'True
                     Caption         =   "Begin"
                     Height          =   195
                     Left            =   180
                     TabIndex        =   19
                     Top             =   60
                     WhatsThisHelpID =   184
                     Width           =   405
                  End
               End
            End
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            ToolTipText     =   "Scroll to select action to perform on extension"
            Top             =   240
            WhatsThisHelpID =   176
            Width           =   5175
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Sample "
         Height          =   700
         Left            =   0
         TabIndex        =   1
         Top             =   5040
         WhatsThisHelpID =   178
         Width           =   4335
         Begin VB.Label laidep 
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   120
            TabIndex        =   3
            ToolTipText     =   "Sample of what prefix will be"
            Top             =   195
            WhatsThisHelpID =   178
            Width           =   4095
         End
         Begin VB.Label laides 
            BackStyle       =   0  'Transparent
            Height          =   210
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "Sample of what extension will be"
            Top             =   405
            WhatsThisHelpID =   178
            Width           =   4095
         End
      End
      Begin VB.Label etat3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Batch file not active"
         Height          =   195
         Left            =   4410
         TabIndex        =   98
         ToolTipText     =   "See options"
         Top             =   5535
         WhatsThisHelpID =   179
         Width           =   1410
      End
      Begin VB.Label etat2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Log file not active"
         Height          =   195
         Left            =   4410
         TabIndex        =   97
         ToolTipText     =   "See options"
         Top             =   5325
         WhatsThisHelpID =   179
         Width           =   1410
      End
      Begin VB.Label etat1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Undo file not active"
         Height          =   195
         Left            =   4410
         TabIndex        =   96
         ToolTipText     =   "See options"
         Top             =   5130
         WhatsThisHelpID =   179
         Width           =   1410
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
