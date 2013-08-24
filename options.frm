VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Begin VB.Form doptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   6705
   ClientLeft      =   2775
   ClientTop       =   1665
   ClientWidth     =   7485
   ControlBox      =   0   'False
   HelpContextID   =   14
   Icon            =   "options.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   255
      Left            =   -960
      Sorted          =   -1  'True
      TabIndex        =   164
      Top             =   6600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6210
      HelpContextID   =   14
      Left            =   45
      TabIndex        =   140
      Top             =   60
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   10954
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "options.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame32"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Files Filters"
      TabPicture(1)   =   "options.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Text7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "List2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdDown"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdUp"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command10"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame49"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "options.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame12"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame13"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame22"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame25"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame27"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame31"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame33"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame41"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Frame42"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Frame46"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Display"
      TabPicture(3)   =   "options.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame20"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame28"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame26"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame34"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Bin"
      TabPicture(4)   =   "options.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame18"
      Tab(4).Control(1)=   "Frame17"
      Tab(4).Control(2)=   "Frame16"
      Tab(4).Control(3)=   "Frame15"
      Tab(4).Control(4)=   "Frame14"
      Tab(4).Control(5)=   "Label6"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Settings"
      TabPicture(5)   =   "options.frx":0098
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame43"
      Tab(5).Control(1)=   "Frame23"
      Tab(5).Control(2)=   "Frame21"
      Tab(5).ControlCount=   3
      TabCaption(6)   =   "Multimedia"
      TabPicture(6)   =   "options.frx":00B4
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame24"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Frame35"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame44"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "Frame45"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "Frame48"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "Frame47"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).ControlCount=   6
      TabCaption(7)   =   "Lists"
      TabPicture(7)   =   "options.frx":00D0
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Label16"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "Label11"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "Line1"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "Line2"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "Line3"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "Line4"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "Frame4"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "Frame7"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "Frame19"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "Frame6"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).Control(10)=   "Frame29"
      Tab(7).Control(10).Enabled=   0   'False
      Tab(7).Control(11)=   "Frame30"
      Tab(7).Control(11).Enabled=   0   'False
      Tab(7).ControlCount=   12
      TabCaption(8)   =   "Reg. Expressions"
      TabPicture(8)   =   "options.frx":00EC
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame37"
      Tab(8).Control(0).Enabled=   0   'False
      Tab(8).Control(1)=   "Frame38"
      Tab(8).Control(1).Enabled=   0   'False
      Tab(8).ControlCount=   2
      Begin VB.Frame Frame47 
         Caption         =   "EXIF Options "
         Height          =   615
         Left            =   -74880
         TabIndex        =   220
         Top             =   5460
         Width           =   7095
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   3660
            TabIndex        =   224
            Top             =   220
            Width           =   675
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   1440
            TabIndex        =   223
            Top             =   220
            Width           =   675
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Delimiter for time"
            Height          =   195
            Left            =   2340
            TabIndex        =   225
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Delimiter for date"
            Height          =   195
            Left            =   120
            TabIndex        =   222
            Top             =   270
            Width           =   1185
         End
      End
      Begin VB.Frame Frame49 
         Caption         =   "Define files extensions to view in the text preview tab "
         Height          =   2415
         Left            =   -74880
         TabIndex        =   213
         Top             =   3480
         Width           =   7095
         Begin VB.CommandButton Command18 
            Caption         =   "Delete"
            Height          =   375
            HelpContextID   =   14
            Left            =   2160
            TabIndex        =   218
            Top             =   900
            Width           =   1140
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Add"
            Height          =   375
            HelpContextID   =   14
            Left            =   2160
            TabIndex        =   217
            Top             =   480
            Width           =   1140
         End
         Begin VB.ListBox List5 
            Height          =   1380
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   216
            Top             =   840
            Width           =   1905
         End
         Begin VB.TextBox Text17 
            Height          =   285
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   214
            Top             =   480
            Width           =   1905
         End
         Begin VB.Label Label25 
            Caption         =   "Remember, to view a file you can also drop it from the files list to the text zone."
            Height          =   495
            Left            =   3480
            TabIndex        =   219
            Top             =   540
            Width           =   3375
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Extension"
            Height          =   195
            Left            =   120
            TabIndex        =   215
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.Frame Frame48 
         Caption         =   "How to preview pictures ? "
         Height          =   1425
         Left            =   -70320
         TabIndex        =   212
         Top             =   3990
         Width           =   2535
         Begin VB.OptionButton Option19 
            Caption         =   "Best Fit"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   103
            Top             =   1020
            Width           =   915
         End
         Begin VB.OptionButton Option19 
            Caption         =   "Stretch"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   102
            Top             =   675
            Width           =   1035
         End
         Begin VB.OptionButton Option19 
            Caption         =   "Real size"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   101
            Top             =   315
            Width           =   1035
         End
      End
      Begin VB.Frame Frame46 
         Caption         =   "Chinese/Japanese/Korean users "
         Height          =   645
         Left            =   -72300
         TabIndex        =   210
         Top             =   1080
         Width           =   4515
         Begin VB.CheckBox Check38 
            Caption         =   "Check this option"
            Height          =   255
            Left            =   120
            TabIndex        =   211
            Top             =   260
            Width           =   1575
         End
      End
      Begin VB.Frame Frame45 
         Caption         =   "Ogg options "
         Height          =   1425
         Left            =   -74865
         TabIndex        =   207
         Top             =   3990
         Width           =   4485
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   3120
            TabIndex        =   100
            Text            =   " - "
            Top             =   1020
            Width           =   375
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Combine tags together"
            Height          =   255
            Index           =   4
            Left            =   2310
            TabIndex        =   99
            Top             =   750
            Width           =   1965
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Use filled tag"
            Height          =   255
            Index           =   3
            Left            =   2310
            TabIndex        =   98
            Top             =   510
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Use the longest tag"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   97
            Top             =   990
            Width           =   1785
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Use the last tag"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   96
            Top             =   750
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Use the first tag"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   95
            Top             =   510
            Width           =   1455
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "with"
            Height          =   195
            Left            =   2700
            TabIndex        =   209
            Top             =   1110
            Width           =   285
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "When a tag is present several times"
            Height          =   195
            Left            =   210
            TabIndex        =   208
            Top             =   260
            Width           =   2520
         End
      End
      Begin VB.Frame Frame44 
         Caption         =   "Mp3, Vqf, Ogg and WMA options "
         Height          =   2175
         Left            =   -72030
         TabIndex        =   203
         Top             =   1740
         Width           =   4260
         Begin VB.CheckBox Check41 
            Caption         =   "Show only filled tags"
            Height          =   255
            Left            =   120
            TabIndex        =   221
            ToolTipText     =   "In the list of tags"
            Top             =   540
            Width           =   1815
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2340
            TabIndex        =   93
            Text            =   "0"
            Top             =   780
            Width           =   375
         End
         Begin VB.CheckBox Check40 
            Caption         =   "Remove multiple spaces"
            Height          =   255
            Left            =   1680
            TabIndex        =   92
            Top             =   300
            Width           =   2055
         End
         Begin VB.CheckBox Check39 
            Caption         =   "Separate words"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            ToolTipText     =   "It uses upper cases and lower cases to separate words"
            Top             =   300
            Width           =   1455
         End
         Begin VB.ListBox List4 
            Height          =   660
            IntegralHeight  =   0   'False
            ItemData        =   "options.frx":0108
            Left            =   120
            List            =   "options.frx":011B
            TabIndex        =   94
            Top             =   1380
            Width           =   2775
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "zeros"
            Height          =   195
            Left            =   2790
            TabIndex        =   206
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Complete tracks numbers with"
            Height          =   195
            Left            =   120
            TabIndex        =   205
            Top             =   840
            Width           =   2115
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Convert tags to"
            Height          =   195
            Left            =   120
            TabIndex        =   204
            Top             =   1140
            Width           =   1080
         End
      End
      Begin VB.Frame Frame43 
         Caption         =   "Default files "
         Height          =   1095
         Left            =   -74880
         TabIndex        =   202
         Top             =   2640
         Width           =   7095
         Begin VB.CommandButton Command16 
            Caption         =   "..."
            Height          =   285
            HelpContextID   =   14
            Left            =   6720
            TabIndex        =   85
            ToolTipText     =   "Click to select a file"
            Top             =   600
            Width           =   285
         End
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   3120
            TabIndex        =   84
            Top             =   600
            Width           =   3495
         End
         Begin VB.CommandButton Command15 
            Caption         =   "..."
            Height          =   285
            HelpContextID   =   14
            Left            =   6720
            TabIndex        =   82
            ToolTipText     =   "Click to select a file"
            Top             =   240
            Width           =   285
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   3120
            TabIndex        =   81
            Top             =   240
            Width           =   3495
         End
         Begin VB.CheckBox Check37 
            Caption         =   "Use a default file for abbreviations"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   620
            Width           =   2775
         End
         Begin VB.CheckBox Check36 
            Caption         =   "Use a default file for cyclic selections"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame42 
         Caption         =   "Starting spaces "
         Height          =   540
         Left            =   -70560
         TabIndex        =   201
         Top             =   3120
         Width           =   2745
         Begin VB.CheckBox Check35 
            Caption         =   "Automatically remove them"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   215
            Width           =   2415
         End
      End
      Begin VB.Frame Frame41 
         Caption         =   "Rules "
         Height          =   990
         Left            =   -70560
         TabIndex        =   200
         Top             =   3720
         Width           =   2745
         Begin VB.CheckBox Check21 
            Caption         =   "Use case insensitive string comparison"
            Height          =   435
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   2475
         End
      End
      Begin VB.Frame Frame38 
         Caption         =   "Back References "
         Height          =   930
         Left            =   -74880
         TabIndex        =   196
         Top             =   2940
         Width           =   7095
         Begin VB.OptionButton Option18 
            Caption         =   "Use $1, $2, $3... $9 notation"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   133
            Top             =   585
            Width           =   2355
         End
         Begin VB.OptionButton Option18 
            Caption         =   "Use \1, \2, \3... \9 notation"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   132
            Top             =   315
            Width           =   2310
         End
      End
      Begin VB.Frame Frame37 
         Caption         =   "Engine to use "
         Height          =   2370
         Left            =   -74880
         TabIndex        =   195
         Top             =   450
         Width           =   7095
         Begin VB.ComboBox Combo6 
            Height          =   315
            ItemData        =   "options.frx":0177
            Left            =   4320
            List            =   "options.frx":0196
            Style           =   2  'Dropdown List
            TabIndex        =   124
            Top             =   200
            Width           =   1695
         End
         Begin VB.Frame Frame40 
            Caption         =   "Common options "
            Height          =   555
            Left            =   135
            TabIndex        =   198
            Top             =   1680
            Width           =   6780
            Begin VB.CheckBox Check33 
               Caption         =   "Extended"
               Height          =   195
               Left            =   105
               TabIndex        =   131
               Top             =   255
               Width           =   1050
            End
         End
         Begin VB.Frame Frame39 
            Caption         =   "Perl Compatible regular expressions options "
            Height          =   780
            Left            =   315
            TabIndex        =   197
            Top             =   795
            Width           =   4575
            Begin VB.CheckBox Check34 
               Caption         =   "Ungreedy"
               Height          =   195
               Left            =   3195
               TabIndex        =   130
               Top             =   240
               Width           =   1035
            End
            Begin VB.CheckBox Check32 
               Caption         =   "Dot all"
               Height          =   195
               Left            =   1500
               TabIndex        =   129
               Top             =   495
               Width           =   840
            End
            Begin VB.CheckBox Check31 
               Caption         =   "Dollar end only"
               Height          =   195
               Left            =   1500
               TabIndex        =   128
               Top             =   270
               Width           =   1485
            End
            Begin VB.CheckBox Check30 
               Caption         =   "Anchored"
               Height          =   195
               Left            =   165
               TabIndex        =   127
               Top             =   495
               Width           =   1020
            End
            Begin VB.CheckBox Check29 
               Caption         =   "Extra"
               Height          =   195
               Left            =   165
               TabIndex        =   126
               Top             =   270
               Width           =   1035
            End
         End
         Begin VB.OptionButton Option17 
            Caption         =   "Use Perl Compatible regular expressions"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   125
            Top             =   525
            Width           =   3210
         End
         Begin VB.OptionButton Option17 
            Caption         =   "Use Unix regular expressions"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   123
            Top             =   240
            Width           =   2385
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Syntax to use"
            Height          =   195
            Left            =   3240
            TabIndex        =   199
            Top             =   240
            Width           =   960
         End
      End
      Begin VB.Frame Frame35 
         Caption         =   "Mp3 Tags "
         Height          =   2175
         Left            =   -74865
         TabIndex        =   193
         Top             =   1740
         Width           =   2775
         Begin VB.Frame Frame36 
            Caption         =   "When both tags are present "
            Height          =   975
            Left            =   300
            TabIndex        =   194
            Top             =   885
            Width           =   2340
            Begin VB.OptionButton Option16 
               Caption         =   "Set priority to filled tags"
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   90
               Top             =   615
               Width           =   1980
            End
            Begin VB.OptionButton Option16 
               Caption         =   "Set priority to V2 tags"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   89
               Top             =   315
               Width           =   1860
            End
         End
         Begin VB.OptionButton Option15 
            Caption         =   "Use V2 tags"
            Height          =   210
            Index           =   1
            Left            =   120
            TabIndex        =   88
            Top             =   615
            Width           =   1200
         End
         Begin VB.OptionButton Option15 
            Caption         =   "Use V1 tags"
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   87
            Top             =   300
            Width           =   1215
         End
      End
      Begin VB.Frame Frame34 
         Caption         =   "Preview Window "
         Height          =   855
         Left            =   -74880
         TabIndex        =   191
         Top             =   4680
         Width           =   7095
         Begin VB.CheckBox Check27 
            Caption         =   "Show only files whose names change"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   480
            Width           =   3015
         End
         Begin VB.CheckBox Check25 
            Caption         =   "Show gridlines"
            Height          =   255
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   64
            Top             =   225
            Width           =   1395
         End
      End
      Begin VB.Frame Frame23 
         Caption         =   "Autosave "
         Height          =   630
         Left            =   -74880
         TabIndex        =   190
         Top             =   1950
         Width           =   7095
         Begin VB.CheckBox Check12 
            Caption         =   "Use autosave"
            Height          =   255
            HelpContextID   =   14
            Left            =   150
            TabIndex        =   79
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame26 
         Caption         =   "Current Directory "
         Height          =   615
         Left            =   -74880
         TabIndex        =   189
         Top             =   3960
         Width           =   7095
         Begin VB.CheckBox Check15 
            Caption         =   "Show path in caption"
            Height          =   255
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   63
            Top             =   225
            Width           =   1875
         End
      End
      Begin VB.Frame Frame33 
         Caption         =   "Free Form "
         Height          =   540
         Left            =   -70560
         TabIndex        =   188
         Top             =   2475
         Width           =   2745
         Begin VB.CheckBox Check22 
            Caption         =   "Remember last command"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   "When you re open THE Rename it will recall your last command"
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame32 
         Caption         =   "Find long file names "
         Height          =   555
         Left            =   120
         TabIndex        =   186
         Top             =   5580
         Width           =   7095
         Begin VB.OptionButton Option14 
            Caption         =   "Filename only"
            Height          =   255
            Index           =   2
            Left            =   5460
            TabIndex        =   24
            Top             =   240
            Width           =   1275
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Path only"
            Height          =   255
            Index           =   1
            Left            =   2730
            TabIndex        =   23
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option14 
            Caption         =   "Check Path+Filename"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame31 
         Caption         =   "Separator for lists "
         Height          =   735
         Left            =   -74880
         TabIndex        =   185
         Top             =   5350
         Width           =   7095
         Begin VB.CheckBox Check20 
            Caption         =   "Remove   """
            Height          =   195
            Left            =   4080
            TabIndex        =   54
            ToolTipText     =   "When they are present"
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   2205
            TabIndex        =   53
            Top             =   240
            Width           =   495
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Other"
            Height          =   195
            Index           =   1
            Left            =   1455
            TabIndex        =   52
            Top             =   300
            Width           =   720
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Tabulation"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   300
            Width           =   1080
         End
      End
      Begin VB.Frame Frame30 
         Caption         =   "Include "
         Height          =   915
         Left            =   -74880
         TabIndex        =   184
         Top             =   4920
         Width           =   7095
         Begin VB.CheckBox Check19 
            Caption         =   "Virtual Folders"
            Height          =   195
            Left            =   135
            TabIndex        =   122
            Top             =   540
            Width           =   1335
         End
         Begin VB.CheckBox Check17 
            Caption         =   "Hidden Folders"
            Height          =   255
            Left            =   135
            TabIndex        =   121
            Top             =   225
            Width           =   1395
         End
      End
      Begin VB.Frame Frame29 
         Caption         =   "Include "
         Height          =   555
         Left            =   -74880
         TabIndex        =   183
         Top             =   4080
         Width           =   7095
         Begin VB.OptionButton Option12 
            Caption         =   "Folders only"
            Height          =   195
            Index           =   2
            Left            =   5400
            TabIndex        =   120
            Top             =   240
            Width           =   1155
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Files and Folders"
            Height          =   195
            Index           =   1
            Left            =   2640
            TabIndex        =   119
            Top             =   240
            Width           =   1515
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Files only"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   118
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Files to include "
         Height          =   1170
         Left            =   -71640
         TabIndex        =   181
         Top             =   645
         Width           =   3855
         Begin VB.CheckBox Check3 
            Caption         =   "Hidden"
            Height          =   255
            HelpContextID   =   14
            Left            =   180
            TabIndex        =   105
            ToolTipText     =   "Hidden files will be (or not) include in list containing files"
            Top             =   240
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox Check4 
            Caption         =   "System"
            Height          =   255
            HelpContextID   =   14
            Left            =   180
            TabIndex        =   106
            ToolTipText     =   "System files will be (or not) include in list containing files"
            Top             =   540
            Value           =   1  'Checked
            Width           =   810
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Read Only"
            Height          =   255
            HelpContextID   =   14
            Left            =   180
            TabIndex        =   107
            ToolTipText     =   "Read Only files will be (or not) include in list containing files"
            Top             =   840
            Value           =   1  'Checked
            Width           =   1095
         End
      End
      Begin VB.Frame Frame19 
         Caption         =   "Misc "
         Height          =   1080
         Left            =   -74880
         TabIndex        =   179
         Top             =   2940
         Width           =   7095
         Begin VB.CheckBox Check24 
            Caption         =   "Automatically select all files"
            Height          =   195
            Left            =   3375
            TabIndex        =   117
            Top             =   720
            Width           =   2235
         End
         Begin VB.CheckBox Check26 
            Caption         =   "Use Natural sort"
            Height          =   255
            Left            =   135
            TabIndex        =   114
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Auto arrange column header"
            Height          =   255
            HelpContextID   =   14
            Left            =   135
            TabIndex        =   112
            Top             =   225
            Value           =   1  'Checked
            Width           =   2400
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Full Row Select"
            Height          =   255
            HelpContextID   =   14
            Left            =   3375
            TabIndex        =   115
            Top             =   225
            Width           =   1470
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Show GridLines"
            Height          =   255
            HelpContextID   =   14
            Left            =   135
            TabIndex        =   113
            Top             =   473
            Width           =   2715
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Save columns width"
            Height          =   255
            HelpContextID   =   14
            Left            =   3375
            TabIndex        =   116
            Top             =   473
            Width           =   1875
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Files date "
         Height          =   1095
         Left            =   -74880
         TabIndex        =   178
         Top             =   1800
         Width           =   7095
         Begin VB.OptionButton Option7 
            Caption         =   "Last access"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   110
            Top             =   720
            Width           =   1335
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Last modified"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   109
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Creation date"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Personalize format"
            Height          =   285
            Left            =   2400
            TabIndex        =   111
            ToolTipText     =   "Enable you to personalize the display date"
            Top             =   450
            Width           =   1590
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "When double click on files list "
         Height          =   1170
         Left            =   -74880
         TabIndex        =   177
         Top             =   645
         Width           =   3135
         Begin VB.ListBox List1 
            Height          =   840
            HelpContextID   =   14
            ItemData        =   "options.frx":01E4
            Left            =   120
            List            =   "options.frx":01FA
            TabIndex        =   104
            Top             =   240
            Width           =   2940
         End
      End
      Begin VB.Frame Frame28 
         Caption         =   "Other "
         Height          =   615
         Left            =   -74880
         TabIndex        =   176
         Top             =   3240
         Width           =   7095
         Begin VB.CheckBox Check18 
            Caption         =   "Use Office 97 flat toolbar buttons"
            Height          =   255
            Left            =   135
            TabIndex        =   62
            Top             =   225
            Width           =   2640
         End
      End
      Begin VB.Frame Frame27 
         Caption         =   "Invalid characters"
         Height          =   540
         Left            =   -71040
         TabIndex        =   175
         Top             =   4755
         Width           =   3240
         Begin VB.CheckBox Check16 
            Caption         =   "Remove them"
            Height          =   255
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   50
            ToolTipText     =   "Invalid characters will be automatically removed from filenames"
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Startup directory "
         Height          =   930
         Left            =   120
         TabIndex        =   173
         Top             =   350
         Width           =   7095
         Begin VB.CheckBox Check23 
            Caption         =   "Remember last visited folder"
            Height          =   255
            Left            =   2040
            TabIndex        =   1
            ToolTipText     =   "THE Rename will open the last visited folder"
            Top             =   320
            Width           =   2295
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Select startup directory"
            Height          =   315
            HelpContextID   =   14
            Left            =   90
            TabIndex        =   0
            Top             =   285
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "<none>"
            Height          =   255
            Left            =   120
            TabIndex        =   174
            ToolTipText     =   "This is the current startup directory"
            Top             =   645
            Width           =   6855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Window's position "
         Height          =   1635
         Left            =   -74880
         TabIndex        =   172
         Top             =   360
         Width           =   7095
         Begin VB.CommandButton Command13 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6600
            TabIndex        =   57
            ToolTipText     =   "Restore THE Rename window to a normal position"
            Top             =   240
            Width           =   300
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            ItemData        =   "options.frx":0239
            Left            =   2400
            List            =   "options.frx":0246
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   1150
            Width           =   1335
         End
         Begin VB.CheckBox Check28 
            Caption         =   "Remember window size"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Save windows position when ending"
            Height          =   240
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   55
            ToolTipText     =   "Program will reposition itself to its former position"
            Top             =   270
            Width           =   2940
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Center windows on screen"
            Height          =   240
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   56
            ToolTipText     =   "Program will center itself in your screen"
            Top             =   540
            Width           =   2220
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Select program's startup state"
            Height          =   195
            Left            =   120
            TabIndex        =   192
            Top             =   1200
            Width           =   2085
         End
      End
      Begin VB.Frame Frame25 
         Caption         =   "Counters format "
         Height          =   645
         Left            =   -74880
         TabIndex        =   171
         Top             =   1800
         Width           =   7095
         Begin VB.CheckBox Check14 
            Caption         =   "Use lower cases for letter counters"
            Height          =   255
            HelpContextID   =   14
            Left            =   3900
            TabIndex        =   38
            Top             =   270
            Width           =   2775
         End
         Begin VB.CheckBox Check13 
            Caption         =   "Complete counters with 0s"
            Height          =   255
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   37
            Top             =   270
            Width           =   2175
         End
      End
      Begin VB.Frame Frame24 
         Caption         =   "Pictures width and Height format "
         Height          =   1230
         Left            =   -74865
         TabIndex        =   169
         Top             =   450
         Width           =   7095
         Begin VB.TextBox Text11 
            Height          =   285
            HelpContextID   =   14
            Left            =   135
            TabIndex        =   86
            Text            =   "%w%x%h%"
            Top             =   720
            Width           =   4995
         End
         Begin VB.Label Label15 
            Caption         =   "Type the exact format you would like to use. %w% represents picture's with and %h% represents picture's height. Unit is pixel."
            Height          =   420
            Left            =   135
            TabIndex        =   170
            Top             =   225
            Width           =   6840
         End
      End
      Begin VB.Frame Frame22 
         Caption         =   "Characters to delimit words"
         Height          =   690
         Left            =   -74880
         TabIndex        =   167
         Top             =   360
         Width           =   7095
         Begin VB.TextBox Text12 
            Height          =   285
            HelpContextID   =   14
            Left            =   5520
            TabIndex        =   35
            Text            =   " "
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox Text9 
            Height          =   285
            HelpContextID   =   14
            Left            =   1980
            TabIndex        =   34
            Text            =   " "
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "For tokens"
            Height          =   195
            Left            =   4620
            TabIndex        =   187
            Top             =   270
            Width           =   750
         End
         Begin VB.Label Label10 
            Caption         =   "For the Capitalize option"
            Height          =   240
            Left            =   180
            TabIndex        =   168
            Top             =   270
            Width           =   1815
         End
      End
      Begin VB.Frame Frame21 
         Caption         =   "Save settings"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   165
         Top             =   360
         Width           =   7095
         Begin VB.CommandButton Command11 
            Caption         =   "Select directory"
            Height          =   285
            HelpContextID   =   14
            Left            =   2520
            TabIndex        =   78
            Top             =   540
            Width           =   1410
         End
         Begin VB.OptionButton Option11 
            Caption         =   "Save to a specific directory"
            Height          =   330
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   77
            Top             =   495
            Width           =   2295
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Save file in current directory"
            Height          =   285
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   76
            Top             =   225
            Width           =   2310
         End
         Begin VB.Label Label9 
            Height          =   645
            Left            =   135
            TabIndex        =   166
            Top             =   855
            Width           =   6840
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "&Sort"
         Height          =   375
         HelpContextID   =   14
         Left            =   -72840
         TabIndex        =   31
         ToolTipText     =   "Sort list"
         Top             =   2040
         Width           =   1140
      End
      Begin VB.CommandButton cmdUp 
         Enabled         =   0   'False
         Height          =   330
         HelpContextID   =   14
         Left            =   -72840
         Picture         =   "options.frx":0268
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Move Up"
         Top             =   2520
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdDown 
         Enabled         =   0   'False
         Height          =   330
         HelpContextID   =   14
         Left            =   -72840
         Picture         =   "options.frx":036A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Move Down"
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Frame Frame20 
         Caption         =   "Startup default options"
         Height          =   1065
         Left            =   -74880
         TabIndex        =   161
         Top             =   2145
         Width           =   7095
         Begin VB.ComboBox Combo4 
            Height          =   315
            HelpContextID   =   14
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   600
            Width           =   5640
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            HelpContextID   =   14
            ItemData        =   "options.frx":046C
            Left            =   1215
            List            =   "options.frx":046E
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   210
            Width           =   5640
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "For extension"
            Height          =   195
            Left            =   135
            TabIndex        =   163
            Top             =   615
            Width           =   945
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "For prefix"
            Height          =   195
            Left            =   135
            TabIndex        =   162
            Top             =   240
            Width           =   645
         End
      End
      Begin VB.Frame Frame18 
         Caption         =   "Confirm make directory "
         Height          =   600
         Left            =   -74880
         TabIndex        =   160
         ToolTipText     =   "Gets if is user is prompted before creating any needed directories"
         Top             =   3360
         Width           =   7095
         Begin VB.OptionButton Option9 
            Caption         =   "No"
            Height          =   240
            HelpContextID   =   14
            Index           =   9
            Left            =   1560
            TabIndex        =   75
            Top             =   270
            Width           =   690
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Yes"
            Height          =   240
            HelpContextID   =   14
            Index           =   8
            Left            =   180
            TabIndex        =   74
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.Frame Frame17 
         Caption         =   "Allow undo "
         Height          =   600
         Left            =   -74880
         TabIndex        =   159
         ToolTipText     =   "Sets if Windows stores undo information (if possible)"
         Top             =   2670
         Width           =   7095
         Begin VB.OptionButton Option9 
            Caption         =   "Yes"
            Height          =   240
            HelpContextID   =   14
            Index           =   7
            Left            =   180
            TabIndex        =   72
            Top             =   270
            Width           =   690
         End
         Begin VB.OptionButton Option9 
            Caption         =   "No"
            Height          =   240
            HelpContextID   =   14
            Index           =   6
            Left            =   1530
            TabIndex        =   73
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Silent Mode "
         Height          =   600
         Left            =   -74880
         TabIndex        =   158
         ToolTipText     =   "If Yes, when you copy big files, a popup window will appear"
         Top             =   2025
         Width           =   7095
         Begin VB.OptionButton Option9 
            Caption         =   "Yes"
            Height          =   240
            HelpContextID   =   14
            Index           =   5
            Left            =   180
            TabIndex        =   70
            Top             =   270
            Width           =   690
         End
         Begin VB.OptionButton Option9 
            Caption         =   "No"
            Height          =   240
            HelpContextID   =   14
            Index           =   4
            Left            =   1530
            TabIndex        =   71
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Rename on collision "
         Height          =   600
         Left            =   -74880
         TabIndex        =   157
         ToolTipText     =   "Specifies if destination should be renamed if file with same name already exists"
         Top             =   1365
         Width           =   7095
         Begin VB.OptionButton Option9 
            Caption         =   "Yes"
            Height          =   240
            HelpContextID   =   14
            Index           =   3
            Left            =   180
            TabIndex        =   68
            Top             =   270
            Width           =   690
         End
         Begin VB.OptionButton Option9 
            Caption         =   "No"
            Height          =   240
            HelpContextID   =   14
            Index           =   2
            Left            =   1530
            TabIndex        =   69
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Confirm operation "
         Height          =   600
         Left            =   -74880
         TabIndex        =   156
         ToolTipText     =   "Sets if user is prompted for confirmation"
         Top             =   720
         Width           =   7095
         Begin VB.OptionButton Option9 
            Caption         =   "No"
            Height          =   240
            HelpContextID   =   14
            Index           =   1
            Left            =   1530
            TabIndex        =   67
            Top             =   270
            Width           =   690
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Yes"
            Height          =   240
            HelpContextID   =   14
            Index           =   0
            Left            =   180
            TabIndex        =   66
            Top             =   270
            Width           =   690
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "History "
         Height          =   540
         Left            =   -74880
         TabIndex        =   154
         Top             =   4755
         Width           =   3705
         Begin VB.CheckBox Check8 
            Caption         =   "Use history"
            Height          =   195
            HelpContextID   =   14
            Left            =   150
            TabIndex        =   49
            Top             =   250
            Width           =   1140
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Search && Replace and Abbreviations "
         Height          =   990
         Left            =   -74880
         TabIndex        =   153
         Top             =   3720
         Width           =   4215
         Begin VB.OptionButton Option8 
            Caption         =   "Execute it before and after process"
            Height          =   240
            HelpContextID   =   14
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   675
            Width           =   2820
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Execute it after process"
            Height          =   240
            HelpContextID   =   14
            Index           =   1
            Left            =   120
            TabIndex        =   46
            Top             =   450
            Width           =   1995
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Execute it before renaming files"
            Height          =   240
            HelpContextID   =   14
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   225
            Width           =   2715
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Counters and recursive mode "
         Height          =   1185
         Left            =   -74880
         TabIndex        =   152
         Top             =   2475
         Width           =   4215
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Height          =   285
            HelpContextID   =   14
            Left            =   2220
            TabIndex        =   42
            Top             =   810
            Width           =   780
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Don't restart counters"
            Height          =   240
            HelpContextID   =   14
            Index           =   0
            Left            =   120
            TabIndex        =   39
            Top             =   270
            Width           =   1860
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Restart counters when full path changes"
            Height          =   240
            HelpContextID   =   14
            Index           =   1
            Left            =   120
            TabIndex        =   40
            Top             =   540
            Width           =   3200
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Restart counters at level"
            Height          =   240
            HelpContextID   =   14
            Index           =   2
            Left            =   120
            TabIndex        =   41
            Top             =   855
            Width           =   2040
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Directory report && Html report "
         Height          =   645
         Left            =   -74880
         TabIndex        =   151
         Top             =   1080
         Width           =   2475
         Begin VB.CommandButton Command14 
            Caption         =   "Options..."
            Height          =   315
            Left            =   180
            TabIndex        =   36
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Delete"
         Height          =   375
         HelpContextID   =   14
         Left            =   -72840
         TabIndex        =   30
         Top             =   1605
         Width           =   1140
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Add"
         Height          =   375
         HelpContextID   =   14
         Left            =   -72840
         TabIndex        =   29
         Top             =   1170
         Width           =   1140
      End
      Begin VB.ListBox List2 
         Height          =   2205
         HelpContextID   =   14
         ItemData        =   "options.frx":0470
         Left            =   -74880
         List            =   "options.frx":0472
         TabIndex        =   28
         Top             =   1125
         Width           =   1905
      End
      Begin VB.TextBox Text7 
         Height          =   285
         HelpContextID   =   14
         Left            =   -74880
         TabIndex        =   27
         Top             =   750
         Width           =   1905
      End
      Begin VB.Frame Frame8 
         Caption         =   "Date and time format "
         Height          =   690
         Left            =   120
         TabIndex        =   147
         Top             =   4860
         Width           =   7095
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "options.frx":0474
            Left            =   4950
            List            =   "options.frx":0481
            TabIndex        =   21
            Text            =   "Long"
            Top             =   225
            Width           =   1740
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "options.frx":049A
            Left            =   630
            List            =   "options.frx":04B6
            TabIndex        =   20
            Text            =   "YYYYMMDD"
            Top             =   225
            Width           =   1740
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Time"
            Height          =   195
            Left            =   4455
            TabIndex        =   149
            Top             =   270
            Width           =   345
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Date"
            Height          =   195
            Left            =   135
            TabIndex        =   148
            Top             =   270
            Width           =   345
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Create a batch dos file instead of renaming files "
         Height          =   645
         Left            =   120
         TabIndex        =   146
         Top             =   2685
         Width           =   7095
         Begin VB.CommandButton Command9 
            Caption         =   "..."
            Height          =   285
            HelpContextID   =   14
            Left            =   6735
            TabIndex        =   13
            ToolTipText     =   "Browse for folder"
            Top             =   270
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Yes"
            Height          =   240
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "It's a batch file for Ms Dos (ver 7) that supports long files names, so, don't forget to add .bat at the end of the name"
            Top             =   270
            Width           =   690
         End
         Begin VB.TextBox Text1 
            Height          =   285
            HelpContextID   =   14
            Left            =   2025
            TabIndex        =   12
            Text            =   "rename.bat"
            ToolTipText     =   "This batch file is for Ms Dos 7 so it support long filenames. Don't forget to add .bat"
            Top             =   270
            Visible         =   0   'False
            Width           =   4635
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Other "
         Height          =   1305
         Left            =   120
         TabIndex        =   145
         Top             =   1305
         Width           =   7095
         Begin VB.CommandButton Command8 
            Caption         =   "..."
            Height          =   285
            HelpContextID   =   14
            Left            =   6735
            TabIndex        =   7
            ToolTipText     =   "Browse for folder"
            Top             =   600
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CommandButton Command7 
            Caption         =   "..."
            Height          =   285
            HelpContextID   =   14
            Left            =   6735
            TabIndex        =   4
            ToolTipText     =   "Browse for folder"
            Top             =   240
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Shutdown when finished"
            Height          =   195
            HelpContextID   =   14
            Left            =   4800
            TabIndex        =   10
            ToolTipText     =   "THE Rename will shutdown when the rename operation will be finished"
            Top             =   990
            Width           =   2130
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Copy files"
            Height          =   195
            HelpContextID   =   14
            Left            =   1530
            TabIndex        =   9
            ToolTipText     =   "Files will be copied"
            Top             =   990
            Width           =   1050
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Rename files"
            Height          =   195
            HelpContextID   =   14
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Files will be renamed"
            Top             =   990
            Width           =   1230
         End
         Begin VB.TextBox Text3 
            Height          =   285
            HelpContextID   =   14
            Left            =   2040
            TabIndex        =   3
            Text            =   "undo.bat"
            ToolTipText     =   "It will be a batch file for Ms Dos, so don't forget to add .bat to file's name"
            Top             =   270
            Visible         =   0   'False
            Width           =   4635
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Create an undo file"
            Height          =   195
            HelpContextID   =   14
            Index           =   0
            Left            =   120
            TabIndex        =   2
            ToolTipText     =   "It will be a batch file for Ms Dos, so don't forget to add .bat to file's name"
            Top             =   315
            Width           =   1635
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Create a log file"
            Height          =   195
            HelpContextID   =   14
            Index           =   1
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "This log file will contains date, time, original file's name and new name."
            Top             =   630
            Width           =   1410
         End
         Begin VB.TextBox Text2 
            Height          =   285
            HelpContextID   =   14
            Left            =   2025
            TabIndex        =   6
            Text            =   "log.txt"
            ToolTipText     =   "This log file will contain date, time, original file's name an new name."
            Top             =   600
            Visible         =   0   'False
            Width           =   4635
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Launch a program ...  "
         Height          =   1380
         Left            =   120
         TabIndex        =   141
         Top             =   3405
         Width           =   7095
         Begin VB.TextBox Text6 
            Height          =   285
            HelpContextID   =   14
            Left            =   2025
            TabIndex        =   14
            ToolTipText     =   "Enter a program to launch before  each file is renamed, include %1 to have file's name as parameter for program"
            Top             =   240
            Width           =   4635
         End
         Begin VB.CommandButton Command3 
            Caption         =   "..."
            Height          =   285
            HelpContextID   =   14
            Left            =   6735
            TabIndex        =   15
            ToolTipText     =   "Click to select a program"
            Top             =   225
            Width           =   285
         End
         Begin VB.CommandButton Command2 
            Caption         =   "..."
            Height          =   285
            HelpContextID   =   14
            Left            =   6735
            TabIndex        =   19
            ToolTipText     =   "Click to select a program"
            Top             =   945
            Width           =   285
         End
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   285
            HelpContextID   =   14
            Left            =   6735
            TabIndex        =   17
            ToolTipText     =   "Click to select a program"
            Top             =   570
            Width           =   285
         End
         Begin VB.TextBox Text5 
            Height          =   285
            HelpContextID   =   14
            Left            =   2025
            TabIndex        =   18
            ToolTipText     =   "Enter a program to launch when process is finished, include %1 to have file's name as parameter for program "
            Top             =   945
            Width           =   4635
         End
         Begin VB.TextBox Text4 
            Height          =   285
            HelpContextID   =   14
            Left            =   2025
            TabIndex        =   16
            ToolTipText     =   "Enter a program to launch after each file is renamed, include %1 to have file's name as parameter for program"
            Top             =   570
            Width           =   4635
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Before each file"
            Height          =   195
            Left            =   135
            TabIndex        =   144
            Top             =   270
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "At the end"
            Height          =   195
            Left            =   135
            TabIndex        =   143
            Top             =   990
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "After each file"
            Height          =   195
            Left            =   135
            TabIndex        =   142
            Top             =   630
            Width           =   975
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         X1              =   -73560
         X2              =   -68000
         Y1              =   4815
         Y2              =   4815
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         X1              =   -73560
         X2              =   -68000
         Y1              =   4800
         Y2              =   4800
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   -73920
         X2              =   -68000
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   -73920
         X2              =   -68000
         Y1              =   500
         Y2              =   500
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Directory List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   182
         Top             =   4680
         Width           =   1140
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Files List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   180
         Top             =   400
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "All these option will affect the bin and the multiple directory copy button"
         Height          =   195
         Left            =   -74910
         TabIndex        =   155
         Top             =   450
         Width           =   4995
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Filename  (i.e *.eps or *.eps;*.dat)"
         Height          =   195
         Left            =   -74880
         TabIndex        =   150
         Top             =   480
         Width           =   2340
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   136
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   139
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   138
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Sample 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   137
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   14
      Left            =   5160
      TabIndex        =   25
      ToolTipText     =   "Cancel your selection"
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   14
      Left            =   6360
      TabIndex        =   26
      ToolTipText     =   "Save settings for ALL sessions"
      Top             =   6360
      Width           =   1095
   End
End
Attribute VB_Name = "doptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim OldAbbrev As String
 Dim OldCyclic As String
 Dim DateFormat2 As Integer
 Dim OggOption As Integer
 Dim setr As Integer
 Dim Ouverture As Boolean
 Dim FilesToInclude2 As Integer
 Dim restart2 As Integer
 Dim CheckLongFileNameOption2 As Integer
 Dim FiltreActif As String
Private Sub Check1_Click(Index As Integer)
 If Check1(0).Value = 0 Then
    Text3.Visible = False
    Command7.Visible = False
 Else
    Text3.Visible = True
    Command7.Visible = True
 End If
 
 If Check1(1).Value = 0 Then
    Text2.Visible = False
    Command8.Visible = False
 Else
    Text2.Visible = True
    Command8.Visible = True
 End If
End Sub

Private Sub Check2_Click()
 If Check2.Value = 0 Then
    Text1.Visible = False
    Command9.Visible = False
 Else
    Text1.Visible = True
    Command9.Visible = True
 End If
End Sub

Private Sub Check36_Click()
    etat1
End Sub

Private Sub Check37_Click()
    etat2
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDown_Click()
    ButtonDown List2
End Sub

Private Sub cmdOK_Click()
 Dim i As Integer
 Dim vnb As Long
 Dim tlibdate(3) As String
 If Check36.Value = 1 And Trim$(Text13.Text) = "" Then
    MsgBox "Eror, you want to use an default cylic file but you did not specify it"
    SSTab1.Tab = 5
    Text13.SetFocus
    Exit Sub
 End If
 
 If Check37.Value = 1 And Trim$(Text14.Text) = "" Then
    MsgBox "Eror, you want to use an default abbreviations file but you did not specify it"
    SSTab1.Tab = 5
    Text14.SetFocus
    Exit Sub
 End If
 
 LesOptions.TextToView = ""
 For i = 0 To List5.ListCount - 1
    LesOptions.TextToView = LesOptions.TextToView + List5.List(i) + "|"
 Next
 LesOptions.TextToView = Left$(LesOptions.TextToView, Len(LesOptions.TextToView) - 1)
 If Option19(0).Value = True Then
    LesOptions.PicturesPreview = 0
 Else
    If Option19(1).Value = True Then
        LesOptions.PicturesPreview = 1
    Else
        LesOptions.PicturesPreview = 2
    End If
 End If
 LesOptions.ExifDelimD = Text18.Text
 LesOptions.ExifDelimT = Text19.Text
 LesOptions.ChineseKorean = Check38.Value
 LesOptions.RemoveEmptyTags = Check41.Value
 LesOptions.OggOpt1 = OggOption
 LesOptions.OggOpt2 = Text16.Text
 LesOptions.Mp3VqfOpt1 = Check39.Value
 LesOptions.Mp3VqfOpt2 = Check40.Value
 LesOptions.Mp3VqfOpt3 = List4.ListIndex
 LesOptions.Mp3VqfOpt4 = Text15.Text
 LesOptions.RxSyntax = Combo6.ListIndex
 LesOptions.SelectAllFiles = Check24.Value
 LesOptions.RulesOpt1 = Check21.Value
 If Option18(0).Value = True Then
    LesOptions.BackRefNotation = 0
 Else
    LesOptions.BackRefNotation = 1
 End If
 LesOptions.RegExOpt1 = Check33.Value
 LesOptions.PCRE1 = Check29.Value
 LesOptions.PCRE2 = Check30.Value
 LesOptions.PCRE3 = Check31.Value
 LesOptions.PCRE4 = Check32.Value
 LesOptions.PCRE5 = Check34.Value
 
 LesOptions.UseDefaultCyclicFile = Check36.Value
 LesOptions.DefaultCyclicFile = Text13.Text
 LesOptions.UseDefaultAbbrevFile = Check37.Value
 LesOptions.DefaultAbbrevFile = Text14.Text
 
 LesOptions.RemoveStartingSpaces = Check35.Value
 If Option17(0).Value = True Then
    LesOptions.RegExpEngine = 0
 Else
    LesOptions.RegExpEngine = 1
 End If
 If Option16(0).Value = True Then
    LesOptions.TagsPriority = 0
 Else
    If Option16(1).Value = True Then
        LesOptions.TagsPriority = 1
    Else
        LesOptions.TagsPriority = 2
    End If
 End If
 If Option15(0).Value = True Then
    LesOptions.TagsVersionToUse = 0
 Else
    LesOptions.TagsVersionToUse = 1
 End If
 LesOptions.FilesToInclude = FilesToInclude2
 If Option13(0).Value = True Then
    LesOptions.ListDelimiter = 9
 Else
    LesOptions.ListDelimiter = Asc(Text10.Text)
 End If
 LesOptions.wWindowState = Combo5.ListIndex
 LesOptions.RememberWSize = Check28.Value
 LesOptions.UseNaturalSort = Check26.Value
 LesOptions.ShowWhenFileNameChange = Check27.Value
 LesOptions.PreviewGridLines = Check25.Value
 LesOptions.RememberLastFolder = Check23.Value
 LesOptions.RememberLastCommand = Check22.Value
 LesOptions.CheckLongFileNameOption = CheckLongFileNameOption2
 LesOptions.CharTokens = Text12.Text
 LesOptions.RemoveGuill = Check20.Value
 LesOptions.ToolbarButtons = Check18.Value
 LesOptions.IncHiddenFolders = Check17.Value
 LesOptions.IncVirtualFolders = Check19.Value
 RENAME.FolderTreeview1(0).VirtualFolders = LesOptions.IncVirtualFolders
 RENAME.FolderTreeview1(0).HiddenFolders = LesOptions.IncHiddenFolders
 If LesOptions.ToolbarButtons = 1 Then
  RENAME.Toolbar1.Style = tbrFlat
 Else
  RENAME.Toolbar1.Style = tbrStandard
 End If
 LesOptions.RemoveIllegals = Check16.Value
 LesOptions.PicturesFormat = Text11.Text
 LesOptions.ShowPathInCaption = Check15.Value
 If LesOptions.ShowPathInCaption = 0 Then
    RENAME.Caption = AncTitre
 Else
    RENAME.Caption = "THE Rename - " + Dir1Path
 End If
 LesOptions.CompleCounters = Check13.Value
 LesOptions.UseLowerInLetterCounters = Check14.Value
 LesOptions.UseAutoSave = Check12.Value
 LesOptions.WordsDelimiters = Text9.Text
 If Option10.Value = True Then
    LesOptions.SettingsDirectory = ""
 Else
    LesOptions.SettingsDirectory = Label9.Caption
 End If
 LesOptions.DefOption1 = Combo3.ListIndex
 LesOptions.DefOption2 = Combo4.ListIndex
 If Check11.Value = 1 Then
    LesOptions.ColumnsWiths = True
 Else
    LesOptions.ColumnsWiths = False
 End If
 
 If Check9.Value = 1 Then
    LesOptions.FullRow = True
 Else
    LesOptions.FullRow = False
 End If
 
 If Check10.Value = 1 Then
    LesOptions.GridLines = True
 Else
    LesOptions.GridLines = False
 End If
 
 If Check7.Value = 1 Then
    LesOptions.ShutDown = True
 Else
    LesOptions.ShutDown = False
 End If
 
 If Option9(0).Value = True Then
    LesOptions.ConfirmOperation = True
 Else
    LesOptions.ConfirmOperation = False
 End If
 
 If Option9(3).Value = True Then
    LesOptions.RenameOnCollision = True
 Else
    LesOptions.RenameOnCollision = False
 End If
 
 If Option9(5).Value = True Then
    LesOptions.SilentMode = True
 Else
    LesOptions.SilentMode = False
 End If

 If Option9(7).Value = True Then
    LesOptions.AllowUndo = True
 Else
    LesOptions.AllowUndo = False
 End If
 
 If Option9(8).Value = True Then
    LesOptions.ConfirmMakeDir = True
 Else
    LesOptions.ConfirmMakeDir = False
 End If
 
 
 If Check8.Value = 1 Then
    LesOptions.UseHistory = True
    RENAME.mhistory.Enabled = True
 Else
    LesOptions.UseHistory = False
    RENAME.mhistory.Enabled = False
 End If
 LesOptions.SearchAndReplace = setr
 tlibdate(1) = "Created"
 tlibdate(2) = "Modified"
 tlibdate(3) = "Access"
 RENAME.ListView1.ColumnHeaders(3).Text = tlibdate(DateFormat2 + 1)
 LesOptions.StartupDir = Trim$(Label4.Caption)
 LesOptions.Dateformat = DateFormat2
 If Check6.Value = 1 Then
    LesOptions.AutoArrange = True
 Else
    LesOptions.AutoArrange = False
 End If
 
 LesOptions.ActDblClick = List1.ListIndex
 LesOptions.RestartCounter = restart2
 LesOptions.LevelRestart = Val(Text8.Text)
 RENAME.Combo5.Clear
 vnb = 0

 LesOptions.FormatDate = Combo1.ListIndex
 LesOptions.FormatTime = Combo2.ListIndex
 
 For i = 0 To List2.ListCount - 1
    RENAME.Combo5.AddItem List2.List(i)
 Next
 RENAME.Combo5.Text = FiltreActif
 
 LesOptions.prog1 = Text6.Text
 LesOptions.prog2 = Text4.Text
 LesOptions.prog3 = Text5.Text
 
 If Check2.Value = 1 Then
    LesOptions.batch = Text1.Text
 Else
    LesOptions.batch = ""
 End If
 
 If Check1(0).Value = 1 Then
    LesOptions.UndoFile = Text3.Text
 Else
    LesOptions.UndoFile = ""
 End If
 
 If Check1(1).Value = 1 Then
    LesOptions.LogFile = Text2.Text
 Else
    LesOptions.LogFile = ""
 End If
 
 If Option1.Value = True Then
    LesOptions.Center0rSave = 2
 Else
    LesOptions.Center0rSave = 1
 End If
  
 If Check1(0).Value = 1 Then
    LesOptions.UndoFile = Text3.Text
 End If
 
 If Option3.Value = True Then
    LesOptions.CopyRename = True
 Else
    LesOptions.CopyRename = False
 End If
  
 If Check3.Value = 1 Then
    LesOptions.Hidden = True
 Else
    LesOptions.Hidden = False
 End If
 
 If Check4.Value = 1 Then
    LesOptions.System = True
 Else
    LesOptions.System = False
 End If
 
 If Check5.Value = 1 Then
    LesOptions.ReadOnly = True
 Else
    LesOptions.ReadOnly = False
 End If
  
' Savesettings
 
 vnb = RENAME.remplissage
 RENAME.état.Panels(3).Text = Trim$(Str$(vnb))
 
 RENAME.ListView1.FullRowSelect = LesOptions.FullRow
 RENAME.ListView1.GridLines = LesOptions.GridLines
 RENAME.ListView2.GridLines = LesOptions.GridLines
 
 ' Test sur les abbréviations
 If LesOptions.UseDefaultAbbrevFile = 1 And (OldAbbrev <> LesOptions.DefaultAbbrevFile) Then
    If MsgBox("You have selected to use an abbreviation file, would you like to load it now ?", vbYesNo, "Abbreviation file") = vbYes Then
        OpenAbbrev Text14.Text
    End If
 Else   ' On n'utilise plus, faut il supprimer ce qu'il y a à l'heure actuelle ?
    If LesOptions.UseDefaultAbbrevFile = 0 Then
        If (OkUseAbbrev = True Or CollAbrev.Count > 0) Then
            If MsgBox("Would you like to remove all you current abbreviations ?", vbYesNo, "Abbreviations") = vbYes Then
                RemoveAbbrev
            End If
        End If
    End If
 End If
 
 ' Test sur le fichier de sélections cycliques
 If LesOptions.UseDefaultCyclicFile = 1 And (OldCyclic <> LesOptions.DefaultCyclicFile) Then
    If MsgBox("You have selected to use a default file for cyclic selections, would you like to load it now ?", vbYesNo, "Cyclic file") = vbYes Then
        OpenCyclic Text13.Text
    End If
 Else   ' On n'utilise plus, faut il supprimer ce qu'il y a à l'heure actuelle ?
    If LesOptions.UseDefaultCyclicFile = 0 Then
        If (UseCylcic = True Or VnbCyclic > 0) Then
            If MsgBox("Would you like to remove your current cyclic selections ?", vbYesNo, "Cyclic selections") = vbYes Then
                ReDim LesCyclic(0)
                VnbCyclic = 0
            End If
        End If
    End If
 End If
 Unload Me
End Sub

Private Sub cmdUp_Click()
    ButtonUp List2
End Sub

Private Sub Combo1_Click()
    If Ouverture = False Then
        If Trim$(Combo1.List(Combo1.ListIndex)) = "Other ..." Then
            QDateTravail = 1
            DtOther.Show 1
        End If
    Else
        Ouverture = False
    End If
End Sub

Private Sub Command1_Click()
 Text4.Text = DialogFile(Me.hWnd, 1, "Open", "*.exe", "EXE" & Chr$(0) & "*.exe" & Chr$(0) & "All files" & Chr$(0) & "*.*", App.Path, "exe")
 Text4.SetFocus
End Sub

Private Sub Command10_Click()
 Dim i As Integer
 List3.Clear
 For i = 0 To List2.ListCount - 1
    List3.AddItem List2.List(i)
 Next
 List2.Visible = False
 List2.Clear
 For i = 0 To List3.ListCount - 1
    List2.AddItem List3.List(i)
 Next
 List2.Visible = True
End Sub

Private Sub Command11_Click()
 Dim szFilename As String
 szFilename = BrowseFolder(Me, "Select a directory")
 If Len(Trim$(szFilename)) = 0 Then Exit Sub
 szFilename = AddBackSlash(Trim$(szFilename))
 Label9.Caption = szFilename + Text3.Text
End Sub

Private Sub Command12_Click()
  QDateTravail = 2
  DtOther.Show 1
End Sub

Private Sub Command13_Click()
    RENAME.WindowState = vbNormal
    RENAME.height = 7740
    RENAME.width = 11085
End Sub

Private Sub Command14_Click()
    doptions2.Show 1
End Sub

Private Sub Command15_Click()
Dim szFilename As String
szFilename = DialogFile(Me.hWnd, 1, "Open Cyclic selections", "cyclic.cyc", "Cyclic File" & Chr$(0) & "*.cyc" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "Cyclic File")
If Trim$(szFilename) = "" Then Exit Sub
Text13.Text = szFilename
Text13.SetFocus
End Sub

Private Sub Command16_Click()
Dim szFilename As String
szFilename = DialogFile(Me.hWnd, 1, "Open Abbreviations", "abbrev.abr", "Abbreviation" & Chr$(0) & "*.abr" & Chr$(0) & "All files" & Chr$(0) & "*.*", LesOptions.SettingsDirectory, "Abbreviation")
If Trim$(szFilename) = "" Then Exit Sub
Text14.Text = szFilename
Text14.SetFocus
End Sub

Private Sub Command17_Click()
 Dim i As Integer
 If Len(Trim$(Text17.Text)) = 0 Then Exit Sub
 For i = 0 To List5.ListCount
  If Trim$(UCase$(List5.List(i))) = Trim$(UCase$(Text17.Text)) Then
   MsgBox "Error, this extension is already in the list"
   Exit Sub
  End If
 Next
 List5.AddItem UCase$(Text17.Text)
 Text17.Text = ""
 Text17.SetFocus
End Sub

Private Sub Command18_Click()
 If List5.ListIndex = -1 Then Exit Sub
 List5.RemoveItem (List5.ListIndex)
End Sub

Private Sub Command2_Click()
 Text5.Text = DialogFile(Me.hWnd, 1, "Open", "*.exe", "EXE" & Chr$(0) & "*.exe" & Chr$(0) & "All files" & Chr$(0) & "*.*", App.Path, "exe")
 Text5.SetFocus
End Sub

Private Sub Command3_Click()
 Text6.Text = DialogFile(Me.hWnd, 1, "Open", "*.exe", "EXE" & Chr$(0) & "*.exe" & Chr$(0) & "All files" & Chr$(0) & "*.*", App.Path, "exe")
 Text6.SetFocus
End Sub
Private Sub Command4_Click()
 Dim i As Integer
 If Len(Trim$(Text7.Text)) = 0 Then
  MsgBox "You must enter a file filter to add it !"
  Exit Sub
 End If
 For i = 0 To List2.ListCount
  If Trim$(UCase$(List2.List(i))) = Trim$(UCase$(Text7.Text)) Then
   MsgBox "Error, this filter is already in the list"
   Exit Sub
  End If
 Next
 List2.AddItem Text7.Text
 Text7.Text = ""
 Text7.SetFocus
End Sub

Private Sub Command5_Click()
 If List2.ListIndex = -1 Then
  Exit Sub
 End If
 If List2.ListCount > 1 Then
  List2.RemoveItem (List2.ListIndex)
 Else
  MsgBox "You must have at least one filter !"
 End If
End Sub

Private Sub Command6_Click()
 Dim szFilename As String
 szFilename = BrowseFolder(Me, "Select startup directory")
 If Len(Trim$(szFilename)) = 0 Then Exit Sub
 Label4 = szFilename
End Sub

Private Sub Command7_Click()
 Dim szFilename As String
 szFilename = BrowseFolder(Me, "Select a directory")
 If Len(Trim$(szFilename)) = 0 Then Exit Sub
 szFilename = AddBackSlash(Trim$(szFilename))
 Text3.Text = szFilename + Text3.Text
 Text3.SetFocus
End Sub

Private Sub Command8_Click()
 Dim szFilename As String
 szFilename = BrowseFolder(Me, "Select a directory")
 If Len(Trim$(szFilename)) = 0 Then Exit Sub
 szFilename = AddBackSlash(Trim$(szFilename))
 Text2.Text = szFilename + Text2.Text
 Text2.SetFocus
End Sub

Private Sub Command9_Click()
 Dim szFilename As String
 szFilename = BrowseFolder(Me, "Select a directory")
 If Len(Trim$(szFilename)) = 0 Then Exit Sub
 szFilename = AddBackSlash(Trim$(szFilename))
 Text1.Text = szFilename + Text1.Text
 Text1.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeTab KeyCode, Shift, SSTab1
    KeyCode = 0
End Sub

Private Sub Form_Load()
 Dim i As Integer
 Dim vnb As Integer
 Dim vtmp As String
 Ouverture = True
 List5.Clear
 FiltreActif = RENAME.Combo5.Text
 vtmp = ""
 vnb = Len(LesOptions.TextToView)
 For i = 1 To vnb
    If Mid$(LesOptions.TextToView, i, 1) <> "|" Then
        vtmp = vtmp + Mid$(LesOptions.TextToView, i, 1)
    Else
        List5.AddItem vtmp
        vtmp = ""
    End If
 Next
 If vtmp <> "" Then
    List5.AddItem vtmp
 End If
 Text18.Text = LesOptions.ExifDelimD
 Text19.Text = LesOptions.ExifDelimT
 Option19(LesOptions.PicturesPreview).Value = True
 Combo6.ListIndex = LesOptions.RxSyntax
 Option18(LesOptions.BackRefNotation).Value = True
 Check38.Value = LesOptions.ChineseKorean
 Check24.Value = LesOptions.SelectAllFiles
 Check21.Value = LesOptions.RulesOpt1
 Check29.Value = LesOptions.PCRE1
 Check30.Value = LesOptions.PCRE2
 Check31.Value = LesOptions.PCRE3
 Check32.Value = LesOptions.PCRE4
 Check34.Value = LesOptions.PCRE5
 Check33.Value = LesOptions.RegExOpt1
 Check35.Value = LesOptions.RemoveStartingSpaces
 Check41.Value = LesOptions.RemoveEmptyTags
 
 Option5(LesOptions.OggOpt1).Value = True
 OggOption = LesOptions.OggOpt1
 etat3
 Text16.Text = LesOptions.OggOpt2
 
 Check39.Value = LesOptions.Mp3VqfOpt1
 Check40.Value = LesOptions.Mp3VqfOpt2
 List4.ListIndex = LesOptions.Mp3VqfOpt3
 Text15.Text = LesOptions.Mp3VqfOpt4
 Check36.Value = LesOptions.UseDefaultCyclicFile
 Text13.Text = LesOptions.DefaultCyclicFile
 OldCyclic = LesOptions.DefaultCyclicFile
 etat1
 
 Check37.Value = LesOptions.UseDefaultAbbrevFile
 Text14.Text = LesOptions.DefaultAbbrevFile
 OldAbbrev = LesOptions.DefaultAbbrevFile
 etat2
 
 Option15(LesOptions.TagsVersionToUse).Value = True
 Option15_Click LesOptions.TagsVersionToUse
 If LesOptions.TagsPriority = 2 Then
    LesOptions.TagsPriority = 0
 End If
 Option16(LesOptions.TagsPriority).Value = True
 Option17(LesOptions.RegExpEngine).Value = True
 Option17_Click LesOptions.RegExpEngine
 If LesOptions.ListDelimiter = 9 Then
    Option13(0).Value = True
 Else
    Option13(1).Value = True
    Text10.Text = Chr$(LesOptions.ListDelimiter)
 End If
 Combo5.ListIndex = LesOptions.wWindowState
 Check28.Value = LesOptions.RememberWSize
 Check27.Value = LesOptions.ShowWhenFileNameChange
 Check25.Value = LesOptions.PreviewGridLines
 Check26.Value = LesOptions.UseNaturalSort
 Option14(LesOptions.CheckLongFileNameOption).Value = True ' Option à utiliser pour controler la longueur des fichiers
 CheckLongFileNameOption2 = LesOptions.CheckLongFileNameOption
 Option12(LesOptions.FilesToInclude).Value = True ' Files to include
 Text12.Text = LesOptions.CharTokens
 Check23.Value = LesOptions.RememberLastFolder
 Check22.Value = LesOptions.RememberLastCommand
 Check20.Value = LesOptions.RemoveGuill
 Check17.Value = LesOptions.IncHiddenFolders
 Check19.Value = LesOptions.IncVirtualFolders
 FilesToInclude2 = LesOptions.FilesToInclude
 Check13.Value = LesOptions.CompleCounters
 Check16.Value = LesOptions.RemoveIllegals
 Check15.Value = LesOptions.ShowPathInCaption
 Check14.Value = LesOptions.UseLowerInLetterCounters
 Check18.Value = LesOptions.ToolbarButtons
 Text9.Text = LesOptions.WordsDelimiters
 Check12.Value = LesOptions.UseAutoSave
 Text11.Text = LesOptions.PicturesFormat
 If Len(Trim$(LesOptions.SettingsDirectory)) > 0 Then
  Label9.Caption = LesOptions.SettingsDirectory
  Option10.Value = False
  Option11.Value = True
 Else
  Option10.Value = True
  Option11.Value = False
 End If
 
 vnb = RENAME.Combo1.ListCount - 1
 For i = 0 To vnb
  Combo3.AddItem RENAME.Combo1.List(i)
 Next
 For i = 0 To RENAME.Combo2.ListCount - 1
  Combo4.AddItem RENAME.Combo2.List(i)
 Next
 Combo3.ListIndex = LesOptions.DefOption1
 Combo4.ListIndex = LesOptions.DefOption2
 If LesOptions.ColumnsWiths = True Then
  Check11.Value = 1
 Else
  Check11.Value = 0
 End If
 
 If LesOptions.FullRow = True Then
  Check9.Value = 1
 Else
  Check9.Value = 0
 End If
 
 If LesOptions.GridLines = True Then
  Check10.Value = 1
 Else
  Check10.Value = 0
 End If
 If LesOptions.ShutDown = True Then
  Check7.Value = 1
 Else
  Check7.Value = 0
 End If
 If LesOptions.ConfirmOperation = True Then
  Option9(0).Value = True
 Else
  Option9(1).Value = True
 End If
 
 If LesOptions.RenameOnCollision = True Then
  Option9(3).Value = True
 Else
  Option9(2).Value = True
 End If
 
 If LesOptions.SilentMode = True Then
  Option9(5).Value = True
 Else
  Option9(4).Value = True
 End If

 If LesOptions.AllowUndo = True Then
  Option9(7).Value = True
 Else
  Option9(6).Value = True
 End If
 
 If LesOptions.ConfirmMakeDir = True Then
  Option9(8).Value = True
 Else
  Option9(9).Value = True
 End If
 
 If LesOptions.UseHistory = True Then
  Check8.Value = 1
 Else
  Check8.Value = 0
 End If
 Option8(LesOptions.SearchAndReplace).Value = True
 Label4.Caption = LesOptions.StartupDir
 Option7(LesOptions.Dateformat).Value = True
 If LesOptions.AutoArrange = True Then
  Check6.Value = 1
 Else
  Check6.Value = 0
 End If
 restart2 = LesOptions.RestartCounter
 Option6(LesOptions.RestartCounter).Value = True
 Text8.Text = LesOptions.LevelRestart
 List2.Clear
 
 Combo1.ListIndex = LesOptions.FormatDate
 Combo2.ListIndex = LesOptions.FormatTime
 
 For i = 0 To RENAME.Combo5.ListCount - 1
  List2.AddItem RENAME.Combo5.List(i)
 Next
 
 List1.ListIndex = LesOptions.ActDblClick
 Text6.Text = LesOptions.prog1
 Text4.Text = LesOptions.prog2
 Text5.Text = LesOptions.prog3
 
 Text3.Text = LesOptions.UndoFile
 If Len(LesOptions.UndoFile) > 0 Then
  Check1(0).Value = 1
 Else
  Check1(0).Value = 0
 End If
 
 Text2.Text = LesOptions.LogFile
 If Len(LesOptions.LogFile) > 0 Then
  Check1(1).Value = 1
 Else
  Check1(1).Value = 0
 End If
  
 Text1.Text = LesOptions.batch
 If Len(Trim$(LesOptions.batch)) > 0 Then
  Check2.Value = 1
 Else
  Check2.Value = 0
 End If
 If LesOptions.Center0rSave = 1 Then
  Option1.Value = False
  Option2.Value = True
 Else
  Option1.Value = True
  Option2.Value = False
 End If
 
 If Len(LesOptions.UndoFile) <> 0 Then
  Check1(0).Value = 1
  Text3.Text = LesOptions.UndoFile
 Else
  Check1(0).Value = 0
  Text3.Text = ""
 End If
 
 If LesOptions.Hidden = True Then
  Check3.Value = 1
 Else
  Check3.Value = 0
 End If
 
 If LesOptions.System = True Then
  Check4.Value = 1
 Else
  Check4.Value = 0
 End If
 
 If LesOptions.ReadOnly = True Then
  Check5.Value = 1
 Else
  Check5.Value = 0
 End If
 
 If Len(LesOptions.LogFile) <> 0 Then
  Check1(1).Value = 1
  Text2.Text = LesOptions.LogFile
 Else
  Check1(1).Value = 0
  Text2.Text = ""
 End If

 If LesOptions.CopyRename = True Then
  Option3.Value = True
  Option4.Value = False
 Else
  Option3.Value = False
  Option4.Value = True
 End If
 
 Text1.Text = LesOptions.batch
End Sub
Private Sub List2_Click()
    SetListButtons List2, cmdUp, cmdDown
End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then Command5_Click
End Sub
Private Sub List5_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then Command18_Click
End Sub

Private Sub Option12_Click(Index As Integer)
    FilesToInclude2 = Index
End Sub

Private Sub Option14_Click(Index As Integer)
    CheckLongFileNameOption2 = Index
End Sub

Private Sub Option15_Click(Index As Integer)
    If Index = 0 Then
        Frame36.Enabled = False
        Option16(0).Enabled = False
        Option16(1).Enabled = False
    Else
        Frame36.Enabled = True
        Option16(0).Enabled = True
        Option16(1).Enabled = True
    End If
End Sub

Private Sub Option17_Click(Index As Integer)
    If Index = 0 Then
        Frame39.Enabled = False
        Check29.Enabled = False
        Check30.Enabled = False
        Check31.Enabled = False
        Check32.Enabled = False
        Check34.Enabled = False
        Label18.Enabled = True
        Combo6.Enabled = True
    Else
        Frame39.Enabled = True
        Check29.Enabled = True
        Check30.Enabled = True
        Check31.Enabled = True
        Check32.Enabled = True
        Check34.Enabled = True
        Label18.Enabled = False
        Combo6.Enabled = False
    End If
End Sub

Private Sub Option5_Click(Index As Integer)
    etat3
    OggOption = Index
End Sub

Private Sub Option6_Click(Index As Integer)
    restart2 = Index
End Sub

Private Sub Option7_Click(Index As Integer)
    DateFormat2 = Index
End Sub

Private Sub Option8_Click(Index As Integer)
    setr = Index
End Sub
Private Sub Text1_GotFocus()
    SelAll Text1
End Sub
Private Sub Text10_GotFocus()
    SelAll Text10
End Sub
Private Sub Text11_GotFocus()
    SelAll Text11
End Sub

Private Sub Text12_GotFocus()
    SelAll Text12
End Sub

Private Sub Text2_GotFocus()
    SelAll Text2
End Sub

Private Sub Text3_GotFocus()
    SelAll Text3
End Sub
Private Sub Text4_GotFocus()
    SelAll Text4
End Sub
Private Sub Text5_GotFocus()
    SelAll Text5
End Sub

Private Sub Text6_GotFocus()
    SelAll Text6
End Sub

Private Sub Text7_GotFocus()
    SelAll Text7
End Sub
Private Sub Text8_GotFocus()
    SelAll Text8
End Sub

Private Sub Text9_GotFocus()
    SelAll Text9
End Sub

Private Sub etat1()
    If Check36.Value = 1 Then
        Text13.Visible = True
        Command15.Visible = True
    Else
        Text13.Visible = False
        Text13.Text = ""
        Command15.Visible = False
    End If
End Sub

Private Sub etat2()
    If Check37.Value = 1 Then
        Text14.Visible = True
        Command16.Visible = True
    Else
        Text14.Visible = False
        Text14.Text = ""
        Command16.Visible = False
    End If
End Sub

Private Sub etat3()
If Option5(4).Value = True Then
    Text16.Enabled = True
    Label23.Enabled = True
Else
    Text16.Enabled = False
    Label23.Enabled = False
End If
End Sub
