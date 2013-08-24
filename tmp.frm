   AutoRedraw      =   -1  
   
   ClientHeight    =   7095
   ClientLeft      =   165
   ClientTop       =   -795
   ClientWidth     =   10995
   
   KeyPreview      =   -1  
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   10995
   WhatsThisButton =   -1  
   WhatsThisHelp   =   -1  
   Begin VB.PictureBox panelcmd 
      AutoRedraw      =   -1  
      BorderStyle     =   0  
      Height          =   1920
      Left            =   120
      ScaleHeight     =   1920
      ScaleWidth      =   5685
      TabIndex        =   20
      Top             =   7620
      Visible         =   0   
      Width           =   5685
      Begin VB.CommandButton Command14 
         Height          =   330
         Left            =   5340
            Style           =   1  
         TabIndex        =   133
         
         Top             =   45
         WhatsThisHelpID =   250
         Width           =   360
      End
      Begin VB.ListBox Combo6 
         Height          =   450
         Left            =   1140
         Sorted          =   -1  
         TabIndex        =   31
         Top             =   405
         Visible         =   0   
         Width           =   3195
      End
      Begin VB.CommandButton Command6 
         Height          =   300
         HelpContextID   =   168
         Left            =   5280
         :04D4
         Style           =   1  
         TabIndex        =   30
         
         Top             =   1200
         WhatsThisHelpID =   198
         Width           =   300
      End
      Begin VB.TextBox cmdtxt3 
         Alignment       =   1  
         Height          =   285
         HelpContextID   =   41
         Left            =   5070
         TabIndex        =   26
         
         
         Top             =   1530
         WhatsThisHelpID =   211
         Width           =   375
      End
      Begin VB.TextBox cmdtxt1 
         Alignment       =   1  
         Height          =   285
         HelpContextID   =   41
         Left            =   3960
         TabIndex        =   25
         
         
         Top             =   1215
         WhatsThisHelpID =   209
         Width           =   375
      End
      Begin VB.TextBox cmdtxt2 
         Alignment       =   1  
         Height          =   285
         HelpContextID   =   41
         Left            =   4830
         TabIndex        =   24
         
         
         Top             =   1215
         WhatsThisHelpID =   210
         Width           =   375
      End
      Begin VB.CommandButton cmdclear 
         Height          =   330
         HelpContextID   =   41
         Left            =   4970
         :05BE
         Style           =   1  
         TabIndex        =   23
         
         Top             =   45
         WhatsThisHelpID =   205
         Width           =   360
      End
      Begin VB.ListBox listcmd 
         Height          =   1230
         ItemData        =   "rename.frx":0788
         Left            =   45
         List            =   "rename.frx":078A
         Sorted          =   -1  
         TabIndex        =   22
         
         Top             =   540
         Width           =   2175
      End
      Begin VB.TextBox txtlang 
         Height          =   330
         HelpContextID   =   41
         Left            =   45
         TabIndex        =   21
         
         Top             =   45
         WhatsThisHelpID =   206
         Width           =   4890
      End
      Begin VB.Label lang4 
         AutoSize        =   -1  
         
         Height          =   195
         Left            =   2310
         TabIndex        =   29
         Top             =   1575
         WhatsThisHelpID =   211
         Width           =   2655
      End
      Begin VB.Label lang3 
         AutoSize        =   -1  
         
         Height          =   195
         Left            =   4380
         TabIndex        =   28
         Top             =   1245
         WhatsThisHelpID =   210
         Width           =   330
      End
      Begin VB.Label lang2 
         AutoSize        =   -1  
         
         Height          =   195
         Left            =   2310
         TabIndex        =   27
         Top             =   1245
         WhatsThisHelpID =   209
         Width           =   1500
      End
   End
   Begin TabDlg.SSTab TabGen 
      Height          =   6260
      Left            =   4785
      TabIndex        =   32
      Top             =   480
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      
      TabPicture(0)   =   "rename.frx":078C
      Tab(0).ControlEnabled=   0   
      Tab(0).Control(0)=   "FolderTreeview1(0)"
      Tab(0).ControlCount=   1
      
      TabPicture(1)   =   "rename.frx":07A8
      Tab(1).ControlEnabled=   -1  
      Tab(1).Control(0)=   "FrameDroite"
      Tab(1).Control(0).Enabled=   0   
      Tab(1).ControlCount=   1
      
      TabPicture(2)   =   "rename.frx":07C4
      Tab(2).ControlEnabled=   0   
      Tab(2).Control(0)=   "LvMP3"
      Tab(2).ControlCount=   1
      
      TabPicture(3)   =   "rename.frx":07E0
      Tab(3).ControlEnabled=   0   
      Tab(3).Control(0)=   "Acdsee"
      Tab(3).ControlCount=   1
      Begin MSComctlLib.ListView LvMP3 
         Height          =   5820
         Left            =   -74940
         TabIndex        =   134
         Top             =   360
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   10266
         View            =   3
         Sorted          =   -1  
         LabelWrap       =   -1  
         HideSelection   =   0   
         FullRowSelect   =   -1  
         GridLines       =   -1  
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.PictureBox FrameDroite 
         AutoRedraw      =   -1  
         BorderStyle     =   0  
         Height          =   5775
         Left            =   120
         ScaleHeight     =   5775
         ScaleWidth      =   5955
         TabIndex        =   33
         Top             =   400
         Width           =   5955
         Begin VB.Frame Frame1 
            
            Height          =   2600
            Left            =   0
            TabIndex        =   79
            Top             =   0
            Width           =   5895
            Begin VB.CommandButton Command3 
               
               Height          =   315
               Left            =   5400
               TabIndex        =   128
               
               Top             =   240
               WhatsThisHelpID =   177
               Width           =   255
            End
            Begin VB.PictureBox PanelPrefix 
               AutoRedraw      =   -1  
               BorderStyle     =   0  
               Height          =   1815
               Left            =   90
               ScaleHeight     =   1815
               ScaleWidth      =   5715
               TabIndex        =   81
               Top             =   630
               Width           =   5715
               Begin VB.CommandButton Command1 
                  
                  Height          =   300
                  HelpContextID   =   110
                  Left            =   3060
                  TabIndex        =   83
                  
                  Top             =   1400
                  WhatsThisHelpID =   180
                  Width           =   1635
               End
               Begin VB.CommandButton Command10 
                  
                  Height          =   300
                  HelpContextID   =   110
                  Left            =   4800
                  TabIndex        =   82
                  
                  Top             =   1400
                  WhatsThisHelpID =   181
                  Width           =   870
               End
               Begin TabDlg.SSTab SSTab1 
                  Height          =   1335
                  Left            =   0
                  TabIndex        =   84
                  Top             =   0
                  Width           =   5685
                  _ExtentX        =   10028
                  _ExtentY        =   2355
                  _Version        =   393216
                  Style           =   1
                  Tabs            =   8
                  TabsPerRow      =   8
                  TabHeight       =   520
                  
                  
                  Tab(0).ControlEnabled=   -1  
                  
                  Tab(0).Control(0).Enabled=   0   
                  
                  Tab(0).Control(1).Enabled=   0   
                  Tab(0).ControlCount=   2
                  
                  :0818
                  Tab(1).ControlEnabled=   0   
                  
                  
                  Tab(1).ControlCount=   2
                  
                  :0834
                  Tab(2).ControlEnabled=   0   
                  
                  
                  Tab(2).ControlCount=   2
                  
                  :0850
                  Tab(3).ControlEnabled=   0   
                  
                  
                  Tab(3).ControlCount=   2
                  
                  :086C
                  Tab(4).ControlEnabled=   0   
                  
                  Tab(4).ControlCount=   1
                  
                  :0888
                  Tab(5).ControlEnabled=   0   
                  
                  
                  
                  
                  
                  
                  Tab(5).ControlCount=   6
                  
                  :08A4
                  Tab(6).ControlEnabled=   0   
                  
                  
                  Tab(6).ControlCount=   2
                  
                  :08C0
                  Tab(7).ControlEnabled=   0   
                  
                  Tab(7).ControlCount=   1
                  Begin VB.CommandButton Command8 
                     Height          =   300
                     HelpContextID   =   168
                     Left            =   -69675
                     :08DC
                     Style           =   1  
                     TabIndex        =   127
                     
                     Top             =   495
                     Visible         =   0   
                     WhatsThisHelpID =   198
                     Width           =   300
                  End
                  Begin VB.CommandButton Command5 
                     
                     Height          =   375
                     Left            =   -72720
                     TabIndex        =   126
                     Top             =   600
                     WhatsThisHelpID =   204
                     Width           =   1455
                  End
                  Begin VB.CheckBox Check1 
                     
                     Height          =   255
                     HelpContextID   =   162
                     Left            =   -74880
                     TabIndex        =   125
                     
                     Top             =   360
                     WhatsThisHelpID =   202
                     Width           =   2445
                  End
                  Begin VB.PictureBox Picture8 
                     BorderStyle     =   0  
                     Height          =   300
                     Left            =   -74640
                     ScaleHeight     =   300
                     ScaleWidth      =   4815
                     TabIndex        =   121
                     Top             =   720
                     Visible         =   0   
                     WhatsThisHelpID =   203
                     Width           =   4815
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        HelpContextID   =   162
                        Index           =   17
                        Left            =   15
                        TabIndex        =   124
                        
                        Top             =   0
                        Value           =   -1  
                        WhatsThisHelpID =   203
                        Width           =   1365
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        HelpContextID   =   162
                        Index           =   16
                        Left            =   1650
                        TabIndex        =   123
                        
                        Top             =   0
                        WhatsThisHelpID =   203
                        Width           =   1395
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        HelpContextID   =   162
                        Index           =   15
                        Left            =   3375
                        TabIndex        =   122
                        
                        Top             =   0
                        WhatsThisHelpID =   203
                        Width           =   2295
                     End
                  End
                  Begin VB.CheckBox Check3 
                     
                     Height          =   255
                     Left            =   120
                     TabIndex        =   120
                     
                     Top             =   360
                     WhatsThisHelpID =   183
                     Width           =   1365
                  End
                  Begin VB.CheckBox Check7 
                     
                     Height          =   255
                     Left            =   -74880
                     TabIndex        =   119
                     
                     Top             =   360
                     WhatsThisHelpID =   193
                     Width           =   1320
                  End
                  Begin VB.PictureBox Picture7 
                     BorderStyle     =   0  
                     Height          =   285
                     Left            =   -74640
                     ScaleHeight     =   285
                     ScaleWidth      =   4695
                     TabIndex        =   115
                     Top             =   720
                     Visible         =   0   
                     WhatsThisHelpID =   194
                     Width           =   4695
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   9
                        Left            =   15
                        TabIndex        =   118
                        
                        Top             =   0
                        Value           =   -1  
                        WhatsThisHelpID =   194
                        Width           =   1365
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   10
                        Left            =   1650
                        TabIndex        =   117
                        
                        Top             =   0
                        WhatsThisHelpID =   194
                        Width           =   1440
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   11
                        Left            =   3375
                        TabIndex        =   116
                        
                        Top             =   0
                        WhatsThisHelpID =   194
                        Width           =   2295
                     End
                  End
                  Begin VB.CheckBox Check6 
                     
                     Height          =   255
                     Left            =   -74880
                     TabIndex        =   114
                     
                     Top             =   360
                     WhatsThisHelpID =   191
                     Width           =   1365
                  End
                  Begin VB.PictureBox Picture6 
                     BorderStyle     =   0  
                     Height          =   240
                     Left            =   -74640
                     ScaleHeight     =   240
                     ScaleWidth      =   5055
                     TabIndex        =   110
                     Top             =   720
                     Visible         =   0   
                     WhatsThisHelpID =   192
                     Width           =   5055
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   6
                        Left            =   3375
                        TabIndex        =   113
                        
                        Top             =   0
                        WhatsThisHelpID =   192
                        Width           =   1575
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   7
                        Left            =   1650
                        TabIndex        =   112
                        
                        Top             =   0
                        WhatsThisHelpID =   192
                        Width           =   1410
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   8
                        Left            =   15
                        TabIndex        =   111
                        
                        Top             =   0
                        Value           =   -1  
                        WhatsThisHelpID =   192
                        Width           =   1365
                     End
                  End
                  Begin VB.PictureBox Picture5 
                     BorderStyle     =   0  
                     Height          =   240
                     Left            =   -74640
                     ScaleHeight     =   240
                     ScaleWidth      =   5055
                     TabIndex        =   106
                     Top             =   720
                     Visible         =   0   
                     WhatsThisHelpID =   190
                     Width           =   5055
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   5
                        Left            =   3375
                        TabIndex        =   109
                        
                        Top             =   0
                        WhatsThisHelpID =   190
                        Width           =   1470
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   3
                        Left            =   15
                        TabIndex        =   108
                        
                        Top             =   0
                        Value           =   -1  
                        WhatsThisHelpID =   190
                        Width           =   1320
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   4
                        Left            =   1650
                        TabIndex        =   107
                        
                        Top             =   0
                        WhatsThisHelpID =   190
                        Width           =   1440
                     End
                  End
                  Begin VB.CheckBox Check5 
                     
                     Height          =   255
                     Left            =   -74880
                     TabIndex        =   105
                     
                     Top             =   360
                     WhatsThisHelpID =   189
                     Width           =   1320
                  End
                  Begin VB.OptionButton Option1 
                     
                     Height          =   195
                     Index           =   0
                     Left            =   -74880
                     TabIndex        =   104
                     
                     Top             =   525
                     WhatsThisHelpID =   196
                     Width           =   1635
                  End
                  Begin VB.TextBox Text2 
                     Height          =   285
                     Left            =   -73200
                     TabIndex        =   103
                     
                     Top             =   480
                     Visible         =   0   
                     WhatsThisHelpID =   197
                     Width           =   2775
                  End
                  Begin VB.PictureBox Picture1 
                     AutoRedraw      =   -1  
                     AutoSize        =   -1  
                     BorderStyle     =   0  
                     Height          =   255
                     Left            =   -71040
                     ScaleHeight     =   255
                     ScaleWidth      =   1335
                     TabIndex        =   100
                     Top             =   900
                     Width           =   1335
                     Begin VB.OptionButton Option2 
                        
                        Height          =   195
                        Index           =   1
                        Left            =   720
                        TabIndex        =   102
                        
                        Top             =   0
                        Visible         =   0   
                        WhatsThisHelpID =   201
                        Width           =   615
                     End
                     Begin VB.OptionButton Option2 
                        
                        Height          =   195
                        Index           =   0
                        Left            =   0
                        TabIndex        =   101
                        
                        Top             =   0
                        Value           =   -1  
                        Visible         =   0   
                        WhatsThisHelpID =   201
                        Width           =   735
                     End
                  End
                  Begin VB.OptionButton Option1 
                     
                     Height          =   195
                     Index           =   1
                     Left            =   -74880
                     TabIndex        =   99
                     
                     Top             =   870
                     WhatsThisHelpID =   199
                     Width           =   960
                  End
                  Begin VB.TextBox Text14 
                     Height          =   285
                     Left            =   -73200
                     TabIndex        =   98
                     
                     Top             =   825
                     Visible         =   0   
                     WhatsThisHelpID =   200
                     Width           =   2055
                  End
                  Begin VB.CommandButton Command19 
                     
                     Height          =   375
                     Left            =   -72720
                     TabIndex        =   97
                     Top             =   600
                     WhatsThisHelpID =   195
                     Width           =   1455
                  End
                  Begin VB.PictureBox onglcounter 
                     AutoRedraw      =   -1  
                     BorderStyle     =   0  
                     Height          =   580
                     Left            =   120
                     ScaleHeight     =   585
                     ScaleWidth      =   5475
                     TabIndex        =   85
                     Top             =   600
                     Visible         =   0   
                     Width           =   5475
                     Begin VB.TextBox Text3 
                        Alignment       =   1  
                        Height          =   285
                        Left            =   795
                        TabIndex        =   93
                        
                        
                        Top             =   35
                        WhatsThisHelpID =   184
                        Width           =   495
                     End
                     Begin VB.TextBox Text4 
                        Alignment       =   1  
                        Height          =   285
                        Left            =   1980
                        TabIndex        =   92
                        
                        
                        Top             =   35
                        WhatsThisHelpID =   185
                        Width           =   495
                     End
                     Begin VB.TextBox Text5 
                        Alignment       =   1  
                        Height          =   285
                        Left            =   3285
                        TabIndex        =   91
                        
                        
                        Top             =   35
                        WhatsThisHelpID =   186
                        Width           =   495
                     End
                     Begin VB.PictureBox Picture2 
                        BorderStyle     =   0  
                        Height          =   255
                        Left            =   210
                        ScaleHeight     =   255
                        ScaleWidth      =   4815
                        TabIndex        =   87
                        Top             =   360
                        WhatsThisHelpID =   188
                        Width           =   4815
                        Begin VB.OptionButton Option3 
                           
                           Height          =   255
                           Index           =   2
                           Left            =   3375
                           TabIndex        =   90
                           
                           Top             =   0
                           WhatsThisHelpID =   188
                           Width           =   2295
                        End
                        Begin VB.OptionButton Option3 
                           
                           Height          =   255
                           Index           =   1
                           Left            =   1665
                           TabIndex        =   89
                           
                           Top             =   0
                           WhatsThisHelpID =   188
                           Width           =   1395
                        End
                        Begin VB.OptionButton Option3 
                           
                           Height          =   255
                           Index           =   0
                           Left            =   15
                           TabIndex        =   88
                           
                           Top             =   0
                           Value           =   -1  
                           WhatsThisHelpID =   188
                           Width           =   1455
                        End
                     End
                     Begin VB.ComboBox Combo3 
                        Height          =   315
                        ItemData        =   "rename.frx":09C6
                        Left            =   4095
                        List            =   "rename.frx":09D9
                        Style           =   2  'Dropdown List
                        TabIndex        =   86
                        
                        Top             =   35
                        WhatsThisHelpID =   187
                        Width           =   1300
                     End
                     Begin VB.Label Label4 
                        AutoSize        =   -1  
                        
                        Height          =   195
                        Left            =   225
                        TabIndex        =   96
                        Top             =   60
                        WhatsThisHelpID =   184
                        Width           =   405
                     End
                     Begin VB.Label Label5 
                        AutoSize        =   -1  
                        
                        Height          =   195
                        Left            =   1455
                        TabIndex        =   95
                        Top             =   60
                        WhatsThisHelpID =   185
                        Width           =   330
                     End
                     Begin VB.Label Label6 
                        AutoSize        =   -1  
                        
                        Height          =   195
                        Left            =   2730
                        TabIndex        =   94
                        Top             =   60
                        WhatsThisHelpID =   186
                        Width           =   390
                     End
                  End
               End
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "rename.frx":0A0B
               Left            =   120
               List            =   "rename.frx":0A0D
               Sorted          =   -1  
               Style           =   2  
               TabIndex        =   80
               
               Top             =   240
               WhatsThisHelpID =   175
               Width           =   5175
            End
         End
         Begin VB.Frame Frame2 
            
            Height          =   2380
            Left            =   0
            TabIndex        =   37
            Top             =   2640
            Width           =   5895
            Begin VB.CommandButton Command4 
               
               Height          =   315
               Left            =   5400
               TabIndex        =   78
               
               Top             =   240
               WhatsThisHelpID =   177
               Width           =   255
            End
            Begin VB.PictureBox PanelExt 
               AutoRedraw      =   -1  
               BorderStyle     =   0  
               Height          =   1650
               Left            =   120
               ScaleHeight     =   1650
               ScaleWidth      =   5715
               TabIndex        =   39
               Top             =   650
               Width           =   5715
               Begin VB.CommandButton Command2 
                  
                  Height          =   300
                  HelpContextID   =   110
                  Left            =   3060
                  TabIndex        =   41
                  
                  Top             =   1350
                  WhatsThisHelpID =   180
                  Width           =   1635
               End
               Begin VB.CommandButton Command11 
                  
                  Height          =   300
                  HelpContextID   =   110
                  Left            =   4800
                  TabIndex        =   40
                  
                  Top             =   1350
                  WhatsThisHelpID =   181
                  Width           =   870
               End
               Begin TabDlg.SSTab SSTab2 
                  Height          =   1300
                  Left            =   0
                  TabIndex        =   42
                  Top             =   0
                  Width           =   5685
                  _ExtentX        =   10028
                  _ExtentY        =   2302
                  _Version        =   393216
                  Style           =   1
                  Tabs            =   5
                  TabsPerRow      =   5
                  TabHeight       =   520
                  
                  :0A0F
                  Tab(0).ControlEnabled=   -1  
                  
                  Tab(0).Control(0).Enabled=   0   
                  
                  Tab(0).Control(1).Enabled=   0   
                  Tab(0).ControlCount=   2
                  
                  :0A2B
                  Tab(1).ControlEnabled=   0   
                  
                  
                  Tab(1).ControlCount=   2
                  
                  :0A47
                  Tab(2).ControlEnabled=   0   
                  
                  
                  Tab(2).ControlCount=   2
                  
                  :0A63
                  Tab(3).ControlEnabled=   0   
                  
                  
                  Tab(3).ControlCount=   2
                  
                  :0A7F
                  Tab(4).ControlEnabled=   0   
                  
                  
                  
                  
                  
                  Tab(4).ControlCount=   5
                  Begin VB.CheckBox Check12 
                     
                     Height          =   255
                     Left            =   -74880
                     TabIndex        =   77
                     
                     Top             =   360
                     WhatsThisHelpID =   189
                     Width           =   1320
                  End
                  Begin VB.CheckBox Check4 
                     
                     Height          =   255
                     Left            =   -74880
                     TabIndex        =   76
                     
                     Top             =   360
                     WhatsThisHelpID =   193
                     Width           =   1275
                  End
                  Begin VB.PictureBox Picture4 
                     BorderStyle     =   0  
                     Height          =   240
                     Left            =   -74640
                     ScaleHeight     =   240
                     ScaleWidth      =   5055
                     TabIndex        =   72
                     Top             =   720
                     Visible         =   0   
                     WhatsThisHelpID =   194
                     Width           =   5055
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   12
                        Left            =   3375
                        TabIndex        =   75
                        
                        Top             =   0
                        WhatsThisHelpID =   194
                        Width           =   2295
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   13
                        Left            =   1650
                        TabIndex        =   74
                        
                        Top             =   0
                        WhatsThisHelpID =   194
                        Width           =   1440
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   14
                        Left            =   15
                        TabIndex        =   73
                        
                        Top             =   0
                        Value           =   -1  
                        WhatsThisHelpID =   194
                        Width           =   1365
                     End
                  End
                  Begin VB.CheckBox Check13 
                     
                     Height          =   255
                     Left            =   -74880
                     TabIndex        =   71
                     
                     Top             =   360
                     WhatsThisHelpID =   191
                     Width           =   1320
                  End
                  Begin VB.PictureBox Picture14 
                     BorderStyle     =   0  
                     Height          =   285
                     Left            =   -74640
                     ScaleHeight     =   285
                     ScaleWidth      =   5055
                     TabIndex        =   67
                     Top             =   720
                     Visible         =   0   
                     WhatsThisHelpID =   192
                     Width           =   5055
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   30
                        Left            =   15
                        TabIndex        =   70
                        
                        Top             =   0
                        Value           =   -1  
                        WhatsThisHelpID =   192
                        Width           =   1365
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   31
                        Left            =   1650
                        TabIndex        =   69
                        
                        Top             =   0
                        WhatsThisHelpID =   192
                        Width           =   1440
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   32
                        Left            =   3375
                        TabIndex        =   68
                        
                        Top             =   0
                        WhatsThisHelpID =   192
                        Width           =   2295
                     End
                  End
                  Begin VB.PictureBox Picture13 
                     BorderStyle     =   0  
                     Height          =   240
                     Left            =   -74640
                     ScaleHeight     =   240
                     ScaleWidth      =   5055
                     TabIndex        =   63
                     Top             =   720
                     Visible         =   0   
                     WhatsThisHelpID =   190
                     Width           =   5055
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   27
                        Left            =   3375
                        TabIndex        =   66
                        
                        Top             =   0
                        WhatsThisHelpID =   190
                        Width           =   2295
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   28
                        Left            =   1650
                        TabIndex        =   65
                        
                        Top             =   0
                        WhatsThisHelpID =   190
                        Width           =   1395
                     End
                     Begin VB.OptionButton Option3 
                        
                        Height          =   255
                        Index           =   29
                        Left            =   15
                        TabIndex        =   64
                        
                        Top             =   0
                        Value           =   -1  
                        WhatsThisHelpID =   190
                        Width           =   1365
                     End
                  End
                  Begin VB.CheckBox Check11 
                     
                     Height          =   255
                     Left            =   120
                     TabIndex        =   62
                     
                     Top             =   360
                     WhatsThisHelpID =   183
                     Width           =   1320
                  End
                  Begin VB.PictureBox Picture3 
                     BorderStyle     =   0  
                     Height          =   240
                     Left            =   -71070
                     ScaleHeight     =   240
                     ScaleWidth      =   1335
                     TabIndex        =   59
                     Top             =   900
                     Width           =   1335
                     Begin VB.OptionButton Option5 
                        
                        Height          =   195
                        Index           =   1
                        Left            =   720
                        TabIndex        =   61
                        
                        Top             =   0
                        Visible         =   0   
                        WhatsThisHelpID =   201
                        Width           =   615
                     End
                     Begin VB.OptionButton Option5 
                        
                        Height          =   195
                        Index           =   0
                        Left            =   0
                        TabIndex        =   60
                        
                        Top             =   0
                        Value           =   -1  
                        Visible         =   0   
                        WhatsThisHelpID =   201
                        Width           =   855
                     End
                  End
                  Begin VB.OptionButton Option4 
                     
                     Height          =   195
                     Index           =   1
                     Left            =   -74880
                     TabIndex        =   58
                     
                     Top             =   870
                     WhatsThisHelpID =   199
                     Width           =   945
                  End
                  Begin VB.OptionButton Option4 
                     
                     Height          =   195
                     Index           =   0
                     Left            =   -74880
                     TabIndex        =   57
                     
                     Top             =   525
                     WhatsThisHelpID =   196
                     Width           =   1635
                  End
                  Begin VB.TextBox Text15 
                     Height          =   285
                     Left            =   -73200
                     TabIndex        =   56
                     
                     Top             =   840
                     Visible         =   0   
                     WhatsThisHelpID =   200
                     Width           =   2055
                  End
                  Begin VB.TextBox Text8 
                     Height          =   285
                     Left            =   -73200
                     TabIndex        =   55
                     
                     Top             =   480
                     Visible         =   0   
                     WhatsThisHelpID =   197
                     Width           =   2775
                  End
                  Begin VB.PictureBox onglcounter2 
                     AutoRedraw      =   -1  
                     BorderStyle     =   0  
                     Height          =   580
                     Left            =   120
                     ScaleHeight     =   585
                     ScaleWidth      =   5475
                     TabIndex        =   43
                     Top             =   600
                     Visible         =   0   
                     Width           =   5475
                     Begin VB.TextBox Text16 
                        Alignment       =   1  
                        Height          =   285
                        Left            =   795
                        TabIndex        =   51
                        
                        
                        Top             =   35
                        WhatsThisHelpID =   184
                        Width           =   495
                     End
                     Begin VB.PictureBox Picture11 
                        BorderStyle     =   0  
                        Height          =   255
                        Left            =   165
                        ScaleHeight     =   255
                        ScaleWidth      =   5055
                        TabIndex        =   47
                        Top             =   345
                        WhatsThisHelpID =   188
                        Width           =   5055
                        Begin VB.OptionButton Option3 
                           
                           Height          =   255
                           Index           =   26
                           Left            =   15
                           TabIndex        =   50
                           
                           Top             =   0
                           Value           =   -1  
                           WhatsThisHelpID =   188
                           Width           =   1365
                        End
                        Begin VB.OptionButton Option3 
                           
                           Height          =   255
                           Index           =   25
                           Left            =   1665
                           TabIndex        =   49
                           
                           Top             =   0
                           WhatsThisHelpID =   188
                           Width           =   1440
                        End
                        Begin VB.OptionButton Option3 
                           
                           Height          =   255
                           Index           =   24
                           Left            =   3375
                           TabIndex        =   48
                           
                           Top             =   0
                           WhatsThisHelpID =   188
                           Width           =   2295
                        End
                     End
                     Begin VB.TextBox Text18 
                        Alignment       =   1  
                        Height          =   285
                        Left            =   3285
                        TabIndex        =   46
                        
                        
                        Top             =   35
                        WhatsThisHelpID =   186
                        Width           =   495
                     End
                     Begin VB.TextBox Text17 
                        Alignment       =   1  
                        Height          =   285
                        Left            =   1980
                        TabIndex        =   45
                        
                        
                        Top             =   35
                        WhatsThisHelpID =   185
                        Width           =   495
                     End
                     Begin VB.ComboBox Combo4 
                        Height          =   315
                        ItemData        =   "rename.frx":0A9B
                        Left            =   4095
                        List            =   "rename.frx":0AAE
                        Style           =   2  
                        TabIndex        =   44
                        
                        Top             =   35
                        WhatsThisHelpID =   187
                        Width           =   1300
                     End
                     Begin VB.Label Label14 
                        AutoSize        =   -1  
                        
                        Height          =   195
                        Left            =   2730
                        TabIndex        =   54
                        Top             =   60
                        WhatsThisHelpID =   186
                        Width           =   390
                     End
                     Begin VB.Label Label13 
                        AutoSize        =   -1  
                        
                        Height          =   195
                        Left            =   1455
                        TabIndex        =   53
                        Top             =   60
                        WhatsThisHelpID =   185
                        Width           =   330
                     End
                     Begin VB.Label Label12 
                        AutoSize        =   -1  
                        
                        Height          =   195
                        Left            =   180
                        TabIndex        =   52
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
               Sorted          =   -1  
               Style           =   2  
               TabIndex        =   38
               
               Top             =   240
               WhatsThisHelpID =   176
               Width           =   5175
            End
         End
         Begin VB.Frame Frame3 
            
            Height          =   700
            Left            =   0
            TabIndex        =   34
            Top             =   5040
            WhatsThisHelpID =   178
            Width           =   4335
            Begin VB.Label laidep 
               BackStyle       =   0  'Transparent
               Height          =   195
               Left            =   120
               TabIndex        =   36
               
               Top             =   195
               WhatsThisHelpID =   178
               Width           =   4095
            End
            Begin VB.Label laides 
               BackStyle       =   0  'Transparent
               Height          =   210
               Left            =   120
               TabIndex        =   35
               
               Top             =   405
               WhatsThisHelpID =   178
               Width           =   4095
            End
         End
         Begin VB.Label etat3 
            Alignment       =   1  
            AutoSize        =   -1  
            
            Height          =   195
            Left            =   4410
            TabIndex        =   131
            
            Top             =   5535
            WhatsThisHelpID =   179
            Width           =   1410
         End
         Begin VB.Label etat2 
            Alignment       =   1  
            AutoSize        =   -1  
            
            Height          =   195
            Left            =   4410
            TabIndex        =   130
            
            Top             =   5325
            WhatsThisHelpID =   179
            Width           =   1410
         End
         Begin VB.Label etat1 
            Alignment       =   1  
            AutoSize        =   -1  
            
            Height          =   195
            Left            =   4410
            TabIndex        =   129
            
            Top             =   5130
            WhatsThisHelpID =   179
            Width           =   1410
         End
      End
      Begin CCRPFolderTV6.FolderTreeview FolderTreeview1 
         Height          =   5820
         Index           =   0
         Left            =   -74940
         TabIndex        =   132
         Top             =   360
         WhatsThisHelpID =   172
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   10266
         IntegralHeight  =   0   
      End
      Begin VB.Image Acdsee 
         BorderStyle     =   1  
         Height          =   4635
         Left            =   -74940
         Stretch         =   -1  
         
         Top             =   360
         Width           =   5985
      End
   End
   Begin VB.PictureBox PanelList 
      AutoRedraw      =   -1  
      BorderStyle     =   0  
      Height          =   4275
      Left            =   5895
      ScaleHeight     =   4275
      ScaleWidth      =   5730
      TabIndex        =   8
      Top             =   7455
      Visible         =   0   
      Width           =   5730
      Begin MSComctlLib.ListView ListView2 
         Height          =   3855
         Left            =   0
         TabIndex        =   14
         
         Top             =   0
         WhatsThisHelpID =   212
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   6800
         View            =   3
         MultiSelect     =   -1  
         LabelWrap       =   -1  
         HideSelection   =   0   
         FullRowSelect   =   -1  
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   
            Italic          =   0   
            Strikethrough   =   0   
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            
            Object.Width           =   5080
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            
            Object.Width           =   5080
         EndProperty
      End
      Begin VB.CommandButton Command16 
         
         Height          =   300
         Left            =   4920
         TabIndex        =   13
         
         Top             =   3900
         WhatsThisHelpID =   217
         Width           =   750
      End
      Begin VB.CommandButton Command15 
         
         Height          =   300
         Left            =   3150
         TabIndex        =   12
         
         Top             =   3900
         WhatsThisHelpID =   216
         Width           =   1725
      End
      Begin VB.CommandButton Command13 
         
         Height          =   300
         Left            =   0
         TabIndex        =   11
         
         Top             =   3900
         WhatsThisHelpID =   213
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         
         Height          =   300
         Left            =   900
         TabIndex        =   10
         
         Top             =   3900
         WhatsThisHelpID =   214
         Width           =   1000
      End
      Begin VB.CommandButton Command7 
         
         Height          =   300
         Left            =   1950
         TabIndex        =   9
         
         Top             =   3900
         WhatsThisHelpID =   215
         Width           =   1155
      End
   End
   Begin VB.PictureBox paneltext 
      AutoRedraw      =   -1  
      BorderStyle     =   0  
      Height          =   1365
      Left            =   11760
      ScaleHeight     =   1365
      ScaleWidth      =   5295
      TabIndex        =   16
      Top             =   7560
      Visible         =   0   
      Width           =   5295
      Begin VB.TextBox Text9 
         Alignment       =   1  
         Height          =   285
         Left            =   2985
         TabIndex        =   17
         
         Top             =   80
         WhatsThisHelpID =   218
         Width           =   420
      End
      Begin VB.Label Label10 
         :0AE0
         Height          =   930
         Left            =   0
         TabIndex        =   19
         Top             =   405
         Width           =   5265
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  
         
         Height          =   195
         Left            =   0
         TabIndex        =   18
         Top             =   135
         Width           =   2655
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      WhatsThisHelpID =   174
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   10
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   11
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            
            Object.
            
            ImageIndex      =   12
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            Object.
            
            ImageIndex      =   13
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            Object.
            
            ImageIndex      =   14
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            Object.
            ImageIndex      =   15
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            Object.
            ImageIndex      =   16
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            Object.
            ImageIndex      =   17
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            Object.
            ImageIndex      =   18
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            
            Object.
            
            ImageIndex      =   20
         EndProperty
      EndProperty
      OLEDropMode     =   1
      Begin VB.ComboBox Combo5 
         Height          =   315
         IntegralHeight  =   0   
         ItemData        =   "rename.frx":0C03
         Left            =   7755
         List            =   "rename.frx":0C05
         TabIndex        =   1
         
         Top             =   20
         WhatsThisHelpID =   182
         Width           =   1455
      End
   End
   Begin VB.ListBox lhistory 
      Enabled         =   0   
      Height          =   285
      IntegralHeight  =   0   
      Left            =   12000
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   
      Width           =   645
   End
   Begin VB.ListBox List3 
      Height          =   255
      ItemData        =   "rename.frx":0C07
      Left            =   12000
      List            =   "rename.frx":0C0E
      TabIndex        =   6
      Top             =   1380
      Visible         =   0   
      Width           =   1320
   End
   Begin VB.ListBox List2 
      Height          =   255
      ItemData        =   "rename.frx":0C23
      Left            =   12195
      List            =   "rename.frx":0C2A
      TabIndex        =   5
      Top             =   -45
      Visible         =   0   
      Width           =   1320
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "rename.frx":0C42
      Left            =   12000
      List            =   "rename.frx":0C49
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   
      Width           =   1320
   End
   Begin MSComctlLib.StatusBar état 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   3
      Top             =   6825
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12911
            MinWidth        =   1058
            Object.
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            
            Object.
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Object.
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Object.
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1693
            MinWidth        =   1411
            TextSave        =   "26/07/2001"
            Object.
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   979
            MinWidth        =   970
            TextSave        =   "08:39"
            Object.
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   
         Italic          =   0   
         Strikethrough   =   0   
      EndProperty
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   285
      Left            =   12000
      TabIndex        =   15
      Top             =   780
      Visible         =   0   
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   503
      View            =   3
      Sorted          =   -1  
      MultiSelect     =   -1  
      LabelWrap       =   -1  
      HideSelection   =   0   
      OLEDropMode     =   1
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         
         Object.Width           =   1586
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         
         Object.Width           =   1586
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         
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
            :0C6D
            
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :11C1
            
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :1715
            
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :1C69
            
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :21BD
            
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :2711
            
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :2C65
            
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :31B9
            
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :370D
            
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :3C61
            
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :41B5
            
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :4709
            
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :4C5D
            
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :51B1
            
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :5705
            
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :5C59
            
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :61AD
            
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :6701
            
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :6C55
            
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            :71A9
            
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5985
      Left            =   30
      TabIndex        =   0
      Top             =   480
      WhatsThisHelpID =   173
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   10557
      View            =   3
      Sorted          =   -1  
      MultiSelect     =   -1  
      LabelWrap       =   -1  
      HideSelection   =   0   
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         
         Object.Width           =   1586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         
         Object.Width           =   706
      EndProperty
   End
   Begin VB.Menu mfile 
      
      HelpContextID   =   13
      Begin VB.Menu mopenset 
         
         HelpContextID   =   50
         Shortcut        =   ^O
      End
      Begin VB.Menu msave 
         
         HelpContextID   =   51
         Shortcut        =   ^S
      End
      Begin VB.Menu msaveas 
         
         HelpContextID   =   52
         Shortcut        =   {F12}
      End
      Begin VB.Menu msep29 
         
      End
      Begin VB.Menu mprintdir 
         
         HelpContextID   =   57
         Shortcut        =   ^P
      End
      Begin VB.Menu HTMLReport 
         
      End
      Begin VB.Menu mfilefind 
         
         HelpContextID   =   60
         Shortcut        =   ^K
      End
      Begin VB.Menu mgreat 
         
         HelpContextID   =   44
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFile 
         
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         
         Index           =   1
         Visible         =   0   
      End
      Begin VB.Menu mnuFile 
         
         Index           =   2
         Visible         =   0   
      End
      Begin VB.Menu mnuFile 
         
         Index           =   3
         Visible         =   0   
      End
      Begin VB.Menu mnuFile 
         
         Index           =   4
         Visible         =   0   
      End
      Begin VB.Menu mnuFile 
         
         Index           =   5
         Visible         =   0   
      End
      Begin VB.Menu mnuFile 
         
         Index           =   6
         Visible         =   0   
      End
      Begin VB.Menu mend 
         
         HelpContextID   =   17
         Index           =   0
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu medit 
      
      HelpContextID   =   143
      Begin VB.Menu msearchpref 
         
         HelpContextID   =   110
         Shortcut        =   {F3}
      End
      Begin VB.Menu msearchext 
         
         HelpContextID   =   110
         Shortcut        =   {F4}
      End
      Begin VB.Menu mglob 
         
         Shortcut        =   ^H
      End
      Begin VB.Menu mregrenam2 
         
         Enabled         =   0   
      End
      Begin VB.Menu msep11 
         
      End
      Begin VB.Menu mabrev 
         
      End
      Begin VB.Menu mrules 
         
         Shortcut        =   ^R
      End
      Begin VB.Menu msep50 
         
      End
      Begin VB.Menu mundo 
         
         Enabled         =   0   
         HelpContextID   =   16
         Shortcut        =   ^Z
      End
      Begin VB.Menu msep6 
         
      End
      Begin VB.Menu mdatetime 
         
         HelpContextID   =   42
         Shortcut        =   ^D
      End
      Begin VB.Menu mattrib 
         
         HelpContextID   =   38
      End
      Begin VB.Menu msep58 
         
      End
      Begin VB.Menu m1selectAll 
         
         HelpContextID   =   149
         Shortcut        =   ^A
      End
      Begin VB.Menu M1Unselect 
         
         HelpContextID   =   150
      End
      Begin VB.Menu M1Invert 
         
         HelpContextID   =   151
      End
      Begin VB.Menu M1Step 
         
         HelpContextID   =   152
      End
      Begin VB.Menu msep501 
         
      End
      Begin VB.Menu mfold 
         
         Begin VB.Menu mgo1 
            
            Shortcut        =   +{F1}
         End
         Begin VB.Menu mgonext 
            
            Shortcut        =   +{F2}
         End
         Begin VB.Menu mgoprev 
            
            Shortcut        =   +{F3}
         End
         Begin VB.Menu mgolast 
            
            Shortcut        =   +{F4}
         End
      End
   End
   Begin VB.Menu mview 
      
      HelpContextID   =   157
      Begin VB.Menu Mrefresh 
         
         HelpContextID   =   158
      End
      Begin VB.Menu msep30 
         
      End
      Begin VB.Menu moptions 
         
         HelpContextID   =   4
         Shortcut        =   ^T
      End
      Begin VB.Menu minfos 
         
         HelpContextID   =   15
         Shortcut        =   ^I
      End
      Begin VB.Menu msep99 
         
      End
      Begin VB.Menu mhistory 
         
         HelpContextID   =   59
      End
      Begin VB.Menu mchangetab 
         
         Shortcut        =   {F9}
      End
      Begin VB.Menu mviewtabs 
         
         Begin VB.Menu mviewmp3tab 
            
         End
         Begin VB.Menu mviewpicturetab 
            
         End
      End
   End
   Begin VB.Menu mrun 
      
      HelpContextID   =   148
      Begin VB.Menu M2Start 
         
         HelpContextID   =   153
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu m2preview 
         
         HelpContextID   =   154
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu m2manually 
         
         HelpContextID   =   155
      End
      Begin VB.Menu msep59 
         
      End
      Begin VB.Menu m2recursive 
         
         HelpContextID   =   156
      End
   End
   Begin VB.Menu mdisk 
      
      HelpContextID   =   61
      Begin VB.Menu mshformat 
         
         HelpContextID   =   62
         Shortcut        =   ^Y
      End
      Begin VB.Menu msetvolabel 
         
         HelpContextID   =   63
         Shortcut        =   ^L
      End
      Begin VB.Menu msepdsk1 
         
      End
      Begin VB.Menu mmap 
         
         HelpContextID   =   64
         Shortcut        =   ^M
      End
      Begin VB.Menu mdisconnect 
         
         HelpContextID   =   65
      End
      Begin VB.Menu mmapped_drives 
         
         HelpContextID   =   161
      End
   End
   Begin VB.Menu mfavorites 
      
      HelpContextID   =   18
      Begin VB.Menu madddirectory 
         
         HelpContextID   =   19
         Shortcut        =   ^E
      End
      Begin VB.Menu morganyze 
         
         HelpContextID   =   20
         Shortcut        =   ^G
      End
      Begin VB.Menu msep 
         
      End
      Begin VB.Menu menufav 
         
         Index           =   0
      End
   End
   Begin VB.Menu mcontextuel 
      
      Visible         =   0   
      Begin VB.Menu maction 
         
         Begin VB.Menu mdelete 
            
         End
         Begin VB.Menu mopen 
            
         End
         Begin VB.Menu mpropertyes 
            
         End
         Begin VB.Menu mprint 
            
         End
      End
      Begin VB.Menu mcreatefold 
         
      End
      Begin VB.Menu mcreatfoldman 
         
      End
      Begin VB.Menu msepchg2 
         
      End
      Begin VB.Menu mnuswap 
         
         Visible         =   0   
      End
      Begin VB.Menu mregrename 
         
         Enabled         =   0   
      End
      Begin VB.Menu mchg1 
         
      End
      Begin VB.Menu mchg2 
         
      End
      Begin VB.Menu msepchg 
         
      End
      Begin VB.Menu MExec 
         
      End
      Begin VB.Menu mexplorer 
         
      End
      Begin VB.Menu msep5 
         
      End
      Begin VB.Menu msearch 
         
      End
      Begin VB.Menu mremdisp 
         
      End
      Begin VB.Menu madd 
         
      End
      Begin VB.Menu mcopy 
         
      End
      Begin VB.Menu msepcont1 
         
      End
      Begin VB.Menu mmove 
         
         Begin VB.Menu mfirst 
            
         End
         Begin VB.Menu mmiddle 
            
         End
         Begin VB.Menu mlast 
            
         End
      End
   End
   Begin VB.Menu mbag 
      
      HelpContextID   =   66
      Begin VB.Menu mcopybag 
         
         HelpContextID   =   66
         Shortcut        =   ^C
      End
      Begin VB.Menu maddbag 
         
         HelpContextID   =   66
         Shortcut        =   {F6}
      End
      Begin VB.Menu mbagsep4 
         
      End
      Begin VB.Menu mcutbag 
         
         HelpContextID   =   66
         Shortcut        =   ^X
      End
      Begin VB.Menu mcutadditive 
         
         HelpContextID   =   66
         Shortcut        =   {F7}
      End
      Begin VB.Menu mbagsep3 
         
      End
      Begin VB.Menu mpastebag 
         
         HelpContextID   =   66
         Shortcut        =   ^V
      End
      Begin VB.Menu mpastekeep 
         
         HelpContextID   =   66
         Shortcut        =   {F8}
      End
      Begin VB.Menu mbagsep1 
         
      End
      Begin VB.Menu mclearbag 
         
         HelpContextID   =   66
         Shortcut        =   ^B
      End
      Begin VB.Menu mviewbag 
         
         HelpContextID   =   66
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mhelp 
      
      Begin VB.Menu mindex 
         
      End
      Begin VB.Menu renameweb 
         
      End
      Begin VB.Menu msep15 
         
      End
      Begin VB.Menu mapropos 
         
      End
   End
   Begin VB.Menu mcontext2 
      
      Visible         =   0   
      Begin VB.Menu mgendir 
         
         Begin VB.Menu mrendirect 
            
         End
         Begin VB.Menu mmakedir 
            
         End
         Begin VB.Menu mdgroupe 
            
         End
         Begin VB.Menu mdelrep 
            
         End
         Begin VB.Menu mchgattdir 
            
         End
      End
      Begin VB.Menu mgoto 
         
      End
      Begin VB.Menu mprop2 
         
      End
      Begin VB.Menu msep55 
         
      End
      Begin VB.Menu maddfavorites 
         
      End
      Begin VB.Menu myourfav 
         
         Begin VB.Menu mnufav 
            
            Index           =   0
         End
      End
      Begin VB.Menu mmdrives 
         
         Begin VB.Menu mnudrives 
            
            Index           =   0
         End
      End
      Begin VB.Menu msep31 
         
      End
      Begin VB.Menu mstartup 
         
      End
      Begin VB.Menu mdosprompthere 
         
      End
      Begin VB.Menu mcopy2 
         
      End
   End
   Begin VB.Menu m3contextuel 
      
      Visible         =   0   
      Begin VB.Menu m3prefix 
         
         Begin VB.Menu m3cmdprefix 
            
            Index           =   0
         End
      End
      Begin VB.Menu m3extension 
         
         Begin VB.Menu m3cmdextension 
            
            Index           =   0
         End
      End
      Begin VB.Menu m3general 
         
         Begin VB.Menu mlang 
            
            Index           =   0
         End
      End
      Begin VB.Menu mmusic 
         
         Begin VB.Menu mimusic 
            
            Index           =   0
         End
      End
      Begin VB.Menu myourcmd 
         
         Begin VB.Menu yourcmd 
            
            Index           =   0
         End
      End
   End
   Begin VB.Menu m4contextuel 
      
      Visible         =   0   
      Begin VB.Menu mrnremoveall 
         
      End
      Begin VB.Menu mrnsep2 
         
      End
      Begin VB.Menu mnewformlist 
         
      End
      Begin VB.Menu msaveonlynewnames 
         
      End
      Begin VB.Menu mrnsep3 
         
      End
      Begin VB.Menu mpastenewnames 
         
      End
      Begin VB.Menu mrnclipboard 
         
      End
   End
   Begin VB.Menu mgenhistory 
      
      Visible         =   0   
      Begin VB.Menu mnuhistory 
         
         Index           =   0
      End
   End
End
Attribute VB_Name = "RENAME"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Déclarations diverses
Private cLstWin As CListSearch
Dim m_oAutoPos As New clsAutoPositioner
Dim cFichier As New cFile
Dim ARemplacer As String
Dim LaPosSauve As Integer
Dim LongSauve As Integer
Dim Buffer As String * 100
Dim SIni As New cInifile
Dim VnbTokensPr As Integer  ' nombre de tokens du préfix
Dim VnbTokensEx As Integer  ' nombre de tokens de l'extension
Dim VnbTokensFo As Integer  ' nombre de tokens du répertoire
Dim TablTokensPr(100) As String ' Le tableau contenant les tokens du prefix
Dim TablTokensEx(100) As String ' Le tableau contenant les tokens de l'extension
Dim TablTokensFo(100) As String ' Le tableau contenant les tokens du répertoire

Rem Les tableaux pour le Free Form
Dim hlplang(154) As Integer ' Contient les numéros de topics du fichier d'aide
Dim langage(154) As String  ' Contient le nom des commandes
Dim LngCmd(154, 2) As Integer ' Premier indice= longueur de la commande à tester (0= pas de test de longeur), 2ième indice=nombre de paramètres de la commande
Dim vnbcmd As Integer  ' Nombre de commandes dans le langage
Dim commandes(300, 4) As String ' Contient les commandes et les paramètres
Dim TemShift As Boolean
Dim FavEncours As Integer
Dim PosEcran As Integer
Dim PbFtv1 As Boolean
' MRU
Private m_cMRU As New cMRUFileList
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

'This message is sent by windows when a menu command is highlighted
'Private Const WM_MENUSELECT = &H11F
'System menu constants
'Private Const SC_RESTORE = &HF120&
'Private Const SC_MOVE = &HF010&
'Private Const SC_SIZE = &HF000&
'Private Const SC_MINIMIZE = &HF020&
'Private Const SC_MAXIMIZE = &HF030&
'Private Const SC_CLOSE = &HF060&

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
 Dim ItemX As ListItem
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
  If annuler = True Then  ' Gestion de l'arrêt du programme
   annuler = False
   ExitLang
   remplissage
   Exit Function
  End If
  fileop.AddSourceFile chemin + OldName
  If RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
   NewName = RemIllegals(NewName)
  End If
  fileop.AddDestFile chemin + NewName
  état.Panels(1). + OldName + " to " + NewName
  état.Panels(2). + Trim$(Str$(vnbtot))
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
    If Len(Trim$(prog1)) <> 0 Then ' Lancer un programme avant de renommer le fichier
     prog11 = prog1
     ExecCmd prog11, chemin + OldName
    End If
    If CopyRename = True Then  ' Renommer les fichier et copier
     If Not fileop.RenameFiles Then
     End If
    Else ' On copie les fichiers, on ne les renomme pas
     If Not fileop.CopyFiles Then
     End If
    End If
   End If
   If Len(Trim$(prog2)) <> 0 Then ' Lancer un programme après avoir renommé le fichier
    prog22 = prog2
    ExecCmd prog22, chemin + NewName
   End If
    DT1.SetFileDateTime (chemin + NewName)
'   If ChgFilesAttr = True Then ' On change les attributs du fichier
'    SetAttr chemin + NewName, lesattributs
    Attr1.ChangeAttr (chemin + NewName)
'   End If
  Else ' On est en preview *********************************************************************
   ' ATTENTION
   Set ItemX = preview.listPreview.ListItems.Add(, , OldName)
   ItemX.SubItems(1) = NewName
   'preview.listPreview.AddItem OldName + "  => " + NewName
   preview.listsav.AddItem NewName
  End If ' Preview ou pas ?
  fileop.ClearSourceFiles
  fileop.ClearDestFiles
 Next i
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
   'vretour = LVSetItemNotSelected(ListView1, i)
    LVSetItemNotSelected ListView1, i
  Else
   'vretour = LVSetItemSelected(ListView1, i)
   LVSetItemSelected ListView1, i
  End If
 Next
 RENAME.MousePointer = 0
 état.Panels(4).Text = Trim$(Str$(LVGetCountSelected(ListView1)))
 ListView1.Visible = True
 letat = False
End Sub

Private Sub MoveRoot()
 If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase$(left$(Dir1Path, 1))) > 0 And Mid$(Dir1Path, 2, 1) = ":" And Mid$(Dir1Path, 3, 1) = "\" Then
  TemMove = False
  FolderTreeview1(0).Visible = False
  FolderTreeview1(0).SelectedFolder = left$(Dir1Path, 3)
  Dir1Path = left$(Dir1Path, 3)
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
chemin = Trim$(Dir1Path)
If right$(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If
If recursive = True Then
 chemin = ""
End If
fileop.ConfirmOperation = False

Rem Les listes pour l'undo
List2.Clear
List3.Clear
fileop.ParentWnd = hWnd
fileop.ClearSourceFiles
fileop.ClearDestFiles
fileop.AllowUndo = False
fileop.ConfirmOperation = False
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
  If RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
   vfichier = RemIllegals(vfichier)
  End If
  fileop.AddDestFile chemin + vfichier
Rem pour le undo
  List2.AddItem chemin + sItem        ' Nom d'origine.
  List3.AddItem chemin + vfichier     ' Nom d'arrivée.
  lhistory.AddItem Trim$(Str$(Time())) & "|" & chemin & "|" & sItem & "|" & vfichier ' Historique
  fileop.RenameFiles
  fileop.ClearSourceFiles
  fileop.ClearDestFiles
 End If
 i = LVGetItemSelected(ListView1, i)
Wend
fin:
vnb2 = remplissage()
état.Panels(1).
LastUseDate = Date
LastUseTime = Time
NumberOFiles = vnbtot
LastDirectory = Dir1Path
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
End Sub

Private Sub StartRename()
Dim Pos1 As Long, Pos2 As Long
Dim sataille As Long, satailles As String
Dim SonInfo As String
Dim sesattr As Integer, simginfo As String, sesattrs As String
Dim TemoinTrop As Boolean ' Temoin indiquant s'il y a trop de fichiers à traiter pour les listbox
Dim i As Long, vnb As Long, val1 As Long, val2 As Long, val3 As Integer, val4 As Long, val5 As Long, val6 As Integer
Dim vprefixe As String, multitache As Integer, vsuffixe As String
Dim vformat1 As Integer, vformat2 As Integer
Dim vnomfichier As String, vNomComplet As String, vnomorig As String
Dim fileop As New CSHFileOp, vnbtot As Long, vaction1 As Integer
Dim vaction2 As Integer, letexte As String, nboctets As Integer
Dim vnomdesti As String, ChaineTempo As String, chemin As String
Dim chemin2 As String, vnbparam As Integer
Dim vMontre As Boolean
Dim sItem As String
Rem **** Les dim pour le langage de commandes ************************
Dim vTemCopy As Boolean, vrai As Boolean, copie As String
Dim TokenVoulu As Integer
Dim litteral As String, vnb1 As Integer, vnb2 As Integer
Dim cmdencours As Integer
Dim chainetempo2 As String, chainetempo3 As String, laboucle As Integer
Dim ChaineTempo4 As String
Dim laboucle2 As Long, laboucle3 As Integer, cmdprefix As String, LaBoucle4 As Integer
Dim cmdextension As String, vnom As String, vtempo As String
Dim ItemX As ListItem
Rem **** Les dim pour le langage de commandes ************************

Rem ******** Variables pour les abréviations *************************
Dim ij As Integer
Dim Match As Boolean
Dim str1 As String
Dim str2 As String
Dim str3 As Integer
Dim vnbcount As Integer
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
Dim LongRandom As Integer
Dim LetterRandom As Long
Dim VraiRandom As Boolean
Dim NomRandom As String
Dim BoucleRandom As Integer

fileop.ConfirmOperation = False
TemoinTrop = False

chemin2 = Trim$(Dir1Path)
If right$(chemin2, 1) <> "\" Then
 chemin2 = chemin2 + "\"
End If
VancRep = ""
multitache = 0
CptFichier = 1
annuler = False
vTemCopy = False

If Len(UndoFile) > 0 Then
 temoin1 = 1
 Close #1
 On Error GoTo Erreur1
 If ExtractPath(UndoFile) = UndoFile Then
  Open chemin2 + UndoFile For Output As #1
 Else
  Open UndoFile For Output As #1
 End If
Else
 temoin1 = 0
End If

If Len(batch) > 0 Then
 temoin2 = 1
 Close #2
 On Error GoTo Erreur2
 If ExtractPath(batch) = batch Then
  Open chemin2 + batch For Output As #2
 Else
  Open batch For Output As #2
 End If
Else
 temoin2 = 0
End If

If Len(LogFile) > 0 Then
 temoin3 = 1
 Close #3
 On Error GoTo Erreur3
 If ExtractPath(LogFile) = LogFile Then
  Open chemin2 + LogFile For Append As #3
 Else
  Open LogFile For Append As #3
 End If
Else
 temoin3 = 0
End If

On Error GoTo ErrGen
If vaction1 <> 13 Then
 nboctets = Val(Text9.Text)
Else
 nboctets = Val(cmdtxt3.Text)
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
 MsgBox "Error, You must select files before renaming them !"
 ExitLang
 Exit Sub
End If

If vnb > 32736 Then
    TemoinTrop = True
    If MsgBox("WARNING, When there are more than 32736 files to rename, THE Rename can't use Preview, History and UNDO. Are you sure you want to go on ?", vbYesNo, "!! WARNING, IMPORTANT !!") = vbNo Then
        RENAME.MousePointer = 0
        If acpreview = True Then
            Unload preview
        End If
        Exit Sub
    End If
End If

If UseHistory = True Then ' on vérifie s'il n'y aura pas trop de fichiers pour l'historique
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
If CompleCounters = 1 Then
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
 If nboctets = 0 Or nboctets > 100 Then
  MsgBox "Numbers of characters to take from file is not valid. Change it !"
  ExitLang
  If vaction1 = 11 Then
   Text9.SetFocus
  Else
   cmdtxt3.SetFocus
  End If
  Exit Sub
 End If
End If

Rem On vérifie que les compteurs seront bons *****************************
If recursive = False Then
 If CompleCounters = 1 Then
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
fileop.ParentWnd = hWnd
fileop.ClearSourceFiles
fileop.ClearDestFiles
fileop.ConfirmOperation = False

Rem ********** Gestion du preview *********************************************
If acpreview = True Then
 If TemoinTrop = False Then
    preview.listPreview.ListItems.Clear
    preview.listsav.Clear
    preview.Command1.Visible = False
    preview.Show 0
 End If
End If

Rem ****************** Analyse en cas de langage ******************************
If vaction1 = 13 Then
 If Len(Trim$(txtlang.Text)) = 0 Then
  MsgBox "Warning, expression is empty !"
  ExitLang
  txtlang.SetFocus
  Exit Sub
 End If
 For i = 1 To 300
  commandes(i, 1) = ""
  commandes(i, 2) = ""
  commandes(i, 3) = ""
 Next i
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
   If left$(ChaineTempo, 2) <> "<D" And left$(ChaineTempo, 2) <> "<X" And left$(ChaineTempo, 2) <> "<O" Then ' ce n'est pas une commande de compteur
    vrai = False
    For vnb2 = 1 To vnbcmd
     If LngCmd(vnb2, 1) = 0 Then ' c'est une commande sans paramètres
      If UCase$(Trim$(langage(vnb2))) = UCase$(Trim$(ChaineTempo)) Then
       vrai = True
       Exit For
      End If
     Else ' C'est une commande AVEC paramètres *******************************
      If left$(UCase$(Trim$(langage(vnb2))), LngCmd(vnb2, 1)) = left$(UCase$(Trim$(ChaineTempo)), LngCmd(vnb2, 1)) Then
       vrai = True
       Exit For
      End If
     End If
    Next
    If vrai = False Then
     If left$(ChaineTempo, 2) = "<P" Or left$(ChaineTempo, 2) = "<p" Or left$(ChaineTempo, 2) = "<E" Or left$(ChaineTempo, 2) = "<e" Or left$(ChaineTempo, 2) = "<f" Or left$(ChaineTempo, 2) = "<F" Then
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
     If UCase$(left$(Trim$(txtlang.Text), 9)) <> "<COPYFILE" Then
      MsgBox "Error, the <copyfile> command MUST be the first command !"
      ExitLang
      txtlang.SetFocus
      Exit Sub
     End If
     vTemCopy = True
    End If
    If vnb2 <= vnbcmd Then   ' Exception pour les commandes de tokens
        If LngCmd(vnb2, 1) = 0 Then ' c'est une commande sans paramètres
            commandes(cmdencours, 2) = ""                  ' Commande sans paramètre
        Else ' C'est une commande AVEC paramètre
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
            ' Modif du 14/11/2000
            'chainetempo3 = left$(chainetempo, Len(chainetempo) - 1)
            chainetempo3 = left$(ChaineTempo4, Len(ChaineTempo4) - 1)
            ' Modif du 29/03/2001 pour les commandes acceptant un nombre variable de paramètres
            'For laboucle3 = 1 To LngCmd(vnb2, 2) ' Insertion des paramètres dans le tableau de commandes
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
        End If ' Commande avec ou sans paramètres ?
    Else ' Commande de token, il faut la mémoriser comme paramètre de facon à connaitre le token voulu
        commandes(cmdencours, 2) = ChaineTempo
    End If
   Else ' c'est une commande de compteur ********************************************
    vrai = False
    If left$(ChaineTempo, 2) = "<D" Then
     chainetempo2 = "<DDDDD>"
    Else
     If left$(ChaineTempo, 2) = "<X" Then
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
   i = vnb1 + Len(left$(copie, i))
  Else ' C'est un litteral ***************************************************************
   vrai = False
   litteral = ""
   While vrai = False And i <= longueur
    If Mid$(txtlang.Text, i, 1) <> "<" Then
     litteral = litteral + Mid$(txtlang.Text, i, 1) ' Pour les litteraux, il faut prendre le texte original, pas celui qui a été passé en majuscules
    Else
     vrai = True
    End If
    
    If vrai <> True Then
     If i <= longueur Then ' si i reste inférieur à longueur et si on n'a pas déjà demandé à s'arrêter
      i = i + 1
     Else ' On arrive en fin de chaine
      vrai = faux
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
Rem *************************************************************************************************************************************************
Rem **********        Traitement sur les fichies         ********************************************************************************************
Rem *************************************************************************************************************************************************
If right$(Trim$(Dir1Path), 1) <> "\" Then
 chemin = Trim$(Dir1Path) + "\"
Else
 chemin = Trim$(Dir1Path)
End If

RENAME.MousePointer = 11

If vaction1 = 16 Then ' Rename from a list
 vnbtot = FRenameList(chemin, temoin1, temoin2, temoin3)
 GoTo zsuite
End If

i = LVGetItemSelected(ListView1, -1)
While i <> -1
  If annuler = True Then  ' Gestion de l'arrêt du programme
   annuler = False
   ExitLang
   remplissage
   Exit Sub
  End If
  
  sItem = LVGetName(ListView1, i)
  vnb = vnb + 1
  multitache = multitache + 1  ' Permet à l'écran de se rafraichir et de redonner la main
  If multitache = 10 Then
   multitache = 0
   DoEvents
  End If
   
  vprefixe = Prefixe(sItem) ' Le préfixe uniquement
  vsuffixe = Suffixe(sItem) ' Le suffixe uniquement
  vnomfichier = sItem ' Nom complet du fichier avec le chemin s'il existe
   
   vnomorig = vprefixe + "." + vsuffixe
   If recursive = True Then
    chemin = ExtractPath(sItem) ' Si on est en récursif, il faut récupérer le chemin du fichier
   End If
   vNomComplet = chemin + vprefixe + "." + vsuffixe
   
   ' **** Les règles ****************************************************************************
    If LesRegles.NumberOfActiveRules > 0 Then   ' On ne fait les tests que s'il y a des règles d'actives
        ' On prépare les infos sur le fichier
        Set cFichier = Nothing
        If LVGetItemName(ListView1, i, 4) = "File" Then
            cFichier.SetFileName vNomComplet, True  ' Fichier
        Else
            cFichier.SetFileName vNomComplet, False ' Répertoire
        End If
        ' On fait les tests par rapport aux règles
        If Not LesRegles.TestRules(cFichier) Then
            GoTo fin
        End If
    End If
   ' ********************************************************************************************
      
   If RechGlob = True Then
    If SearchAndReplace = 0 Or SearchAndReplace = 2 Then
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
    
   If RechPref = True Then
    If SearchAndReplace = 0 Or SearchAndReplace = 2 Then
     rech1.SourceString = vprefixe
     vprefixe = rech1.BeginSearchAndReplace
     rech1.SourceString = vprefixe
     vprefixe = rech1.BeginReplaceCharacters
    End If
   End If
   
   ' Abréviations
   If OkUseAbbrev = True Then ' il faut utiliser les abbréviations
    If SearchAndReplace = 0 Or SearchAndReplace = 2 Then
        For ij = 1 To CollAbrev.Count   ' Boucle sur toutes les abbréviations de la collection
            If GetToken(CollAbrev.Item(ij), Chr$(254), 7) = "yes" Then ' on utilise des expressions régulières
                'On Error Resume Next
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
                If Match = True Then
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
        Next ij
    
    End If
   End If
   
   If RestartCounter = 1 Then
    If chemin <> VancRep Then
     val1 = Val(Text3.Text)    ' valeur de départ
     val4 = Val(Text16.Text)   ' valeur de départ
     VancRep = chemin
    End If
   ElseIf RestartCounter = 2 Then
    If left$(chemin, At(chemin, "\", LevelRestart)) <> VancRep Then
     val1 = Val(Text3.Text)    ' valeur de départ
     val4 = Val(Text16.Text)   ' valeur de départ
     VancRep = left$(chemin, At(chemin, "\", LevelRestart))
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
     vprefixe = Menage(FmtDate(Now)) + Menage(FmtHeure(Time()))
    
    Case 10 ' Modifier le préfixe
      If Option1(0).Value = True Then ' remplacer par un texte fixe
       vprefixe = Text2.Text
       If UseCylcic = True Then
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
       If Option1(1).Value = True Then
        If Option2(0).Value = True Then ' au début
         vprefixe = Text14.Text + LTrim$(vprefixe)
        Else ' à la fin
         vprefixe = RTrim$(vprefixe) + Text14.Text
        End If
        If UseCylcic = True Then
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
       If Option3(0).Value = True Then  ' Rajouter à gauche
        vprefixe = Compteur(val1, val3, vformat1) + vprefixe
       Else
        If Option3(1).Value = True Then ' Rajouter à droite
         vprefixe = vprefixe + Compteur(val1, val3, vformat1)
        Else ' Remplacer le préfixe par le compteur
         vprefixe = Compteur(val1, val3, vformat1)
        End If
       End If
       val1 = val1 + val2
      End If
      
      If Check5.Value = 1 Then ' ajouter la taille
       If Option3(3).Value = True Then ' Rajouter à gauche
         vprefixe = FileLen(vnomfichier) & vprefixe
       Else
        If Option3(4).Value = True Then ' Rajouter à droite
         vprefixe = vprefixe & FileLen(vnomfichier)
        Else ' Remplacer le préfixe par la taille
         vprefixe = FileLen(vnomfichier)
        End If
       End If
      End If
      
      If Check6.Value = 1 Then ' ajouter la date
       If Option3(8).Value = True Then ' Rajouter à gauche
        vprefixe = Menage(FmtDate(FileDateTime(vNomComplet))) + vprefixe
       Else
        If Option3(7).Value = True Then 'Rajouter à droite
         vprefixe = vprefixe + Menage(FmtDate(FileDateTime(vNomComplet)))
        Else 'Remplacer le préfixe par la date
         vprefixe = Menage(FmtDate(FileDateTime(vNomComplet)))
        End If
       End If
      End If
      
      If Check7.Value = 1 Then ' ajouter l'heure
       If Option3(9).Value = True Then ' Rajouter à gauche
        vprefixe = Menage(FmtHeure(FileDateTime(vNomComplet))) + vprefixe
       Else
        If Option3(10).Value = True Then 'Rajouter à droite
         vprefixe = vprefixe + Menage(FmtHeure(FileDateTime(vNomComplet)))
        Else 'Remplacer le préfixe par la date
         vprefixe = Menage(FmtHeure(FileDateTime(vNomComplet)))
        End If
       End If
      End If
      
      If FolderOk = True Then ' Add folder's name
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
        If LVGetItemName(ListView1, i, 4) = "File" Then
            simginfo = ""
            simginfo = ImgInfo(vNomComplet)
            If Option3(17).Value = True Then ' Rajouter à gauche
                vprefixe = simginfo + vprefixe
            Else
                If Option3(16).Value = True Then 'Rajouter à droite
                    vprefixe = vprefixe + simginfo
                Else 'Remplacer le préfixe par les infos
                    If Len(Trim$(simginfo)) <> 0 Then
                        vprefixe = simginfo
                End If
            End If
        End If
       End If
      End If
      
      If UseMP3 = True Then
       If LVGetItemName(ListView1, i, 4) = "File" Then
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
      
    
    Case 11 ' Remplacer avec le contenu du fichier
     If LVGetItemName(ListView1, i, 4) = "File" Then
         Temoin11 = True
         Open vnomfichier For Binary As #1
         Get 1, , Buffer
         Close #1
         Buffer = left$(Buffer, nboctets)
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
        If LVGetItemName(ListView1, i, 4) = "File" Then
            vprefixe = Menage(GetFontName(vNomComplet))
        End If
    
    Case 13 ' Free form, mini language
     cmdprefix = Prefixe(vnomfichier)
     cmdextension = Suffixe(vnomfichier)
     ' Il faut gérer les recherches et remplacements
     If RechPref = True Then
      If SearchAndReplace = 0 Or SearchAndReplace = 2 Then
       rech1.SourceString = cmdprefix
       cmdprefix = rech1.BeginSearchAndReplace
       rech1.SourceString = cmdprefix
       cmdprefix = rech1.BeginReplaceCharacters
      End If
     End If
     If RechSuff = True Then
      If SearchAndReplace = 0 Or SearchAndReplace = 2 Then
       rech2.SourceString = cmdextension
       cmdextension = rech2.BeginSearchAndReplace
       rech2.SourceString = cmdextension
       cmdextension = rech2.BeginReplaceCharacters
      End If
     End If
     
     For laboucle2 = 1 To CptFichier
      vnom = ""
      CreateTokenTabl vnomorig, chemin ' chargement des tokens pour le fichier courant
      For laboucle = 1 To cmdencours
       Select Case commandes(laboucle, 1)
        Case "-1" ' Commande de token
            TokenVoulu = Val(Mid$(commandes(laboucle, 2), 3, Len(commandes(laboucle, 2)) - 3))
            If left$(UCase$(commandes(laboucle, 2)), 2) = "<P" Then   ' Token sur le prefixe
                If TokenVoulu <= VnbTokensPr Then
                    If At(commandes(laboucle, 2), ",", 1) <> 0 Then ' On a mis un ou des paramètres
                        vtempo = TablTokensPr(TokenVoulu)
                        For LaBoucle4 = 1 To CharOccurs(commandes(laboucle, 2), ",")
                            Select Case Val(Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", LaBoucle4 + 1)))
                                Case 1   ' Capitalize First word only
                                    vtempo = UCase$(left$(vtempo, 1)) + LCase$(Mid$(vtempo, 2))
                                Case 2   ' Capitalize all words
                                    vtempo = MyStrConv(vtempo)
                                Case 3   ' CowBoys
                                    vtempo = CoWbOyS(vtempo)
                                Case 4   ' Invert
                                    vtempo = StrReverse(vtempo)
                                Case 5   ' Lower
                                    vtempo = LCase$(vtempo)
                                Case 6   ' Ltrim
                                    vtempo = LTrim$(vtempo)
                                Case 7   ' Rtrim
                                    vtempo = RTrim$(vtempo)
                                Case 8   ' Trim
                                    vtempo = Trim$(vtempo)
                                Case 9   ' Toggle
                                    vtempo = ToggleCase(vtempo)
                                Case 10  ' Upper
                                    vtempo = UCase$(vtempo)
                            End Select
                        Next
                        vnom = vnom + vtempo
                    Else    ' Token sans paramètre
                        vnom = vnom + TablTokensPr(TokenVoulu)
                    End If
                End If
            Else    ' Token sur l'extension ?
                If left$(UCase$(commandes(laboucle, 2)), 2) = "<E" Then   ' Token sur l'extension
                    If TokenVoulu <= VnbTokensEx Then
                        If At(commandes(laboucle, 2), ",", 1) <> 0 Then ' On a mis un ou des paramètres
                            vtempo = TablTokensEx(TokenVoulu)
                            For LaBoucle4 = 1 To CharOccurs(commandes(laboucle, 2), ",")
                                Select Case Val(Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", LaBoucle4 + 1)))
                                    Case 1   ' Capitalize First word only
                                        vtempo = UCase$(left$(vtempo, 1)) + LCase$(Mid$(vtempo, 2))
                                    Case 2   ' Capitalize all words
                                        vtempo = MyStrConv(vtempo)
                                    Case 3   ' CowBoys
                                        vtempo = CoWbOyS(vtempo)
                                    Case 4   ' Invert
                                        vtempo = StrReverse(vtempo)
                                    Case 5   ' Lower
                                        vtempo = LCase$(vtempo)
                                    Case 6   ' Ltrim
                                        vtempo = LTrim$(vtempo)
                                    Case 7   ' Rtrim
                                        vtempo = RTrim$(vtempo)
                                    Case 8   ' Trim
                                        vtempo = Trim$(vtempo)
                                    Case 9   ' Toggle
                                        vtempo = ToggleCase(vtempo)
                                    Case 10  ' Upper
                                        vtempo = UCase$(vtempo)
                                End Select
                            Next
                            vnom = vnom + vtempo
                        Else    ' Token sans paramètre
                            vnom = vnom + TablTokensEx(TokenVoulu)
                        End If
                    End If
                Else    ' Token pour le répertoire
                    If TokenVoulu <= VnbTokensFo Then
                        If At(commandes(laboucle, 2), ",", 1) <> 0 Then ' On a mis un ou des paramètres
                            vtempo = TablTokensFo(TokenVoulu)
                            For LaBoucle4 = 1 To CharOccurs(commandes(laboucle, 2), ",")
                                Select Case Val(Trim$(GetToken(Replace(commandes(laboucle, 2), ">", ""), ",", LaBoucle4 + 1)))
                                    Case 1   ' Capitalize First word only
                                        vtempo = UCase$(left$(vtempo, 1)) + LCase$(Mid$(vtempo, 2))
                                    Case 2   ' Capitalize all words
                                        vtempo = MyStrConv(vtempo)
                                    Case 3   ' CowBoys
                                        vtempo = CoWbOyS(vtempo)
                                    Case 4   ' Invert
                                        vtempo = StrReverse(vtempo)
                                    Case 5   ' Lower
                                        vtempo = LCase$(vtempo)
                                    Case 6   ' Ltrim
                                        vtempo = LTrim$(vtempo)
                                    Case 7   ' Rtrim
                                        vtempo = RTrim$(vtempo)
                                    Case 8   ' Trim
                                        vtempo = Trim$(vtempo)
                                    Case 9   ' Toggle
                                        vtempo = ToggleCase(vtempo)
                                    Case 10  ' Upper
                                        vtempo = UCase$(vtempo)
                                End Select
                            Next
                            vnom = vnom + vtempo
                        Else    ' Token sans paramètre
                            vnom = vnom + TablTokensFo(TokenVoulu)
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
            If LVGetItemName(ListView1, i, 4) = "File" Then
                Open chemin & cmdprefix & "." & cmdextension For Input As #1
                letexte = Input(nboctets, 1)
                Close #1
                If Len(Trim$(letexte)) = 0 Then
                    MsgBox "Error, unable to rename file " + chemin & cmdprefix & "." & cmdextension + " because file's content is empty !"
                    ExitLang
                    Exit Sub
                End If
                vpos = InStr(Chr$(13) + Chr$(10), letexte)
                If vpos <> 0 Then
                    letexte = left$(letexte, vpos - 1)
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
            If LVGetItemName(ListView1, i, 4) = "File" Then
                vnom = vnom + Menage(GetFontName(chemin & cmdprefix & "." & cmdextension))
            End If
        Case "22" ' <html>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                vnom = vnom + GetHtmlName(chemin & cmdprefix & "." & cmdextension)
            End If
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
            vnom = vnom + left$(cmdextension, Val(commandes(laboucle, 2)))
         Else
            If Len(cmdextension) + Val(commandes(laboucle, 2)) > 0 Then
                vnom = vnom + left$(cmdextension, Len(cmdextension) + Val(commandes(laboucle, 2)))
            Else
                vnom = vnom + cmdextension
            End If
         End If
        Case "34" ' <EXRight>
         If Val(commandes(laboucle, 2)) > 0 Then
            vnom = vnom + right$(cmdextension, Val(commandes(laboucle, 2)))
         Else
            If Len(cmdextension) + Val(commandes(laboucle, 2)) > 0 Then
                vnom = vnom + right$(cmdextension, Len(cmdextension) + Val(commandes(laboucle, 2)))
            Else
                vnom = vnom + cmdextension
            End If
         End If
        Case "35" ' <PRLeft>
         If Val(commandes(laboucle, 2)) > 0 Then
            vnom = vnom + left$(cmdprefix, Val(commandes(laboucle, 2)))
         Else
            If Len(cmdprefix) + Val(commandes(laboucle, 2)) > 0 Then
                vnom = vnom + left$(cmdprefix, Len(cmdprefix) + Val(commandes(laboucle, 2)))
            Else
                vnom = vnom + cmdprefix
            End If
         End If
        Case "36" ' <PRRight>
         If Val(commandes(laboucle, 2)) > 0 Then
            vnom = vnom + right$(cmdprefix, Val(commandes(laboucle, 2)))
         Else
            If Len(cmdprefix) + Val(commandes(laboucle, 2)) > 0 Then
                vnom = vnom + right$(cmdprefix, Len(cmdprefix) + Val(commandes(laboucle, 2)))
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
         vnom = vnom + GetToken(cmdextension, commandes(laboucle, 3), Val(commandes(laboucle, 2)))
        Case "44" ' <PRToken>
         vnom = vnom + GetToken(cmdprefix, commandes(laboucle, 3), Val(commandes(laboucle, 2)))
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
         If sesattr And vbReadOnly Then
          sesattrs = sesattrs + "R"
         End If
         If sesattr And vbHidden Then
          sesattrs = sesattrs + "H"
         End If
         If sesattr And vbSystem Then
          sesattrs = sesattrs + "S"
         End If
         If sesattr And vbArchive Then
          sesattrs = sesattrs + "A"
         End If
         vnom = vnom + sesattrs
        Case "57" ' <ImgInfo>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                simginfo = ImgInfo(chemin & cmdprefix & "." & cmdextension)
                vnom = vnom + simginfo
            End If
        Case "58" ' <CyclicSelection>
         CompteurCyclic = CompteurCyclic + 1
         If CompteurCyclic > VnbCyclic Then
          CompteurCyclic = 1
         End If
         vnom = vnom + LesCyclic(CompteurCyclic)
        Case "59" ' <PRBefore,0>
         If InStr(1, cmdprefix, commandes(laboucle, 2)) - 1 > 0 Then
          vnom = vnom + left$(cmdprefix, InStr(1, cmdprefix, commandes(laboucle, 2)) - 1)
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
          vnom = vnom + left$(cmdextension, InStr(1, cmdextension, commandes(laboucle, 2)) - 1)
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
          If UseLowerInLetterCounters = 1 Then ' passer en minuscules
           NomRandom = LCase$(NomRandom)
          End If
          VraiRandom = FileExists(chemin & vnom & NomRandom & "." & vsuffixe)
         Wend
         vnom = vnom & NomRandom
        Case "67"  ' <EXCapitalFirst>
         vnom = vnom + UCase$(left$(cmdextension, 1)) + LCase$(Mid$(cmdextension, 2))
        Case "68"  ' <PRCapitalFirst>
         vnom = vnom + UCase$(left$(cmdprefix, 1)) + LCase$(Mid$(cmdprefix, 2))
        Case "69" ' <PRCowBoys>
         vnom = vnom + CoWbOyS(cmdprefix)
        Case "70" ' <EXCowBoys>
         vnom = vnom + CoWbOyS(cmdextension)
        Case "71" ' <PRRemMultSp>
         vnom = vnom + RemoveMultipleSpacing(cmdprefix)
        Case "72" ' <EXRemMultSp>
         vnom = vnom + RemoveMultipleSpacing(cmdextension)
        
        Case "73" ' <MP3Title,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Title, laboucle, vnom)
            End If
        Case "74" ' <MP3Artist,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Artist, laboucle, vnom)
            End If
        Case "75" ' <MP3Album,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Album, laboucle, vnom)
            End If
        Case "76" ' <MP3Year,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Year, laboucle, vnom)
            End If
        Case "77" ' <MP3Comment,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Comment, laboucle, vnom)
            End If
        Case "78" ' <MP3Genre,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Genre, laboucle, vnom)
            End If
        Case "79" ' <MP3Band,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Band, laboucle, vnom)
            End If
        Case "80" ' <MP3BMP,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.BPM, laboucle, vnom)
            End If
        Case "81" ' <MP3Composer,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Composer, laboucle, vnom)
            End If
        Case "82" ' <MP3Conductor,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Conductor, laboucle, vnom)
            End If
        Case "83" ' <MP3ContentGroup,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.ContentGroup, laboucle, vnom)
            End If
            
        Case "84" ' <MP3Copyright,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
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
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                vnom = MP3Commands(MusVQF.Author, laboucle, vnom)
            End If
        
        Case "94" ' <VQFBitrate,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                vnom = MP3Commands(MusVQF.Bitrate, laboucle, vnom)
            End If
        
        Case "95" ' <VQFComment,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                vnom = MP3Commands(MusVQF.Comment, laboucle, vnom)
            End If
        
        Case "96" ' <VQFCopyright,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                vnom = MP3Commands(MusVQF.Copyright, laboucle, vnom)
            End If
        
        Case "97" ' <VQFFileSaveAs,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                vnom = MP3Commands(MusVQF.SaveAsFilename, laboucle, vnom)
            End If
        
        Case "98" ' <VQFMonoStereo,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                vnom = MP3Commands(MusVQF.Mono_Stereo, laboucle, vnom)
            End If
        
        Case "99" ' <VQFQuality,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                vnom = MP3Commands(MusVQF.Quality, laboucle, vnom)
            End If
        
        Case "100" ' <VQFSampleRate,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                vnom = MP3Commands(MusVQF.SampleRate, laboucle, vnom)
            End If
        
        Case "101" ' <VQFTitle>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusVQF.GetVQFInfos(vNomComplet, False)
                vnom = MP3Commands(MusVQF.Title, laboucle, vnom)
            End If
        Case "102" ' <MP3EncryptionMethod,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.EncryptionMethod, laboucle, vnom)
            End If
        
        Case "103" ' <MP3Date,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.mDate, laboucle, vnom)
            End If

        Case "104" ' <MP3EncodedBy,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.EncodedBy, laboucle, vnom)
            End If
        
        Case "105" ' <MP3SoftwareEncodingSettings,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.SoftwareEncodingSettings, laboucle, vnom)
            End If
        
        Case "106" ' <MP3FileOwner,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.FileOwner, laboucle, vnom)
            End If
        
        Case "107" ' <MP3FileType,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.FileType, laboucle, vnom)
            End If
        
        Case "108" ' <MP3GroupIdent,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.GroupIdent, laboucle, vnom)
            End If
        
        Case "109" ' <MP3InitialKey,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.InitialKey, laboucle, vnom)
            End If
        
        Case "110" ' <MP3InvolvedPeopleList,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.InvolvedPeopleList, laboucle, vnom)
            End If
        
        Case "111" ' <MP3Isrc,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.ISRC, laboucle, vnom)
            End If
        
        Case "112" ' <MP3Language,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Language, laboucle, vnom)
            End If
        
        Case "113" ' <MP3LinkedInformation,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.LinkedInformation, laboucle, vnom)
            End If
        
        Case "114" ' <MP3Lyricist,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Lyricist, laboucle, vnom)
            End If
        
        Case "115" ' <MP3MediaType,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.MediaType, laboucle, vnom)
            End If
        
        Case "116" ' <MP3MixArtist,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.MixArtist, laboucle, vnom)
            End If
        
        Case "117" ' <MP3NetRadioOwner,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.NetRadioOwner, laboucle, vnom)
            End If
        
        Case "118" ' <MP3NetRadioStation,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.NetRadioStation, laboucle, vnom)
            End If
        
        Case "119" ' <MP3OriginalAlbum,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.OriginalAlbum, laboucle, vnom)
            End If
        
        Case "120" ' <MP3OriginalArtist,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.OriginalArtist, laboucle, vnom)
            End If
        
        Case "121" ' <MP3OriginalFilename,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.OriginalFilename, laboucle, vnom)
            End If
        
        Case "122" ' <MP3OriginalLyricist,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.OriginalLyricist, laboucle, vnom)
            End If
        
        Case "123" ' <MP3OriginalYear,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.OriginalYear, laboucle, vnom)
            End If
        
        Case "124" ' <MP3PartOfASet,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.PartOfASet, laboucle, vnom)
            End If
        
        Case "125" ' <MP3PlayListDelay,,>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
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
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.PopulariMeter, laboucle, vnom)
            End If

        Case "135"  ' <MP3Publisher>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Publisher, laboucle, vnom)
            End If

        Case "136"  ' <MP3RecordingDates>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.RecordingDates, laboucle, vnom)
            End If

        Case "137"  ' <MP3SongLength>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.SongLength, laboucle, vnom)
            End If

        Case "138"  ' <MP3SubTitle>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.SubTitle, laboucle, vnom)
            End If

        Case "139"  ' <MP3SynchronizedLyric>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.SynchronizedLyric, laboucle, vnom)
            End If

        Case "140"  ' <MP3TermsOfUse>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.TermsOfUse, laboucle, vnom)
            End If

        Case "141"  ' <MP3Time>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.Time, laboucle, vnom)
            End If

        Case "142"  ' <MP3TrackNumber>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.TrackNumber, laboucle, vnom)
            End If

        Case "143"  ' <MP3TotalTracks>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.TotalTracks, laboucle, vnom)
            End If

        Case "144"  ' <MP3UnsynchronizedLyric>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.UnsynchronizedLyric, laboucle, vnom)
            End If

        Case "145"  ' <MP3UserText>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.UserText, laboucle, vnom)
            End If

        Case "146"  ' <MP3wwwArtist>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.wwwArtist, laboucle, vnom)
            End If

        Case "147"  ' <MP3wwwAudioFile>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.wwwAudioFile, laboucle, vnom)
            End If

        Case "148"  ' <MP3wwwAudioSource>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.wwwAudioSource, laboucle, vnom)
            End If

        Case "149"  ' <MP3wwwCommercialInfo>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.wwwCommercialInfo, laboucle, vnom)
            End If

        Case "150"  ' <MP3wwwCopyright>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.wwwCopyright, laboucle, vnom)
            End If

        Case "151"  ' <MP3wwwPayment>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.wwwPayment, laboucle, vnom)
            End If

        Case "152"  ' <MP3wwwPublisher>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.wwwPublisher, laboucle, vnom)
            End If

        Case "153"  ' <MP3wwwRadioPage>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.wwwRadioPage, laboucle, vnom)
            End If

        Case "154"  ' <MP3wwwUserURL>
            If LVGetItemName(ListView1, i, 4) = "File" Then
                SonInfo = ""
                SonInfo = MusMP3.GetMP3Infos(vNomComplet, False)
                vnom = MP3Commands(MusMP3.wwwUserURL, laboucle, vnom)
            End If
            
       End Select ' Quelle commande ?
      Next '    Boucle sur les commandes
      If acpreview = True And vTemCopy = True Then
       If TemoinTrop = False Then
        preview.Command5.Visible = False
        vMontre = True
        If ShowWhenFileNameChange = 1 Then
            If Trim$(vnomfichier) = Trim$(chemin & vnom) Then
                vMontre = False
            End If
        End If
        If vMontre Then
            Set ItemX = preview.listPreview.ListItems.Add(, , vnomfichier)
            ItemX.SubItems(1) = chemin & vnom
            preview.listsav.AddItem vnomdesti
        End If
       End If
      Else
       If vTemCopy = True Then
        état.Panels(1). + vnomfichier + " to " + vnom
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
     
    If RechGlob = True Then
        If SearchAndReplace = 1 Or SearchAndReplace = 2 Then
            rech3.SourceString = cmdprefix + "." + cmdextension
            vtmpGlob = rech3.BeginSearchAndReplace
            cmdprefix = Prefixe(vtmpGlob)
            cmdextension = Suffixe(vtmpGlob)
         
            rech3.SourceString = cmdprefix + "." + cmdextension
            vtmpGlob = rech3.BeginReplaceCharacters
            cmdprefix = Prefixe(vtmpGlob)
            cmdextension = Suffixe(vtmpGlob)
        End If
    End If
     
     If RechPref = True Then
      If SearchAndReplace = 1 Or SearchAndReplace = 2 Then
       rech1.SourceString = cmdprefix
       cmdprefix = rech1.BeginSearchAndReplace
       rech1.SourceString = cmdprefix
       cmdprefix = rech1.BeginReplaceCharacters
      End If
     End If
     
     If RechSuff = True Then
      If SearchAndReplace = 0 Or SearchAndReplace = 2 Then
       rech2.SourceString = cmdextension
       cmdextension = rech2.BeginSearchAndReplace
       rech2.SourceString = cmdextension
       cmdextension = rech2.BeginReplaceCharacters
      End If
     End If
     vnom = cmdprefix & "." & cmdextension
    
    Case 14 ' Short name
     If recursive = True Then
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
     vprefixe = UCase$(left$(vprefixe, 1)) + LCase$(Mid$(vprefixe, 2))
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
        If UseLowerInLetterCounters = 1 Then ' passer en minuscules
         NomRandom = LCase$(NomRandom)
        End If
       VraiRandom = FileExists(chemin & NomRandom & "." & vsuffixe)
     Wend
     vprefixe = NomRandom
    
    Case 19 ' CoWbOyS
     vprefixe = CoWbOyS(vprefixe)
     
    Case 20 ' Remove Multiple Spacing
     vprefixe = RemoveMultipleSpacing(vprefixe)
    Case 21 ' Separate Words
     vprefixe = ExtractWords(vprefixe)
   End Select
   
   If RechGlob = True Then
    If SearchAndReplace = 1 Or SearchAndReplace = 2 Then
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
   
   If RechPref = True Then
    If SearchAndReplace = 1 Or SearchAndReplace = 2 Then
     rech1.SourceString = vprefixe
     vprefixe = rech1.BeginSearchAndReplace
     rech1.SourceString = vprefixe
     vprefixe = rech1.BeginReplaceCharacters
    End If
   End If
   
   ' Abréviations
   If OkUseAbbrev = True Then ' il faut utiliser les abbréviations
        For ij = 1 To CollAbrev.Count   ' Boucle sur toutes les abbréviations de la collection
            If GetToken(CollAbrev.Item(ij), Chr$(254), 7) = "yes" Then ' on utilise des expressions régulières
                'On Error Resume Next
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
                If Match = True Then
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
        Next ij
   End If
   

   If RechSuff = True Then
    If SearchAndReplace = 0 Or SearchAndReplace = 2 Then
     rech2.SourceString = vsuffixe
     vsuffixe = rech2.BeginSearchAndReplace
     rech2.SourceString = vsuffixe
     vsuffixe = rech2.BeginReplaceCharacters
    End If
   End If

Rem *************************************************************************
Rem ***************** Action à effectuer sur le suffixe *********************
Rem *************************************************************************
If Frame2.Visible = True Then  ' On essaye de gagner du temps, si les actions possibles sur le suffixe ne sont pas visibles, c'est pas la peine de faire des tests
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
     vsuffixe = Menage(FmtDate(Now)) + Menage(FmtHeure(Time()))
    Case 12 ' Remove internal spaces
     vsuffixe = RInternalSpaces(vsuffixe)
    Case 13 ' CoWbOyS
     vsuffixe = CoWbOyS(vsuffixe)
    Case 14 ' Remove Multiple Spacing
     vsuffixe = RemoveMultipleSpacing(vsuffixe)
    Case 15 ' Separate Words
     vsuffixe = ExtractWords(vsuffixe)
     
    Case 10 ' Modifer l'extension
      If Option4(0).Value = True Then ' remplacer par un texte fixe
       vsuffixe = Text8.Text
      Else
       If Option4(1).Value = True Then ' ajouter un texte fixe
        If Option5(0).Value = True Then ' au début
         vsuffixe = Text15.Text + LTrim$(vsuffixe)
        Else ' à la fin
         vsuffixe = RTrim$(vsuffixe) + Text15.Text
        End If
       End If
      End If
      
      If Check11.Value = 1 Then ' ajouter un compteur
       If Option3(26).Value = True Then  ' Rajouter à gauche
        vsuffixe = Compteur(val4, val6, vformat2) + vsuffixe
       Else
        If Option3(25).Value = True Then ' Rajouter à droite
         vsuffixe = vsuffixe + Compteur(val4, val6, vformat2)
        Else ' Remplacer le préfixe par le compteur
         vsuffixe = Compteur(val4, val6, vformat2)
        End If
       End If
       val4 = val4 + val5
      End If
      
      If Check12.Value = 1 Then ' ajouter la taille
       If Option3(29).Value = True Then ' Rajouter à gauche
        vsuffixe = FileLen(vnomfichier) & vsuffixe
       Else
        If Option3(28).Value = True Then ' Rajouter à droite
         vsuffixe = vsuffixe & FileLen(vnomfichier)
        Else ' Remplacer le préfixe par la taille
         vsuffixe = FileLen(vnomfichier)
        End If
       End If
      End If
      
      If Check13.Value = 1 Then ' ajouter la date
       If Option3(30).Value = True Then ' Rajouter à gauche
        vsuffixe = Menage(FmtDate(FileDateTime(vNomComplet))) + vsuffixe
       Else
        If Option3(31).Value = True Then 'Rajouter à droite
         vsuffixe = vsuffixe + Menage(FmtDate(FileDateTime(vNomComplet)))
        Else 'Remplacer le préfixe par la date
          vsuffixe = Menage(FmtDate(FileDateTime(vNomComplet)))
        End If
       End If
      End If
      
      If Check4.Value = 1 Then ' ajouter l'heure
       If Option3(14).Value = True Then ' Rajouter à gauche
        vsuffixe = Menage(FmtHeure(FileDateTime(vNomComplet))) + vsuffixe
       Else
        If Option3(13).Value = True Then 'Rajouter à droite
         vsuffixe = vsuffixe + Menage(FmtHeure(FileDateTime(vNomComplet)))
        Else 'Remplacer le préfixe par la date
         vsuffixe = Menage(FmtHeure(FileDateTime(vNomComplet)))
        End If
       End If
      End If
      
    End Select
End If ' Si la frame pour l'extension est visible
    
   If RechSuff = True Then
    If SearchAndReplace = 1 Or SearchAndReplace = 2 Then
     rech2.SourceString = vsuffixe
     vsuffixe = rech2.BeginSearchAndReplace
     rech2.SourceString = vsuffixe
     vsuffixe = rech2.BeginReplaceCharacters
    End If
   End If
   
   ' Abréviations
   If OkUseAbbrev = True Then ' il faut utiliser les abbréviations
    If SearchAndReplace = 1 Or SearchAndReplace = 2 Then
        For ij = 1 To CollAbrev.Count   ' Boucle sur toutes les abbréviations de la collection
            If GetToken(CollAbrev.Item(ij), Chr$(254), 7) = "yes" Then ' on utilise des expressions régulières
                'On Error Resume Next
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
                If Match = True Then
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
        Next ij
    End If
   End If
   

' Rajout du 9/7/98
If vTemCopy = False Then
 If Frame2.Visible = True Then  ' On essaye de gagner du temps, si les actions possibles sur le suffixe ne sont pas visibles, c'est pas la peine de faire des tests
    vnomdesti = vprefixe & "." & vsuffixe
    vnom = vnomdesti
 Else  ' La frame de l'extension n'est pas visible
    vnomdesti = vnom
 End If ' Si la frame pour l'extension est visible

fileop.AddSourceFile chemin + vnomorig
If RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
 vnom = RemIllegals(vnom)
 vnomdesti = RemIllegals(vnomdesti)
End If
fileop.AddDestFile chemin + vnom
état.Panels(1). + vnomorig + " to " + vnom
état.Panels(2). + Trim$(Str$(vnbtot))
    
    If acpreview = False Then  ' On n'est pas en preview
     If temoin1 = 1 Then ' Undofile
      If recursive = False Then
       Print #1, "ren " + Chr$(34) + vnom + Chr$(34) + " " + Chr$(34) + vnomorig + Chr$(34)
      End If
     End If
     
     If temoin3 = 1 Then ' Logfile
      Print #3, Str$(Date) + " " + Str$(Time) + " => " + chemin + vnomorig + " is rename to " + chemin + vnom
     End If
     
     If temoin2 = 1 Then ' Batch
      If recursive = False Then
       Print #2, "ren " + Chr$(34) + vnomorig + Chr$(34) + " " + Chr$(34) + vnom + Chr$(34)
      End If
     Else
      ' Les ajouts dans les listes sont pour l'UNDO
      If TemoinTrop = False Then
        List2.AddItem chemin + vnomorig ' Nom d'origine.
        List3.AddItem chemin + vnomdesti   ' Nom d'arrivée.
      End If
      If UseHistory = True Then
       If TemoinTrop = False Then
        lhistory.AddItem Trim$(Str$(Time())) + "|" + chemin + "|" + vnomorig + "|" + vnomdesti ' Historique
       End If
      End If
      If Len(Trim$(prog1)) <> 0 Then ' Lancer un programme avant de renommer le fichier
      prog11 = prog1
       ExecCmd prog11, chemin + vnomorig
      End If
      If CopyRename = True Then  ' Renommer les fichier et copier
       If Not fileop.RenameFiles Then
       End If
      Else ' On copie les fichiers, on ne les renomme pas
       If Not fileop.CopyFiles Then
       End If
      End If
     End If
     If Len(Trim$(prog2)) <> 0 Then ' Lancer un programme après avoir renommé le fichier
      prog22 = prog2
      ExecCmd prog22, chemin + vnomdesti
     End If
     DT1.SetFileDateTime (chemin + vnomdesti)
      
'     If ChgFilesAttr = True Then ' On change les attributs du fichier
      Attr1.ChangeAttr (chemin + vnomdesti)
'      SetAttr chemin + vnomdesti, lesattributs
'     End If
    Else ' On est en preview *********************************************************************
     If recursive = False Then
      If TemoinTrop = False Then
        vMontre = True
        If ShowWhenFileNameChange = 1 Then
            If Trim$(vnomorig) = Trim$(vnomdesti) Then
                vMontre = False
            End If
        End If
        If vMontre Then
            Set ItemX = preview.listPreview.ListItems.Add(, , vnomorig)
            ItemX.SubItems(1) = vnomdesti
            preview.listsav.AddItem vnomdesti
        End If
      End If
     Else
      If TemoinTrop = False Then
        vMontre = True
        If ShowWhenFileNameChange = 1 Then
            If Trim$(vnomfichier) = Trim$(chemin + vnomdesti) Then
                vMontre = False
            End If
        End If
        If vMontre Then
            Set ItemX = preview.listPreview.ListItems.Add(, , vnomfichier)
            ItemX.SubItems(1) = chemin + vnomdesti
            preview.listsav.AddItem vnomdesti
        End If
      End If
     End If
    End If ' Preview ou pas ?
    
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

état.Panels(1).
RENAME.MousePointer = 0
Rem ** Mise à jour des paramètres
LastUseDate = Date
LastUseTime = Time
NumberOFiles = vnbtot
LastDirectory = Dir1Path

If acpreview = True Then
 preview.Command7.Visible = True
 preview.Command1.Visible = True
 preview.Command6.Visible = True
 preview.Command2.Visible = True
 preview.Command3.Visible = True
 preview.Command4.Visible = True
 If vTemCopy = False Then
  preview.Command5.Visible = True
 End If
 If recursive = True Then
  preview.Command5.Visible = False
 End If
Else ' On n'est pas en preview
 If Len(Trim$(prog3)) <> 0 Then ' Lancer un programme après avoir renommé tous les fichiers
  prog33 = prog3
  ExecCmd prog33, ""
 End If
End If
Close #1
Close #2
Close #3

If List2.ListCount > 0 Then
 mundo.Enabled = True
End If

état.Panels(1).
état.Panels(2).

If ShutDown = True And acpreview = False Then
 Unload Me
End If

Exit Sub

Erreur1:
 MsgBox "Error, unable to create the undo file, verify it's name and path (did you only specify a name?)"
 Exit Sub

Erreur2:
 MsgBox "Error, unable to create the batch file, verify it's name path  (did you only specify a name?)"
 Exit Sub

Erreur3:
 MsgBox "Error, unable to create the log file, verify it's name and path  (did you only specify a name?)"
 Exit Sub

ErrGen:
    If Err.Number = 53 Then ' File not found...
        Resume Next
    End If
    If Err.Number = 62 And Temoin11 = True Then
        Resume Next
    End If
 ErreurGrave "StartRename"
 Exit Sub
End Sub

Private Sub StepSelection()
 Dim vretour As String
 Dim i As Long
 Dim pas As Integer
 Dim vnb As Long
 Dim vnb2 As Long
 vnb = 0
 vretour = ""
 vretour = InputBox("Enter step for selection", "Step", "2")
 If vretour = "" Then
  Exit Sub
 End If
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
End Sub

Private Sub Unselect()
 Dim i As Long
 Dim vnb As Long
 letat = True
 RENAME.MousePointer = 11
 ListView1.Visible = False
 vnb = ListView1.ListItems.Count - 1
 For i = 0 To vnb
   'vretour = LVSetItemNotSelected(ListView1, i)
   LVSetItemNotSelected ListView1, i
 Next
 ListView1.Visible = True
 RENAME.MousePointer = 0
 état.Panels(4).
 letat = False
End Sub
Private Sub Acdsee_DblClick()
    Dim vret As Boolean
    Dim vtmp As String
    If recursive = False Then
        vtmp = Dir1Path
        If right(Dir1Path, 1) <> "\" Then
            vtmp = vtmp + "\"
        End If
    Else
        vtmp = ""
    End If
    vret = FViewPict.ChargeImage(vtmp & ListView1.SelectedItem.Text, ListView1.SelectedItem.Text)
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
  If CompleCounters = 0 Then
    If AskQuestion = 0 Then ' Variable cachée dans la base de registres pour éviter la question chiante
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
txtlang.
txtlang.SetFocus
End Sub
Private Sub cmdtxt1_GotFocus()
 SelAll cmdtxt1
End Sub
Private Sub cmdtxt2_GotFocus()
SelAll cmdtxt2
End Sub
Private Sub cmdtxt3_GotFocus()
SelAll cmdtxt3
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
 état.Panels(1).
 Command3.Visible = False
 Frame2.Visible = False
 laides.Visible = False
 panelcmd.Visible = False
' Frame1.height = 5000
 Frame1.height = Frame3.top - Frame1.top
' PanelList.ZOrder 0
 PanelList.Move FrameDroite.left + Frame1.left + PanelPrefix.left, Frame1.top + PanelPrefix.top + FrameDroite.top
 Set PanelList.Container = TabGen
 PanelList.Visible = True
 PanelList.height = Frame1.height - Combo1.height - Combo1.top - 140
 m_oAutoPos.RefreshPositions
 Exit Sub
Else
 Frame1.height = 2600
 PanelList.Visible = False
 Frame2.Visible = True
 laides.Visible = True
End If

If Trim$(Combo1.List(Combo1.ListIndex)) = "Free form" Then
 état.Panels(1).
' panelcmd.ZOrder 0
 'm3contextuel.Visible = True
 'ChargeYourCmd
 panelcmd.Move FrameDroite.left + Frame1.left + PanelPrefix.left, Frame1.top + PanelPrefix.top + FrameDroite.top
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
' paneltext.ZOrder 0
 paneltext.Move FrameDroite.left + Frame1.left + PanelPrefix.left, Frame1.top + PanelPrefix.top + FrameDroite.top
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
  état.Panels(1).
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
  état.Panels(1).
 End If
ModifierLigne
End Sub

Private Sub Combo3_Change()
 If Combo3.List(Combo3.ListIndex) = "Letters" Then
  If Text3. Then
   Text3.
  End If
 End If
End Sub

Private Sub Combo3_Click()
 If Combo3.List(Combo3.ListIndex) = "Letters" Then
  If Text3. Then
   Text3.
  End If
 End If
End Sub

Private Sub Combo4_Change()
 If Combo4.List(Combo4.ListIndex) = "Letters" Then
  If Text16. Then
   Text16.
  End If
 End If
End Sub

Private Sub Combo4_Click()
 If Combo4.List(Combo4.ListIndex) = "Letters" Then
  If Text16. Then
   Text16.
  End If
 End If
End Sub

Private Sub Combo5_Click()
 Dim vnb As Long
 Filtre = Trim$(Combo5.Text)
 If right$(Filtre, 1) = ";" Then
  Filtre = left$(Filtre, Len(Filtre) - 1)
 End If
 Combo5.Text = Filtre
 vnb = remplissage()
 état.Panels(3).Text = Trim$(Str$(vnb))
 état.Panels(4).
End Sub

Private Sub Combo5_GotFocus()
    SelAll Combo5
End Sub

Private Sub Combo5_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim vnb As Long
 If KeyCode = 13 Then
  Filtre = Trim$(Combo5.Text)
  If right$(Filtre, 1) = ";" Then
   Filtre = left$(Filtre, Len(Filtre) - 1)
  End If
  Combo5.Text = Filtre
  vnb = remplissage()
  état.Panels(3).Text = Trim$(Str$(vnb))
  état.Panels(4).
 End If
End Sub
Private Sub Combo6_DblClick()
    ' Le double clic est pris comme une validation
    txtlang.Text = left$(txtlang.Text, LaPosSauve - 1) + Trim$(Mid$(Combo6.List(Combo6.ListIndex), LongSauve + 1)) + Trim$(Mid$(txtlang.Text, txtlang.SelStart))
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
    txtlang.Text = left$(txtlang.Text, LaPosSauve - 1) + Trim$(Mid$(Combo6.List(Combo6.ListIndex), LongSauve + 1)) + Trim$(Mid$(txtlang.Text, txtlang.SelStart))
    txtlang.SelStart = LaPosSauve + Len(Trim$(Mid$(Combo6.List(Combo6.ListIndex), LongSauve + 1))) - 1
    Combo6.Visible = False
    Combo6.Clear
End If
End Sub

Private Sub Combo6_LostFocus()
    Combo6.Visible = False
End Sub

Private Sub Command1_Click()
 msearchpref_Click
End Sub

Private Sub Command10_Click()
 UseMP3 = False
 UseVQF = False
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
 Text2.
 Option1(1).Value = False
 Text14.
 Option2(0).Value = True
 Option2(1).Value = False
 Check3.Value = 0
 Text3.
 Text4.
 Text5.
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
 rech1.ResetSearch
End Sub

Private Sub Command11_Click()
Check11.Value = 0
Text16.
Text17.
Text18.
Combo4.ListIndex = 0
Option3(26).Value = False
Option3(25).Value = False
Option3(24).Value = False
Option4(0).Value = False
Text8.
Text15.
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
rech2.ResetSearch
End Sub

Private Sub Command12_Click()
Dim szFilename As String
Dim chemin As String
Dim ligne As String, chaine1 As String, chaine2 As String
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
If ListDelimiter = 0 Then
 ListDelimiter = 9
End If
Open szFilename For Input As #1
Line Input #1, ligne
While Not EOF(1)
 If Len(Trim$(ligne)) > 0 Then
  If RemoveGuill = 1 Then
    ligne = Replace(ligne, Chr$(34), "")
  End If
  chaine1 = GetToken(ligne, Chr$(ListDelimiter), 1)
  chaine2 = GetToken(ligne, Chr$(ListDelimiter), 2)
  Set itmX = ListView2.ListItems.Add(, , chaine1)
  itmX.Text = chaine1
  itmX.SubItems(1) = chaine2
 End If
 Line Input #1, ligne
Wend
Close #1
 If Len(Trim$(ligne)) > 0 Then
  If RemoveGuill = 1 Then
    ligne = Replace(ligne, Chr$(34), "")
  End If
  chaine1 = GetToken(ligne, Chr$(ListDelimiter), 1)
  chaine2 = GetToken(ligne, Chr$(ListDelimiter), 2)
  Set itmX = ListView2.ListItems.Add(, , chaine1)
  itmX.Text = chaine1
  itmX.SubItems(1) = chaine2
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
Dim r As Long
Dim vnb As Long
ListView2.Visible = False
RENAME.MousePointer = 11
vnb = ListView2.ListItems.Count
For i = vnb To 0 Step -1
  If LVIsSelected(ListView2, i) = True Then
   r = LVRemoveItem(ListView2, i)
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
Dim i As Long, sItem As String
RENAME.MousePointer = 11
ListView2.Visible = False
i = LVGetItemSelected(ListView1, -1)
While i <> -1
 sItem = LVGetName(ListView1, i)
 Set itmX = ListView2.ListItems.Add(, , sItem)
 itmX.Text = sItem
 itmX.SubItems(1) = sItem
 i = LVGetItemSelected(ListView1, i)
Wend
ListView2.Visible = True
RENAME.MousePointer = 0
End Sub

Private Sub Command16_Click()
Dim i As Long
Dim vnb As Long
RENAME.MousePointer = 11
ListView2.Visible = False
vnb = ListView1.ListItems.Count - 1
For i = 0 To vnb
 sItem = LVGetName(ListView1, i)
 Set itmX = ListView2.ListItems.Add(, , sItem)
 itmX.Text = sItem
 itmX.SubItems(1) = sItem
Next i
ListView2.Visible = True
RENAME.MousePointer = 0
End Sub
Private Sub Command19_Click()
 ffolder.Show 1
End Sub

Private Sub Command2_Click()
 msearchext_Click
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
szFilename = DialogFile(Me.hWnd, 2, "Save as", "rename.list", "Text" & Chr$(0) & "*.list" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "list")

RENAME.MousePointer = 11
If Trim$(szFilename) = "" Then
 RENAME.MousePointer = 0
 Exit Sub
End If
Open szFilename For Output As #1
If ListDelimiter = 0 Then
    ListDelimiter = 9
End If
ListView2.Visible = False
vnb = ListView2.ListItems.Count - 1
For i = 0 To vnb
 sItem1 = LVGetName(ListView2, i)
 sItem2 = LVGetItemName(ListView2, i, 1)
 Print #1, sItem1 & Chr$(ListDelimiter) & sItem2
Next i
Close #1
RENAME.MousePointer = 0
ListView2.Visible = True
Beep
End Sub

Private Sub Command8_Click()
 OptionsCyclic = True
 Fcyclic.Show 1
End Sub
Private Sub etat1_Click()
 On Error Resume Next
 PosEcran = PosEcran + 1
 If PosEcran > 5 Then
  PosEcran = 1
 End If
 Select Case PosEcran
  Case 1 ' en haut à gauche
   RENAME.top = 0
   RENAME.left = 0
  Case 2 ' en bas à gauche
   RENAME.top = Screen.height - RENAME.height
   RENAME.left = 0
  Case 3 ' en bas à droite
   RENAME.top = Screen.height - RENAME.height
   RENAME.left = Screen.width - RENAME.width
  Case 4 ' en haut à droite
   RENAME.top = 0
   RENAME.left = Screen.width - RENAME.width
  Case 5 ' au centre
   RENAME.Move (Screen.width - Me.width) / 2, (Screen.height - Me.height) / 2
 End Select
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
 If ShowPathInCaption = 1 Then
  Me. + Dir1Path
 End If
 vnb = 0
 vnb = remplissage()
 état.Panels(3).Text = Trim$(Str$(vnb))
 état.Panels(4).
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
Dim vnb As Long
letat = False
'If CBool(GetKeyState(VK_CONTROL) And 1) Then
'    If (KeyCode > 111 And KeyCode < 123) Then ' Appel aux favoris 48 58
'        menufav_Click (KeyCode - 112)
'    End If
'End If

If KeyCode = 27 Then   ' Esc
    annuler = True
End If
If KeyCode = 116 Then ' Refresh (F5)
    RefreshF5
End If
End Sub
Private Sub Form_Load()
Dim leretour As Integer
Dim mvnb1 As Integer, mvnb2 As Integer, mvnb3 As Integer, vnb As Long
Dim tlibdate(3) As String
Dim vretour As Variant
On Error GoTo ErrGen
Set cLstWin = New CListSearch
Set cLstWin.Client = Combo6
AppPath = App.path
If right$(App.path, 1) <> "\" Then
    AppPath = AppPath + "\"
End If
OkUseAbbrev = False ' Par défaut on n'utilise pas les abréviations
UseMP3 = False
UseVQF = False
PbFtv1 = False
CurrentCommand = 0
LoadTags ' Charge les tags pour les TIF
UseCylcic = False
VnbCyclic = 0
FavEncours = -1
TemShift = False
OptionsCyclic = False
Piège1 = False
' MRU
Dim cR As New cRegistry
cR.ClassKey = HKEY_CURRENT_USER
cR.Section
m_cMRU.Load cR
m_cMRU.MaxFileCount = 5
pDisplayMRU False
' Fin MRU
PosEcran = 0
VnbHistory = 0
RechPref = False
RechSuff = False
FolderOk = False
Folder1 = 0
Folder5 = "1"
Folder2 = 0
Folder3 = 0
Folder6 = " "
Folder4 = 0
tlibdate(1) = "Created"
tlibdate(2) = "Modified"
tlibdate(3) = "Access"
VnbRep = 0
TemMove = False
'AjoutEnCours = True
recursive = False
m2recursive.Checked = False
VancRep = ""
vnb = 0
ChargeVNBCommandes
'msave.Enabled = False  ' Désactivation du menu "save"
lhistory.Clear
LeCancel = False
aOuvrir = False
' mundo.Enabled = False
acpreview = False
vnboptionp = UBound(optionp)
vnboptions = UBound(options)
vnbcmd = UBound(hlplang)  ' Nombre de commandes du langage
zshift = False
App.HelpFile = AppPath + "therename.hlp"
Filtre = "*.*"
TemDelete = False
EtchedLine RENAME, 0, Toolbar1.top + Toolbar1.height + 50, RENAME.width
OtDate(1, 1) = "d"
OtDate(2, 1) = "dd"
OtDate(3, 1) = "ddd"
OtDate(4, 1) = "dddd"
OtDate(5, 1) = "ddddd"
OtDate(6, 1) = "w"
OtDate(7, 1) = "ww"
OtDate(8, 1) = "m"
OtDate(9, 1) = "mm"
OtDate(10, 1) = "mmm"
OtDate(11, 1) = "mmmm"
OtDate(12, 1) = "y"
OtDate(13, 1) = "yy"
OtDate(14, 1) = "yyyy"
OtDate(15, 1) = "h"
OtDate(16, 1) = "hh"
OtDate(17, 1) = "n"
OtDate(18, 1) = "nn"
OtDate(19, 1) = "s"
OtDate(20, 1) = "ss"
OtDate(21, 1) = "ttttt"
OtDate(22, 1) = "AM/PM"
OtDate(23, 1) = "am/pm"
OtDate(1, 2) = "Return the day as a number without any leading zero (1 - 31)."
OtDate(2, 2) = "Return the day as a number with a leading zero (01 for example)"
OtDate(3, 2) = "Return the day as an abbreviation (Sun - Sat)"
OtDate(4, 2) = "Return the day as a full name (Sunday - Saturday)"
OtDate(5, 2) = "Return the date as a complete date (including day, month, and year), formatted according to your  system's short date format setting"
OtDate(6, 2) = "Return the day of the week as a number (1=Sunday ... 7=Saturday)"
OtDate(7, 2) = "Return the week of the year as a number (1 - 54)"
OtDate(8, 2) = "Return the month as a number without any leading zero (1 - 12)"
OtDate(9, 2) = "Return the month as a number with a leading zero (01 - 12)"
OtDate(10, 2) = "Return the month as an abbreviation (Jan - Dec)"
OtDate(11, 2) = "Return the month as a full month name (January - December)"
OtDate(12, 2) = "Return the day of the year as a number (1 - 366)"
OtDate(13, 2) = "Return the year as a 2-digit number (00 - 99)"
OtDate(14, 2) = "Return the year as a 4-digit number (100 - 9999)"
OtDate(15, 2) = "Return the hour as a number without leading zeros (0 - 23)"
OtDate(16, 2) = "Return the hour as a number with leading zeros (00 - 23)"
OtDate(17, 2) = "Return the minute as a number without leading zeros (0 - 59)"
OtDate(18, 2) = "Return the minute as a number with leading zeros (00 - 59)"
OtDate(19, 2) = "Return the second as a number without leading zeros (0 - 59)"
OtDate(20, 2) = "Return the second as a number with leading zeros (00 - 59)"
OtDate(21, 2) = "Return a time as a complete time (including hour, minute, and second)"
OtDate(22, 2) = "Use the 12-hour clock and return an uppercase AM with any hour before noon; display an uppercase PM with any hour between noon and 11:59 P.M"
OtDate(23, 2) = "Use the 12-hour clock and return a lowercase AM with any hour before noon; display a lowercase PM with any hour between noon and 11:59 P.M"

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
langage(9) = "<FileContent>"
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
langage(59) = "<PRBefore,0>"
hlplang(59) = 319
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

LngCmd(1, 1) = 0
LngCmd(2, 1) = 0
LngCmd(3, 1) = 0
LngCmd(4, 1) = 0
LngCmd(5, 1) = 0
LngCmd(6, 1) = 0
LngCmd(7, 1) = 0
LngCmd(8, 1) = 0
LngCmd(9, 1) = 0
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
LngCmd(9, 2) = 0
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
LngCmd(59, 1) = 9
LngCmd(59, 2) = 1
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

' Pour le retaillage
m_oAutoPos.AddAssignment Me.Frame3, Me.FrameDroite, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.etat1, Me.FrameDroite, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.etat2, Me.FrameDroite, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.etat3, Me.FrameDroite, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.Command13, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.Command12, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.Command7, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.Command15, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.Command16, Me.PanelList, tCONTAINER_RELATIVE_POS_BOTTOM
m_oAutoPos.AddAssignment Me.ListView2, Me.PanelList, tCONTAINER_HEIGHT_DELTA_BOTTOM

' Chargement des règles
LesRegles.LoadRulesFromFile AppPath & "Rules.ini"

' Chargement des commandes dans les menus contextuels
For i = 1 To vnbcmd
 listcmd.AddItem langage(i)
 listcmd.ItemData(listcmd.NewIndex) = hlplang(i)    ' Ajout du topic du fichier d'aide a appeler
Next i
LoadAvailableDrives
mvnb1 = -1
mvnb2 = -1
mvnb3 = -1
mvnb4 = -1
For i = 0 To vnbcmd - 1
    If UCase$(left$(listcmd.List(i), 3)) = "<PR" Then
        mvnb1 = mvnb1 + 1
        If mvnb1 <> 0 Then Load m3cmdprefix(mvnb1)
        m3cmdprefix(mvnb1).Caption = listcmd.List(i)
    Else
        If UCase$(left$(listcmd.List(i), 3)) = "<EX" Then
            mvnb2 = mvnb2 + 1
            If mvnb2 <> 0 Then Load m3cmdextension(mvnb2)
            m3cmdextension(mvnb2).Caption = listcmd.List(i)
        Else
            If UCase$(left$(listcmd.List(i), 4)) = "<VQF" Or UCase$(left$(listcmd.List(i), 4)) = "<MP3" Then
                mvnb4 = mvnb4 + 1
                If mvnb4 <> 0 Then Load mimusic(mvnb4)
                mimusic(mvnb4).Caption = listcmd.List(i)
            Else
                mvnb3 = mvnb3 + 1
                If mvnb3 <> 0 Then Load mlang(mvnb3)
                mlang(mvnb3).Caption = listcmd.List(i)
            End If
        End If
    End If
Next i

For i = 1 To vnboptionp
 Combo1.AddItem (optionp(i))
Next
For i = 1 To vnboptions
 Combo2.AddItem (options(i))
Next
AncTitre = Me.Caption
rafraichir = True
vnb = remplissage()
état.Panels(3).Text = Trim$(Str$(vnb))

Rem *** Lecture des paramètres pas défaut.
leretour = LoadPref()
FolderTreeview1(0).VirtualFolders = IncVirtualFolders
FolderTreeview1(0).HiddenFolders = IncHiddenFolders
If ToolbarButtons = 1 Then
 RENAME.Toolbar1.Style = tbrFlat
Else
 RENAME.Toolbar1.Style = tbrStandard
End If
ListView1.ColumnHeaders(3).Text = tlibdate(Dateformat + 1)
Combo1.ListIndex = DefOption1
Combo2.ListIndex = DefOption2
Combo1.Text = Combo1.List(DefOption1)
Combo2.Text = Combo2.List(DefOption2)
If ColumnsWiths = True Then
 ListView1.ColumnHeaders(1).width = WCol1
 ListView1.ColumnHeaders(2).width = WCol2
 ListView1.ColumnHeaders(3).width = WCol3
 ListView1.ColumnHeaders(4).width = WCol4
End If

mviewmp3tab.Checked = ShowMP3Tab
mviewpicturetab.Checked = ShowMusicTab
TabGen.TabVisible(2) = ShowMP3Tab
TabGen.TabVisible(3) = ShowMusicTab

If UseHistory = True Then ' Menu history
 mhistory.Enabled = True
Else
 mhistory.Enabled = False
End If
ListView1.FullRowSelect = FullRow
ListView1.GridLines = GridLines
If Center0rSave = 1 Then ' centrer la fenêtre
    Me.Move (Screen.width - Me.width) / 2, (Screen.height - Me.height) / 2
Else ' Rappeler sa position précédente
    RENAME.left = lLeft
    RENAME.top = lTOp
End If
RENAME.WindowState = wWindowState

If RememberWSize = 1 Then
    If wHeight <> -1 And wWidth <> -1 Then
        RENAME.height = wHeight
        RENAME.width = wWidth
    Else
        RENAME.WindowState = vbMaximized
    End If
End If
NumberOfRuns = NumberOfRuns + 1
NomSettings = ""
 txtlang.Text = DefaultCommand
 If RememberLastCommand = 1 Then
    txtlang.Text = LastCommand
 End If
 ' Mise en place des valeurs de recherche et de remplacement par défaut
 'rech1.SearchString = DefaultSearchPr
 'rech1.ReplaceString = DefaultReplacePr
 'rech2.SearchString = DefaultSearchExt
 'rech2.ReplaceString = DefaultReplaceExt

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
  If StartupOnLastDir = 1 Then ' il faut démarrer dans le dernier répertoire de travail
   FolderTreeview1(0).SelectedFolder = LastDirectory
   Dir1Path = LastDirectory
  Else
    If Trim$(StartupDir) <> "" Then
        FolderTreeview1(0).SelectedFolder = StartupDir
        PbFtv1 = True
        Dir1Path = StartupDir
    End If
    If RememberLastFolder = 1 Then
        If Trim$(LastFolder) <> "" Then
            FolderTreeview1(0).SelectedFolder = LastFolder
            Dir1Path = LastFolder
            PbFtv1 = True
        End If
    End If
    'FolderTreeview1(0).SelectedFolder = StartupDir
    
    If ShowPathInCaption = 1 Then
        Me. + Dir1Path
    End If
  End If
  If UseAutoSave = 1 Then
   aOuvrir = True
   NomSettings = AppPath + "autosave.ren"
   mopenset_Click
  End If
End If
Exit Sub

ErrGen:
ErreurGrave "Form_Load"
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim cR As New cRegistry
 If RememberLastCommand = 1 Then
    LastCommand = txtlang.Text
 End If
 If RememberLastFolder = 1 Then
    LastFolder = FolderTreeview1(0).SelectedFolder
  End If
  wWidth = RENAME.width
  wHeight = RENAME.height

cR.ClassKey = HKEY_CURRENT_USER
cR.Section
m_cMRU.Save cR
Unload preview
lLeft = RENAME.left
lTOp = RENAME.top
resultat = Savesettings()
If UseAutoSave = 1 Then
    NomSettings = AppPath + "autosave.ren"
    msave_Click ' et on lance la sauvegarde dans la répertoire du programme
End If
End Sub

Private Sub Form_Resize()
 On Error Resume Next
 If RENAME.width < 11085 Then
    RENAME.width = 11085
 End If
 If RENAME.height < 7740 Then
    RENAME.height = 7740
 End If
 EtchedLine RENAME, 0, Toolbar1.top + Toolbar1.height + 50, RENAME.width
 ListView1.height = Me.ScaleHeight - état.height - ListView1.top
 TabGen.left = Me.ScaleWidth - TabGen.width - 10
 TabGen.height = ListView1.height
 FolderTreeview1(0).height = TabGen.height - TabGen.TabHeight - 150
 LvMP3.height = FolderTreeview1(0).height
 Acdsee.height = Acdsee.width
 If Acdsee.height > FolderTreeview1(0).height Then
    Acdsee.height = FolderTreeview1(0).height
 End If
 
 ListView1.width = TabGen.left - ListView1.left - 50
 Combo1_Click
 FrameDroite.height = TabGen.height - 500
 m_oAutoPos.RefreshPositions
 If PanelList.Visible = True Then
     Frame1.height = Frame3.top - Frame1.top
    PanelList.height = Frame1.height - Combo1.height - Combo1.top - 140
 End If
  m_oAutoPos.RefreshPositions
End Sub

Private Sub Form_Terminate()
 If acpreview = True Then
  Unload preview
  acpreview = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Dim RetVal As Long
 Dim Flag As Long
 If acpreview = True Then
  Unload preview
  acpreview = False
  End If
 lLeft = RENAME.left
 lTOp = RENAME.top
  wWidth = RENAME.width
  wHeight = RENAME.height
 
 If RememberLastCommand = 1 Then
   LastCommand = txtlang.Text
 End If
 If RememberLastFolder = 1 Then
    LastFolder = FolderTreeview1(0).SelectedFolder
  End If
 
 resultat = Savesettings()
End Sub
Private Sub HTMLReport_Click()
Dim i As Long
Dim sItem As String, vnb As Long, taille As Long
Dim vnb2 As Long
Dim vtempo As String
Dim vrep As String
Dim szFilename As String
Dim Largeur As String
Dim LeLien As String
Dim VnbTagMP3 As Integer
Dim VnbTagVQF As Integer
Dim Boucle1 As Integer
Dim ChMP3 As String
Dim ChVQF As String
Dim MP3Tags(57) As String
Dim VQFTags(9) As String
Dim SonInfo As String
Dim vNomComplet As String

Largeur = "20%"
If IncludePictInfo = 0 Then
    Largeur = "35%"
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
If recursive = False Then
    vrep = Dir1Path
    If right$(vrep, 1) <> "\" Then
        vrep = vrep + "\"
    End If
Else
    vrep = ""
End If
Open szFilename For Output As #1
Print #1, "<!DOCTYPE HTML PUBLIC " + Chr$(34) + "-//W3C//DTD HTML 4.0 Transitional//EN" + Chr$(34) + ">"
Print #1, "<HTML>" + vbCrLf + "<HEAD>" + vbCrLf + "<TITLE>THE Rename</TITLE>" + vbCrLf + "</HEAD>"
Print #1, "<BODY BGCOLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + " #000000" + Chr$(34) + ">"
Print #1, "<H1 ALIGN=" + Chr$(34) + "center" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#0000FF" + Chr$(34) + " SIZE=" + Chr$(34) + "+4" + Chr$(34) + ">THE Rename</FONT></H1>"
Print #1, "<DIV ALIGN=" + Chr$(34) + "CENTER" + Chr$(34) + "><CENTER>"
Print #1, "<P><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + " COLOR=" + Chr$(34) + "#0000FF" + Chr$(34) + ">" + "Content of directory " + Dir1Path + "   on " + Format$(Date, "Long Date") + " at " + Format$(Time, "Long Time") + "</FONT></P>"
Print #1, "<TABLE BORDER=" + Chr$(34) + "3" + Chr$(34) + " CELLSPACING=" + Chr$(34) + "0" + Chr$(34) + " CELLPADDING=" + Chr$(34) + "2" + Chr$(34) + " WIDTH=" + Chr$(34) + "100%" + Chr$(34) + ">"
Print #1, "<TR>"
Print #1, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Name</FONT></TH>"
If HtmlIncFolder = 1 Then
    Print #1, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Folder</FONT></TH>"
End If
If HtmlIncSize = 1 Then
    Print #1, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Size</FONT></TH>"
End If
If HtmlIncDate = 1 Then
    Print #1, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Date</FONT></TH>"
End If
If IncludePictInfo = 1 Then
    Print #1, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Pict Info</FONT></TH>"
End If
If HtmlIncAttr = 1 Then
    Print #1, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">Attrib</FONT></TH>"
End If
If HtmlIncMusic = 1 Then ' Il faut inclure les infos sur les MP3 et sur les VQF
    VnbTagMP3 = NbMP3Tags(ChMP3)
    VnbTagVQF = NbVQFTags(ChVQF)
    For Boucle1 = 1 To VnbTagMP3
        Print #1, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">" + MP3Caption(Boucle1) + "</FONT></TH>"
    Next
    For Boucle1 = 1 To VnbTagVQF
        Print #1, "<TH BGCOLOR=" + Chr$(34) + "#000000" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#FFFFFF" + Chr$(34) + ">" + VQFCaption(Boucle1) + "</FONT></TH>"
    Next
End If
Print #1, "</TR>"
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
    Print #1, "<TR>"
    If IncludeLinks = 1 Then ' Il faut inclure des liens vers les images
        If vtempo <> "&nbsp" Then   ' Seulement si on arrive à lire ses propriétés
            LeLien = "<A HREF=" + Chr$(34) + Prefixe(sItem) & "." & Suffixe(sItem) + Chr$(34) + ">"
            Print #1, "<TD><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp " + LeLien + Prefixe(sItem) & "." & Suffixe(sItem) + "</A></FONT></TD>"
        Else    ' On n'a pas réussi à lire les propriétés de l'image
            Print #1, "<TD><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp " + Prefixe(sItem) & "." & Suffixe(sItem) + "</FONT></TD>"
        End If
    Else    ' Il ne faut pas inclure de liens vers les images.
        Print #1, "<TD><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp " + Prefixe(sItem) & "." & Suffixe(sItem) + "</FONT></TD>"
    End If
    If HtmlIncFolder = 1 Then
        Print #1, "<TD><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp " + ExtractPath(vrep + sItem) + "</FONT></TD>"
    End If
    If HtmlIncSize = 1 Then
        Print #1, "<TD ALIGN=right><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + Format$(LVGetItemName(ListView1, i, 1), "### ### ### ###") + "&nbsp</FONT></TD>"
    End If
    If HtmlIncDate = 1 Then
        Print #1, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + LVGetItemName(ListView1, i, 2) + "</FONT></TD>"
    End If
    If IncludePictInfo = 1 Then
        Print #1, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + vtempo + "</FONT></TD>"
    End If
    vtempo = LVGetItemName(ListView1, i, 3)
    If Trim$(vtempo) = "" Then
        vtempo = "&nbsp;"
    End If
    If HtmlIncAttr = 1 Then
        Print #1, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + vtempo + "</FONT></TD>"
    End If
    
    If HtmlIncMusic = 1 Then ' Il faut inclure les infos sur les MP3 et sur les VQF
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
            VQFTags(2) = Blanc(MusVQF.Bitrate)
            VQFTags(3) = Blanc(MusVQF.Comment)
            VQFTags(4) = Blanc(MusVQF.Copyright)
            VQFTags(5) = Blanc(MusVQF.SaveAsFilename)
            VQFTags(6) = Blanc(MusVQF.Mono_Stereo)
            VQFTags(7) = Blanc(MusVQF.Quality)
            VQFTags(8) = Blanc(MusVQF.SampleRate)
            VQFTags(9) = Blanc(MusVQF.Title)
            For Boucle1 = 1 To VnbTagVQF
                Print #1, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + VQFTags(Val(GetToken(ChVQF, "|", Boucle1))) + "</FONT></TD>"
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
                Print #1, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">" + MP3Tags(Val(GetToken(ChMP3, "|", Boucle1))) + "</FONT></TD>"
            Next
        End If
        If UCase$(Suffixe(sItem)) <> "MP3" And UCase$(Suffixe(sItem)) <> "VQF" Then
            For Boucle1 = 1 To VnbTagMP3 + VnbTagVQF
                Print #1, "<TD ALIGN=center><FONT SIZE=" + Chr$(34) + "-1" + Chr$(34) + ">&nbsp</FONT></TD>"
            Next
        End If
    End If
    Print #1, "</TR>"
    vnb = vnb + 1
    taille = taille + Val(LVGetItemName(ListView1, i, 1))
Next
Print #1, "</TABLE></CENTER></DIV>"
Print #1, "<BR><BR><TABLE BORDER=" + Chr$(34) + "0" + Chr$(34) + " CELLSPACING=" + Chr$(34) + "0" + Chr$(34) + " WIDTH=" + Chr$(34) + "100%" + Chr$(34) + "><TR>"
Print #1, "<TD WIDTH=" + Chr$(34) + "15%" + Chr$(34) + ">Number of files</TD>"
Print #1, "<TD WIDTH=" + Chr$(34) + "85%" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#0000FF" + Chr$(34) + ">" + Trim$(Str$(vnb)) + "</FONT></TD>"
Print #1, "</TR>"
Print #1, "<TR>"
Print #1, "<TD WIDTH=" + Chr$(34) + "15%" + Chr$(34) + ">Total Size</TD>"
Print #1, "<TD WIDTH=" + Chr$(34) + "85%" + Chr$(34) + "><FONT COLOR=" + Chr$(34) + "#0000FF" + Chr$(34) + ">" + Format$(taille, "### ### ### ###") + "</FONT></TD>"
Print #1, "</TR></TABLE><BR><BR><BR>"
Print #1, "</BODY></HTML>"

Close #1
Me.MousePointer = 0
RefreshF5
Exit Sub

ErrorHandler:
 RENAME.MousePointer = 0
 MsgBox "There was a problem while generating the report...!!!"
 Exit Sub

End Sub

Private Sub listcmd_Click()
If Combo6.Visible = True Then
   Combo6.Visible = False
End If
End Sub

Private Sub listcmd_DblClick()
 Dim letexte1 As String, letexte2 As String
 Dim vnewdeb As Integer
 If Len(Trim$(txtlang.Text)) > 0 Then ' S'il y a déjà du texte
  letexte1 = left$(txtlang.Text, txtlang.SelStart)
  letexte2 = Mid$(txtlang.Text, txtlang.SelStart + 1)
  If txtlang.SelLength = Len(Trim$(txtlang.Text)) Then ' si tout est sélectionné, tout effacer !
   txtlang.Text = listcmd.List(listcmd.ListIndex)
   vnewdeb = Len(listcmd.List(listcmd.ListIndex))
  Else ' Tout n'est pas sélectionné, on inssère
   If txtlang.SelLength > 1 Then
    txtlang.SelText = listcmd.List(listcmd.ListIndex)
   Else
    txtlang.Text = letexte1 + listcmd.List(listcmd.ListIndex) + letexte2
    vnewdeb = Len(letexte1) + Len(listcmd.List(listcmd.ListIndex)) ' Postionnement du caret à la fin de ce qui vient d'être insséré
   End If
  End If
 Else ' Il n'y a pas de texte
  txtlang.Text = listcmd.List(listcmd.ListIndex)
  vnewdeb = Len(listcmd.List(listcmd.ListIndex))
 End If
 txtlang.SelStart = vnewdeb
 txtlang.SetFocus
End Sub

Private Sub listcmd_KeyUp(KeyCode As Integer, Shift As Integer)
Dim X As Long
If KeyCode = 112 Then ' F1
    If listcmd.ItemData(listcmd.ListIndex) <> 0 Then
        X = WinHelp(Me.hWnd, App.HelpFile, HELP_CONTEXT, listcmd.ItemData(listcmd.ListIndex))
    End If
End If
End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
 Piège1 = False
 If Not Cancel Then
  On Error GoTo ErrorNewname
  If VancFichier <> Newtext Then
   If RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
    NewString = RemIllegals(NewString)
   End If
   If right$(Trim$(Dir1Path), 1) <> "\" Then
    Name VancFichier As Dir1Path + "\" + NewString
   Else
    Name VancFichier As Dir1Path + NewString
   End If
  End If
  Exit Sub
 End If

ErrorNewname:
  ListView1.SelectedItem.Text = VancFichier
  Exit Sub
 
End Sub

Private Sub ListView1_BeforeLabelEdit(Cancel As Integer)
 On Error GoTo ErrGen
 Piège1 = True
 If right$(Trim$(Dir1Path), 1) <> "\" Then
  VancFichier = Dir1Path + "\" + ListView1.SelectedItem.Text
 Else
  VancFichier = Dir1Path + ListView1.SelectedItem.Text
 End If
Exit Sub
ErrGen:
ErreurGrave "ListView1_BeforeLabelEdit"
End Sub
Private Sub ListView1_Click()
On Error GoTo ErrGen
Dim vnb As Long
Dim Pref As String
Dim vtmp As String
Dim vret As String
If letat = False Then
 vnb = LVGetCountSelected(ListView1)
 état.Panels(4).Text = Trim$(Str$(vnb))
End If
Pref = ""
If ListView1.ListItems.Count > 0 Then
    Pref = UCase$(Suffixe(ListView1.SelectedItem.Text))
End If

If mviewpicturetab.Checked Then
    If Pref = "JPG" Or Pref = "BMP" Or Pref = "GIF" Or Pref = "JPEG" Or Pref = "DIB" Or Pref = "WMF" Or Pref = "EMF" Or Pref = "ICO" Or Pref = "CUR" Then
        Me.MousePointer = vbHourglass
        If recursive = False Then
            vtmp = Dir1Path
            If right(Dir1Path, 1) <> "\" Then
                vtmp = vtmp + "\"
            End If
        Else
            vtmp = ""
        End If
        vret = ImgInfo(vtmp & ListView1.SelectedItem.Text)
        
        If Glob1 <= 397 And Glob2 <= 397 Then
            Acdsee.Stretch = False
        Else
            Acdsee.Stretch = True
            Acdsee.width = 5985
            Acdsee.height = Acdsee.width
        End If
        Acdsee.Picture = LoadPicture(vtmp & ListView1.SelectedItem.Text)
        Me.MousePointer = vbDefault
    End If
End If
If mviewmp3tab.Checked Then
    If Pref = "MP3" Then
        LoadLvMP3
    End If
End If
Exit Sub
ErrGen:
If Err.Number = 481 Then
    MsgBox "Invalid Picture"
    Resume Next
End If
ErreurGrave "ListView1_Click"
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo ErrGen
Dim currSortKey As Integer
RENAME.MousePointer = vbArrowHourglass
   sOrder = Not sOrder
   ListView1.SortKey = ColumnHeader.Index - 1
   ListView1.SortOrder = Abs(Not ListView1.SortOrder = 1)
   Select Case ColumnHeader.Index - 1
      Case 0:
               ListView1.Sorted = True
               If UseNaturalSort = 1 Then
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
RENAME.MousePointer = vbDefault
ListView1.SortKey = ColumnHeader.Index - 1
currSortKey = ListView1.SortKey
Exit Sub
ErrGen:
ErreurGrave "ListView1_ColumnClick"
End Sub

Private Sub ListView1_DblClick()
On Error GoTo ErrGen
 Select Case ActDblClick
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
    MoveFilesUp
 End If
 
 If (tAlt = -127 Or tAlt = -128) And KeyCode = 40 Then ' Déplacer vers le bas
    MoveFilesDown
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
 état.Panels(4).Text = LVGetCountSelected(RENAME.ListView1)
On Error GoTo ErrGen
 If Button = 2 Then
  If recursive = True Then
   madd.Visible = True
  Else
   madd.Visible = False
  End If
  vnb = LVGetCountSelected(ListView1)
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
 Next i
 vrai = True
 VancRep = ExtractPath(Trim$(List1.List(0)))
 For i = 0 To List1.ListCount - 1
  If ExtractPath(Trim$(List1.List(i))) <> VancRep Then
   vrai = False  ' Il faudra se mettre en mode récursif
   Exit For
  End If
 Next i
 ChDir ExtractPath(List1.List(0))
 FolderTreeview1(0).SelectedFolder = ExtractPath(List1.List(0))
 Dir1Path = ExtractPath(List1.List(0))
 
 If vrai = True Then ' On est sur le même répertoire
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
   Next i
  End If
 Else ' On passe en mode récursif
  If Toolbar1.Buttons(13).Value <> 1 Then
   Toolbar1.Buttons(13).Value = tbrPressed
   m2recursive.Checked = True
   MsgBox "Warning, you have dropped files from different directories so I'm going to use the recursive mode"
  End If
  recursive = True
  ListView1.ListItems.Clear
  clsFind.Dateformat = "short Date"
   For i = 0 To List1.ListCount - 1
    strFile = clsFind.Find(List1.List(i), False)
    If Len(strFile) > 0 Then
     If (clsFind.FileAttributes And vbDirectory) = 0 Then
      attributs = clsFind.FileAttributes
      chaine = ""
      If attributs And FILE_ATTRIBUTE_READONLY Then
       chaine = "R"
      End If
      If attributs And FILE_ATTRIBUTE_HIDDEN Then
       chaine = chaine + "H"
      End If
      If attributs And FILE_ATTRIBUTE_SYSTEM Then
       chaine = chaine + "S"
      End If
      If attributs And FILE_ATTRIBUTE_ARCHIVE Then
       chaine = chaine + "A"
      End If
      Set itmX = ListView1.ListItems.Add(, , Trim$(List1.List(i)))
      itmX.Text = Trim$(List1.List(i))
      itmX.SubItems(1) = clsFind.FileSize
      itmX.SubItems(2) = clsFind.GetCreationDate
      itmX.SubItems(3) = chaine
     End If
    End If

   Next i
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
  recursive = True
  m2recursive.Checked = True
 Else
  Toolbar1.Buttons(13).Value = tbrUnpressed
  recursive = False
  m2recursive.Checked = False
 End If
 vnb = remplissage()
 état.Panels(3).Text = Trim$(Str$(vnb))
 état.Panels(4).
End Sub

Private Sub M2Start_Click()
 StartRename
End Sub

Private Sub m3cmdextension_Click(Index As Integer)
 Dim letexte1 As String, letexte2 As String
 Dim vnewdeb As Integer
 If Len(Trim$(txtlang.Text)) > 0 Then ' S'il y a déjà du texte
  letexte1 = left$(txtlang.Text, txtlang.SelStart)
  letexte2 = Mid$(txtlang.Text, txtlang.SelStart + 1)
  If txtlang.SelLength = Len(Trim$(txtlang.Text)) Then ' si tout est sélectionné, tout effacer !
   txtlang.Text = m3cmdextension(Index).Caption
   vnewdeb = Len(m3cmdextension(Index).Caption)
  Else ' Tout n'est pas sélectionné, on inssère
   txtlang.Text = letexte1 + m3cmdextension(Index).Caption + letexte2
   vnewdeb = Len(letexte1) + Len(m3cmdextension(Index).Caption) ' Postionnement du caret à la fin de ce qui vient d'être insséré
  End If
 Else ' Il n'y a pas de texte
  txtlang.Text = m3cmdextension(Index).Caption
  vnewdeb = Len(m3cmdextension(Index).Caption)
 End If
 txtlang.SelStart = vnewdeb
End Sub

Private Sub m3cmdprefix_Click(Index As Integer)
 Dim letexte1 As String, letexte2 As String
 Dim vnewdeb As Integer
 If Len(Trim$(txtlang.Text)) > 0 Then ' S'il y a déjà du texte
  letexte1 = left$(txtlang.Text, txtlang.SelStart)
  letexte2 = Mid$(txtlang.Text, txtlang.SelStart + 1)
  If txtlang.SelLength = Len(Trim$(txtlang.Text)) Then ' si tout est sélectionné, tout effacer !
   txtlang.Text = m3cmdprefix(Index).Caption
   vnewdeb = Len(m3cmdprefix(Index).Caption)
  Else ' Tout n'est pas sélectionné, on inssère
   txtlang.Text = letexte1 + m3cmdprefix(Index).Caption + letexte2
   vnewdeb = Len(letexte1) + Len(m3cmdprefix(Index).Caption) ' Postionnement du caret à la fin de ce qui vient d'être insséré
  End If
 Else ' Il n'y a pas de texte
  txtlang.Text = m3cmdprefix(Index).Caption
  vnewdeb = Len(m3cmdprefix(Index).Caption)
 End If
 txtlang.SelStart = vnewdeb
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
 If attributs And FILE_ATTRIBUTE_READONLY Then
  chaine = "R"
 End If
 If attributs And FILE_ATTRIBUTE_HIDDEN Then
  chaine = chaine + "H"
 End If
 If attributs And FILE_ATTRIBUTE_SYSTEM Then
  chaine = chaine + "S"
 End If
 If attributs And FILE_ATTRIBUTE_ARCHIVE Then
  chaine = chaine + "A"
 End If
 itmX.SubItems(3) = chaine
 état.Panels(3).Text = Trim$(Str$(ListView1.ListItems.Count + 1))
End Sub

Private Sub maddbag_Click()
Dim vret As Integer
vret = Bag(2)
End Sub
Private Sub madddirectory_Click()
Dim repvoulu As String

repvoulu = Trim$(Dir1Path)
For i = 1 To 20
 If fav(i) = repvoulu Then
  Beep
  MsgBox "This directory is already in your favorites"
  Exit Sub
 End If
Next i

For i = 20 To 2 Step -1
 fav(i) = fav(i - 1)
Next i
fav(1) = repvoulu

For i = 0 To 19
 RENAME.menufav(i). + Chr$(65 + i) + " " + fav(i + 1)
 RENAME.mnufav(i). + Chr$(65 + i) + " " + fav(i + 1)
Next i
End Sub

Private Sub maddfavorites_Click()
 madddirectory_Click
End Sub

Private Sub mapropos_Click()
 about.Show 1
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
Dim i As Long
Dim sItem As String
i = 0
chemin = Trim$(Dir1Path)
If right$(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If
If recursive = True Then
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
état.Panels(4).
End Sub

Private Sub mchg2_Click()
' Changes files attributes now
Dim i As Long
Dim sItem As String
i = 0
chemin = Trim$(Dir1Path)
If right$(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If
If recursive = True Then
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
état.Panels(4).
End Sub

Private Sub mchgattdir_Click()
' Changes files attributes now
Dim i As Long
Dim sItem As String
i = 0
chemin = Trim$(Dir1Path)
If right$(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If

AttrEncours = 4
attributs.Show 1

Attr4.ChangeAttr (chemin)
vnb = remplissage()
état.Panels(3).Text = Trim$(Str$(vnb))
état.Panels(4).

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
                If recursive = True Then
                    vtmp = ExtractPath(LVGetName(ListView1, i))
                Else
                    vtmp = Dir1Path
                End If
            Case 4  ' Path + Full Path name
                If recursive = True Then
                    vtmp = LVGetName(ListView1, i)
                Else
                    If right$(Dir1Path, 1) <> "\" Then
                        vtmp = Dir1Path + "\" + LVGetName(ListView1, i)
                    Else
                        vtmp = Dir1Path + LVGetName(ListView1, i)
                    End If
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
                If recursive = True Then
                    vtmp = ExtractPath(LVGetName(ListView1, i))
                Else
                    vtmp = Dir1Path
                End If
            Case 4  ' Path + Full Path name
                If recursive = True Then
                    vtmp = LVGetName(ListView1, i)
                Else
                    If right$(Dir1Path, 1) <> "\" Then
                        vtmp = Dir1Path + "\" + LVGetName(ListView1, i)
                    Else
                        vtmp = Dir1Path + LVGetName(ListView1, i)
                    End If
                End If
        End Select
        chaine = chaine + vtmp + vbCrLf
        i = LVGetItemSelected(ListView1, i)
    Wend
End If
Clipboard.SetText chaine
chaine = ""
RENAME.MousePointer = 0
End Sub

Private Sub mcopy2_Click()
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

If Misc5 = 0 Then ' All
    vnb = ListView1.ListItems.Count
    For i = 0 To ListView1.ListItems.Count - 1
        MoveCopyFile i, vnb, i + 1
    Next
Else    ' Selected
    vnb = LVGetCountSelected(ListView1)
    i = LVGetItemSelected(ListView1, -1)
    While i <> -1
        vnbren = vnbren + 1
        MoveCopyFile i, vnb, vnbren
        i = LVGetItemSelected(ListView1, i)
    Wend
End If
If Misc4 = 2 Then
    RefreshF5
End If
état.Panels(1).
état.Panels(2).
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
i = 0
fileop.ParentWnd = hWnd
fileop.ClearSourceFiles
fileop.ClearDestFiles
chemin = Trim$(Dir1Path)
If right$(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If
If recursive = True Then
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
état.Panels(4).
End Sub

Private Sub mdelrep_Click()
Dim fileop As New CSHFileOp
fileop.ParentWnd = hWnd
fileop.ClearSourceFiles
fileop.ClearDestFiles
fileop.AddSourceFile Dir1Path
If Not fileop.DeleteFiles Then
End If
remplissage
End Sub

Private Sub mdgroupe_Click()
 fmdgroupe.Show 1
 RefreshF5
End Sub

Private Sub mdisconnect_Click()
 Dim r As Long
 r = WNetDisconnectDialog(RENAME.hWnd, RESOURCETYPE_DISK)
End Sub

Private Sub mdosprompthere_Click()
 Dim i
 On Error Resume Next
 ChDrive left$(Dir1Path, 3)
 ChDir Dir1Path
 i = Shell("command.com", 1)
End Sub

Private Sub mend_Click(Index As Integer)
Dim RetVal As Long
 Dim Flag As Long
 lLeft = RENAME.left
 lTOp = RENAME.top
  If RememberLastCommand = 1 Then
    LastCommand = txtlang.Text
 End If
 If RememberLastFolder = 1 Then
    LastFolder = FolderTreeview1(0).SelectedFolder
  End If
 resultat = Savesettings()
 If UseAutoSave = 1 Then
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

Private Sub mexplorer_Click()
FileExecutor Me.hWnd, Dir1Path, "Explore"
End Sub
Private Sub mfilefind_Click()
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
Dim vretour
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
  état.Panels(1). & strFile
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
    état.Panels(1). & vnomorig
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
Next i

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

Private Sub mindex_Click()
'SendKeys "{F1}"
Dim X As Long
X = WinHelp(Me.hWnd, App.HelpFile, HELP_CONTEXT, 1)
End Sub

Private Sub minfos_Click()
 Infos.Show 1
End Sub
Private Sub mlang_Click(Index As Integer)
 Dim letexte1 As String, letexte2 As String
 Dim vnewdeb As Integer
 If Len(Trim$(txtlang.Text)) > 0 Then ' S'il y a déjà du texte
  letexte1 = left$(txtlang.Text, txtlang.SelStart)
  letexte2 = Mid$(txtlang.Text, txtlang.SelStart + 1)
  If txtlang.SelLength = Len(Trim$(txtlang.Text)) Then ' si tout est sélectionné, tout effacer !
   txtlang.Text = mlang(Index).Caption
   vnewdeb = Len(mlang(Index).Caption)
  Else ' Tout n'est pas sélectionné, on inssère
   txtlang.Text = letexte1 + mlang(Index).Caption + letexte2
   vnewdeb = Len(letexte1) + Len(mlang(Index).Caption) ' Postionnement du caret à la fin de ce qui vient d'être insséré
  End If
 Else ' Il n'y a pas de texte
  txtlang.Text = mlang(Index).Caption
  vnewdeb = Len(mlang(Index).Caption)
 End If
 txtlang.SelStart = vnewdeb
End Sub

Private Sub mlast_Click()
  ListView1.ListItems(ListView1.ListItems.Count).Selected = True
  ListView1.ListItems(ListView1.ListItems.Count).EnsureVisible
End Sub

Private Sub mmakedir_Click()
 Dim rep As String
 Dim chem As String
 Dim tmp As String
 Dim vnb As Integer
 rep = ""
 chem = Dir1Path
 If right$(chem, 1) <> "\" Then
  chem = chem + "\"
 End If
 rep = InputBox("Enter directory name", "Make directory", chem)
 If rep = "" Then
  Exit Sub
 End If
 On Error Resume Next
 vnb = CreateNestedFoldersByPath(rep)
 RefreshF5
End Sub
Private Sub mmap_Click()
 Dim r As Long
 r = WNetConnectionDialog(RENAME.hWnd, RESOURCETYPE_DISK)
End Sub

Private Sub mmapped_drives_Click()
 FMappedDrives.Show 1
End Sub

Private Sub mmiddle_Click()
 ListView1.ListItems(Int(ListView1.ListItems.Count / 2)).Selected = True
 ListView1.ListItems(Int(ListView1.ListItems.Count / 2)).EnsureVisible
End Sub

Private Sub mnewformlist_Click()
Dim szFilename As String
Dim chemin As String
Dim ligne As String, chaine1 As String, chaine2 As String
Dim i As Integer
i = 0
If ListView2.ListItems.Count = 0 Then
 MsgBox "error, this can only be used on a list containing files"
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
If ListDelimiter = 0 Then
    ListDelimiter = 9
End If
Open szFilename For Input As #1
Line Input #1, ligne
While Not EOF(1)
 If Len(Trim$(ligne)) > 0 Then
  i = i + 1
  If i <= ListView2.ListItems.Count Then
    If RemoveGuill = 1 Then
      ligne = Replace(ligne, Chr$(34), "")
    End If
    ListView2.ListItems.Item(i).Text = ligne
  End If
 End If
 Line Input #1, ligne
Wend
Close #1
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
'If AjoutEnCours = False Then
 ChangeRepHistorique Index
'Else
' AjoutEnCours = False
'End If
End Sub
Private Sub mnuswap_Click()
' Inverse le nom de 2 fichiers sélectionnés (et seulement de 2)
Dim fileop As New CSHFileOp
Dim i As Long
Dim vnom1 As String
Dim vnom2 As String
Dim chemin1 As String
Dim chemin2 As String
If right$(Trim$(Dir1Path), 1) <> "\" Then
 chemin1 = Trim$(Dir1Path) + "\"
Else
 chemin1 = Trim$(Dir1Path)
End If
If right$(Trim$(Dir1Path), 1) <> "\" Then
 chemin2 = Trim$(Dir1Path) + "\"
Else
 chemin2 = Trim$(Dir1Path)
End If

i = LVGetItemSelected(ListView1, -1)
vnom1 = LVGetName(ListView1, i)
If recursive = True Then
  chemin1 = ExtractPath(vnom1) ' Si on est en récursif, il faut récupérer le chemin du fichier
  If right$(Trim$(chemin1), 1) <> "\" Then
   chemin1 = chemin1 + "\"
  End If
End If
i = LVGetItemSelected(ListView1, i)
vnom2 = LVGetName(ListView1, i)
If recursive = True Then
 chemin2 = ExtractPath(vnom2) ' Si on est en récursif, il faut récupérer le chemin du fichier
 If right$(Trim$(chemin2), 1) <> "\" Then
  chemin2 = chemin2 + "\"
 End If
End If
fileop.ParentWnd = hWnd
fileop.ConfirmOperation = False
fileop.ClearSourceFiles
fileop.ClearDestFiles
' on renomme le premier fichier avec un nom bidon
fileop.AddSourceFile chemin1 + Prefixe(vnom1) & "." & Suffixe(vnom1)
fileop.AddDestFile chemin1 + "$hthouzard$"
fileop.RenameFiles
' on renomme le deuxième fichier
fileop.ClearSourceFiles
fileop.ClearDestFiles
fileop.AddSourceFile chemin2 + Prefixe(vnom2) & "." & Suffixe(vnom2)
fileop.AddDestFile chemin2 + Prefixe(vnom1) & "." & Suffixe(vnom1)
fileop.RenameFiles
' on re renomme le premier
fileop.ClearSourceFiles
fileop.ClearDestFiles
fileop.AddSourceFile chemin1 + "$hthouzard$"
fileop.AddDestFile chemin1 + Prefixe(vnom2) & "." & Suffixe(vnom2)
fileop.RenameFiles
' et on raffaichit
SendKeys "{F5}"
End Sub
Private Sub mopen_Click()
Dim chemin As String
chemin = Trim$(Dir1Path)
If right$(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If
FileExecutor Me.hWnd, chemin + ListView1.ListItems(ListView1.SelectedItem.Index), "Open"
End Sub

Private Sub mopenset_Click()
Dim Version As String
On Error GoTo errloadset

If aOuvrir = False Then
 If Len(Trim$(SettingsDirectory)) > 0 Then
  szFilename = DialogFile(Me.hWnd, 1, "Open settings", "settings.ren", "Rename" & Chr$(0) & "*.ren" & Chr$(0) & "All files" & Chr$(0) & "*.*", SettingsDirectory, "ren")
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
txtlang., szFilename, "txtlang")
Combo1.ListIndex = Val(LoadSet("General", szFilename, "Combo1"))
cmdtxt1., szFilename, "CmdTxt1")
cmdtxt2., szFilename, "CmdTxt2")
cmdtxt3., szFilename, "CmdTxt3")
Text9., szFilename, "Txt9")
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
Text2., szFilename, "Text2")
Option1(1).Value = Val(LoadSet("General", szFilename, "Option1_1"))
Text14., szFilename, "Text14")
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
Text3., szFilename, "Text3")
Text4., szFilename, "Text4")
Text5., szFilename, "Text5")
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
Text16., szFilename, "Text16")
Text17., szFilename, "Text17")
Text18., szFilename, "Text18")
Combo4.ListIndex = Val(LoadSet("General", szFilename, "Combo4"))
Option3(26).Value = Val(LoadSet("General", szFilename, "Option3_26"))
Option3(25).Value = Val(LoadSet("General", szFilename, "Option3_25"))
Option3(24).Value = Val(LoadSet("General", szFilename, "Option3_24"))
Check4.Value = Val(LoadSet("General", szFilename, "Check4"))
Option3(14).Value = Val(LoadSet("General", szFilename, "Option3_14"))
Option3(13).Value = Val(LoadSet("General", szFilename, "Option3_13"))
Option3(12).Value = Val(LoadSet("General", szFilename, "Option3_12"))
Option4(0).Value = Val(LoadSet("General", szFilename, "Option4_0"))
Text8., szFilename, "Text8")
Option4(1).Value = Val(LoadSet("General", szFilename, "Option4_1"))
Text15., szFilename, "Text15")
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

DT1.CreateDFixed = IIf(Trim(LoadSet("Date and Time", szFilename, "CreateDFixed")) <> "", LoadSet("Date and Time", szFilename, "CreateDFixed"), 0)

DT1.CreateDInc1 = Val(LoadSet("Date and Time", szFilename, "CreateDInc1"))
DT1.CreateDInc2 = Val(LoadSet("Date and Time", szFilename, "CreateDInc2"))
DT1.CreateDInc3 = Val(LoadSet("Date and Time", szFilename, "CreateDInc3"))
DT1.CreateTOption = Val(LoadSet("Date and Time", szFilename, "CreateTOption"))
DT1.CreateTFixed = LoadSet("Date and Time", szFilename, "CreateTFixed")
DT1.AccessDOption = Val(LoadSet("Date and Time", szFilename, "AccessDOption"))

DT1.AccessDFixed = IIf(Trim(LoadSet("Date and Time", szFilename, "AccessDFixed")) <> "", LoadSet("Date and Time", szFilename, "AccessDFixed"), 0)

DT1.AccessDInc1 = Val(LoadSet("Date and Time", szFilename, "AccessDInc1"))
DT1.AccessDInc2 = Val(LoadSet("Date and Time", szFilename, "AccessDInc2"))
DT1.AccessDInc3 = Val(LoadSet("Date and Time", szFilename, "AccessDInc3"))
DT1.AccessTOption = Val(LoadSet("Date and Time", szFilename, "AccessTOption"))
DT1.AccessTFixed = LoadSet("Date and Time", szFilename, "AccessTFixed")
DT1.WriteDOption = Val(LoadSet("Date and Time", szFilename, "WriteDOption"))

DT1.WriteDFixed = IIf(Trim(LoadSet("Date and Time", szFilename, "WriteDFixed")) <> "", LoadSet("Date and Time", szFilename, "WriteDFixed"), 0)

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

UseMP3 = Val(LoadSet("MP3", szFilename, "Use"))
MusMP3.Rule = LoadSet("MP3", szFilename, "Rule")
MusMP3.PlaceWhereToPut = Val(LoadSet("MP3", szFilename, "PlaceWhereToPut"))
MusMP3.DefaultArtistToUse = LoadSet("MP3", szFilename, "DefaultArtistToUse")
MusMP3.DefaultYearToUse = LoadSet("MP3", szFilename, "DefaultYearToUse")
MusMP3.DefaultGenreToUse = LoadSet("MP3", szFilename, "DefaultGenreToUse")
MusMP3.DefaultAlbumToUse = LoadSet("MP3", szFilename, "DefaultAlbumToUse")
MusMP3.DefaultTitleToUse = LoadSet("MP3", szFilename, "DefaultTitleToUse")

' Infos sur les VQF
UseVQF = Val(LoadSet("VQF", szFilename, "Use"))
MusVQF.Rule = LoadSet("VQF", szFilename, "Rule")
MusVQF.PlaceWhereToPut = Val(LoadSet("VQF", szFilename, "PlaceWhereToPut"))
MusVQF.DefaultArtistToUse = LoadSet("VQF", szFilename, "DefaultArtistToUse")
MusVQF.DefaultTitle = LoadSet("VQF", szFilename, "DefaultTitle")

' ********************************************************************************
' MRU
SuiteLoad:
m_cMRU.AddFile szFilename
pDisplayMRU True

On Error GoTo Erreur2
'TemMove = False
If Trim$(LoadSet("Folder", szFilename, "CurrentFolder")) <> "<DoNotRestore>" Then
    If Mid$(Dir1Path, 2, 1) = ":" Then
        ChDrive left$(Dir1Path, 1)
    End If
    ChDir Dir1Path
    FolderTreeview1(0).SelectedFolder = Dir1Path
End If
état.Panels(1).
Exit Sub

errloadset:
 MsgBox "There was an error during the process. May be the settings file is corrupted or can't be found. Settings files older than this version can't be opened ... sorry !"
 Exit Sub

Erreur2:
 MsgBox "Error. The program was unable to connect to" + vbCrLf + Dir1Path
 Exit Sub
End Sub

Private Sub moptions_Click()
 doptions.Show 1
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

Private Sub mprint_Click()
Dim chemin As String
chemin = Trim$(Dir1Path)
If right$(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If
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
If recursive = False Then
    vrep = Dir1Path
    If right$(vrep, 1) <> "\" Then
        vrep = vrep + "\"
    End If
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

If DirectoryReport <> 1 Then
    Printer.Print "Filename" + Chr$(9) + Chr$(9) + "Size" + Chr$(9) + "Date" + Chr$(9) + Chr$(9) + "Attributes" + Chr$(9) + Chr$(9) + "Pict Info"
    Printer.Print " "
End If
For i = 0 To ListView1.ListItems.Count - 1
 sItem = LVGetName(ListView1, i)
 If DirectoryReport = 1 Then
  Printer.Print sItem
 Else
  If IncludePictInfo = 1 Then
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
chemin = Trim$(Dir1Path)
If right$(chemin, 1) <> "\" Then
 chemin = chemin + "\"
End If
If recursive = True Then
 chemin = ""
End If
fName = chemin + ListView1.ListItems(ListView1.SelectedItem.Index)
ShowProperties fName, Me
End Sub
Private Sub Mrefresh_Click()
  SendKeys "{F5}"
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
Dim r As Boolean, i As Long
vnb = Val(état.Panels(3).Text)
vnb2 = ListView1.ListItems.Count
For i = vnb2 To 0 Step -1
  If LVIsSelected(ListView1, i) = True Then
   r = LVRemoveItem(ListView1, i)
  vnb = vnb - 1
 End If
Next
état.Panels(3).Text = Trim$(Str$(vnb))
état.Panels(4).
End Sub

Private Sub mrendirect_Click()
Dim NewName As String
Dim fileop As New CSHFileOp
fileop.ParentWnd = hWnd
fileop.ClearSourceFiles
fileop.ClearDestFiles
fileop.ConfirmOperation = False

NewName = InputBox("Select a new name for " + Dir1Path, "Rename directory", Dir1Path)
If Trim$(NewName) = "" Then
 Exit Sub
End If
fileop.AddSourceFile Dir1Path
If RemoveIllegals = 1 Then ' Il faut vérifier qu'il n'y a pas de caractères illégaux et les virer
 NewName = RemIllegals(NewName, True)
End If
fileop.AddDestFile NewName
If Not fileop.RenameFiles Then
End If
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
Next i
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
 If Len(Trim$(SettingsDirectory)) > 0 Then
  szFilename = DialogFile(Me.hWnd, 2, "Save settings as", "settings.ren", "Rename" & Chr$(0) & "*.ren" & Chr$(0) & "All files" & Chr$(0) & "*.*", SettingsDirectory, "ren")
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
 SavSet "General", szFilename, "CmdTxt3", cmdtxt3.Text
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

' extension ***********************
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

' Infos sur les VQF
 SavSet "VQF", szFilename, "Use", UseVQF
 SavSet "VQF", szFilename, "PlaceWhereToPut", MusVQF.PlaceWhereToPut
 SavSet "VQF", szFilename, "DefaultArtistToUse", MusVQF.DefaultArtistToUse
 SavSet "VQF", szFilename, "DefaultTitle", MusVQF.DefaultTitle
 SavSet "VQF", szFilename, "Rule", MusVQF.Rule
RefreshF5
état.Panels(1).
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
Dim szFilename As String
Dim sItem1 As String, sItem2 As String
Dim i As Long
Dim vnb As Long
szFilename = DialogFile(Me.hWnd, 2, "Save as", "rename.list", "Text" & Chr$(0) & "*.list" & Chr$(0) & "All files" & Chr$(0) & "*.*", Dir1Path, "list")

RENAME.MousePointer = 11
If Trim$(szFilename) = "" Then
 RENAME.MousePointer = 0
 Exit Sub
End If
Open szFilename For Output As #1
ListView2.Visible = False
vnb = ListView2.ListItems.Count - 1
For i = 0 To vnb
 sItem1 = LVGetName(ListView2, i)
 Print #1, sItem1
Next i
Close #1
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
 StartupDir = Dir1Path
End Sub

Private Sub mundo_Click()
 Dim vnb As Integer, vnom1 As String, vnom2 As String, i As Integer
 Dim fileop As New CSHFileOp
 fileop.ParentWnd = hWnd
 fileop.ClearSourceFiles
 fileop.ClearDestFiles
 fileop.AllowUndo = False
 fileop.ConfirmOperation = False
 vnb = List2.ListCount
 RENAME.MousePointer = 11
 
 For i = 0 To vnb - 1
  vnom1 = Trim$(List2.List(i)) ' Nom d'origine
  vnom2 = Trim$(List3.List(i)) ' Nouveau nom
  état.Panels(1). + vnom2 + " => " + vnom1
  état.Panels(2). + Trim$(Str$(vnb))
  fileop.AddSourceFile vnom2
  fileop.AddDestFile vnom1
  fileop.RenameFiles
  fileop.ClearSourceFiles
  fileop.ClearDestFiles
 Next i
 
 On Error GoTo ErrUndo
 If Len(UndoFile) > 0 Then
  Open UndoFile For Input As #1
  Close #1
  Dim vretour As Integer
  vretour = MsgBox("An undo file =>" + UndoFile + "<= have been created, would you like to delete it ?", vbOKCancel, "Delete the undo file ?")
  If vretour = vbOK Then
    Kill UndoFile
  End If
 End If

ErrUndo:
 état.Panels(1).
 remplissage
 RENAME.MousePointer = 0
 List2.Clear
 List3.Clear
 mundo.Enabled = False
End Sub
Private Sub mviewbag_Click()
 fBag.Show 1
End Sub

Private Sub mviewmp3tab_Click()
    MTetat1
End Sub

Private Sub mviewpicturetab_Click()
    MTetat2
End Sub

Private Sub Option1_Click(Index As Integer)
 If Index = 0 Then
  Text2.Visible = True
  Command8.top = 495
  Command8.Visible = True
  Text14.Visible = False
  Option2(0).Visible = False
  Option2(1).Visible = False
'  Text2.SetFocus
 Else
  Text2.Visible = False
  Text14.Visible = True
  Command8.top = 850
  Command8.Visible = True
  Option2(0).Visible = True
  Option2(1).Visible = True
'  Text14.SetFocus
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
 If recursive = True Then  ' Mode récursif ***************************************************************************
    remplissage = srecursive()
    If AutoArrange = True Then
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
    If SelectAllFiles = 1 Then
        SelectAll
    End If
    Exit Function
 End If
 
 If rafraichir = False Then
  rafraichir = True
  Exit Function
 End If
 clsFind.Dateformat = "short Date"
 ListView1.ListItems.Clear
 nbfichiers = 0
 chemin = Trim$(Dir1Path)
 If right$(chemin, 1) <> "\" Then
  chemin = chemin + "\"
 End If
 Filtre = Trim$(Filtre)
 ' Suppression des caractères en trop
 If right$(Filtre, 1) = ";" Then
  Filtre = left$(Filtre, Len(Filtre) - 1)
 End If
 If left$(Filtre, 1) = ";" Then
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
            Select Case FilesToInclude
                Case 0 ' Files Only
                    If (clsFind.FileAttributes And vbDirectory) = 0 Then ' Si ce n'est pas un répertoire
                        If InStr(Suffixe(extrait), "*") = 0 Then  ' Test de correspondance sur l'intégralité du masque de sélection
                            If UCase$(Suffixe(strFile)) <> UCase$(Suffixe(extrait)) Then
                                afficher = False
                            End If
                        End If
                        attributs = clsFind.FileAttributes
                        chaine = ""
                        If (attributs And FILE_ATTRIBUTE_READONLY) And ReadOnly = False Then
                            afficher = False
                        End If
                        If (attributs And FILE_ATTRIBUTE_HIDDEN) And Hidden = False Then
                            afficher = False
                        End If
                        If (attributs And FILE_ATTRIBUTE_SYSTEM) And System = False Then
                            afficher = False
                        End If
                        If afficher = True Then
                            If attributs And FILE_ATTRIBUTE_READONLY Then
                                chaine = "R"
                            End If
                            If attributs And FILE_ATTRIBUTE_HIDDEN Then
                                chaine = chaine + "H"
                            End If
                            If attributs And FILE_ATTRIBUTE_SYSTEM Then
                                chaine = chaine + "S"
                            End If
                            If attributs And FILE_ATTRIBUTE_ARCHIVE Then
                                chaine = chaine + "A"
                            End If
                            If chaine = "" Then
                                chaine = " "
                            End If
                            nbfichiers = nbfichiers + 1
                            Set itmX = ListView1.ListItems.Add(, , strFile)
                            itmX.Text = strFile
                            itmX.SubItems(1) = clsFind.FileSize
                            Select Case Dateformat
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
                    If Trim$(strFile) = "." Or Trim$(strFile) = ".." Then
                        afficher = False
                    End If
                    If (attributs And FILE_ATTRIBUTE_READONLY) And ReadOnly = False Then
                        afficher = False
                    End If
                    If (attributs And FILE_ATTRIBUTE_HIDDEN) And Hidden = False Then
                        afficher = False
                    End If
                    If (attributs And FILE_ATTRIBUTE_SYSTEM) And System = False Then
                        afficher = False
                    End If
                    If afficher = True Then
                        If attributs And FILE_ATTRIBUTE_READONLY Then
                            chaine = "R"
                        End If
                        If attributs And FILE_ATTRIBUTE_HIDDEN Then
                            chaine = chaine + "H"
                        End If
                        If attributs And FILE_ATTRIBUTE_SYSTEM Then
                            chaine = chaine + "S"
                        End If
                        If attributs And FILE_ATTRIBUTE_ARCHIVE Then
                            chaine = chaine + "A"
                        End If
                        If chaine = "" Then
                            chaine = " "
                        End If
                        nbfichiers = nbfichiers + 1
                        If (clsFind.FileAttributes And vbDirectory) <> 0 Then ' Si c'est un répertoire
                            Set itmX = ListView1.ListItems.Add(, , strFile)
                            itmX.Bold = True ' *****************************************************************************
                            itmX.SubItems(4) = "Dir"  ' Type répertoire
                        Else ' Ce n'est pas un répertoire
                            Set itmX = ListView1.ListItems.Add(, , strFile)
                            itmX.SubItems(4) = "File"  ' Type fichier
                        End If
                        itmX.Text = strFile
                        itmX.SubItems(1) = clsFind.FileSize
                        Select Case Dateformat
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
                        If Trim$(strFile) = "." Or Trim$(strFile) = ".." Then
                            afficher = False
                        End If
                        If (attributs And FILE_ATTRIBUTE_READONLY) And ReadOnly = False Then
                            afficher = False
                        End If
                        If (attributs And FILE_ATTRIBUTE_HIDDEN) And Hidden = False Then
                            afficher = False
                        End If
                        If (attributs And FILE_ATTRIBUTE_SYSTEM) And System = False Then
                            afficher = False
                        End If
                        If afficher = True Then
                            If attributs And FILE_ATTRIBUTE_READONLY Then
                                chaine = "R"
                            End If
                            If attributs And FILE_ATTRIBUTE_HIDDEN Then
                                chaine = chaine + "H"
                            End If
                            If attributs And FILE_ATTRIBUTE_SYSTEM Then
                                chaine = chaine + "S"
                            End If
                            If attributs And FILE_ATTRIBUTE_ARCHIVE Then
                                chaine = chaine + "A"
                            End If
                            If chaine = "" Then
                                chaine = " "
                            End If
                            nbfichiers = nbfichiers + 1
                            Set itmX = ListView1.ListItems.Add(, , strFile)
                            itmX.Bold = True ' *****************************************************************************
                            itmX.Text = strFile
                            itmX.SubItems(1) = clsFind.FileSize
                            Select Case Dateformat
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
 Next i
RENAME.MousePointer = 0
If AutoArrange = True Then
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
If SelectAllFiles = 1 Then
    SelectAll
End If
remplissage = nbfichiers
Exit Function

ErrGen:
'MsgBox "Ligne de l'erreur=" & Err.Line
ErreurGrave "Remplissage"
End Function

Private Sub ModifierLigne()
 Dim vrai As Boolean
 vrai = False
 If Trim$(Combo1.List(Combo1.ListIndex)) <> "Modify prefix" Or Trim$(Combo2.List(Combo2.ListIndex)) <> "Modify extension" Then
  état.Panels(1).
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
 If vrai = False Then
  chnmodifier = chnmodifier + "<nothing>"
 End If
 vrai = False
 
 ' 2-Extension
 If right$(chnmodifier, 1) = "/" Then
  chnmodifier = left$(chnmodifier, Len(chnmodifier) - 1)
 End If
 
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
 If vrai = False Then
  chnmodifier = chnmodifier + "<nothing>"
 End If
 
 If right$(chnmodifier, 1) = "/" Then
  chnmodifier = left$(chnmodifier, Len(chnmodifier) - 1)
  vrai = True
 End If
 état.Panels(1).Text = chnmodifier
End Sub

Private Sub renameweb_Click()
BrowseTo ("http://www.herve-thouzard.com/therename.phtml")
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
Private Sub Text16_GotFocus()
SelAll Text16
End Sub
Private Sub Text17_GotFocus()
SelAll Text17
End Sub
Private Sub Text18_GotFocus()
SelAll Text18
End Sub

Private Sub Text2_Change()
 CharInterdits Text2.Text
 ModifierLigne
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

Private Sub Text8_Change()
 CharInterdits Text8.Text
 ModifierLigne
End Sub

Private Sub Text8_GotFocus()
SelAll Text8
End Sub
Private Sub Text9_GotFocus()
SelAll Text9
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
    recursive = False
    m2recursive.Checked = False
   Else
    recursive = True
    m2recursive.Checked = True
   End If
   vnb = remplissage()
   état.Panels(3).Text = Trim$(Str$(vnb))
   état.Panels(4).
   
  Case 15 ' Up
   MoveUp
  Case 16 ' Root
   MoveRoot
  Case 18
   If VnbHistory > 0 Then
    PopupMenu mgenhistory, , Button.left, Toolbar1.top + Toolbar1.height
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
 fileop.ParentWnd = hWnd
 fileop.ConfirmOperation = ConfirmOperation
 fileop.RenameOnCollision = RenameOnCollision
 fileop.SilentMode = SilentMode
 fileop.AllowUndo = AllowUndo
 fileop.ConfirmMakeDir = ConfirmMakeDir
  
 If VnbRep > 0 Then
  If Data.GetFormat(vbCFFiles) Then ' Ce sont des fichiers qui proviennent de l'explorateur ou d'une autre fenêtre que celle de THE Rename
   For i = 1 To Data.Files.Count
     For j = 1 To VnbRep
      fileop.ClearSourceFiles
      fileop.ClearDestFiles
      fileop.AddSourceFile Data.Files(i)
      fileop.AddDestFile LesRepertoires(j)
      état.Panels(1). + Data.Files(i) + " to " + LesRepertoires(j)
      vtempo = LesRepertoires(j)
      If right$(vtempo, 1) <> "\" Then
       vtempo = vtempo + "\"
      End If
      DoEvents
      If fileop.CopyFiles Then
       DT3.SetFileDateTime (vtempo + Prefixe(Data.Files(i)) & "." & Suffixe(Data.Files(i)))
       Attr3.ChangeAttr (vtempo + Prefixe(Data.Files(i)) & "." & Suffixe(Data.Files(i)))
      End If
     Next j
   Next i
  Else
   chemin = Dir1Path
   If right$(chemin, 1) <> "\" Then
    chemin = chemin + "\"
   End If
   If recursive = True Then
    chemin = ""
   End If
   If Screen.ActiveControl.Name = "ListView1" Then
    i = LVGetItemSelected(ListView1, -1)
    While i <> -1
     sItem = LVGetName(ListView1, i)
     For j = 1 To VnbRep
      fileop.ClearSourceFiles
      fileop.ClearDestFiles
      fileop.AddSourceFile chemin & sItem
      fileop.AddDestFile LesRepertoires(j)
      état.Panels(1). + sItem + " to " + LesRepertoires(j)
      DoEvents
      vtempo = LesRepertoires(j)
      If right$(vtempo, 1) <> "\" Then
       vtempo = vtempo + "\"
      End If
      If fileop.CopyFiles Then
       DT3.SetFileDateTime (vtempo & sItem)
       Attr3.ChangeAttr (vtempo & sItem)
      End If
     Next j
     i = LVGetItemSelected(ListView1, i)
    Wend
   End If ' Si on est sur le listview
   état.Panels(1).
  End If
 End If
End Sub

Private Sub txtlang_Click()
    If Combo6.Visible = True Then
        Combo6.Visible = False
    End If
End Sub

Private Sub txtlang_GotFocus()
    SelAll txtlang
End Sub

Private Sub txtlang_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim tControle As Integer
    Dim Phrase As String ' Contient la commande en cours (complète)
    Dim PosDeb As Integer
    Dim PosFin As Integer
    Dim PosSauve As Integer
    Dim extrait As String
    Dim longueur As Integer
    Dim i As Integer
    Dim sValue As String
    Dim chemin As String
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
                    .path = chemin
                    .Section = "Commands"
                    . & Trim$(Str$(CurrentCommand))
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
                    .path = chemin
                    .Section = "Commands"
                    . & Trim$(Str$(CurrentCommand))
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
                    .path = chemin
                    .Section = "Commands"
                    . & Trim$(Str$(CurrentCommand))
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
                    .path = chemin
                    .Section = "Commands"
                    . & Trim$(Str$(CurrentCommand))
                    sValue = .Value
                End With
                If Trim$(sValue) <> "" Then
                    txtlang.Text = sValue
                End If
            Case 32 ' Barre d'espace
                txtlang.Text = left$(txtlang.Text, txtlang.SelStart - 1) + Mid$(txtlang.Text, txtlang.SelStart)
                Phrase = txtlang.Text
                ' 1) Il faut déterminer sur quel mot on est positionné
                PosDeb = txtlang.SelStart
                PosFin = InStrRev(left$(Phrase, txtlang.SelStart), "<")
                If PosFin = 0 Then
                    Exit Sub
                End If
                Combo6.Clear
                ' on extrait ce qui a été tapé
                extrait = RTrim$(UCase$(Mid$(Phrase, PosFin, txtlang.SelStart - PosFin + 1)))
                longueur = Len(extrait)
                For i = 0 To vnbcmd - 1 ' Boucle sur les commandes du langage
                    If UCase$(left$(listcmd.List(i), longueur)) = extrait Then    ' La commande ressemble à ce qui a été tapé
                        Combo6.AddItem listcmd.List(i)  ' on l'ajoute dans le combo puisqu'il correspond
                    End If
                Next
                If Combo6.ListCount = 0 Then ' La recherche n'a rien donnée
                    Exit Sub    ' On peut s'en aller
                End If
                PosSauve = txtlang.SelStart
                LongSauve = longueur
                If Combo6.ListCount > 1 Then ' Il y a eu plusieurs occurences trouvées, il va falloir l'afficher
                    Combo6.top = txtlang.top + txtlang.height + 10
                    LaPosSauve = txtlang.SelStart
                    ARemplacer = extrait
                    Combo6.Visible = True
                    Combo6.ListIndex = 0
                    Combo6.SetFocus
                Else ' Il n'y a eu qu'une seule occurence de trouvée, on la met directement dans le texte
                    txtlang.Text = left$(Phrase, txtlang.SelStart - 1) + Trim$(Mid$(Combo6.List(0), longueur + 1)) + Trim$(Mid$(Phrase, txtlang.SelStart))
                    txtlang.SelStart = PosSauve + Len(Trim$(Mid$(Combo6.List(0), longueur + 1))) - 1
                    Combo6.Clear ' et à la fin on efface son contenu (par propreté)
                End If
                ' Combo6.Visible = False
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
Dim intDirsFound As Integer, vntItem As Variant
Dim szFilename As String, i As Integer, vnb As Integer, extrait As String

'rafraichir = False
recursive = True
m2recursive.Checked = True
ListView1.ListItems.Clear
nbfichiers = 0
szFilename = Trim$(Dir1Path)
If right$(szFilename, 1) <> "\" Then
 szFilename = szFilename + "\"
End If
Filtre = Trim$(Filtre)
If right$(Filtre, 1) = ";" Then
 Filtre = left$(Filtre, Len(Filtre) - 1)
End If
If left$(Filtre, 1) = ";" Then
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
   Next vntItem
  End If
 Next i

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
  Dim r As Long, lpBuffer  As String * 256, longueur As Long, i As Integer
  Dim n&
  On Error GoTo ErrGen
  i = 0
  lpBuffer = Space$(256)
  longueur = Len(lpBuffer)
  r = GetLogicalDriveStrings(longueur, lpBuffer)
  If r Then
    Do
     n = InStr(lpBuffer, Chr$(0))
     If n > 1 Then
      If i <> 0 Then Load mnudrives(i)
      mnudrives(i).Caption = left$(lpBuffer, n - 1)
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
 Dim sItem As String
 Dim actions As String
 Dim i As Long
 Dim vnb2 As Long
 Dim chemin As String
 Dim qfaire As String
 Dim fileop As New CSHFileOp
 Dim chemin2 As String
 Dim chemin3 As String
 fileop.ParentWnd = hWnd
 fileop.ConfirmOperation = ConfirmOperation
 fileop.RenameOnCollision = RenameOnCollision
 fileop.SilentMode = SilentMode
 fileop.AllowUndo = AllowUndo
 fileop.ConfirmMakeDir = ConfirmMakeDir
 If Piège1 = True Then ' Le truc pour éviter le piège des Ctrl C, Ctrl X, Ctrl V (et autres) quand on est en édition de nom de fichier (à la main) sur le listview principal
    GoTo suite
 End If
 RENAME.MousePointer = 11
 chemin = Dir1Path
 If right$(chemin, 1) <> "\" Then
  chemin = chemin + "\"
 End If
 chemin2 = chemin
 If recursive = True Then
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
    itmX.Text = actions
    itmX.SubItems(1) = chemin + LVGetName(ListView1, i)
    itmX.SubItems(2) = LVGetItemName(ListView1, i, 1)
    itmX.SubItems(3) = LVGetItemName(ListView1, i, 2)
    itmX.SubItems(4) = LVGetItemName(ListView1, i, 3)
    i = LVGetItemSelected(ListView1, i)
   Wend
  Else ' coller
   vnb2 = ListView3.ListItems.Count - 1
   For i = 0 To vnb2
    qfaire = Trim$(LVGetName(ListView3, i))
    fileop.ClearSourceFiles
    fileop.ClearDestFiles
    fileop.AddSourceFile LVGetItemName(ListView3, i, 1)
    fileop.AddDestFile chemin2
    état.Panels(2). + Trim$(Str$(ListView3.ListItems.Count))
    If qfaire = "Cut to bin" Or qfaire = "Cut additive" Then
     état.Panels(1). + LVGetItemName(ListView3, i, 1) + " to " + chemin2
     If fileop.MoveFiles Then
      chemin3 = chemin2
      If right$(chemin3, 1) <> "\" Then
       chemin3 = chemin3 + "\"
      End If
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView3, i, 1)) + "." + Suffixe(LVGetItemName(ListView3, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
     End If
    Else
     état.Panels(1). + LVGetItemName(ListView3, i, 1) + " to " + chemin2
     If fileop.CopyFiles Then
      chemin3 = chemin2
      If right$(chemin3, 1) <> "\" Then
       chemin3 = chemin3 + "\"
      End If
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView3, i, 1)) + "." + Suffixe(LVGetItemName(ListView3, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
     End If
    End If
   Next i
   If action = 5 Then
    ListView3.ListItems.Clear
   End If
   vnb = remplissage()
   état.Panels(1).
   état.Panels(3).Text = Trim$(Str$(vnb))
   état.Panels(4).
  End If
 
 Case "FolderTreeview1"
  If action > 0 And action < 5 Then ' Copy ou Cut peut importe que se soit additive ou pas, le ménage est déjà fait
    Set itmX = ListView3.ListItems.Add(, , actions)
    itmX.Text = actions
    itmX.SubItems(1) = Dir1Path
  Else ' Coller
   For i = 0 To ListView3.ListItems.Count - 1
    qfaire = Trim$(LVGetName(ListView3, i))
    fileop.ClearSourceFiles
    fileop.ClearDestFiles
    fileop.AddSourceFile LVGetItemName(ListView3, i, 1)
    fileop.AddDestFile chemin2
    état.Panels(2). + Trim$(Str$(ListView3.ListItems.Count))
    If qfaire = "Cut to bin" Or qfaire = "Cut additive" Then
     état.Panels(1). + LVGetItemName(ListView3, i, 1) + " to " + chemin2
     If fileop.MoveFiles Then
      chemin3 = chemin2
      If right$(chemin3, 1) <> "\" Then
       chemin3 = chemin3 + "\"
      End If
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView3, i, 1)) + "." + Suffixe(LVGetItemName(ListView3, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
     End If
    Else
     état.Panels(1). + LVGetItemName(ListView3, i, 1) + " to " + chemin2
     If fileop.CopyFiles Then
      chemin3 = chemin2
      If right$(chemin3, 1) <> "\" Then
       chemin3 = chemin3 + "\"
      End If
      chemin3 = chemin3 + Prefixe(LVGetItemName(ListView3, i, 1)) + "." + Suffixe(LVGetItemName(ListView3, i, 1))
      DT2.SetFileDateTime (chemin3)
      Attr2.ChangeAttr (chemin3)
     End If
    End If
   Next i
   If action = 5 Then
    ListView3.ListItems.Clear
   End If
   vnb = remplissage()
   état.Panels(1).
   état.Panels(3).Text = Trim$(Str$(vnb))
   état.Panels(4).
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
    Next iFile
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
  Next i
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
 .path = chemin
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
    .path = chemin
    .Section = LSection
    .Key = cle
    sValue = .Value
End With
sValue = Replace(sValue, Chr$(255), " ")
LoadSet = sValue
End Function

Private Sub RefreshF5()
  FolderTreeview1(0).Refresh
  vnb = remplissage()
  état.Panels(3).Text = Trim$(Str$(vnb))
  état.Panels(4).
End Sub

Private Sub ChargeVNBCommandes()
    Dim sValue As String
    Dim chemin As String
    chemin = AppPath
    chemin = chemin + "commands.ini"
    sValue = ""
    With SIni
     .path = chemin
     .Section = "General"
     .
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
            Next i
        End If
        yourcmd(0).
        txtlang.Enabled = False
        sValue = ""
        With SIni
            .path = chemin
            .Section = "General"
            .
            sValue = .Value
        End With
        vnbcmd = Val(sValue)
        For i = 0 To vnbcmd - 1
         With SIni
            .path = chemin
            .Section = "Commands"
            . & Trim$(Str$(i + 1))
            sValue = .Value
            End With
            If i <> 0 Then
                Load yourcmd(i)
            End If
            yourcmd(i).Caption = sValue
        Next
End Sub
Function GetToken1(sTarget As String, sSeps As String) As String
' Note that sSave and iStart must be static from call to call
' If first call, make copy of string
Static sSave As String, iStart As Integer
If sTarget <> "" Then
    iStart = 1
    sSave = sTarget
End If
' Find start of next token
Dim iNew As Integer
iNew = StrSpan1(Mid$(sSave, iStart, Len(sSave)), sSeps)
If iNew Then
    ' Set position to start of token
    iStart = iNew + iStart - 1
Else
    ' If no new token, return empty string
    GetToken1 = ""
    Exit Function
End If

' Find end of token
iNew = StrBreak1(Mid$(sSave, iStart, Len(sSave)), sSeps)
If iNew Then
    ' Set position to end of token
    iNew = iStart + iNew - 1
Else
    ' If no end of token, set to end of string
    iNew = Len(sSave) + 1
End If
' Cut token out of sTarget string
GetToken1 = Mid$(sSave, iStart, iNew - iStart)
' Set new starting position
iStart = iNew
End Function
Function StrSpan1(sTarget As String, sSeps As String) As Integer
    Dim cTarget As Integer, iStart As Integer
    cTarget = Len(sTarget)
    iStart = 1
    ' Look for start of token (character that isn’t a separator)
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1))
        If iStart > cTarget Then
            StrSpan1 = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrSpan1 = iStart
End Function
Function StrBreak1(sTarget As String, sSeps As String) As Integer
    Dim cTarget As Integer, iStart As Integer
    cTarget = Len(sTarget)
    iStart = 1
    'Look for end of token (first character that is a separator)
    Do While InStr(sSeps, Mid$(sTarget, iStart, 1)) = 0
        If iStart > cTarget Then
            StrBreak1 = 0
            Exit Function
        Else
            iStart = iStart + 1
        End If
    Loop
    StrBreak1 = iStart
End Function

Private Sub CreateTokenTabl(FileName As String, chemin As String)
' Remplit les tableaux des tokens avec les tokens du préfixe ET du suffixe
    Dim zPrefixe As String
    Dim zSuffixe As String
    Dim zchemin As String
    Dim sToken As String
    Dim i As Integer
    VnbTokensFo = 0
    If FileName = "" Then
        VnbTokensPr = 0
        VnbTokensEx = 0
    End If
    zchemin = Replace(chemin, ":", "")
    zPrefixe = Prefixe(FileName)
    zSuffixe = Suffixe(FileName)
    ' 1) réinitialisation des tableaux
    For i = 1 To 100
        TablTokensPr(i) = ""    ' Tokens du préfix
        TablTokensEx(i) = ""    ' Tokens de l'extension
        TablTokensFo(i) = ""    ' Tokens du répertoire
    Next
    ' 2) Les Tokens du prefix
    i = 0
    sToken = GetToken1(zPrefixe, CharTokens)
    Do While sToken <> ""
        i = i + 1
        If i > 100 Then
            Exit Do
        End If
        TablTokensPr(i) = sToken
        sToken = GetToken1("", CharTokens)
    Loop
    VnbTokensPr = i
    ' 3) Les Tokens de l'extension
    i = 0
    sToken = GetToken1(zSuffixe, CharTokens)
    Do While sToken <> ""
        i = i + 1
        If i > 100 Then
            Exit Do
        End If
        TablTokensEx(i) = sToken
        sToken = GetToken1("", CharTokens)
    Loop
    VnbTokensEx = i
    ' 4) Les Tokens du répertoire
    i = 0
    sToken = GetToken1(zchemin, "\")
    Do While sToken <> ""
        i = i + 1
        If i > 100 Then
            Exit Do
        End If
        TablTokensFo(i) = sToken
        sToken = GetToken1("", "\")
    Loop
    VnbTokensFo = i
End Sub

'Private Sub VQFInfo()
'Dim i As Long, chemin As String, vnb As Integer, vnb2 As Long
'Dim vnbtot As Integer
'Dim vfichier As String, sItem As String
'Dim Vin As Integer
'Dim Buffer As String
'Dim TailleEntete As Integer
'Dim position As Integer
'Dim tmp As String
'Dim tmp2 As Integer
'chemin = Trim$(Dir1Path)
'If right$(chemin, 1) <> "\" Then
' chemin = chemin + "\"
'End If
'If recursive = True Then
' chemin = ""
'End If
'
'vnb = 0
'vnbtot = 0
'vnb = LVGetCountSelected(ListView1)
'vnbtot = vnb
'
'If vnb = 0 Then
' MsgBox "You must select files before renaming them !"
' Exit Sub
'End If
'
'i = LVGetItemSelected(ListView1, -1)
'While i <> -1
' sItem = LVGetName(ListView1, i)
' Vin = FreeFile
' Open sItem For Binary As #Vin
' Buffer = Space$(4)
' Get #Vin, 1, Buffer
' If UCase$(Buffer) <> "TWIN" Then
'    Exit Sub
' End If
' Buffer = Space$(1)
' Get #Vin, 16, Buffer
' TailleEntete = Asc(Buffer)
' If TailleEntete = 0 Then
'    Exit Sub
' End If
' ' Rendu là, on connait la longueur de l'entête et on sait où le lire
' Buffer = Space$(TailleEntete)
' Get #Vin, 17, Buffer
' Close #Vin
' ' Il ne reste "plus" qu'à analyzer le contenu de cet entête.
' position = 1
' While position <= TailleEntete
'    tmp = UCase$(Mid$(Buffer, position, 4))
'    Select Case tmp
'        Case "COMM" ' Informations sur le fichier
'            MsgBox "Qualité sonore=" & Asc(Mid$(Buffer, position + 7, 1))
'            MsgBox "Mono ou stéréo=" & Asc(Mid$(Buffer, position + 11, 1))
'            MsgBox "Bitrate=" & Asc(Mid$(Buffer, position + 15, 1))
'            MsgBox "Sample rate=" & Asc(Mid$(Buffer, position + 19, 1))
'            ' Il ne reste plus qu'à sauter ce TAG
'            position = position + 24
'        Case "NAME" ' Titre de la chanson
'            ' Il faut extraire la longueur de la chaine contenant le titre
'            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
'' merde !
'            If tmp2 <> 0 Then
'                MsgBox "Titre de la chanson=" & Mid$(Buffer, position + 8, tmp2)
'            End If
'            ' Il ne reste plus qu'à sauter ce TAG
'            position = position + 8 + tmp2
'        Case "COMT" ' Commentaire
'            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
'            If tmp2 <> 0 Then
'                MsgBox "Commentaire=" & Mid$(Buffer, position + 8, tmp2)
'            Else
'                MsgBox "Zone commentaire vide"
'            End If
'            ' Il ne reste plus qu'à sauter ce TAG
'            position = position + 8 + tmp2
'        Case "AUTH" ' Nom de l'auteur
'            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
'            If tmp2 <> 0 Then
'                MsgBox "Auteur=" & Mid$(Buffer, position + 8, tmp2)
'            Else
'                MsgBox "Zone de l'auteur vide"
'            End If
'            ' Il ne reste plus qu'à sauter ce TAG
'            position = position + 8 + tmp2
'        Case "(C) " ' copyright
'            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
'            If tmp2 <> 0 Then
'                MsgBox "Copyright=" & Mid$(Buffer, position + 8, tmp2)
'            Else
'                MsgBox "Zone du Copyright vide"
'            End If
'            ' Il ne reste plus qu'à sauter ce TAG
'            position = position + 8 + tmp2
'        Case "FILE" ' Nom par défaut à donner au fichier lors de sa sauvegarde
'            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
'            If tmp2 <> 0 Then
'                MsgBox "File=" & Mid$(Buffer, position + 8, tmp2)
'            Else
'                MsgBox "Zone du File vide"
'            End If
'            ' Il ne reste plus qu'à sauter ce TAG
'            position = position + 8 + tmp2
'        Case "DSIZ" ' Inutilisé dans le programme
'            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
'            ' On saute directement ce TAG puisqu'il n'est pas utilisé
'            position = position + 8 + tmp2
'        Case Else
'            Exit Sub ' On n'a pas du bien interpreter les données
'    End Select
' Wend
'
' i = LVGetItemSelected(ListView1, i)
'Wend
'End Sub

Private Sub MoveCopyFile(i As Long, vnbit As Long, vnbren As Long)
Dim vtmp As String, vtmp2 As String
Dim fileop As New CSHFileOp
Dim chemin As String
Dim ChemDest As String
Dim NomFichier As String
Dim j As Integer
On Error Resume Next

If recursive = False Then
    chemin = Dir1Path
    If right$(chemin, 1) <> "\" Then
        chemin = chemin + "\"
    End If
End If
fileop.ClearSourceFiles
fileop.ClearDestFiles
NomFichier = LVGetName(ListView1, i)
Select Case Misc3
    Case 0  ' Prefix
        vtmp = Prefixe(NomFichier)
    Case 1  ' Extension
        vtmp = Suffixe(NomFichier)
End Select

If Misc6 = 1 Then    ' Stop to numemric character
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
    
If Misc7 = 1 Then ' Replace _ with space
    vtmp = Replace(vtmp, "_", " ")
    vtmp = Trim$(vtmp)
End If
    
If Misc8 = 1 Then ' Capitalize all words
    vtmp = MyStrConv(vtmp)
End If

' Création du répertoire
If recursive = True Then
    chemin = ExtractPath(NomFichier)
    If right$(chemin, 1) <> "\" Then
        chemin = chemin + "\"
    End If
End If
ChemDest = chemin + vtmp
MkDir (ChemDest)
Select Case Misc4
    Case 0  ' Just create
    Case 1  ' Copy
        If recursive = False Then
            fileop.AddSourceFile chemin + NomFichier
        Else
            fileop.AddSourceFile NomFichier
        End If
        fileop.AddDestFile ChemDest + "\" + NomFichier
        état.Panels(1). + chemin + NomFichier + " to " + ChemDest + "\" + NomFichier
        état.Panels(2). + Trim$(Str$(vnbit))
        fileop.CopyFiles
    Case 2  ' Move
        If recursive = False Then
            fileop.AddSourceFile chemin + NomFichier
        Else
            fileop.AddSourceFile NomFichier
        End If
        fileop.AddDestFile ChemDest + "\" + NomFichier
        état.Panels(1). + chemin + NomFichier + " to " + ChemDest + "\" + NomFichier
        état.Panels(2). + Trim$(Str$(vnbit))
        fileop.MoveFiles
End Select
End Sub
' ********************************************************************************************
' Déplace les fichers sélectionnés vers le haut
' ********************************************************************************************
Private Sub MoveFilesUp()
Dim vnb As Long
Dim i As Long
Dim CollFiles As New Collection
Dim vret As Boolean
On Error Resume Next
ListView1.Sorted = False
i = 0
i = LVGetItemSelected(ListView1, -1)
While i <> -1
    vnb = vnb + 1
    CollFiles.Add i + 1, Str$(i + 1)
    LVSetItemNotSelected ListView1, i
    i = LVGetItemSelected(ListView1, i)
Wend
If vnb = 0 Then
    Exit Sub
End If
RENAME.MousePointer = vbHourglass
ListView1.Visible = False
For i = 1 To vnb
    vret = MoveRow(ListView1, CollFiles.Item(i), CollFiles.Item(i) - 1)
Next
For i = 1 To vnb
    LVSetItemSelected ListView1, CollFiles.Item(i) - 2
Next
ListView1.Visible = True
RENAME.MousePointer = vbDefault
ListView1.SetFocus
End Sub

' ********************************************************************************************
' Déplace les fichers sélectionnés vers le bas
' ********************************************************************************************
Private Sub MoveFilesDown()
Dim vnb As Long
Dim i As Long
Dim CollFiles As New Collection
Dim vret As Boolean
On Error Resume Next
ListView1.Sorted = False
i = 0
i = LVGetItemSelected(ListView1, -1)
While i <> -1
    vnb = vnb + 1
    CollFiles.Add i + 1, Str$(i + 1)
    LVSetItemNotSelected ListView1, i
    i = LVGetItemSelected(ListView1, i)
Wend
If vnb = 0 Then
    Exit Sub
End If
RENAME.MousePointer = vbHourglass
ListView1.Visible = False
For i = vnb To 1 Step -1
    vret = MoveRow(ListView1, CollFiles.Item(i), CollFiles.Item(i) + 1)
Next
For i = 1 To vnb
    LVSetItemSelected ListView1, CollFiles.Item(i)
Next
ListView1.Visible = True
RENAME.MousePointer = vbDefault
ListView1.SetFocus
End Sub

Private Sub SRegRename(i As Long)
Dim vtmp As String
Dim fileop As New CSHFileOp
Dim chemin As String
Dim NomFichier As String
Dim ReturnString As String
On Error Resume Next

vtmp = LVGetName(ListView1, i)
If recursive = False Then
    chemin = Dir1Path
Else
    chemin = ExtractPath(vtmp)
End If

If right$(chemin, 1) <> "\" Then
    chemin = chemin + "\"
End If

NomFichier = Prefixe(vtmp) & "." & Suffixe(vtmp)
If RegSub(NomFichier, LChaine1 + Chr$(0), LChaine2 + Chr$(0), ReturnString, LOption2, LOption3, 0, 0) Then
    fileop.ParentWnd = hWnd
    fileop.ClearSourceFiles
    fileop.ClearDestFiles
    fileop.ConfirmOperation = False
    fileop.AddSourceFile chemin + NomFichier
    fileop.AddDestFile chemin + ReturnString
    If CopyRename = True Then
        fileop.RenameFiles
    Else
        fileop.CopyFiles
    End If
    If UseHistory = True Then
        lhistory.AddItem Trim$(Str$(Time())) + "|" + chemin + "|" + NomFichier + "|" + ReturnString ' Historique
    End If
    List2.AddItem chemin + NomFichier       ' Nom d'origine.
    List3.AddItem chemin + ReturnString     ' Nom d'arrivée.
End If
End Sub

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
    .path = chemin
    .Section = "MP3Tags"
    .
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "MP3Tags"
        . & Trim$(Str$(i))
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
Dim chemin As String
Dim vnbcmd As Integer
Dim i As Integer
Dim vnb As Integer
Dim sValue As String
lachaine = ""
chemin = AppPath + "Music.ini"
With SIni
    .Section = "VQFTags"
    .
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "VQFTags"
        . & Trim$(Str$(i))
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
    .path = chemin
    .Section = "MP3Tags"
    .
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "MP3Tags"
        . & Trim$(Str$(i))
        sValue = .Value
    End With
    If Val(GetToken(sValue, "|", 1)) = 1 Then
        vnb = vnb + 1
        If vnb = lindex Then
            MP3, 2)
            Exit Function
        End If
    End If
Next
MP3
End Function

Private Function VQFCaption(lindex As Integer) As String
Dim chemin As String
Dim vnbcmd As Integer
Dim i As Integer
Dim vnb As Integer
Dim sValue As String
chemin = AppPath + "Music.ini"
With SIni
    .path = chemin
    .Section = "VQFTags"
    .
    sValue = .Value
End With
vnbcmd = Val(sValue)
For i = 1 To vnbcmd
    With SIni
        .Section = "VQFTags"
        . & Trim$(Str$(i))
        sValue = .Value
    End With
    If Val(GetToken(sValue, "|", 1)) = 1 Then
        vnb = vnb + 1
        If vnb = lindex Then
            VQF, 2)
            Exit Function
        End If
    End If
Next
VQF
End Function

Private Function Blanc(lachaine As String) As String
If Trim$(lachaine) = "" Then
    Blanc = "&nbsp"
Else
    Blanc = lachaine
End If
End Function

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

Private Sub LoadLvMP3()
Dim vNomFic As String
Dim SonInfo As String
Dim colonne As ColumnHeader
vNomFic = Dir1Path
If right$(vNomFic, 1) <> "\" Then
    vNomFic = vNomFic + "\"
End If

vNomFic = vNomFic + ListView1.SelectedItem.Text
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

Set colonne = ListView1.ColumnHeaders.Item(1)
AutoSizeColumnHeader LvMP3, colonne, True
Set colonne = ListView1.ColumnHeaders.Item(2)
AutoSizeColumnHeader LvMP3, colonne, True

End Sub

Private Sub AddOneTag(Donnee As String, texte As String)
Dim itmX As ListItem
Set itmX = LvMP3.ListItems.Add(, , texte)
itmX.SubItems(1) = Donnee
End Sub

Private Sub MTetat1()
    If mviewmp3tab.Checked = True Then
        mviewmp3tab.Checked = False
    Else
        mviewmp3tab.Checked = True
    End If
    ShowMP3Tab = mviewmp3tab.Checked
    TabGen.TabVisible(2) = mviewmp3tab.Checked
End Sub

Private Sub MTetat2()
    If mviewpicturetab.Checked = True Then
        mviewpicturetab.Checked = False
    Else
        mviewpicturetab.Checked = True
    End If
    ShowMusicTab = mviewpicturetab.Checked
    TabGen.TabVisible(3) = mviewpicturetab.Checked
End Sub
