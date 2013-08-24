VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Begin VB.Form FMP3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sounds information"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5715
   ControlBox      =   0   'False
   HelpContextID   =   77
   Icon            =   "FMP3.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   1725
      TabIndex        =   11
      ToolTipText     =   "Don't use sounds information"
      Top             =   5430
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      Left            =   2895
      TabIndex        =   12
      ToolTipText     =   "Use sounds information with these parameters"
      Top             =   5430
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   42
      ToolTipText     =   "Use F9 and Shift+F9 to change the current tab"
      Top             =   60
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "MP3"
      TabPicture(0)   =   "FMP3.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Picture5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Combo2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "List1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdclear"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text6"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text8"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "VQF"
      TabPicture(1)   =   "FMP3.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text3"
      Tab(1).Control(1)=   "Text4"
      Tab(1).Control(2)=   "Text9"
      Tab(1).Control(3)=   "Picture1"
      Tab(1).Control(4)=   "List2"
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(8)=   "Label7"
      Tab(1).Control(9)=   "Label6"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Ogg"
      TabPicture(2)   =   "FMP3.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture2"
      Tab(2).Control(1)=   "Combo3"
      Tab(2).Control(2)=   "List3"
      Tab(2).Control(3)=   "Command2"
      Tab(2).Control(4)=   "Text11"
      Tab(2).Control(5)=   "Text10"
      Tab(2).Control(6)=   "Text7"
      Tab(2).Control(7)=   "Combo1"
      Tab(2).Control(8)=   "Label18"
      Tab(2).Control(9)=   "Label16"
      Tab(2).Control(10)=   "Label15"
      Tab(2).Control(11)=   "Label14"
      Tab(2).Control(12)=   "Label13"
      Tab(2).Control(13)=   "Label10"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "WMA"
      TabPicture(3)   =   "FMP3.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label17"
      Tab(3).Control(1)=   "Label19"
      Tab(3).Control(2)=   "Label20"
      Tab(3).Control(3)=   "Label21"
      Tab(3).Control(4)=   "Label22"
      Tab(3).Control(5)=   "Label23"
      Tab(3).Control(6)=   "Label24"
      Tab(3).Control(7)=   "Combo4"
      Tab(3).Control(8)=   "Text12"
      Tab(3).Control(9)=   "Text13"
      Tab(3).Control(10)=   "Text14"
      Tab(3).Control(11)=   "Command3"
      Tab(3).Control(12)=   "List4"
      Tab(3).Control(13)=   "Text15"
      Tab(3).Control(14)=   "Combo5"
      Tab(3).Control(15)=   "Picture3"
      Tab(3).ControlCount=   16
      TabCaption(4)   =   "WAV"
      TabPicture(4)   =   "FMP3.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -74760
         ScaleHeight     =   240
         ScaleWidth      =   4950
         TabIndex        =   63
         Top             =   1260
         Width           =   4950
         Begin VB.OptionButton Option4 
            Caption         =   "Add to the right"
            Height          =   255
            Index           =   1
            Left            =   1650
            TabIndex        =   34
            ToolTipText     =   "Add information to the right"
            Top             =   0
            Width           =   1440
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Add to the left"
            Height          =   255
            Index           =   0
            Left            =   15
            TabIndex        =   33
            ToolTipText     =   "Add information to the left"
            Top             =   0
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Replace prefix"
            Height          =   255
            Index           =   2
            Left            =   3375
            TabIndex        =   35
            ToolTipText     =   "Replace prefix"
            Top             =   0
            Width           =   1470
         End
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "FMP3.frx":0098
         Left            =   -73170
         List            =   "FMP3.frx":0258
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   3996
         Width           =   2145
      End
      Begin VB.TextBox Text15 
         Height          =   285
         Left            =   -73170
         TabIndex        =   38
         Top             =   3603
         Width           =   540
      End
      Begin VB.ListBox List4 
         Height          =   1230
         ItemData        =   "FMP3.frx":0829
         Left            =   -74790
         List            =   "FMP3.frx":084E
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   1890
         Width           =   2940
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   330
         Left            =   -70305
         TabIndex        =   32
         ToolTipText     =   "Clear command line"
         Top             =   810
         Width           =   615
      End
      Begin VB.TextBox Text14 
         Height          =   285
         Left            =   -73170
         TabIndex        =   40
         Top             =   4419
         Width           =   3015
      End
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   -73170
         TabIndex        =   41
         Top             =   4815
         Width           =   3015
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   -73170
         TabIndex        =   37
         Top             =   3240
         Width           =   3015
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   -74790
         TabIndex        =   31
         Top             =   825
         Width           =   4380
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -74760
         ScaleHeight     =   240
         ScaleWidth      =   4950
         TabIndex        =   56
         Top             =   1260
         Width           =   4950
         Begin VB.OptionButton Option1 
            Caption         =   "Add to the right"
            Height          =   255
            Index           =   1
            Left            =   1650
            TabIndex        =   24
            ToolTipText     =   "Add information to the right"
            Top             =   0
            Width           =   1440
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Add to the left"
            Height          =   255
            Index           =   0
            Left            =   15
            TabIndex        =   23
            ToolTipText     =   "Add information to the left"
            Top             =   0
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Replace prefix"
            Height          =   255
            Index           =   2
            Left            =   3375
            TabIndex        =   25
            ToolTipText     =   "Replace prefix"
            Top             =   0
            Width           =   1470
         End
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "FMP3.frx":08C7
         Left            =   -73170
         List            =   "FMP3.frx":0A87
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3603
         Width           =   2145
      End
      Begin VB.ListBox List3 
         Height          =   1230
         ItemData        =   "FMP3.frx":1058
         Left            =   -74790
         List            =   "FMP3.frx":10B6
         Sorted          =   -1  'True
         TabIndex        =   26
         Top             =   1890
         Width           =   2940
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
         Height          =   330
         Left            =   -70305
         TabIndex        =   22
         ToolTipText     =   "Clear command line"
         Top             =   810
         Width           =   615
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   -73170
         TabIndex        =   29
         Top             =   3996
         Width           =   3015
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   -73170
         TabIndex        =   30
         Top             =   4419
         Width           =   3015
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   -73170
         TabIndex        =   27
         Top             =   3240
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   -74790
         TabIndex        =   21
         Top             =   825
         Width           =   4380
      End
      Begin VB.ComboBox Text3 
         Height          =   315
         Left            =   -74790
         TabIndex        =   13
         Top             =   825
         Width           =   4380
      End
      Begin VB.ComboBox Text1 
         Height          =   315
         Left            =   210
         TabIndex        =   0
         Top             =   825
         Width           =   4380
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   -73170
         TabIndex        =   20
         Top             =   3603
         Width           =   3015
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   -73170
         TabIndex        =   19
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1830
         TabIndex        =   6
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1830
         TabIndex        =   10
         Top             =   4815
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1830
         TabIndex        =   9
         Top             =   4419
         Width           =   3015
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -74760
         ScaleHeight     =   240
         ScaleWidth      =   4950
         TabIndex        =   49
         Top             =   1260
         Width           =   4950
         Begin VB.OptionButton Option3 
            Caption         =   "Add to the right"
            Height          =   255
            Index           =   1
            Left            =   1650
            TabIndex        =   16
            ToolTipText     =   "Add information to the right"
            Top             =   0
            Width           =   1440
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Add to the left"
            Height          =   255
            Index           =   0
            Left            =   15
            TabIndex        =   15
            ToolTipText     =   "Add information to the left"
            Top             =   0
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Replace prefix"
            Height          =   255
            Index           =   2
            Left            =   3375
            TabIndex        =   17
            ToolTipText     =   "Replace prefix"
            Top             =   0
            Width           =   1470
         End
      End
      Begin VB.ListBox List2 
         Height          =   1230
         ItemData        =   "FMP3.frx":123D
         Left            =   -74790
         List            =   "FMP3.frx":125C
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   1890
         Width           =   2940
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Clear"
         Height          =   330
         Left            =   -70305
         TabIndex        =   14
         ToolTipText     =   "Clear command line"
         Top             =   810
         Width           =   615
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "Clear"
         Height          =   330
         Left            =   4695
         TabIndex        =   1
         ToolTipText     =   "Clear command line"
         Top             =   810
         Width           =   615
      End
      Begin VB.ListBox List1 
         Height          =   1230
         ItemData        =   "FMP3.frx":12CC
         Left            =   210
         List            =   "FMP3.frx":137B
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1890
         Width           =   2940
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1830
         TabIndex        =   7
         Top             =   3603
         Width           =   540
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "FMP3.frx":16BF
         Left            =   1830
         List            =   "FMP3.frx":187F
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3996
         Width           =   2145
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   240
         ScaleHeight     =   240
         ScaleWidth      =   4950
         TabIndex        =   43
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
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Default genre to use"
         Height          =   195
         Left            =   -74790
         TabIndex        =   70
         Top             =   4065
         Width           =   1440
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Default year to use"
         Height          =   195
         Left            =   -74790
         TabIndex        =   69
         Top             =   3622
         Width           =   1335
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Select commands to extract information from your WMA files"
         Height          =   195
         Left            =   -74775
         TabIndex        =   68
         Top             =   525
         Width           =   4230
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Default artist"
         Height          =   195
         Left            =   -74790
         TabIndex        =   67
         Top             =   3250
         Width           =   885
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "List of available information"
         Height          =   195
         Left            =   -74790
         TabIndex        =   66
         Top             =   1635
         Width           =   1905
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Default album to use"
         Height          =   195
         Left            =   -74790
         TabIndex        =   65
         Top             =   4455
         Width           =   1455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Default title to use"
         Height          =   195
         Left            =   -74790
         TabIndex        =   64
         Top             =   4875
         Width           =   1275
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Default genre to use"
         Height          =   195
         Left            =   -74790
         TabIndex        =   62
         Top             =   3622
         Width           =   1440
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Select commands to extract information from your Ogg file"
         Height          =   195
         Left            =   -74775
         TabIndex        =   61
         Top             =   525
         Width           =   4050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Default artist"
         Height          =   195
         Left            =   -74790
         TabIndex        =   60
         Top             =   3250
         Width           =   885
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "List of available information"
         Height          =   195
         Left            =   -74790
         TabIndex        =   59
         Top             =   1635
         Width           =   1905
      End
      Begin VB.Label Label13 
         Caption         =   "Default album to use"
         Height          =   255
         Left            =   -74790
         TabIndex        =   58
         Top             =   4065
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Default title to use"
         Height          =   195
         Left            =   -74790
         TabIndex        =   57
         Top             =   4455
         Width           =   1275
      End
      Begin VB.Label Label9 
         Caption         =   "Default title to use"
         Height          =   255
         Left            =   -74790
         TabIndex        =   55
         Top             =   3622
         Width           =   1575
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Default title to use"
         Height          =   195
         Left            =   210
         TabIndex        =   54
         Top             =   4875
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Default album to use"
         Height          =   195
         Left            =   210
         TabIndex        =   53
         Top             =   4455
         Width           =   1455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Select commands to extract information from your VQF file"
         Height          =   195
         Left            =   -74775
         TabIndex        =   52
         Top             =   525
         Width           =   4065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Default artist"
         Height          =   195
         Left            =   -74790
         TabIndex        =   51
         Top             =   3250
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "List of available information"
         Height          =   195
         Left            =   -74790
         TabIndex        =   50
         Top             =   1635
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "List of available information"
         Height          =   195
         Left            =   210
         TabIndex        =   48
         Top             =   1635
         Width           =   1905
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Default artist"
         Height          =   195
         Left            =   210
         TabIndex        =   47
         Top             =   3250
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select commands to extract information from your MP3 file"
         Height          =   195
         Left            =   225
         TabIndex        =   46
         Top             =   525
         Width           =   4080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Default year to use"
         Height          =   195
         Left            =   210
         TabIndex        =   45
         Top             =   3622
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Default genre to use"
         Height          =   195
         Left            =   210
         TabIndex        =   44
         Top             =   4065
         Width           =   1440
      End
   End
End
Attribute VB_Name = "FMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index1 As Integer ' Pour les MP3
Dim Index2 As Integer ' Pour les VQF
Dim index3 As Integer ' Pour les ogg
Dim index4 As Integer ' Pour les WMA
Dim cHist11 As New cHistory
Dim cHist12 As New cHistory
Dim cHist16 As New cHistory
Dim cHist17 As New cHistory

Private Sub cmdCancel_Click()
    UseMP3 = False
    UseVQF = False
    UseOGG = False
    UseWMA = False
    Unload Me
End Sub

Private Sub cmdclear_Click()
    Text1.Text = ""
End Sub

Private Sub cmdOK_Click()
 cHist11.AddNewItem Text1.Text
 cHist12.AddNewItem Text3.Text
 cHist16.AddNewItem Combo1.Text
 cHist17.AddNewItem Combo4.Text
 
 If Trim$(Text1.Text) <> "" Then    ' on règle le cas du MP3
    UseMP3 = True
    With MusMP3
        .Rule = Text1.Text
        .PlaceWhereToPut = Index1
        .DefaultArtistToUse = Text8.Text
        .DefaultYearToUse = Text2.Text
        .DefaultGenreToUse = Combo2.List(Combo2.ListIndex)
        .DefaultAlbumToUse = Text5.Text
        .DefaultTitleToUse = Text6.Text
    End With
 Else
    UseMP3 = False
 End If
 
 If Trim$(Combo4.Text) <> "" Then    ' on règle le cas du WMA
    UseWMA = True
    With MusWMA
        .Rule = Combo4.Text
        .PlaceWhereToPut = index4
        .DefaultArtistToUse = Text12.Text
        .DefaultYearToUse = Text15.Text
        .DefaultGenreToUse = Combo5.List(Combo5.ListIndex)
        .DefaultAlbumToUse = Text14.Text
        .DefaultTitleToUse = Text13.Text
    End With
 Else
    UseWMA = False
 End If
 
 If Trim$(Text3.Text) <> "" Then    ' et on passe au VQF
    UseVQF = True
    With MusVQF
        .Rule = Text3.Text
        .PlaceWhereToPut = Index2
        .DefaultArtistToUse = Text9.Text
        .DefaultTitle = Text4.Text
    End With
 Else
    UseVQF = False
 End If
 
 If Trim$(Combo1.Text) <> "" Then   ' ensuite on passe au format ogg
    UseOGG = True
    With MusOgg
        .Rule = Combo1.Text
        .PlaceWhereToPut = index3
        .DefaultAlbumToUse = Text11.Text
        .DefaultArtistToUse = Text7.Text
        .DefaultGenreToUse = Combo3.List(Combo3.ListIndex)
        .DefaultTitleToUse = Text10.Text
    End With
 Else
    UseOGG = False
 End If
 
 Unload Me
End Sub
Private Sub Command1_Click()
    Text3.Text = ""
End Sub

Private Sub Command2_Click()
    Combo1.Text = ""
End Sub

Private Sub Command3_Click()
    Combo4.Text = ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeTab KeyCode, Shift, SSTab1
    KeyCode = 0
End Sub

Private Sub Form_Load()
 Dim i As Integer
 cHist11.sKey = "MP3String"
 cHist11.Items Text1
 cHist12.sKey = "VQFString"
 cHist12.Items Text3
 cHist16.sKey = "OGGString"
 cHist16.Items Combo1
 cHist17.sKey = "WMAString"
 cHist17.Items Combo4
 SSTab1.TabVisible(4) = False
' On régle son cas au MP3
 Text1.Text = MusMP3.Rule
 Option2(MusMP3.PlaceWhereToPut).Value = True
 Text8.Text = MusMP3.DefaultArtistToUse
 Text2.Text = MusMP3.DefaultYearToUse
 For i = 0 To Combo2.ListCount - 1
  If Trim$(Combo2.List(i)) = Trim$(MusMP3.DefaultGenreToUse) Then
   Combo2.Text = Combo2.List(i)
   Combo2.ListIndex = i
   i = Combo2.ListCount - 1 ' sortie brutale
  End If
 Next
 Text5.Text = MusMP3.DefaultAlbumToUse
 Text6.Text = MusMP3.DefaultTitleToUse

' Ensuite au WMA
 Combo4.Text = MusWMA.Rule
 Option4(MusWMA.PlaceWhereToPut).Value = True
 Text12.Text = MusWMA.DefaultArtistToUse
 Text15.Text = MusWMA.DefaultYearToUse
 For i = 0 To Combo2.ListCount - 1
  If Trim$(Combo5.List(i)) = Trim$(MusWMA.DefaultGenreToUse) Then
   Combo5.Text = Combo5.List(i)
   Combo5.ListIndex = i
   i = Combo5.ListCount - 1 ' sortie brutale
  End If
 Next
 Text14.Text = MusWMA.DefaultAlbumToUse
 Text13.Text = MusWMA.DefaultTitleToUse

' Ensuite au format ogg
 Combo1.Text = MusOgg.Rule
 Option1(MusOgg.PlaceWhereToPut).Value = True
 Text7.Text = MusOgg.DefaultArtistToUse
 For i = 0 To Combo3.ListCount - 1
  If Trim$(Combo3.List(i)) = Trim$(MusOgg.DefaultGenreToUse) Then
   Combo3.Text = Combo3.List(i)
   Combo3.ListIndex = i
   i = Combo3.ListCount - 1 ' sortie brutale
  End If
 Next
 Text11.Text = MusOgg.DefaultAlbumToUse
 Text10.Text = MusOgg.DefaultTitleToUse
' C'est au VQF d'y passer...
 Text3.Text = MusVQF.Rule
 Option3(MusVQF.PlaceWhereToPut).Value = True
 Text9.Text = MusVQF.DefaultArtistToUse
 Text4.Text = MusVQF.DefaultTitle
End Sub
Private Sub List1_DblClick()
    InsertTextInTextBox Text1, List1
End Sub
Private Sub List2_DblClick()
    InsertTextInTextBox Text3, List2
End Sub
Private Sub List3_DblClick()
    InsertTextInTextBox Combo1, List3
End Sub
Private Sub List4_Click()
     InsertTextInTextBox Combo4, List4
End Sub
Private Sub Option1_Click(Index As Integer)
    index3 = Index
End Sub
Private Sub Option2_Click(Index As Integer)
    Index1 = Index
End Sub
Private Sub Option3_Click(Index As Integer)
    Index2 = Index
End Sub
Private Sub Option4_Click(Index As Integer)
    index4 = Index
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
Private Sub Text13_GotFocus()
    SelAll Text13
End Sub
Private Sub Text14_GotFocus()
    SelAll Text14
End Sub
Private Sub Text15_GotFocus()
    SelAll Text15
End Sub
Private Sub Text2_GotFocus()
    SelAll Text2
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
