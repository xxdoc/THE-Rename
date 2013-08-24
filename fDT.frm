VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{753FEE6F-A545-4EAA-AAC8-87512ED29F21}#3.0#0"; "ccrpDtp6.ocx"
Begin VB.Form fDT 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Date and Time"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4980
   ControlBox      =   0   'False
   HelpContextID   =   42
   Icon            =   "fDT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   42
      Left            =   2640
      TabIndex        =   42
      ToolTipText     =   "Date and time will not be modified"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   42
      Left            =   3825
      TabIndex        =   43
      ToolTipText     =   "Date and time will be modified"
      Top             =   4200
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      HelpContextID   =   42
      Left            =   90
      TabIndex        =   44
      ToolTipText     =   "Set date to a fixed value"
      Top             =   45
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Creation"
      TabPicture(0)   =   "fDT.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Modified"
      TabPicture(1)   =   "fDT.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(1)=   "Frame3"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Last Access"
      TabPicture(2)   =   "fDT.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame6 
         Caption         =   "Time "
         Height          =   1200
         Left            =   -74865
         TabIndex        =   50
         Top             =   2190
         Width           =   4605
         Begin VB.ComboBox Combo12 
            Height          =   315
            ItemData        =   "fDT.frx":0060
            Left            =   1860
            List            =   "fDT.frx":006D
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   780
            Width           =   2555
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Use EXIF time"
            Height          =   285
            HelpContextID   =   42
            Index           =   2
            Left            =   90
            TabIndex        =   40
            ToolTipText     =   "Set the time to a fixed value"
            Top             =   840
            Width           =   1515
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Keep"
            Height          =   225
            HelpContextID   =   42
            Index           =   0
            Left            =   90
            TabIndex        =   37
            ToolTipText     =   "Don't modify time"
            Top             =   270
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Set to..."
            Height          =   285
            HelpContextID   =   42
            Index           =   1
            Left            =   90
            TabIndex        =   38
            ToolTipText     =   "Set the time to a fixed value"
            Top             =   555
            Width           =   885
         End
         Begin CCRPDTP6.ccrpDtp Heure3 
            Height          =   345
            Left            =   1845
            TabIndex        =   39
            Top             =   500
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Min             =   -109205
            Max             =   2958465
            CCRPVer         =   1
            Var             =   "fDT.frx":00A0
            XD              =   "fDT.frx":00D4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "18:24:29"
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Date "
         Height          =   1755
         Left            =   -74865
         TabIndex        =   49
         Top             =   360
         Width           =   4605
         Begin VB.ComboBox Combo10 
            Height          =   315
            ItemData        =   "fDT.frx":0130
            Left            =   1845
            List            =   "fDT.frx":013D
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1320
            Width           =   2555
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Use EXIF date"
            Height          =   240
            HelpContextID   =   42
            Index           =   3
            Left            =   90
            TabIndex        =   35
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   1380
            Width           =   1425
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Keep"
            Height          =   285
            HelpContextID   =   42
            Index           =   0
            Left            =   90
            TabIndex        =   28
            ToolTipText     =   "Don't modify date"
            Top             =   270
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Set to..."
            Height          =   285
            HelpContextID   =   42
            Index           =   1
            Left            =   90
            TabIndex        =   29
            ToolTipText     =   "Set date to a fixed value"
            Top             =   630
            Width           =   1005
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Increase/Decrease"
            Height          =   240
            HelpContextID   =   42
            Index           =   2
            Left            =   90
            TabIndex        =   31
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   990
            Width           =   1725
         End
         Begin VB.ComboBox Combo6 
            Height          =   315
            HelpContextID   =   42
            ItemData        =   "fDT.frx":0170
            Left            =   1845
            List            =   "fDT.frx":017A
            Style           =   2  'Dropdown List
            TabIndex        =   32
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   945
            Width           =   1050
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            HelpContextID   =   42
            Left            =   2895
            TabIndex        =   33
            Text            =   "1"
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   945
            Width           =   645
         End
         Begin VB.ComboBox Combo5 
            Height          =   315
            HelpContextID   =   42
            ItemData        =   "fDT.frx":0192
            Left            =   3540
            List            =   "fDT.frx":019F
            Style           =   2  'Dropdown List
            TabIndex        =   34
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   945
            Width           =   870
         End
         Begin CCRPDTP6.ccrpDtp Calendrier3 
            Height          =   360
            Left            =   1845
            TabIndex        =   30
            Top             =   540
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   635
            Min             =   -109205
            Max             =   2958465
            CCRPVer         =   1
            Var             =   "fDT.frx":01B5
            XD              =   "fDT.frx":01E9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "24/10/2002"
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Time "
         Height          =   1200
         Left            =   -74865
         TabIndex        =   48
         Top             =   2190
         Width           =   4605
         Begin VB.ComboBox Combo11 
            Height          =   315
            ItemData        =   "fDT.frx":0245
            Left            =   1860
            List            =   "fDT.frx":0252
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   780
            Width           =   2555
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Use EXIF time"
            Height          =   285
            HelpContextID   =   42
            Index           =   2
            Left            =   90
            TabIndex        =   26
            ToolTipText     =   "Set the time to a fixed value"
            Top             =   840
            Width           =   1425
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Keep"
            Height          =   225
            HelpContextID   =   42
            Index           =   0
            Left            =   90
            TabIndex        =   23
            ToolTipText     =   "Don't modify time"
            Top             =   270
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Set to..."
            Height          =   285
            HelpContextID   =   42
            Index           =   1
            Left            =   90
            TabIndex        =   24
            ToolTipText     =   "Set the time to a fixed value"
            Top             =   555
            Width           =   885
         End
         Begin CCRPDTP6.ccrpDtp Heure2 
            Height          =   345
            Left            =   1845
            TabIndex        =   25
            Top             =   500
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Min             =   -109205
            Max             =   2958465
            CCRPVer         =   1
            Var             =   "fDT.frx":0285
            XD              =   "fDT.frx":02B9
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "18:24:29"
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Date "
         Height          =   1755
         Left            =   -74865
         TabIndex        =   47
         Top             =   360
         Width           =   4605
         Begin VB.ComboBox Combo9 
            Height          =   315
            ItemData        =   "fDT.frx":0315
            Left            =   1845
            List            =   "fDT.frx":0322
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1320
            Width           =   2555
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Use EXIF date"
            Height          =   240
            HelpContextID   =   42
            Index           =   3
            Left            =   90
            TabIndex        =   21
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   1380
            Width           =   1425
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Keep"
            Height          =   285
            HelpContextID   =   42
            Index           =   0
            Left            =   90
            TabIndex        =   14
            ToolTipText     =   "Don't modify date"
            Top             =   270
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Set to..."
            Height          =   285
            HelpContextID   =   42
            Index           =   1
            Left            =   90
            TabIndex        =   15
            ToolTipText     =   "Set date to a fixed value"
            Top             =   630
            Width           =   1005
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Increase/Decrease"
            Height          =   240
            HelpContextID   =   42
            Index           =   2
            Left            =   90
            TabIndex        =   17
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   990
            Width           =   1725
         End
         Begin VB.ComboBox Combo4 
            Height          =   315
            HelpContextID   =   42
            ItemData        =   "fDT.frx":0355
            Left            =   1845
            List            =   "fDT.frx":035F
            Style           =   2  'Dropdown List
            TabIndex        =   18
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   945
            Width           =   1050
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            HelpContextID   =   42
            Left            =   2895
            TabIndex        =   19
            Text            =   "1"
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   945
            Width           =   645
         End
         Begin VB.ComboBox Combo3 
            Height          =   315
            HelpContextID   =   42
            ItemData        =   "fDT.frx":0377
            Left            =   3540
            List            =   "fDT.frx":0384
            Style           =   2  'Dropdown List
            TabIndex        =   20
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   945
            Width           =   870
         End
         Begin CCRPDTP6.ccrpDtp Calendrier2 
            Height          =   360
            Left            =   1845
            TabIndex        =   16
            Top             =   540
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   635
            Min             =   -109205
            Max             =   2958465
            CCRPVer         =   1
            Var             =   "fDT.frx":039A
            XD              =   "fDT.frx":03CE
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "24/10/2002"
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Time "
         Height          =   1200
         Left            =   135
         TabIndex        =   46
         Top             =   2190
         Width           =   4605
         Begin VB.ComboBox Combo8 
            Height          =   315
            ItemData        =   "fDT.frx":042A
            Left            =   1860
            List            =   "fDT.frx":0437
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   780
            Width           =   2555
         End
         Begin CCRPDTP6.ccrpDtp Heure1 
            Height          =   340
            Left            =   1845
            TabIndex        =   11
            Top             =   500
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   609
            Min             =   -109205
            Max             =   2958465
            CCRPVer         =   1
            Var             =   "fDT.frx":046A
            XD              =   "fDT.frx":049E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "18:24:29"
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Use EXIF time"
            Height          =   285
            HelpContextID   =   42
            Index           =   2
            Left            =   90
            TabIndex        =   12
            ToolTipText     =   "If the EXIF time does not exists, the current time will stay unchanged"
            Top             =   840
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Set to..."
            Height          =   285
            HelpContextID   =   42
            Index           =   1
            Left            =   90
            TabIndex        =   10
            ToolTipText     =   "Set the time to a fixed value"
            Top             =   555
            Width           =   960
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Keep"
            Height          =   225
            HelpContextID   =   42
            Index           =   0
            Left            =   90
            TabIndex        =   9
            ToolTipText     =   "Don't modify time"
            Top             =   270
            Value           =   -1  'True
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Date "
         Height          =   1755
         Left            =   135
         TabIndex        =   45
         Top             =   360
         Width           =   4605
         Begin CCRPDTP6.ccrpDtp Calendrier1 
            Height          =   360
            Left            =   1845
            TabIndex        =   2
            Top             =   540
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   635
            Min             =   -109205
            Max             =   2958465
            CCRPVer         =   1
            Var             =   "fDT.frx":04FA
            XD              =   "fDT.frx":052E
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "24/10/2002"
         End
         Begin VB.ComboBox Combo7 
            Height          =   315
            ItemData        =   "fDT.frx":058A
            Left            =   1845
            List            =   "fDT.frx":0597
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1320
            Width           =   2555
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Use EXIF date"
            Height          =   240
            HelpContextID   =   42
            Index           =   3
            Left            =   90
            TabIndex        =   7
            ToolTipText     =   "If the EXIF date does not exists, the current date will stay unchanged"
            Top             =   1380
            Width           =   1425
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            HelpContextID   =   42
            ItemData        =   "fDT.frx":05CA
            Left            =   3540
            List            =   "fDT.frx":05D7
            Style           =   2  'Dropdown List
            TabIndex        =   6
            TabStop         =   0   'False
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   945
            Width           =   870
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            HelpContextID   =   42
            Left            =   2895
            TabIndex        =   5
            Text            =   "1"
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   945
            Width           =   645
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            HelpContextID   =   42
            ItemData        =   "fDT.frx":05ED
            Left            =   1845
            List            =   "fDT.frx":05F7
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   945
            Width           =   1050
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Increase/Decrease"
            Height          =   240
            HelpContextID   =   42
            Index           =   2
            Left            =   90
            TabIndex        =   3
            ToolTipText     =   "Use this option to increase or decrease date"
            Top             =   990
            Width           =   1725
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Set to..."
            Height          =   285
            HelpContextID   =   42
            Index           =   1
            Left            =   90
            TabIndex        =   1
            ToolTipText     =   "Set date to a fixed value"
            Top             =   630
            Width           =   900
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Keep"
            Height          =   285
            HelpContextID   =   42
            Index           =   0
            Left            =   90
            TabIndex        =   0
            ToolTipText     =   "Don't modify date"
            Top             =   270
            Value           =   -1  'True
            Width           =   735
         End
      End
   End
   Begin VB.Label Warning 
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   51
      Top             =   3780
      Width           =   4815
   End
End
Attribute VB_Name = "fDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Ouverture As Boolean
Dim ind1 As Integer
Dim ind2 As Integer
Dim ind3 As Integer
Dim ind4 As Integer
Dim ind5 As Integer
Dim ind6 As Integer
Dim LaDateTime As New CDateTime
Private Sub cmdCancel_Click()
 LaDateTime.DTOk = False
 Unload Me
End Sub

Private Sub Command1_Click()
    With LaDateTime
        .DTOk = True
        .CreateDOption = ind1
        .CreateTOption = ind2
        .AccessDOption = ind5
        .AccessTOption = ind6
        .WriteDOption = ind3
        .WriteTOption = ind4
    
        .CreateDFixed = Calendrier1.Value
        .CreateDInc1 = Combo1.ListIndex
        .CreateDInc2 = Val(Text1.Text)
        .CreateDInc3 = Combo2.ListIndex
        .CreateTFixed = Heure1.Value
    
        .AccessDFixed = Calendrier2.Value
        .AccessDInc1 = Combo6.ListIndex
        .AccessDInc2 = Val(Text3.Text)
        .AccessDInc3 = Combo5.ListIndex
        .AccessTFixed = Heure2.Value
    
        .WriteDFixed = Calendrier3.Value
        .WriteDInc1 = Combo4.ListIndex
        .WriteDInc2 = Val(Text2.Text)
        .WriteDInc3 = Combo3.ListIndex
        .WriteTFixed = Heure3.Value
    
        .CreateDExif = Combo7.ListIndex
        .CreateTExif = Combo8.ListIndex
        .AccessDExif = Combo9.ListIndex
        .AccessTExif = Combo11.ListIndex
        .WriteDExif = Combo10.ListIndex
        .WriteTExif = Combo12.ListIndex
    End With
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeTab KeyCode, Shift, SSTab1
    KeyCode = 0
End Sub

Private Sub Form_Load()
Ouverture = True
 Select Case DTEnCours
  Case 1 ' 1= Renommer
   Set LaDateTime = DT1
   Warning.Caption = "Warning files will have their date and time changed only when you will rename files"
  Case 2 ' 2=bin
   Set LaDateTime = DT2
   Warning.Caption = "Warning files will have their date and time changed only when you will copy or paste files"
  Case 3 ' 3=bouton copie multiple
   Set LaDateTime = DT3
   Warning.Caption = "Warning files will have their date and time changed only when you will drop files on the multiple copy button files"
  Case 4 ' 4=Action immédiate
   Set LaDateTime = DT4
   Warning.Caption = ""
 End Select
 ' Date de création
    Option1(LaDateTime.CreateDOption).Value = True
    Calendrier1.Value = LaDateTime.CreateDFixed
    Combo1.ListIndex = LaDateTime.CreateDInc1
    Text1.Text = LaDateTime.CreateDInc2
    Combo2.ListIndex = LaDateTime.CreateDInc3
    Option2(LaDateTime.CreateTOption).Value = True
    Heure1.Value = LaDateTime.CreateTFixed
    Combo7.ListIndex = LaDateTime.CreateDExif
    Combo8.ListIndex = LaDateTime.CreateTExif
  
 ' Date de modification
    Option5(LaDateTime.AccessDOption).Value = True
    Calendrier2.Value = LaDateTime.AccessDFixed
    Combo6.ListIndex = LaDateTime.AccessDInc1
    Text3.Text = LaDateTime.AccessDInc2
    Combo5.ListIndex = LaDateTime.AccessDInc3
    Option6(LaDateTime.AccessTOption).Value = True
    Heure2.Value = LaDateTime.AccessTFixed
    Combo9.ListIndex = LaDateTime.AccessDExif
    Combo11.ListIndex = LaDateTime.AccessTExif
    
 ' Date de dernier accès
    Option3(LaDateTime.WriteDOption).Value = True
    Calendrier3.Value = LaDateTime.WriteDFixed
    Combo4.ListIndex = LaDateTime.WriteDInc1
    Text2.Text = LaDateTime.WriteDInc2
    Combo3.ListIndex = LaDateTime.WriteDInc3
    Option4(LaDateTime.WriteTOption).Value = True
    Heure3.Value = LaDateTime.WriteTFixed
    Combo10.ListIndex = LaDateTime.WriteDExif
    Combo12.ListIndex = LaDateTime.WriteTExif
    
 Option1_Click (LaDateTime.CreateDOption)
 Option2_Click (LaDateTime.CreateTOption)
 Option5_Click (LaDateTime.AccessDOption)
 Option6_Click (LaDateTime.AccessTOption)
 Option3_Click (LaDateTime.WriteDOption)
 Option4_Click (LaDateTime.WriteTOption)
 Ouverture = False
End Sub
Private Sub Option1_Click(Index As Integer)
Combo1.Visible = False
Text1.Visible = False
Combo2.Visible = False
Combo7.Visible = False
Calendrier1.Visible = False
Select Case Index
    Case 0
    Case 1
        Calendrier1.Visible = True
    Case 2
        Combo1.Visible = True
        Text1.Visible = True
        Combo2.Visible = True
    Case 3
        Combo7.Visible = True
End Select
If Ouverture = False Then
    ind1 = Index
End If
If Combo7.ListIndex = -1 Then
    Combo7.ListIndex = 0
End If
End Sub

Private Sub Option2_Click(Index As Integer)
Combo8.Visible = False
If Index = 0 Or Index = 2 Then
    Heure1.Visible = False
Else
    Heure1.Visible = True
End If
If Index = 2 Then
    Combo8.Visible = True
End If
If Ouverture = False Then
    ind2 = Index
End If
End Sub

Private Sub Option3_Click(Index As Integer)
Combo4.Visible = False
Text2.Visible = False
Combo3.Visible = False
Combo9.Visible = False
Calendrier2.Visible = False
Select Case Index
    Case 0
    Case 1
        Calendrier2.Visible = True
    Case 2
        Combo4.Visible = True
        Text2.Visible = True
        Combo3.Visible = True
    Case 3
        Combo9.Visible = True
End Select
If Ouverture = False Then
    ind3 = Index
End If
If Combo9.ListIndex = -1 Then
    Combo9.ListIndex = 0
End If
End Sub

Private Sub Option4_Click(Index As Integer)
Combo11.Visible = False
If Index = 0 Or Index = 2 Then
    Heure2.Visible = False
Else
    Heure2.Visible = True
End If

If Index = 2 Then
    Combo11.Visible = True
End If

If Ouverture = False Then
    ind4 = Index
End If
End Sub

Private Sub Option5_Click(Index As Integer)
Combo6.Visible = False
Text3.Visible = False
Combo5.Visible = False
Combo10.Visible = False
Calendrier3.Visible = False
Select Case Index
    Case 0
    Case 1
        Calendrier3.Visible = True
    Case 2
        Combo6.Visible = True
        Text3.Visible = True
        Combo5.Visible = True
    Case 3
        Combo10.Visible = True
End Select
If Ouverture = False Then
    ind5 = Index
End If
If Combo10.ListIndex = -1 Then
    Combo10.ListIndex = 0
End If
End Sub

Private Sub Option6_Click(Index As Integer)
Combo12.Visible = False
If Index = 0 Or Index = 2 Then
    Heure3.Visible = False
Else
    Heure3.Visible = True
End If

If Index = 2 Then
    Combo12.Visible = True
End If

If Ouverture = False Then
    ind6 = Index
End If
End Sub

Private Sub Text1_GotFocus()
    SelAll Text1
End Sub

Private Sub Text2_GotFocus()
    SelAll Text2
End Sub

Private Sub Text3_GotFocus()
    SelAll Text3
End Sub
