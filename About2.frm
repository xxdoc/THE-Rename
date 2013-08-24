VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Begin VB.Form About2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   ControlBox      =   0   'False
   Icon            =   "About2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4410
      Left            =   1275
      TabIndex        =   3
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   7779
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "THE Rename"
      TabPicture(0)   =   "About2.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Version info"
      TabPicture(1)   =   "About2.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LV1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Credits"
      TabPicture(2)   =   "About2.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cancer"
      TabPicture(3)   =   "About2.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text2"
      Tab(3).Control(1)=   "Label4"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Donation"
      TabPicture(4)   =   "About2.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command1"
      Tab(4).Control(1)=   "Donation2"
      Tab(4).Control(2)=   "Donation1"
      Tab(4).Control(3)=   "Line1(3)"
      Tab(4).Control(4)=   "Line1(2)"
      Tab(4).ControlCount=   5
      TabCaption(5)   =   "Other programs"
      TabPicture(5)   =   "About2.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label10"
      Tab(5).Control(1)=   "Label9"
      Tab(5).Control(2)=   "Label8"
      Tab(5).Control(3)=   "Label7"
      Tab(5).Control(4)=   "Label6"
      Tab(5).ControlCount=   5
      Begin VB.CommandButton Command1 
         Caption         =   "Donation"
         Height          =   345
         Left            =   -73050
         TabIndex        =   13
         Top             =   3930
         Width           =   1695
      End
      Begin VB.TextBox Donation2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1755
         Left            =   -74940
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2100
         Width           =   5685
      End
      Begin VB.TextBox Donation1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   -74940
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   5625
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   3525
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Text            =   "About2.frx":03B2
         Top             =   450
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "About2.frx":0A4E
         Top             =   420
         Width           =   5595
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   3825
         Left            =   -74880
         TabIndex        =   5
         Top             =   450
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6747
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Resource"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox Label1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   400
         Width           =   5640
      End
      Begin VB.Label Label10 
         Caption         =   $"About2.frx":1AC1
         Height          =   855
         Left            =   -74220
         MouseIcon       =   "About2.frx":1BDC
         MousePointer    =   99  'Custom
         TabIndex        =   19
         Top             =   2400
         Width           =   4815
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FontView"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74760
         MouseIcon       =   "About2.frx":1EE6
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   2100
         Width           =   660
      End
      Begin VB.Label Label8 
         Caption         =   $"About2.frx":21F0
         Height          =   675
         Left            =   -74220
         MouseIcon       =   "About2.frx":2297
         MousePointer    =   99  'Custom
         TabIndex        =   17
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "UnzipThemAll"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74760
         MouseIcon       =   "About2.frx":25A1
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Here is a list of my other freeware programs :"
         Height          =   195
         Left            =   -74760
         TabIndex        =   15
         Top             =   480
         Width           =   3120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   3
         X1              =   -74880
         X2              =   -69550
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   -74880
         X2              =   -69550
         Y1              =   2055
         Y2              =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Click here to see project's home page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -73320
         MouseIcon       =   "About2.frx":28AB
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   4035
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5520
      TabIndex        =   2
      Top             =   4890
      Width           =   1260
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   120
      Picture         =   "About2.frx":2BB5
      ScaleHeight     =   1485
      ScaleWidth      =   855
      TabIndex        =   0
      Top             =   360
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Check on line if newer version available"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      MouseIcon       =   "About2.frx":6E7B
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   5040
      Width           =   2805
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Email me"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3960
      MouseIcon       =   "About2.frx":7185
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4770
      Width           =   630
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Click here to see THE Rename's home page"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   240
      MouseIcon       =   "About2.frx":748F
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4770
      Width           =   3150
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   219
      X2              =   6967
      Y1              =   4665
      Y2              =   4665
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   219
      X2              =   6967
      Y1              =   4650
      Y2              =   4650
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "In memory of my brother Xavier 1966-2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   1980
      Width           =   1200
   End
End
Attribute VB_Name = "About2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDLLVersion Lib "therename.dll" (ByVal Version As String) As String
Private Declare Function GetOggDLLVersion Lib "renogg.dll" (ByVal VersionDLL As String) As String
Private Declare Function GetWMADLLversion Lib "renmm.dll" (ByVal VersionDLL As String) As Long
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    BrowseTo ("https://www.yaskifo.com/order.asp?ID=106302")
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ChangeTab KeyCode, Shift, SSTab1
    KeyCode = 0
End Sub

Private Sub Form_Load()
Dim colonne As ColumnHeader
Dim VersionDLL As String
Dim vtmp As Long
VersionDLL = String$(20, Chr$(0))
VersionDLL = GetDLLVersion(VersionDLL)
Label1.Text = "THE Rename - November 2002 - by Hervé Thouzard" + vbCrLf + vbCrLf + "This is a freeware program. You can use it in any situation (professional and personal use)" + vbCrLf + vbCrLf + "Feel Free to copy and distribute it only if you don't ask money for it !" + vbCrLf + vbCrLf + "Send me your comments, suggestions or bug reports." + vbCrLf + vbCrLf + "You can subscribe to my mailing list to be informed of a new release. Simply go to THE Rename home page to subscribe." + vbCrLf + vbCrLf + "Many thanks to Ferran Pou and to Andy Schmidt for correcting my English and many thanks to Dave Mullins and Ivo Koudela for suggesting so many ideas. Finally and last thanks to Philip Hazel for PCRE, to the GNU guys for RX, to the id3lib guys for the id3lib library and to FoxBat for its DLL." + vbCrLf + "PCRE Version 3.9, Rx version 1.5, id3lib Version 3.8.0pre1" + vbCrLf + "vorbis-sdk-1.0"
AddOneTag App.Major, "Major version number"
AddOneTag App.Minor, "Minor version number"
AddOneTag App.Revision, "Revision"
AddOneTag App.Comments, "Comment"
AddOneTag App.CompanyName, "Company name"
AddOneTag App.EXEName, "Exe name"
AddOneTag App.FileDescription, "File Description"
AddOneTag App.LegalCopyright, "Legal Copyright"
AddOneTag App.LegalTrademarks, "Legal Trademarks"
AddOneTag App.ProductName, "Product Name"
AddOneTag VersionDLL, "therename DLL Version"
VersionDLL = String$(20, Chr$(0))
vtmp = Replace(GetWMADLLversion(VersionDLL), Chr$(0), "")
AddOneTag VersionDLL, "Multimedia DLL Version"
VersionDLL = String$(20, Chr$(0))
VersionDLL = GetOggDLLVersion(VersionDLL)
AddOneTag VersionDLL, "renogg DLL Version"
Set colonne = LV1.ColumnHeaders.Item(1)
AutoSizeColumnHeader LV1, colonne, True
Set colonne = LV1.ColumnHeaders.Item(2)
AutoSizeColumnHeader LV1, colonne, True
Donation1.Text = "THE Rename is freeware and will stay freeware" + vbCrLf + "There is no obligation to pay any amount of money." + vbCrLf + "But since domain registrations and bandwidth is never free," + vbCrLf + "and supporting THE Rename can take up a lot of my time," + vbCrLf + "I would appreciate any donations you can give." + vbCrLf + vbCrLf + "Donations are handled by Yaskifo's secure payment system." + vbCrLf + "Click on the button to donate"
Donation2.Text = "THE Rename est gratuit et restera gratuit" + vbCrLf + "Il n'y a pas d'obligation de payer quoi que ce soit." + vbCrLf + "Mais comme la réservation d'un nom de domaine et l'hébergement ne sont jamais gratuits, et comme le support de THE Rename peut me prendre beaucoup de temps, j'apprécierais une quelconque donation que vous pourriez me faire." + vbCrLf + vbCrLf + "Les donations sont gérées par le système de paiements sécuriés de Yaskifo." + vbCrLf + "Appuyer sur le bouton pour faire une donation"
End Sub

Private Sub AddOneTag(Donnee As String, texte As String)
Dim itmX As ListItem
Set itmX = LV1.ListItems.Add(, , texte)
itmX.SubItems(1) = Donnee
End Sub

Private Sub Label10_Click()
    BrowseTo ("http://www.herve-thouzard.com/fontview.phtml")
End Sub

Private Sub Label2_Click()
    BrowseTo ("mailto:herve@herve-thouzard.com")
End Sub

Private Sub Label4_Click()
     BrowseTo ("http://members.ud.com/vypc/cancer/")
End Sub

Private Sub Label5_Click()
    BrowseTo ("http://www.herve-thouzard.com/verren.phtml?version=" & App.Major & App.Minor & App.Revision)
End Sub

Private Sub Label7_Click()
    BrowseTo ("http://www.herve-thouzard.com/unzipthemall.phtml")
End Sub

Private Sub Label8_Click()
    BrowseTo ("http://www.herve-thouzard.com/unzipthemall.phtml")
End Sub

Private Sub Label9_Click()
    BrowseTo ("http://www.herve-thouzard.com/fontview.phtml")
End Sub

Private Sub lblTitle_Click()
    BrowseTo ("http://www.herve-thouzard.com/therename.phtml")
End Sub

Private Sub LV1_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub
