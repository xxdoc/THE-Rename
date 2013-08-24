VERSION 5.00
Begin VB.Form about 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About ..."
   ClientHeight    =   4500
   ClientLeft      =   2400
   ClientTop       =   3780
   ClientWidth     =   6825
   ClipControls    =   0   'False
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105.981
   ScaleMode       =   0  'User
   ScaleWidth      =   6409.028
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   3270
      Left            =   1515
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   104
      Width           =   5220
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1545
      Left            =   315
      Picture         =   "about.frx":000C
      ScaleHeight     =   1485
      ScaleWidth      =   855
      TabIndex        =   4
      Top             =   300
      Width           =   915
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "about.frx":42D2
      Left            =   1080
      List            =   "about.frx":4570
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Combo1"
      ToolTipText     =   "And all the other i forgotted"
      Top             =   4065
      Width           =   4215
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5445
      TabIndex        =   0
      Top             =   4080
      Width           =   1260
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
      Left            =   195
      TabIndex        =   5
      Top             =   1920
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Thanks too :"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   4125
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   56.343
      X2              =   6393.064
      Y1              =   2722.91
      Y2              =   2722.91
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Click here to see THE Rename's home page"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2235
      MouseIcon       =   "about.frx":554B
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3525
      Width           =   4005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   56.343
      X2              =   6393.064
      Y1              =   2733.263
      Y2              =   2733.263
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
  Unload Me
End Sub
Private Sub Form_Load()
Combo1.ListIndex = 0
Label1.Text = vbCrLf + "THE Rename Version 2.0c - August 2001 - by Hervé Thouzard" + vbCrLf + vbCrLf + "This is a freeware program. You can use it in any situation (professional and personal use)" + vbCrLf + "Feel Free to copy and distribute it only if you don't ask money for it !" + vbCrLf + "Send me your comments, suggestions or bug reports at herve@herve-thouzard.com" + vbCrLf + "You can subscribe to my mailing list to be informed of a new release. Simply go to THE Rename home page to subscribe." + vbCrLf + "Many thanks to Ferran Pou and to Andy Schmidt for correcting my English and many thanks to Dave Mullins and Ivo Koudela for suggesting so many ideas. Finaly and last thanks to Philip Hazel for PCRE, to the GNU guys for RX and to the id3lib guys for the id3lib library." + vbCrLf + vbCrLf + "PCRE Version 3.4, Rx version 1.5, id3lib Version 3.8.0pre1"
End Sub
Private Sub lblTitle_Click()
 BrowseTo ("http://www.herve-thouzard.com/therename.phtml")
End Sub
