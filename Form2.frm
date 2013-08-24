VERSION 5.00
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "CCRPFTV6.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5820
   ScaleWidth      =   6225
   Begin CCRPFolderTV6.FolderTreeview FolderTreeview1 
      Height          =   5820
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   172
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   10266
      IntegralHeight  =   0   'False
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
FolderTreeview1(0).left = 0
FolderTreeview1(0).top = 0
FolderTreeview1(0).height = Form2.ScaleHeight
FolderTreeview1(0).width = Form2.ScaleWidth
End Sub
