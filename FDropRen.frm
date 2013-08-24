VERSION 5.00
Begin VB.Form FDropRen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Drop Files to rename them"
   ClientHeight    =   750
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   1170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   750
   ScaleWidth      =   1170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Menu mfile 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu mchgcmd 
         Caption         =   "Change &command to use"
      End
      Begin VB.Menu mend 
         Caption         =   "&End"
      End
   End
End
Attribute VB_Name = "FDropRen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const SWP_WNDFLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Sub Form_Load()
    SetTopmost Me, True
End Sub

Private Sub SetTopmost(frm As Form, bTopmost As Boolean)
     Dim i As Long
     If bTopmost = True Then
          i = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS)
     Else
          i = SetWindowPos(frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_WNDFLAGS)
     End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mfile
    End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim vnb As Long, i As Integer, vnb1 As Integer, vnb2 As Long
Dim fichier As String, vrai As Boolean
Dim VancRep As String
' Variables utilisées dans le cas ou les fichiers droppés ne sont pas du même répertoire
Dim clsFind As New clsFindFile, chaine As String
Dim strFile As String, attributs As Long

List1.Clear
vnb = 0
If Data.GetFormat(vbCFFiles) Then
    For i = 1 To Data.Files.Count
        List1.AddItem Data.Files(i)
    Next
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SetTopmost Me, False
End Sub
