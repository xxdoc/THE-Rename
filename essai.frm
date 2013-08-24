VERSION 5.00
Begin VB.Form essai 
   Caption         =   "Form1"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   7410
   ScaleWidth      =   11325
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   2400
      Left            =   3420
      TabIndex        =   11
      Top             =   4230
      Width           =   2625
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Analyze"
      Height          =   330
      Left            =   3420
      TabIndex        =   10
      Top             =   3825
      Width           =   1860
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   270
      TabIndex        =   9
      Top             =   4005
      Width           =   3075
   End
   Begin VB.PictureBox panelcmd 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2130
      Left            =   45
      ScaleHeight     =   2130
      ScaleWidth      =   6225
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   6225
      Begin VB.TextBox cmdtxt3 
         Height          =   285
         Left            =   5310
         TabIndex        =   14
         Text            =   "25"
         ToolTipText     =   "Enter begin's value"
         Top             =   1530
         Width           =   495
      End
      Begin VB.TextBox cmdtxt1 
         Height          =   285
         Left            =   3420
         TabIndex        =   5
         Text            =   "1"
         ToolTipText     =   "Enter begin's value"
         Top             =   1215
         Width           =   495
      End
      Begin VB.TextBox cmdtxt2 
         Height          =   285
         Left            =   4635
         TabIndex        =   4
         Text            =   "1"
         ToolTipText     =   "Enter increment value (1 for example)"
         Top             =   1215
         Width           =   495
      End
      Begin VB.CommandButton cmdclear 
         Caption         =   "Clear"
         Height          =   375
         Left            =   5265
         TabIndex        =   3
         ToolTipText     =   "Clear command"
         Top             =   45
         Width           =   735
      End
      Begin VB.ListBox listcmd 
         Height          =   1425
         ItemData        =   "essai.frx":0000
         Left            =   90
         List            =   "essai.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Double click to add command, one click to have help"
         Top             =   540
         Width           =   2175
      End
      Begin VB.TextBox txtlang 
         Height          =   330
         Left            =   45
         TabIndex        =   1
         Text            =   "img<ddddd>.<EXLower>"
         ToolTipText     =   "Type your command or pick one from the list"
         Top             =   45
         Width           =   5145
      End
      Begin VB.Label lang4 
         AutoSize        =   -1  'True
         Caption         =   "Maximum of characters to take from file"
         Height          =   195
         Left            =   2430
         TabIndex        =   13
         Top             =   1575
         Width           =   2760
      End
      Begin VB.Label langhlp 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Left            =   2430
         TabIndex        =   12
         ToolTipText     =   "Help for command"
         Top             =   585
         Width           =   90
      End
      Begin VB.Label lang1 
         AutoSize        =   -1  'True
         Caption         =   "Options for counters"
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
         Left            =   2385
         TabIndex        =   8
         Top             =   945
         Width           =   1740
      End
      Begin VB.Label lang3 
         AutoSize        =   -1  'True
         Caption         =   "Step"
         Height          =   195
         Left            =   4080
         TabIndex        =   7
         Top             =   1245
         Width           =   330
      End
      Begin VB.Label lang2 
         AutoSize        =   -1  'True
         Caption         =   "Begin value"
         Height          =   195
         Left            =   2385
         TabIndex        =   6
         Top             =   1245
         Width           =   840
      End
   End
End
Attribute VB_Name = "essai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hlplang(21) As String
Dim langage(21) As String
Private Sub cmdclear_Click()
txtlang.Text = ""
txtlang.SetFocus
End Sub

Private Sub Command2_Click()
 
 If Len(Trim(txtlang.Text)) = 0 Then
  resultat = MsgBox("Error, expression is empty !", vbOKOnly)
  Exit Sub
 End If
 txtlang.Text = Trim(txtlang.Text)
 copie = UCase(txtlang.Text)
 vnb1 = charoccurs(copie, "<")
 vnb2 = charoccurs(copie, ">")
 If vnb1 <> vnb2 Then
  If vnb1 > vnb2 Then
   resultat = MsgBox("Error, number of '>' is different from number of '<'", vbOKOnly)
  Else
   resultat = MsgBox("Error, number of '<' is different from number of '>'", vbOKOnly)
  End If
  Exit Sub
 End If
 
 ' Théoriquement il n'y a plus d'erreurs, on peut commencer à lancer l'analyse
 cmdencours = 0
 posdeb = 1
 longueur = Len(txtlang.Text)
 i = 1
 
 While i <= longueur
  If Mid(copie, i, 1) = "<" Then ' C'est une commande *******************************
   chainetempo = Mid(copie, i)
   vnb1 = at(chainetempo, ">", 1)
   chainetempo = Mid(chainetempo, 1, vnb1)
   ' Bon, on sait qu'on est sur une commande, on l'a, il ne reste plus qu'à savoir laquelle c'est
   cmdencours = cmdencours + 1
   If Left(chainetempo, 2) <> "<D" And Left(chainetempo, 2) <> "<H" And Left(chainetempo, 2) <> "<O" Then ' ce n'est pas une commande de compteur
    vrai = False
    For vnb2 = 1 To vnbcmd
     If UCase(Trim(langage(vnb2))) = UCase(Trim(chainetempo)) Then
      vrai = True
      Exit For
     End If
    Next
    If vrai = False Then
     resultat = MsgBox("Error, " + chainetempo + " is not a valid command !", vbOKOnly)
     Exit Sub
    End If
    cmdencours = cmdencours + 1
    commandes(cmdencours, 1) = Trim(Str(vnb2))     ' L'indice de la commande
    commandes(cmdencours, 2) = ""                  ' Commande sans paramètre
   Else ' c'est une commande de compteur
    vrai = False
    If Left(chainetempo, 2) = "<D" Then
     chainetempo2 = "<DDDDD>"
    Else
     If Left(chainetempo, 2) = "<H" Then
      chainetempo2 = "<HHHHH>"
     Else
      chainetempo2 = "<OOOOO>"
     End If
    End If
    For vnb2 = 1 To vnbcmd
     If UCase(Trim(langage(vnb2))) = UCase(Trim(chainetempo2)) Then
      vrai = True
      Exit For
     End If
    Next
    If vrai = False Then
     resultat = MsgBox("Error, " + chainetempo + " is not a valid command !", vbOKOnly)
     Exit Sub
    End If
    cmdencours = cmdencours + 1
    commandes(cmdencours, 1) = Trim(Str(vnb2))     ' L'indice de la commande
    commandes(cmdencours, 2) = Trim(Str(Len(chainetempo) - 2))       ' Commande sans paramètre
   End If
   i = vnb1 + Len(Left(copie, i))
  Else ' C'est un litteral ***************************************************************
   vrai = False
   litteral = ""
   While vrai = False And i <= longueur
    If Mid(txtlang.Text, i, 1) <> "<" Then
     litteral = litteral + Mid(txtlang.Text, i, 1) ' Pour les litteraux, il faut prendre le texte original, pas celui qui a été passé en majuscules
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
End Sub

Private Sub Form_Load()
langage(1) = "<curext>"
hlplang(1) = "Take current file's extension"
langage(2) = "<curprefix>"
hlplang(2) = "Take current file's prefix"
langage(3) = "<ddddd>"
hlplang(3) = "Add a counter in decimal, number 'd' indicate counter siez"
langage(4) = "<EXCapital>"
hlplang(4) = "Capitalize extension"
langage(5) = "<EXInvert>"
hlplang(5) = "Invert extension's letters"
langage(6) = "<EXLower>"
hlplang(6) = "Convert extension to lower cases"
langage(7) = "<EXToggle>"
hlplang(7) = "Toggle cases of extension"
langage(8) = "<EXUpper>"
hlplang(8) = "Convert extension to upper cases"
langage(9) = "<filecontent>"
hlplang(9) = "Take file's content"
langage(10) = "<filedate>"
hlplang(10) = "Take file's date"
langage(11) = "<filetime>"
hlplang(11) = "Take file's time"
langage(12) = "<hhhhh>"
hlplang(12) = "Add a counter in hexadecimal, number of 'h' indicate counter size"
langage(13) = "<ooooo>"
hlplang(13) = "Add a counter in octal, number of 'o' indicate counter size"
langage(14) = "<PRCapital>"
hlplang(14) = "Convert prefix to capitals"
langage(15) = "<PRInvert>"
hlplang(15) = "Invert letters in prefix"
langage(16) = "<PRLower>"
hlplang(16) = "Convert prefix to lower cases"
langage(17) = "<PRToggle>"
hlplang(17) = "Toggle cases of prefix"
langage(18) = "<PRUpper>"
hlplang(18) = "Convert prefix to upper cases"
langage(19) = "<systdate>"
hlplang(19) = "Take system date"
langage(20) = "<systtime>"
hlplang(20) = "Take system time"
langage(21) = "<ttfname>"
hlplang(21) = "Take internal name for a truetype file"

For i = 1 To 21
 listcmd.AddItem langage(i)
Next i
End Sub


Private Sub listcmd_Click()
 langhlp = hlplang(listcmd.ListIndex + 1)
End Sub

Private Sub listcmd_DblClick()
 Dim debut As Integer
 Dim fin As Integer
 Dim letexte1 As String
 Dim letexte2 As String
 Dim vancpos As Long
 vancpos = txtlang.SelStart
 If Len(Trim$(txtlang.Text)) > 0 Then
  letexte1 = Left$(txtlang.Text, txtlang.SelStart)
  letexte2 = Mid$(txtlang.Text, txtlang.SelStart + 1)
  If txtlang.SelLength = Len(Trim(txtlang.Text)) Then
   txtlang.Text = listcmd.List(listcmd.ListIndex)
  Else
   txtlang.Text = letexte1 + listcmd.List(listcmd.ListIndex) + letexte2
  End If
 Else
  txtlang.Text = listcmd.List(listcmd.ListIndex)
 End If
 txtlang.SelStart = vancpos
 txtlang.SetFocus
End Sub

