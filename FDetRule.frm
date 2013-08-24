VERSION 5.00
Object = "{753FEE6F-A545-4EAA-AAC8-87512ED29F21}#3.0#0"; "ccrpDtp6.ocx"
Begin VB.Form FDetRule 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rules"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5865
   ControlBox      =   0   'False
   HelpContextID   =   472
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   300
      HelpContextID   =   38
      Left            =   3000
      TabIndex        =   5
      ToolTipText     =   "Save rule"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   300
      HelpContextID   =   38
      Left            =   1807
      TabIndex        =   4
      ToolTipText     =   "Don't save rule"
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   4140
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   52
      TabIndex        =   2
      Top             =   3360
      Width           =   5775
   End
   Begin VB.ListBox List2 
      Height          =   1035
      ItemData        =   "FDetRule.frx":0000
      Left            =   52
      List            =   "FDetRule.frx":0002
      TabIndex        =   1
      Top             =   1860
      Width           =   5775
   End
   Begin VB.ListBox List1 
      Height          =   1035
      ItemData        =   "FDetRule.frx":0004
      Left            =   37
      List            =   "FDetRule.frx":0006
      TabIndex        =   0
      Top             =   420
      Width           =   5775
   End
   Begin CCRPDTP6.ccrpDtp Calendrier1 
      Height          =   360
      Left            =   4380
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   635
      Min             =   -109205
      Max             =   2958465
      CCRPVer         =   1
      Var             =   "FDetRule.frx":0008
      XD              =   "FDetRule.frx":003C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "09/10/2002"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "4) Select a name for this rule"
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Top             =   3900
      Width           =   2010
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "3) Type the value to test"
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   3060
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "2) Select the condition"
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   1620
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1) Select the element you want to test"
      Height          =   195
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   2670
   End
End
Attribute VB_Name = "FDetRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Conditions(13) As String
Dim LeMasque As Integer
Dim Masques(13) As String
Dim Tests(16) As String
Dim vok As Boolean
Dim maction As Integer
Dim OneRule As New Rule
Dim TheRules As New Rules
Dim mTitre As String
Dim mIndice As Integer
Dim Relation(13) As String

Public Function GetRule(action As Integer, Regles As Rules, indice As Integer, Optional Title As String) As Boolean
    ' Action :
    ' 1=Add (new), 2=Modify
    If Not IsMissing(Title) Then mTitre = Title
    maction = action
    mIndice = indice
    Set TheRules = Nothing
    Set TheRules = Regles
    If action = 2 Then
        Set OneRule = Regles.GetRule(indice)
    End If
    Me.Show vbModal
    GetRule = vok
End Function

Private Sub cmdCancel_Click()
    vok = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim i As Integer
Dim vnb As Integer
Dim RglTempo As New Rule
vok = True
Dim vtmp As Variant
If List1.ListIndex = -1 Then    ' Pas d'élément sélectionné
    MsgBox "Please, select the element you want to test or press the 'Cancel' button !"
    List1.SetFocus
    Exit Sub
End If

If List2.ListIndex = -1 Then    ' Pas de condition sélectionnée
    MsgBox "Please, select the condition or press the 'Cancel' button !"
    List2.SetFocus
    Exit Sub
End If

' test sur du numérique
If LeMasque = 1 Then
    If Trim$(Text1.Text) = "" Then   ' Valeur vide
        MsgBox "Please, enter a numeric value to test or press the 'Cancel' button "
        Text1.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(Text1.Text) Then ' Valeur incorrecte
        MsgBox "Please, enter a valid numeric value or press the 'Cancel' button !"
        Text1.SetFocus
        Exit Sub
    End If
End If

' Test sur date mais pas de date sélectionnée
If LeMasque = 3 Then
    If Calendrier1.Value = "" Then
        MsgBox "Please, select a date to test or press the 'Cancel' button "
        Calendrier1.SetFocus
        Exit Sub
    End If
End If

' Test pour des valeurs texte mais vide
If LeMasque = 2 And Trim$(Text1.Text) = "" Then
    If MsgBox("Are you sure you want to have a blank test value ?", vbYesNo, "Warning") = vbNo Then
        Text1.SetFocus
        Exit Sub
    End If
End If

If Trim$(Text2.Text) = "" Then
    MsgBox "Please, enter a name for this rule"
    Text2.SetFocus
    Exit Sub
End If

' Si on est encore là c'est que tout est bon
Select Case maction
    Case 2   ' Modify
        vnb = TheRules.RulesCount
        For i = 1 To vnb
            If i <> mIndice Then
                Set RglTempo = TheRules.GetRule(i)
                If Text2.Text = RglTempo.RuleName Then
                    MsgBox "Warning, a rule already exist with this name, change it !"
                    Text2.SetFocus
                    Exit Sub
                End If
            End If
        Next
        OneRule.RuleType = List1.ListIndex
        OneRule.RuleCondition = List2.ListIndex
        OneRule.RuleName = Text2.Text
        If Calendrier1.Visible = True Then
            OneRule.RuleTestValue = Calendrier1.Value
            vtmp = Calendrier1.Value
        Else
            OneRule.RuleTestValue = Text1.Text
            vtmp = Text1.Text
        End If
        OneRule.RuleDescription = List1.List(List1.ListIndex) & " " & List2.List(List2.ListIndex) & " " & vtmp
    Case 1  ' Add
        Set OneRule = Nothing
        OneRule.RuleActive = True
        OneRule.RuleCondition = List2.ListIndex
        OneRule.RuleType = List1.ListIndex
        OneRule.RuleName = Text2.Text
        If Calendrier1.Visible = True Then
            OneRule.RuleTestValue = Calendrier1.Value
            vtmp = Calendrier1.Value
        Else
            OneRule.RuleTestValue = Text1.Text
            vtmp = Text1.Text
        End If
        OneRule.RuleDescription = List1.List(List1.ListIndex) & " " & List2.List(List2.ListIndex) & " " & vtmp
        If TheRules.AddRule(OneRule) = 0 Then
            MsgBox "Sorry, a rule already exist with this name, change it."
            Text2.SetFocus
            Exit Sub
        End If
End Select
Unload Me
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Dim vnb As Integer
    vok = False
    Me.Caption = mTitre
    Conditions(1) = "File's size"
    Masques(1) = "1,1"
    Conditions(2) = "Prefix"
    Masques(2) = "1,2"
    Conditions(3) = "Extension"
    Masques(3) = "1,2"
    Conditions(4) = "Path"
    Masques(4) = "1,2"
    Conditions(5) = "Complete filename (path + prefix + extension)"
    Masques(5) = "1,2"
    Conditions(6) = "Creation's date"
    Masques(6) = "1,3"
    Conditions(7) = "Last modification's date"
    Masques(7) = "1,3"
    Conditions(8) = "Last update date"
    Masques(8) = "1,3"
    Conditions(9) = "Archive attribute"
    Masques(9) = "0,0"
    Conditions(10) = "Read Only Attribute"
    Masques(10) = "0,0"
    Conditions(11) = "Hidden Attribute"
    Masques(11) = "0,0"
    Conditions(12) = "System Attribute"
    Masques(12) = "0,0"
    Conditions(13) = "TypeOf (File or Folder ?)"
    Masques(13) = "0,0"
    Tests(1) = ">"
    Tests(2) = ">="
    Tests(3) = "<"
    Tests(4) = "<="
    Tests(5) = "=="
    Tests(6) = "!="
    Tests(7) = "begin with"
    Tests(8) = "contains"
    Tests(9) = "does not contains"
    Tests(10) = "is equal to"
    Tests(11) = "is different from"
    Tests(12) = "is set"
    Tests(13) = "is not set"
    Tests(14) = "isFile"
    Tests(15) = "isFolder"
    Tests(16) = "is ended by"
    Relation(1) = "1,2,3,4,5,6"     ' Size
    Relation(2) = "7,8,9,10,11,16"  ' Prefix
    Relation(3) = Relation(2)       ' Extension
    Relation(4) = Relation(2)       ' Path
    Relation(5) = Relation(2)       ' Complete filename
    Relation(6) = Relation(1)       ' Creation date
    Relation(7) = Relation(6)       ' Modif date
    Relation(8) = Relation(6)       ' last update date
    Relation(9) = "12,13"           ' Archive Attr
    Relation(10) = Relation(9)      ' Read Only Attr
    Relation(11) = Relation(9)      ' System Attr
    Relation(12) = Relation(9)      ' Hidden Attr
    Relation(13) = "14,15"          ' TypeOf
    List1.Clear
    vnb = UBound(Conditions)
    For i = 1 To vnb
        List1.AddItem Conditions(i)
    Next
    If maction = 2 Then ' Modify
        List1.ListIndex = OneRule.RuleType
        MetCond OneRule.RuleType
        List2.ListIndex = OneRule.RuleCondition
        Text1.Text = OneRule.RuleTestValue
        Calendrier1.Value = OneRule.RuleTestValue
        Text2.Text = OneRule.RuleName
    End If
End Sub

Private Sub MetCond(Index As Integer)
Dim vnb As Integer
Dim i As Integer
Dim j As Integer
Dim vnb2 As Integer
Dim vtmp As String
vnb = UBound(Conditions)
vtmp = List1.List(Index)
For i = 1 To vnb
    If Conditions(i) = vtmp Then
        ' On commence par rajouter toutes les conditions possibles
        List2.Clear
        vnb2 = CharOccurs(Relation(i), ",")
        vnb2 = vnb2 + 1
        For j = 1 To vnb2
            List2.AddItem Tests(GetToken(Relation(i), ",", j))
        Next
        ' Ensuite on gère l'état de la zone de test
        If GetToken(Masques(i), ",", 1) = "1" Then
            Label3.Visible = True
            Text1.Visible = True
        Else
            Label3.Visible = False
            Text1.Visible = False
        End If
        If GetToken(Masques(i), ",", 2) = "3" Then
            Text1.Visible = False
            Calendrier1.Left = Text1.Left
            Calendrier1.Top = Text1.Top
            Calendrier1.Visible = True
            Label3.Caption = "3) Select a date"
        Else
            Label3.Caption = "3) Type the value to test"
            Calendrier1.Visible = False
            Calendrier1.Visible = False
        End If
        ' et on termine en mettant le masque de saisie pour la zone de test
        LeMasque = Val(GetToken(Masques(i), ",", 2))
        
        i = vnb
    End If
Next
End Sub

Private Sub List1_Click()
 MetCond List1.ListIndex
End Sub

