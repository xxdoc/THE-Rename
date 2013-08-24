Attribute VB_Name = "RegExp"
Option Explicit
Private Declare Function RxInterf Lib "therename.dll" (ByVal Pattern As String, ByVal Source As String, ByVal CorrTot As String, ByVal Substr As String, ByVal erreur As String, ByVal OptExt As Long, ByVal iCase As Long, ByVal SyntaxToUse As Long) As Integer
Private Declare Function PcreInterf Lib "therename.dll" (ByVal Pattern As String, ByVal Source As String, ByVal CorrTot As String, ByVal Substr As String, ByVal erreur As String, ByVal OptExt As Long, ByVal iCase As Long, ByVal OptAnchor As Long, ByVal OptDollarEnd As Long, ByVal OptDotAll As Long, ByVal OptExtra As Long, ByVal OptUnGreedy As Long) As Integer

Public Function RegSub(InputString As String, PatternString As String, Substitution As String, ReturnString As String, MatchCase As Integer, NbRemplacements As Integer, FromPosition As Integer, ToPosition As Integer) As Boolean
Dim Pattern As String       ' Contient la pattern (l'expression régulière)
Dim Source As String        ' Contient la chaine source (la chaine de recherche)
Dim CorrTot As String       ' Contient la correspondance totale
Dim Substr As String        ' Contient la liste des sous chaines
Dim erreur As String        ' Message d'erreur renvoyé par la DLL
Dim Option1 As Long         ' Extended
Dim Option2 As Long         ' Match Case
Dim Option3 As Long         ' Anchor
Dim Option4 As Long         ' Dollar End
Dim Option5 As Long         ' Dot All
Dim Option6 As Long         ' Extra
Dim Option7 As Long         ' UnGreedy
Dim retour As Integer       ' Valeur de retour de la fonction de la DLL
Dim Rempl As String         ' Chaine de remplacement
Dim NbRempl As Integer      ' Nombre de remplacements à faire
Dim NbRemplFait As Integer  ' Nombre de remplacements faits
Dim ok As Boolean           ' Pour boucler
Dim CorrTotDeb As Integer
Dim CorrTotFin As Integer
Dim VnbBackRef As Integer   ' Contient le nombre de backreference de l'expression de remplacement
Dim vtmp As String
Dim vdest As String
Dim vmax As Integer         ' Contient le nombre de parenthèses dans l'expression source
Dim vmaxbackref As Integer  ' Contient le numéro de la backreference de plus haut niveau
Dim ExprBack As String

If LesOptions.BackRefNotation = 0 Then
    ExprBack = "\"
Else
    ExprBack = "$"
End If

Rempl = Substitution
NbRempl = NbRemplacements
Pattern = PatternString + Chr$(0)
Source = InputString + Chr$(0)
CorrTot = Space$(256) + Chr$(0)
Substr = Space$(256) + Chr$(0)
erreur = Space$(256) + Chr$(0)

' Mise en place des options
Option1 = LesOptions.RegExOpt1     ' Extended
Option2 = MatchCase     ' Match Case
Option3 = LesOptions.PCRE2         ' Anchor
Option4 = LesOptions.PCRE3         ' Dollar End
Option5 = LesOptions.PCRE4         ' Dot all
Option6 = LesOptions.PCRE1         ' Extra
Option7 = LesOptions.PCRE5         ' Ungreedy

' controle des parenthèses et des backreferences
VnbBackRef = NbBackRef(Rempl)
If VnbBackRef > 0 Then
    vmax = NbParenth(Pattern)
    vmaxbackref = MaxBackRef(Rempl)
    If vmaxbackref > vmax Then
        If vmax = 0 Then
            MsgBox "Error, you are calling the back reference named " & ExprBack & vmaxbackref & " but there are no parenthesis in the pattern string !", vbOKOnly, "Error !"
            RegSub = False
            Exit Function
        Else
            MsgBox "Error, you are calling the back reference named " & ExprBack & vmaxbackref & " but there are only " & vmax & " parenthesis in the pattern string !", vbOKOnly, "Error !"
            RegSub = False
            Exit Function
        End If
    End If
End If

If FromPosition <> 0 And ToPosition <> 0 Then
    Source = Mid$(InputString, FromPosition, ToPosition - FromPosition) + Chr$(0)
End If

If LesOptions.RegExpEngine = 0 Then    ' Utilisation de RX
    retour = RxInterf(Pattern, Source, CorrTot, Substr, erreur, Option1, Option2, LesOptions.RxSyntax)
Else    ' Utilisation de PCRE
    retour = PcreInterf(Pattern, Source, CorrTot, Substr, erreur, Option1, Option2, Option3, Option4, Option5, Option6, Option7)
End If

If retour = -1 Then
    MsgBox erreur, , "Error while compiling pattern"
    Exit Function
End If
If retour = -2 Then
'    MsgBox "Pattern does not match"
    ReturnString = InputString
    RegSub = True
    Exit Function
End If
If retour = -3 Then
    MsgBox "RegEx run out of memory"
    RegSub = False
    Exit Function
End If
ok = True

vtmp = Left$(Source, Len(Source) - 1)
While (ok)
    CorrTotDeb = Val(GetToken(CorrTot, "-", 1)) + 1
    CorrTotFin = Val(GetToken(CorrTot, "-", 2))
    vdest = vdest + Remplace(vtmp, Rempl, CorrTotDeb, CorrTotFin, Substr, retour, Mid$(vtmp, CorrTotDeb, (CorrTotFin - CorrTotDeb) + 1))
    vtmp = Mid$(vtmp, CorrTotFin + 1)
    NbRemplFait = NbRemplFait + 1
    If (NbRemplFait >= NbRempl) Then
        ok = False
    End If
    If Len(vtmp) = 0 Then
        ok = False
    End If
    If (ok) Then
        If LesOptions.RegExpEngine = 0 Then    ' Utilisation de RX
            retour = RxInterf(Pattern, vtmp, CorrTot, Substr, erreur, Option1, Option2, LesOptions.RxSyntax)
        Else
            retour = PcreInterf(Pattern, vtmp, CorrTot, Substr, erreur, Option1, Option2, Option3, Option4, Option5, Option6, Option7)
        End If
        If retour <= 0 Then
            ok = False
        End If
    End If
Wend
vdest = vdest + vtmp

If FromPosition <> 0 And ToPosition <> 0 Then
    ReturnString = Mid$(InputString, 1, FromPosition - 1) + vdest + Mid$(InputString, ToPosition + 1)
Else
    ReturnString = vdest
End If
RegSub = True
End Function

' ******************************************************************************************
' Remplace le contenu d'une chaine par une autre (par rapport à des positions)
' ******************************************************************************************
Private Function Remplace(Source As String, Rempl As String, debut As Integer, fin As Integer, Substr As String, VnbExpr As Integer, CorrTot As String) As String
    Dim i As Integer
    Dim PosDeb As Integer
    Dim PosFin As Integer
    Dim vtmp As String
    Dim ExprBack As String
    If LesOptions.BackRefNotation = 0 Then
        ExprBack = "\"
    Else
        ExprBack = "$"
    End If
    If VnbExpr = 1 Then
        Remplace = Left$(Source, debut - 1) + Rempl  '  + mid$(Source, fin + 1))
    Else    ' Expression avec backreferences
        Remplace = Left$(Source, debut - 1) + Rempl  '  + mid$(Source, fin + 1))
        If NbBackRef(Rempl) > 0 Then
            For i = 1 To VnbExpr - 1
                vtmp = GetToken(Substr, "|", i)
                PosDeb = Val(GetToken(vtmp, "-", 1))
                PosFin = Val(GetToken(vtmp, "-", 2))
                Remplace = Replace(Remplace, ExprBack + Trim$(Str$(i)), Mid$(Source, PosDeb + 1, PosFin - PosDeb))
            Next
            ' Remplacement de \0
            Remplace = Replace(Remplace, ExprBack + "0", CorrTot)
            For i = 0 To 9  ' code pour supprimer les backreferences qui pourraient rester et qui ne sont pas affectées à des chaines (par exemple on a appelé \3 mais ce la ne coprrespond à aucune chaine)
                Remplace = Replace(Remplace, ExprBack + Trim$(Str$(i)), "")
            Next
        End If
    End If
End Function
' ******************************************************************************************
' Renvoie le nombre de backreference d'une expression
' ******************************************************************************************
Private Function NbBackRef(lachaine As String) As Integer
    Dim vret As Integer
    Dim i As Integer
    Dim vnb As Integer
    Dim ExprBack As String
    If LesOptions.BackRefNotation = 0 Then
        ExprBack = "\"
    Else
        ExprBack = "$"
    End If
    
    vnb = Len(lachaine)
    vret = 0
    For i = 1 To vnb
        If Mid$(lachaine, i, 1) = ExprBack Then
            If i <> vnb Then
                If IsNumeric(Mid$(lachaine, i + 1, 1)) Then
                    vret = vret + 1
                End If
            End If
        End If
    Next
    NbBackRef = vret
End Function

' ******************************************************************************************
' Renvoie la valeur de la nième backreference
' ******************************************************************************************
Private Function ValBackRef(lachaine As String, n As Integer) As Integer
    Dim vret As Integer
    Dim vret2 As Integer
    Dim i As Integer
    Dim vnb As Integer
    Dim ExprBack As String
    If LesOptions.BackRefNotation = 0 Then
        ExprBack = "\"
    Else
        ExprBack = "$"
    End If
    vnb = Len(lachaine)
    vret = 0
    For i = 1 To vnb
        If Mid$(lachaine, i, 1) = ExprBack Then
            If i <> vnb Then
                If IsNumeric(Mid$(lachaine, i + 1, 1)) Then
                    vret = vret + 1
                    If vret = n Then
                        vret2 = Val(Mid$(lachaine, i + 1, 1))
                    End If
                End If
            End If
        End If
    Next
    ValBackRef = vret2
End Function

' ******************************************************************************************
' Renvoie la valeur de la backref la plus élevée
' ******************************************************************************************
Private Function MaxBackRef(lachaine As String) As Integer
    Dim vret As Integer
    Dim i As Integer
    Dim vnb As Integer
    Dim ExprBack As String
    If LesOptions.BackRefNotation = 0 Then
        ExprBack = "\"
    Else
        ExprBack = "$"
    End If
    vnb = Len(lachaine)
    vret = 0
    For i = 1 To vnb
        If Mid$(lachaine, i, 1) = ExprBack Then
            If i <> vnb Then
                If IsNumeric(Mid$(lachaine, i + 1, 1)) Then
                    If Val(Mid$(lachaine, i + 1, 1)) > vret Then
                        vret = Val(Mid$(lachaine, i + 1, 1))
                    End If
                End If
            End If
        End If
    Next
    MaxBackRef = vret
End Function
' ******************************************************************************************
' Renvoie le nombre de parenthèses (ouvrantes et fermantes) d'une expression
' ******************************************************************************************
Private Function NbParenth(lachaine As String) As Integer
    Dim vnb As Integer
    Dim i As Integer
    Dim ouvert As Boolean
    Dim vret As Integer
    Dim extr As String
    vnb = Len(lachaine)
    
    vret = 0
    For i = 1 To vnb
        extr = Mid$(lachaine, i, 1)
        If extr = "(" Then
            ouvert = True
        End If
        If extr = ")" Then
            If ouvert Then
                vret = vret + 1
                ouvert = False
            End If
        End If
    Next
    NbParenth = vret
End Function

