VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A class module to make search and replace on strings and characters"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarSourceString As String
Private mvarSearchString As String
Private mvarReplaceString As String
Private mvarReplaceAll As Boolean
Private mvarMatchCase As Boolean
Private mvarNumberOfSubst As Integer
Private mvarSearchFromLeft As Boolean
Private mvarSearchFrom As Integer
Private mvarSearchTo As Integer
Private mvarSearchCharacters As String
Private mvarReplaceCharacters As String
Private mvarUseRegExp As Integer
Public Property Let UseRegExp(ByVal vData As Integer)
    mvarUseRegExp = vData
End Property
Public Property Get UseRegExp() As Integer
    UseRegExp = mvarUseRegExp
End Property
Public Function BeginReplaceCharacters() As String
 If mvarSearchCharacters = "" Then
  BeginReplaceCharacters = mvarSourceString
  Exit Function
 End If
 If mvarSourceString = "" Then Exit Function
 Dim i As Integer, l As Integer, cH As String, extrait As String
 Dim Pos As Integer
 l = Len(mvarSourceString)
 cH = ""
 For i = 1 To l
  extrait = Mid$(mvarSourceString, i, 1)
  Pos = InStr(mvarSearchCharacters, extrait)
  If Pos <> 0 Then ' Remplacement des caractères
   cH = cH + Mid$(mvarReplaceCharacters, Pos, 1)
  Else
   cH = cH + extrait
  End If
 Next
 BeginReplaceCharacters = cH
End Function
Public Property Let ReplaceCharacters(ByVal vData As String)
Attribute ReplaceCharacters.VB_Description = "Contains the list of all characters to replace"
    mvarReplaceCharacters = vData
End Property
Public Property Get ReplaceCharacters() As String
    ReplaceCharacters = mvarReplaceCharacters
End Property
Public Property Let SearchCharacters(ByVal vData As String)
Attribute SearchCharacters.VB_Description = "Contains the list of all characters to search for"
    mvarSearchCharacters = vData
End Property
Public Property Get SearchCharacters() As String
    SearchCharacters = mvarSearchCharacters
End Property
Public Function BeginSearchAndReplace() As String
Attribute BeginSearchAndReplace.VB_Description = "Begin the search and replace action"
Attribute BeginSearchAndReplace.VB_UserMemId = 0
If mvarSearchString = "" Then
 BeginSearchAndReplace = mvarSourceString
 Exit Function
End If
If mvarSourceString = "" Then Exit Function
Dim t As String
Dim sOrig As String, zexact As Integer, Final As String, tempo As String
Dim tempo1 As String, tempo2 As String
Dim tmp As Integer
Dim ModeComp As Integer
If mvarUseRegExp = 1 Then ' On utilise des expressions régulières
   Dim Match As Boolean
   Dim str1 As String
   On Error Resume Next
   Dim RegCase As Integer
   Dim vnbremp As Integer
   If mvarMatchCase Then
        RegCase = 1
   Else
        RegCase = 0
   End If
    If mvarReplaceAll Then
        vnbremp = 9999
   Else
        vnbremp = mvarNumberOfSubst
   End If
   Match = RegSub(mvarSourceString, mvarSearchString, mvarReplaceString, str1, RegCase, vnbremp, mvarSearchFrom, mvarSearchTo)
   If Match = True Then
        BeginSearchAndReplace = str1 ' Si ca a marché, on renvoie l'expression retournée par la librairie
    Else
        BeginSearchAndReplace = mvarSourceString ' Si ca n'a pas marché, on renvoie la chaine originale
    End If
   Exit Function
End If

sOrig = mvarSourceString
If mvarSearchFromLeft = False Then ' Search from right
 mvarSourceString = StrReverse(mvarSourceString)
 mvarSearchString = StrReverse(mvarSearchString)
 mvarReplaceString = StrReverse(mvarReplaceString)
End If

If mvarSearchFrom <> 0 And mvarSearchTo <> 0 Then ' From ... to ...
 If mvarSearchTo < mvarSearchFrom Then
  BeginSearchAndReplace = ""
  Exit Function
 End If
 mvarSourceString = Mid$(mvarSourceString, mvarSearchFrom, (mvarSearchTo - mvarSearchFrom) + 1)
End If

If mvarMatchCase = True Then
 zexact = 0
Else
 zexact = 1
End If

If mvarReplaceAll Then
    tmp = -1
Else
    tmp = mvarNumberOfSubst
End If

If mvarMatchCase = True Then
    ModeComp = vbBinaryCompare
Else
    ModeComp = vbTextCompare
End If

t = Replace(mvarSourceString, mvarSearchString, mvarReplaceString, 1, tmp, ModeComp)

Final = t ' & mvarSourceString

If mvarSearchFromLeft = False Then ' Search from right
 Final = StrReverse(Final)
 ' Faut pas oublier de tout remettre dans le bon sens ... !
 mvarSourceString = StrReverse(mvarSourceString)
 mvarSearchString = StrReverse(mvarSearchString)
 mvarReplaceString = StrReverse(mvarReplaceString)
End If

If mvarSearchFrom <> 0 And mvarSearchTo <> 0 Then ' From ... to ...
 If mvarSearchFromLeft = True Then
  Final = Left$(sOrig, mvarSearchFrom - 1) + Final + Mid$(sOrig, mvarSearchTo + 1)
 Else
  tempo = sOrig
  tempo = StrReverse(tempo)
  tempo1 = Left$(tempo, mvarSearchFrom - 1)
  tempo2 = Mid$(tempo, mvarSearchTo + 1)
  tempo1 = StrReverse(tempo1)
  tempo2 = StrReverse(tempo2)
  Final = tempo2 + Final + tempo1
 End If
End If

BeginSearchAndReplace = Final
End Function
Public Property Let NumberOfSubst(ByVal vData As Integer)
    mvarNumberOfSubst = vData
End Property
Public Property Get NumberOfSubst() As Integer
    NumberOfSubst = mvarNumberOfSubst
End Property
Public Property Let SearchTo(ByVal vData As Integer)
Attribute SearchTo.VB_Description = "Indicate the ending position of the search in the original string, optional"
    mvarSearchTo = vData
End Property
Public Property Get SearchTo() As Integer
    SearchTo = mvarSearchTo
End Property
Public Property Let SearchFrom(ByVal vData As Integer)
Attribute SearchFrom.VB_Description = "Indicate the begin position of the search in the original string, optional"
    mvarSearchFrom = vData
End Property
Public Property Get SearchFrom() As Integer
    SearchFrom = mvarSearchFrom
End Property
Public Property Let SearchFromleft(ByVal vData As Boolean)
Attribute SearchFromleft.VB_Description = "If True, the search will begin from the left of the string, else the search will begin from the right"
    mvarSearchFromLeft = vData
End Property
Public Property Get SearchFromleft() As Boolean
    SearchFromleft = mvarSearchFromLeft
End Property
Public Property Let MatchCase(ByVal vData As Boolean)
Attribute MatchCase.VB_Description = "If True, the search and replace will be base on exact case"
    mvarMatchCase = vData
End Property
Public Property Get MatchCase() As Boolean
    MatchCase = mvarMatchCase
End Property
Public Property Let ReplaceAll(ByVal vData As Boolean)
Attribute ReplaceAll.VB_Description = "If true, the class will replace all occurences of the search string with the replace string, else this will jsut replace the first occurence"
    mvarReplaceAll = vData
End Property
Public Property Get ReplaceAll() As Boolean
    ReplaceAll = mvarReplaceAll
End Property
Public Property Let ReplaceString(ByVal vData As String)
Attribute ReplaceString.VB_Description = "Contains the string to replace"
    mvarReplaceString = vData
End Property
Public Property Get ReplaceString() As String
    ReplaceString = mvarReplaceString
End Property
Public Property Let SearchString(ByVal vData As String)
Attribute SearchString.VB_Description = "Contains the string to search in the source string"
    mvarSearchString = vData
End Property
Public Property Get SearchString() As String
    SearchString = mvarSearchString
End Property
Public Property Let SourceString(ByVal vData As String)
Attribute SourceString.VB_Description = "Contains the source string, the sting where to search for"
    mvarSourceString = vData
End Property
Private Sub Class_Initialize()
 mvarReplaceAll = True
 mvarMatchCase = False
 mvarSearchFromLeft = True
 mvarSearchFrom = 0
 mvarSearchTo = 0
 mvarUseRegExp = 0
 mvarNumberOfSubst = 1
End Sub
Public Sub ResetSearch()
 mvarReplaceAll = True
 mvarMatchCase = False
 mvarSearchFromLeft = True
 mvarSearchFrom = 0
 mvarSearchTo = 0
 mvarSourceString = ""
 mvarSearchString = ""
 mvarReplaceString = ""
 mvarReplaceCharacters = ""
 mvarSearchCharacters = ""
 mvarUseRegExp = 0
 mvarNumberOfSubst = 1
End Sub
