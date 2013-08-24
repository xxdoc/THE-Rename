VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParseString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************
' clsParseString
'********************************************
'Input: String,Delimiter
'Ouput: Token array of parsed string
'
'Movefisrt = Moves to first token
'Movelast = Moves to last token
'MoveNext = Move to next token
'MovePrevious = Move to previous token
'BOS = Beginning of token string
'EOS = End of token string
'Count = Returns number of tokens
'********************************************
Option Explicit
Option Compare Text
Private Tokens As Variant
Private Positions As Variant
Private TokensFound As Integer
Private Index As Integer

'Retrieve and array of tokens delimited by a list of tokens
'This function takes the following arguments
'1 - The string to be searched (this argument is passed byval to preserve it in the calling function)
'2 - A REFERENCE to an array fo strings in which the tokens are stored
'3 - A list of token delimitors.
'    If this string is ommited the delimitors are set to " " and vbTab
'The function returns the number of tokens found

Public Sub ParseDelimitedString(ByVal SearchString As String, Optional TokenList As String = " " & vbTab)
Dim StringLength As Integer
'Dim TokensFound As Integer
Dim StartPos As Integer
Dim machaine As String

'Initialize local variables
'ParseDelimitedString = 0
TokensFound = 0
Index = 0

StartPos = skipDelimitor(SearchString, TokenList)
If StartPos = 0 Then
  'Empty string or just demilitors.
  Tokens = Array()
  Positions = Array()
  Exit Sub
ElseIf StartPos > 1 Then
  'String starts with delimitors. Skip initial delimitors
  SearchString = Right$(SearchString, Len(SearchString) - StartPos + 1)
End If
StringLength = Len(SearchString)

If StringLength = 0 Then
  'Empty string
  Tokens = Array()
  Positions = Array()
  Exit Sub
End If

'make Tokens somewhat large to avoid redim-ming to often
ReDim Tokens(50)
ReDim Positions(50)



machaine = SearchString
Positions(TokensFound) = (Len(machaine) - Len(SearchString)) + 1

Do While getNextToken(SearchString, Tokens(TokensFound), TokenList)
  TokensFound = TokensFound + 1
  If TokensFound > UBound(Tokens) Then
    'Running out of space.
    ReDim Preserve Tokens(TokensFound * 2)
    ReDim Preserve Positions(TokensFound * 2)
  End If
  Positions(TokensFound) = (Len(machaine) - Len(SearchString)) + 1
Loop
ReDim Preserve Tokens(TokensFound)
ReDim Preserve Positions(TokensFound)
TokensFound = TokensFound + 1

End Sub

'Get the next token from a string and remove token+delimitors from the string
'This function takes the following arguments:
'1 - The string to be searched. THIS STRING WILL BE CHANGED BY THIS FUNCTION
'2 - The string to contain the token
'3 - A list of token delimitors.
'    If this string is ommited the delimitors are set to " " and vbTab
'The function returns True if another token can befound in the string and
'return False if this is the last token in the string

Private Function getNextToken(ByRef SearchString As String, ByRef Token As Variant, Optional TokenList As String = " " & vbTab) As Boolean
Dim StartPos As Integer
Dim DelimitorPos As Integer

DelimitorPos = FindDelimitor(SearchString, TokenList)
If DelimitorPos = 0 Then
  Token = SearchString
  getNextToken = False
  Exit Function
Else
  'Found a delimitor.
  'Store string in Tokens
  Token = Left$(SearchString, DelimitorPos - 1)
  SearchString = Right$(SearchString, Len(SearchString) - DelimitorPos + 1)
  StartPos = skipDelimitor(SearchString, TokenList)
  If StartPos > 1 Then
    'Skip delimitor characters
    SearchString = Right$(SearchString, Len(SearchString) - StartPos + 1)
    getNextToken = True
    Exit Function
  Else
    If StartPos = 0 Then
      'only delimitors left
    Else
      'No non-delimitor characters left in the string
      Token = SearchString
    End If
    getNextToken = False
    Exit Function
  End If
End If
End Function


'Find the position of the first token delimitor character in a string.
'This function takes the following arguments:
'1 - The string to be searched
'2 - A list of token delimitors.
'    If this string is ommited the delimitors are set to " " and vbTab
'The function returns the position of the found character or
'zero if the character was not found

Private Function FindDelimitor(SearchString As String, Optional TokenList As String = " " & vbTab) As Integer

Dim StringLength As Integer
Dim Counter As Integer
Dim vtmp As String

StringLength = Len(SearchString)
For Counter = 1 To StringLength
  vtmp = Mid$(SearchString, Counter, 1)
  If InStr(TokenList, vtmp) > 0 Then
    FindDelimitor = Counter
    Exit Function
  End If
Next
FindDelimitor = 0
End Function


'Find the position of the first character, which is NOT a token delimitor character, in a string.
'This function takes the following arguments:
'1 - The string to be searched
'2 - A list of token delimitors.
'    If this string is ommited the delimitors are set to " " and vbTab
'The function returns the position of the found character or
'zero if the character was not found

Private Function skipDelimitor(SearchString As String, Optional TokenList As String = " " & vbTab) As Integer

Dim StringLength As Integer
Dim Counter As Integer
Dim vtmp As String

StringLength = Len(SearchString)
For Counter = 1 To StringLength
  vtmp = Mid$(SearchString, Counter, 1)
  If InStr(TokenList, vtmp) = 0 Then
    skipDelimitor = Counter
    Exit Function
  End If
Next
skipDelimitor = 0
End Function

Public Property Get BOS() As Boolean
'Set Begining of String
    If Index = 0 Then
        BOS = True
    Else
        BOS = False
    End If
End Property

Public Property Get EOS() As Boolean
'Set End of String
    If Index = TokensFound Then
        EOS = True
    Else
        EOS = False
    End If
End Property

Private Sub Class_Initialize()
'Initialize the Class
    Index = -1
End Sub

Public Property Get Count() As Integer
'Return token count
    Count = TokensFound
End Property

Public Property Get Pos() As Variant
'Return the current token's position
    If Index > -1 And Index < TokensFound Then
        Pos = Positions(Index)
    ElseIf Index > -1 Then
        Index = TokensFound - 1
        Pos = Positions(Index)
    End If
End Property

Public Property Get Token() As Variant
'Return the current token
    If Index > -1 And Index < TokensFound Then
        Token = Tokens(Index)
    ElseIf Index > -1 Then
        Index = TokensFound - 1
        Token = Tokens(Index)
    End If
End Property

Public Sub MovePrevious()
'Move to the Previous Token
    If Index > 0 Then
        Index = Index - 1
    End If
End Sub

Public Sub MoveNext()
'Move to the Next Token
    If Index < TokensFound Then
        Index = Index + 1
    End If
End Sub

Public Sub MoveFirst()
'Move to the First Token
    If Index <> -1 Then
        Index = 0
    End If
End Sub

Public Sub MoveLast()
'Move to the Last Token
    If Index <> -1 Then
        Index = TokensFound - 1
    End If
End Sub

Public Property Get TokenX(ReqIndex As Integer) As Variant
'Return the specified Token
    If Index > -1 Then
        If ReqIndex <= UBound(Tokens) Then
            TokenX = Tokens(ReqIndex)
        Else
            TokenX = ""
        End If
    End If
End Property
