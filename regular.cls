VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegularExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
' Name: RegularExpression
' Description:This is a class module that performs regular expression searches in a string.
' By: visual basic prof. edition
' Inputs:Use the Init method to initialize to a specific regular expression (which will be
' precompiled), then use Match to check if a string contains such a substring.
' Returns:None
' Assumes:
' b) The syntax for range (e.g. [a-z]) is the same as for operator Like.
' c) Not calling Init, or passing an empty pattern, will result in an "Illegal function call" error.
' Side Effects:None
'****************************************************************

Option Explicit
'     'PRIVATE
'     '? = edOptional; + = edMulti; * = edOptional or edMulti
Private Enum RegExpStateTypes
edOptional = 65536
edMulti = 131072
edModifierMask = edOptional Or edMulti
edCharacter = 0
edBracketed = 262144 'for example, [a-z]
edAny = 524288
End Enum


Private Type StateStack
       State As Long
       Posi As Long
       MinPosi As Long
End Type

Private mStack() As StateStack
Private mCompiled() As Long
Private mNStates As Long
Private mPattern As String
Private mAnchorBeginning As Boolean
Private mAnchorEnd As Boolean
Private mMinLength As Long

Private Sub AddState(ByVal Flags As Long, ByVal CharOrPosi As Long)


              If mNStates = UBound(mCompiled) Then
                     ReDim Preserve mCompiled(1 To mNStates + 10) As Long
              End If

       mNStates = mNStates + 1
       mCompiled(mNStates) = CharOrPosi Or Flags
End Sub


Public Sub Init(RegExp As String)

       Dim StackSize As Long, Posi As Long, EndPosi As Long
       '     'Initialize member variables
       mPattern = RegExp
       mNStates = 0
       mMinLength = 0
       ReDim mCompiled(1 To 10) As Long
       Posi = 1
EndPosi = Len(RegExp)


If Left(mPattern, 1) = "^" Then
       Posi = Posi + 1
       mAnchorBeginning = True
End If


If Right(mPattern, 1) = "$" And Right(mPattern, 2) <> "\$" Then
EndPosi = EndPosi - 1

mAnchorEnd = True
End If


Do Until Posi > EndPosi
Select Case Mid$(mPattern, Posi, 1)
Case "."
AddState edAny, 0
Posi = Posi + 1
Case "\"
AddState edCharacter, Asc(Mid$(mPattern, Posi + 1, 1))
Posi = Posi + 2
Case "["
AddState edBracketed, Posi
Posi = RangeParse(Posi)

If Posi = -1 Then Err.Raise 5
       Case Else
       AddState edCharacter, Asc(Mid$(mPattern, Posi, 1))
       Posi = Posi + 1
End Select

'     'check for modifiers (?, +, *)
Select Case Mid$(mPattern, Posi, 1)
Case "?"
mCompiled(mNStates) = mCompiled(mNStates) Or edOptional
StackSize = StackSize + 1
Posi = Posi + 1
Case "+"
mCompiled(mNStates) = mCompiled(mNStates) Or edMulti
StackSize = StackSize + 1
Posi = Posi + 1
mMinLength = mMinLength + 1
Case "*"
mCompiled(mNStates) = mCompiled(mNStates) Or edMulti Or edOptional
StackSize = StackSize + 1
Posi = Posi + 1
Case Else
mMinLength = mMinLength + 1
End Select

Loop

'     'Minimize wasted memory by dimensioning exact arrays
ReDim Preserve mCompiled(1 To mNStates) As Long
ReDim mStack(1 To StackSize) As StateStack
End Sub


Public Function Match(ByRef FromX As Long, ByRef ToX As Long, Text As String) As Boolean
Dim Match As Boolean
Dim CurState As Long
Dim State As Long
Dim SP As Long
Dim LenText As Long

       If mNStates = 0 Then Err.Raise 5
              LenText = Len(Text)

                     For FromX = FromX To IIf(mAnchorBeginning, 1, LenText - mMinLength)
                            ToX = FromX
                            State = 1
                            SP = 0

                                   Do

                                                 If State > mNStates Then

                                                               If (Not mAnchorEnd) Or (ToX > LenText) Then
                                                                      'ToX is pointing the first character PAST the matched string
                                                                      '
                                                                      ToX = ToX - 1
                                                                      MatchRight = True
                                                                      Exit Function
                                                               End If

                                                 End If

                                          GoSub MatchState

                                                 If Match Then

                                                               If CurState And edModifierMask Then
                                                                      '     'create a new item in the backtrack stack
                                                                      SP = SP + 1
                                                                      mStack(SP).MinPosi = IIf(CurState And edOptional, ToX, ToX + 1)

                                                                             If (CurState And (edAny Or edMulti)) = (edAny Or edMulti) Then
                                                                                    'When matching .* and .+, we don't need to check the whole s
                                                                                    '     tring
                                                                                    ToX = LenText + 1
                                                                             ElseIf CurState And edMulti Then
                                                                                    '     '+ or *, try to get as far as possible

                                                                                           Do
                                                                                                  ToX = ToX + 1
                                                                                                  GoSub MatchState
                                                                                           Loop Until Not Match

                                                                             Else
                                                                                    '     '?, you only have to look one character ahead
                                                                                    ToX = ToX + 1
                                                                             End If

                                                                      State = State + 1
                                                                      mStack(SP).Posi = ToX
                                                                      mStack(SP).State = State
                                                               Else
                                                                      '     'no +, *, nor ?, just advance to the next state
                                                                      ToX = ToX + 1
                                                                      State = State + 1
                                                               End If

                                                 ElseIf CurState And edOptional Then
                                                        '     'not matched, but it was optional... no problem
                                                        State = State + 1
                                                 Else
                                                        '     'backtrack - find the next usable item in the stack

                                                               For SP = SP To 1 Step -1

                                                                             If mStack(SP).Posi > mStack(SP).MinPosi Then Exit For
                                                                             Next SP


                                                                             If SP = 0 Then Exit Do
                                                                                    mStack(SP).Posi = mStack(SP).Posi - 1
                                                                                    ToX = mStack(SP).Posi
                                                                                    State = mStack(SP).State
                                                                             End If

                                                               Loop

                                                 Next FromX

                                          Exit Function
MatchState:
                                          CurState = mCompiled(State)

                                                 If ToX > LenText Then
                                                        Match = False
                                                 ElseIf CurState And edAny Then
                                                        Match = True
                                                 ElseIf CurState And edBracketed Then
                                                        Match = RangeMatch(CurState And 65535, Mid$(Text, ToX, 1))
                                                 Else
                                                        Match = (CurState And 65535) = Asc(Mid$(Text, ToX, 1))
                                                 End If

Return
End Function


Private Function RangeMatch(Posi As Long, ch As String) As Boolean
       RangeMatch = ch Like Mid$(mPattern, Posi, InStr(Posi, mPattern, "]") - Posi + 1)
End Function

'Return the end of the range (e.g. [a-z]) starting at positi
'     on Posi.
'     'Return -1 if the regular expression is not well formed.

Private Function RangeParse(Posi As Long) As Long

       Dim EndPosi As Long
EndPosi = InStr(Posi, mPattern, "]")

'     'Try using operator Like and check if an error occurs
On Error Resume Next

If "a" Like Mid(mPattern, Posi, EndPosi - Posi + 1) Then:

              If Err Then
                     RangeParse = -1
                     Err.Clear
              Else
                     RangeParse = EndPosi + 1
              End If

End Function
