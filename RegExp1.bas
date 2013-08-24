Attribute VB_Name = "RegExp1"
Public Declare Function RegExpIndx Lib "THRegexp.dll" Alias "_OolRegExprVB@16" (ByVal inputString As String, ByVal pattern As String, ByRef subExpIndx() As Integer, ByVal noCase As Long) As Long
Public Const E_UNEXPECTED = &H8000FFFF
Public Const E_OUTOFMEMORY = &H8007000E
Public Const E_INVALIDARG = &H80070057
Public Const S_OK = 0
Public Const S_FAIL = 1

Public Const E_REGEXPNOEXP = &H80006001           ' The Regular Expression string was NULL
Public Const E_REGEXPTOOBIG = &H80006002          ' The Regular Expression was too big
Public Const E_REGEXPTOOMANYPAREN = &H80006003    ' Too many ()'s
Public Const E_REGEXPUNMATCHPAREN = &H80006004    ' Unmatched ()'s
Public Const E_REGEXPSTARPLUSEMPT = &H80006005    ' *+ operand could be empty
Public Const E_REGEXPNESTED = &H80006006          ' nested *?+
Public Const E_REGEXPINVALIDBRKRANGE = &H80006007 ' invalid [] range
Public Const E_REGEXPUNMATCHBRACKET = &H80006008  ' unmatched []
Public Const E_REGEXPOPFOLLOWSNOTHING = &H80006009     ' ?+* follows nothing
Public Const E_REGEXPTRAILINGSLASHS = &H8000600A  ' trailing backslashes

Public Function RegExpErrStr(errcode As Long) As String
 Select Case errcode
    Case E_UNEXPECTED
        RegExpErrStr = "Unexpected Error"
    Case E_OUTOFMEMORY
        RegExpErrStr = "Out of Memory"
    Case E_INVALIDARG
        RegExpErrStr = "Invalid Argument"
    Case E_REGEXPTOOBIG
        RegExpErrStr = "The Regular Expression was too big"
    Case E_REGEXPTOOMANYPAREN
        RegExpErrStr = "Too many ()'s in the Regular Expression"
    Case E_REGEXPUNMATCHPAREN
        RegExpErrStr = "Unmatched ()'s in the Regular Expression"
    Case E_REGEXPSTARPLUSEMPT
        RegExpErrStr = "Possible problem with *+"
    Case E_REGEXPNESTED
        RegExpErrStr = "Nested *?+ in Regular Expression"
    Case E_REGEXPINVALIDBRKRANGE
        RegExpErrStr = "Invalid [] range in Regular Expression"
    Case E_REGEXPUNMATCHBRACKET
        RegExpErrStr = "Unmatched [] in Regular Expression"
    Case E_REGEXPOPFOLLOWSNOTHING
        RegExpErrStr = "?+* follows nothing in Regular Expression"
    Case E_REGEXPTRAILINGSLASHS
        RegExpErrStr = "Trailing backslashes in Regular Expression"
    Case Else
        RegExpErrStr = "Unknown Error"
 End Select
End Function

Public Function RegExpErrCode(errcode As Long) As Long
 Select Case errcode
    Case E_UNEXPECTED
        RegExpErrCode = 1
    Case E_OUTOFMEMORY
        RegExpErrCode = 2
    Case E_INVALIDARG
        RegExpErrCode = 3
    Case E_REGEXPTOOBIG
        RegExpErrCode = 4
    Case E_REGEXPTOOMANYPAREN
        RegExpErrCode = 5
    Case E_REGEXPUNMATCHPAREN
        RegExpErrCode = 6
    Case E_REGEXPSTARPLUSEMPT
        RegExpErrCode = 7
    Case E_REGEXPNESTED
        RegExpErrCode = 8
    Case E_REGEXPINVALIDBRKRANGE
        RegExpErrCode = 9
    Case E_REGEXPUNMATCHBRACKET
        RegExpErrCode = 10
    Case E_REGEXPOPFOLLOWSNOTHING
        RegExpErrCode = 11
    Case E_REGEXPTRAILINGSLASHS
        RegExpErrCode = 12
    Case Else
        RegExpErrCode = 13
 End Select
End Function

Public Function RegSub(inputString As String, patternString As String, substr As String, returnString As String, Optional pos As Variant) As Boolean
    If IsMissing(pos) Then
        pos = 0
    End If
    Dim subPos As Integer
    subPos = CInt(pos)
    Dim indx(2, 1) As Integer
    Dim fromEnd As Integer
    Dim res As Long
    res = RegExpIndx(inputString, patternString, indx(), 1)
    If FAILED(res) Then
        GoTo ErrorHandler
    End If
    If res = S_OK Then
        If (indx(0, i) < 1 Or indx(1, i) < 1) Then
            returnString = substr
        Else
            fromEnd = Len(inputString) - indx(1, subPos) + 1
            returnString = left(inputString, indx(0, subPos) - 1) & substr & right(inputString, fromEnd)
        End If
        RegSub = True
    Else
        RegSub = False
    End If
    Exit Function

ErrorHandler:
    ' Raise an exception
    Err.Raise vbObjectError + RegExpErrCode(res), "RegExp", RegExpErrStr(res)
    RegSub = False
End Function
Public Function RegExp(inputString As String, patternString As String, ParamArray subExpresions() As Variant) As Boolean
    Dim indx(2, 1) As Integer
    Dim i As Integer, argNum As Integer, indxnum As Integer
    Dim res As Long
    
    res = RegExpIndx(inputString, patternString, indx(), 1)
    If FAILED(res) Then
        GoTo ErrorHandler
    End If
    If res = S_OK Then
        i = 0
        argNum = UBound(subExpresions, 1)
        indxnum = UBound(indx, 2)
        While i < argNum + 1 And i < indxnum + 1
            If (indx(0, i) < 1 Or indx(1, i) < 1) Then
                subExpresions(i) = ""
            Else
               subExpresions(i) = Mid(inputString, indx(0, i), indx(1, i) - indx(0, i))
            End If
            i = i + 1
        Wend
        RegExp = True
    Else
        RegExp = False
    End If
    Exit Function

ErrorHandler:
    ' Raise an exception
    Err.Raise vbObjectError + RegExpErrCode(res), "RegExp", RegExpErrStr(res), App.HelpFile, RegExpErrCode(res)
    RegExp = False
End Function
Public Function FAILED(hresult As Long) As Boolean
    If hresult < 0 Then
        FAILED = True
    Else
        FAILED = False
    End If
End Function
Public Function SUCCEDED(hresult As Long) As Boolean
    If hresult >= 0 Then
        SUCCEDED = True
    Else
        SUCCEDED = False
    End If
End Function
