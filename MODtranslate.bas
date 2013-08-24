Attribute VB_Name = "MODTranslate"
'===============================================================================
'   send comments to
'   fredjust@hotmail.com
'   http://fred.just.free.fr/
'   http://go.to/fredjust
'===============================================================================

Option Explicit


Dim tempo As String
Dim Obj As Control
Dim ObjIndex As Long

Global gstrFileLNG As String
Global translation As New Collection


'Fonctions de lectures du fichier .INI
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'==================================================================================
'   Save all caption, tooltiptext for each object in a from
'==================================================================================
Public Sub SaveAllName(ByVal aForm As Form, ByVal Filename As String)

    On Error Resume Next
    
    WriteIniFile Filename, aForm.Name, "Form.Caption", aForm.Caption

    For Each Obj In aForm

        tempo = ""
        Err.Clear
        ObjIndex = Obj.Index
        
        If Err.Number <> 0 Then ' if the object is not indexed
            
            tempo = "" 'don't delete
            tempo = Obj.Caption

            If tempo <> "" Then
                ' write Caption of the object
                WriteIniFile Filename, aForm.Name, CStr(Obj.Name) & ".Caption", CStr(Obj.Caption)
            End If
            
            
            tempo = "" 'don't delete
            tempo = Obj.ToolTipText

            If tempo <> "" Then
                ' write ToolTipText of the object
                WriteIniFile Filename, aForm.Name, Obj.Name & ".ToolTipText", Obj.ToolTipText
            End If

        Else        ' if the object is indexed
            
            tempo = "" 'don't delete
            tempo = Obj.Caption

            If tempo <> "" Then
                ' write Caption of the object
                WriteIniFile Filename, aForm.Name, Obj.Name & "(" & Obj.Index & ").Caption", Obj.Caption
            End If

            tempo = "" 'don't delete
            tempo = Obj.ToolTipText
            
            If tempo <> "" Then
                ' write ToolTipText of the object
                WriteIniFile Filename, aForm.Name, Obj.Name & "(" & Obj.Index & ").ToolTipText", Obj.ToolTipText
            End If
        End If
    Next
End Sub


'==================================================================================
'   Load and change Caption and toolTipText
'==================================================================================
Public Sub LoadAllName(aForm As Form, ByVal Filename As String)

    On Error Resume Next
    gstrFileLNG = Filename
    aForm.Caption = ReadIniFile(Filename, aForm.Name, aForm.Caption, aForm.Caption)
    For Each Obj In aForm
        Err.Clear
        ObjIndex = Obj.Index
        'if the objet is indexed
        If Err.Number = 0 Then
            ' change caption of object
            Obj.Caption = ReadIniFile(Filename, aForm.Name, Obj.Name & "(" & Obj.Index & ").Caption", Obj.Caption)
            ' change tooltiptext of object
            Obj.ToolTipText = ReadIniFile(Filename, aForm.Name, Obj.Name & "(" & Obj.Index & ").ToolTipText", Obj.ToolTipText)
        Else
            ' change caption of object
            Obj.Caption = ReadIniFile(Filename, aForm.Name, Obj.Name & ".Caption", Obj.Caption)
            ' change tooltiptext of object
            Obj.ToolTipText = ReadIniFile(Filename, aForm.Name, Obj.Name & ".ToolTipText", Obj.ToolTipText)
        End If
    Next
End Sub

'===============================================================================
'   return the translation of a sentence msg1,msg2,msg3
'===============================================================================
Public Function Translate(msgX As String) As String
   Translate = ReadIniFile(gstrFileLNG, "MSG", msgX, "")
End Function


'===============================================================================
'   Save the message
'===============================================================================
Public Sub SaveMessage(ByVal Filename As String)
Dim i As Long
Dim phrase
    ' write all sentence
    i = 1
    For Each phrase In translation
        WriteIniFile Filename, "MSG", "msg" & CStr(i), CStr(phrase)
        i = i + 1
    Next
End Sub



'===============================================================================
'
'===============================================================================
Function ReadIniFile(ByVal strIniFile As String, ByVal strSection As String, _
        ByVal strKey As String, Optional ByVal strDefault As String) As String
        
Dim szBuffer As String
Dim iLen As Integer

     szBuffer = String(255, Chr(0))
     iLen = GetPrivateProfileString(strSection, strKey, strDefault, szBuffer, Len(szBuffer), strIniFile)
     ReadIniFile = Left$(szBuffer, iLen)
     
End Function

'===============================================================================
'
'===============================================================================
Function WriteIniFile(ByVal strIniFile As String, strSection As String, strKey As String, v As String) As Long
    WriteIniFile = WritePrivateProfileString(strSection, ByVal strKey, ByVal v, strIniFile)
End Function

