VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Gestion des règles"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mrules As New Collection
Private mvarRulesType As Integer
Private mvarUseRules As Boolean
Private mvarNumberOfActiveRules As Integer ' Nombre de règles actives
Public Property Get NumberOfActiveRules() As Integer
Attribute NumberOfActiveRules.VB_Description = "Renvoie le nombre de règles actives"
    NumberOfActiveRules = mvarNumberOfActiveRules
End Property
' Sauvegarde les règles courantes vers un fichier
Public Function SaveRulesToFile(Filename As String) As Boolean
Attribute SaveRulesToFile.VB_Description = "Sauvegarde les règles dans un fichier"
    Dim RglSav As New Rule
    Dim i As Integer
    Dim vnb As Integer
    Dim SIni As New cInifile
    Dim vtmp As String
    Dim valret As Boolean
    valret = False
    
    If InStr(UCase$(Filename), ".INI") = 0 Then
        Filename = Filename + ".ini"
    End If
    Kill Filename
    vnb = mrules.Count
'    If vnb > 0 Then
        valret = True
        With SIni
            .Path = Filename
            .Section = "General"
            .Key = "NumberOfRules"
            .Value = Trim$(Str$(vnb))
        End With
        With SIni
            .Section = "General"
            .Key = "RuleType"
            .Value = Trim$(Str$(mvarRulesType))
        End With
    
        For i = 1 To vnb
            Set RglSav = mrules.Item(i)
            vtmp = ""
            If RglSav.RuleActive Then
                vtmp = vtmp + "1|"
            Else
                vtmp = vtmp + "0|"
            End If
            vtmp = vtmp & RglSav.RuleCondition & "|"
            vtmp = vtmp & RglSav.RuleDescription & "|"
            vtmp = vtmp & RglSav.RuleIndex & "|"
            vtmp = vtmp & RglSav.RuleName & "|"
            vtmp = vtmp & RglSav.RuleTestValue & "|"
            vtmp = vtmp & RglSav.RuleType & "|"
            With SIni
                .Section = "Rules"
                .Key = "Rule" + Trim$(Str$(i))
                .Value = vtmp
            End With
        Next
'    End If
    SaveRulesToFile = valret
End Function
' Charge les règles depuis un fichier
Public Function LoadRulesFromFile(Filename As String) As Boolean
Attribute LoadRulesFromFile.VB_Description = "Charge des règles depuis un fichier"
    Dim RglSav As New Rule
    Dim i As Integer
    Dim vnb As Integer
    Dim SIni As New cInifile
    Dim vtmp As String
    Dim valret As Boolean
    valret = False
    If InStr(UCase$(Filename), ".INI") = 0 Then
        Filename = Filename + ".ini"
    End If
    ClearRules
    With SIni
        .Path = Filename
        .Section = "General"
        .Key = "NumberOfRules"
        vnb = Val(.Value)
    End With

    For i = 1 To vnb
        valret = True
        With SIni
            .Section = "General"
            .Key = "RuleType"
            mvarRulesType = Val(.Value)
        End With
        
        Set RglSav = Nothing
        With SIni
            .Section = "Rules"
            .Key = "Rule" + Trim$(Str$(i))
            vtmp = .Value
        End With
        If GetToken(vtmp, "|", 1) = "1" Then
            RglSav.RuleActive = True
        Else
            RglSav.RuleActive = False
        End If
        RglSav.RuleCondition = Val(GetToken(vtmp, "|", 2))
        RglSav.RuleDescription = GetToken(vtmp, "|", 3)
        RglSav.RuleIndex = Val(GetToken(vtmp, "|", 4))
        RglSav.RuleName = GetToken(vtmp, "|", 5)
        RglSav.RuleTestValue = GetToken(vtmp, "|", 6)
        RglSav.RuleType = Val(GetToken(vtmp, "|", 7))
        AddRule RglSav
    Next
    LoadRulesFromFile = valret
End Function
' Indique s'il faut utiliser des règles ou pas
Public Property Let UseRules(ByVal vData As Boolean)
Attribute UseRules.VB_Description = "Indique s'il faut utiliser les règles ou pas"
    mvarUseRules = vData
End Property
Public Property Get UseRules() As Boolean
    UseRules = mvarUseRules
End Property
' Indique le type de règle
Public Property Let RulesType(ByVal vData As Integer)
Attribute RulesType.VB_Description = "Indique le type des règles"
    mvarRulesType = vData
End Property
Public Property Get RulesType() As Integer
    RulesType = mvarRulesType
End Property
' Renvoie une règle
Public Function GetRule(RuleNumber As Integer) As Rule
Attribute GetRule.VB_Description = "Renvoie une règle selon son numéro"
    If RuleNumber >= 1 And RuleNumber <= mrules.Count Then
        Set GetRule = mrules.Item(RuleNumber)
    Else
        Set GetRule = Nothing
    End If
End Function
' Efface toutes les règles
Public Sub ClearRules()
Attribute ClearRules.VB_Description = "Supprime toutes les règles"
    Do While mrules.Count > 0
        mrules.Remove 1
    Loop
    mvarNumberOfActiveRules = 0
End Sub
' Renvoie le nombre de règles
Public Function RulesCount() As Integer
Attribute RulesCount.VB_Description = "Renvoie lenombre de règles"
    RulesCount = mrules.Count
End Function
' Supprime une règle à partir de son numéro
Public Sub RemoveRule(RuleNumber As Integer)
Attribute RemoveRule.VB_Description = "Supprime une règle"
    Dim zOneRule As New Rule
    If RuleNumber <= mrules.Count Then
        Set zOneRule = GetRule(RuleNumber)
        If zOneRule.RuleActive = True Then
            mvarNumberOfActiveRules = mvarNumberOfActiveRules - 1
        End If
        Set zOneRule = Nothing
        mrules.Remove RuleNumber
    End If
End Sub
' Ajoute une règle
Public Function AddRule(Regle As Rule) As Integer
Attribute AddRule.VB_Description = "Ajoute une règle"
Dim clé As String
Dim i As Integer
Dim a As Rule
Dim zOneRule As New Rule
    clé = Regle.RuleName
    For i = 1 To mrules.Count
        Set a = mrules(i)
        If a.RuleName = clé Then
            AddRule = 0
            Exit Function
        End If
    Next
    mrules.Add Regle, clé
    Set zOneRule = mrules.Item(mrules.Count)
    zOneRule.RuleIndex = mrules.Count
    If Regle.RuleActive Then
        mvarNumberOfActiveRules = mvarNumberOfActiveRules + 1
    End If
    AddRule = mrules.Count
End Function
' Fait les tests sur un fichier selon les règles courantes
Public Function TestRules(LeFichier As cFile) As Boolean
    Dim i As Integer
    Dim vnb As Integer
    Dim vok As Boolean
    Dim OptTxt As Integer
    Dim UneRegle As New Rule
    Dim vtmp1 As String
    Dim vtmp2 As String
    If LesOptions.RulesOpt1 = 1 Then   ' case insensitive
        OptTxt = vbTextCompare
    Else                    ' case sensitive
        OptTxt = vbBinaryCompare
    End If
    
    vok = True
    vnb = mrules.Count
    For i = 1 To vnb    ' Boucle sur les règles
        Set UneRegle = Nothing
        Set UneRegle = GetRule(i)
        If UneRegle.RuleActive Then
            Select Case UneRegle.RuleType
                Case 0  ' File's size   ************************************************************************************************
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' >
                            If LeFichier.FileSize > UneRegle.RuleTestValue Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' >=
                            If LeFichier.FileSize >= UneRegle.RuleTestValue Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 3  ' <
                            If LeFichier.FileSize < UneRegle.RuleTestValue Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 4  ' <=
                            If LeFichier.FileSize <= UneRegle.RuleTestValue Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 5  ' ==
                            If LeFichier.FileSize = UneRegle.RuleTestValue Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 6  ' !=
                            If LeFichier.FileSize <> UneRegle.RuleTestValue Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 1  ' Prefix    ********************************************************************************************************
                    If LesOptions.RulesOpt1 = 1 Then   ' case insensitive
                        vtmp1 = UCase$(LeFichier.FilePrefix)
                        vtmp2 = UCase$(UneRegle.RuleTestValue)
                    Else                    ' case sensitive
                        vtmp1 = LeFichier.FilePrefix
                        vtmp2 = UneRegle.RuleTestValue
                    End If
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' Begin with
                            If Left$(vtmp1, Len(UneRegle.RuleTestValue)) = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' contains
                            If InStr(1, vtmp1, vtmp2, OptTxt) <> 0 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 3  ' does not contains
                            If InStr(1, vtmp1, vtmp2, OptTxt) = 0 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 4  ' is equal to
                            If vtmp1 = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 5  ' is different from
                            If vtmp1 <> vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 6  ' is ended by
                            If Right$(vtmp1, Len(vtmp2)) = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 2  ' Extension ******************************************************************************************************
                    If LesOptions.RulesOpt1 = 1 Then   ' case insensitive
                        vtmp1 = UCase$(LeFichier.FileExtension)
                        vtmp2 = UCase$(UneRegle.RuleTestValue)
                    Else                    ' case sensitive
                        vtmp1 = LeFichier.FileExtension
                        vtmp2 = UneRegle.RuleTestValue
                    End If
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' Begin with
                            If Left$(vtmp1, Len(UneRegle.RuleTestValue)) = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' contains
                            If InStr(1, vtmp1, vtmp2, OptTxt) <> 0 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 3  ' does not contains
                            If InStr(1, vtmp1, vtmp2, OptTxt) = 0 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 4  ' is equal to
                            If vtmp1 = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 5  ' is different from
                            If vtmp1 <> vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 6  ' is ended by
                            If Right$(vtmp1, Len(vtmp2)) = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 3  ' Path  ***************************************************************************************
                    If LesOptions.RulesOpt1 = 1 Then   ' case insensitive
                        vtmp1 = UCase$(LeFichier.FilePath)
                        vtmp2 = UCase$(UneRegle.RuleTestValue)
                    Else                    ' case sensitive
                        vtmp1 = LeFichier.FilePath
                        vtmp2 = UneRegle.RuleTestValue
                    End If
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' Begin with
                            If Left$(vtmp1, Len(UneRegle.RuleTestValue)) = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' contains
                            If InStr(1, vtmp1, vtmp2, OptTxt) <> 0 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 3  ' does not contains
                            If InStr(1, vtmp1, vtmp2, OptTxt) = 0 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 4  ' is equal to
                            If vtmp1 = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 5  ' is different from
                            If vtmp1 <> vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 6  ' is ended by
                            If Right$(vtmp1, Len(vtmp2)) = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 4  ' Complete filename (path + prefixe + extension) ************************************************************************
                    If LesOptions.RulesOpt1 = 1 Then   ' case insensitive
                        vtmp1 = UCase$(LeFichier.CompleteFileName)
                        vtmp2 = UCase$(UneRegle.RuleTestValue)
                    Else                    ' case sensitive
                        vtmp1 = LeFichier.CompleteFileName
                        vtmp2 = UneRegle.RuleTestValue
                    End If
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' Begin with
                            If Left$(vtmp1, Len(UneRegle.RuleTestValue)) = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' contains
                            If InStr(1, vtmp1, vtmp2, OptTxt) <> 0 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 3  ' does not contains
                            If InStr(1, vtmp1, vtmp2, OptTxt) = 0 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 4  ' is equal to
                            If vtmp1 = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 5  ' is different from
                            If vtmp1 <> vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 6  ' is ended by
                            If Right$(vtmp1, Len(vtmp2)) = vtmp2 Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 5  ' Creation's date ****************************************************************************************
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' >
                            If LeFichier.CreationDate > CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' >=
                            If LeFichier.CreationDate >= CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 3  ' <
                            If LeFichier.CreationDate < CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 4  ' <=
                            If LeFichier.CreationDate <= CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 5  ' ==
                            If LeFichier.CreationDate = CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 6  ' !=
                            If LeFichier.CreationDate <> CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                    
                Case 6  ' Last modification's date *******************************************************************************
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' >
                            If LeFichier.LastAccessDate > CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' >=
                            If LeFichier.LastAccessDate >= CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 3  ' <
                            If LeFichier.LastAccessDate < CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 4  ' <=
                            If LeFichier.LastAccessDate <= CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 5  ' ==
                            If LeFichier.LastAccessDate = CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 6  ' !=
                            If LeFichier.LastAccessDate <> CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 7  ' Last update date ***************************************************************************************
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' >
                            If LeFichier.LastUpdateDate > CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' >=
                            If LeFichier.LastUpdateDate >= CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 3  ' <
                            If LeFichier.LastUpdateDate < CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 4  ' <=
                            If LeFichier.LastUpdateDate <= CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 5  ' ==
                            If LeFichier.LastUpdateDate = CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 6  ' !=
                            If LeFichier.LastUpdateDate <> CDate(UneRegle.RuleTestValue) Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 8  ' Archive attribut ***************************************************************************************
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' is set
                            If LeFichier.AttrArchive = True Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' is not set
                            If LeFichier.AttrArchive = False Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 9 ' Read Only Attribut *************************************************************************************
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' is set
                            If LeFichier.AttrReadOnly = True Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' is not set
                            If LeFichier.AttrReadOnly = False Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 10 ' Hidden Attribut ****************************************************************************************
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' is set
                            If LeFichier.AttrHidden = True Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' is not set
                            If LeFichier.AttrHidden = False Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 11 ' System Attribut ****************************************************************************************
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' is set
                            If LeFichier.AttrSystem = True Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' is not set
                            If LeFichier.AttrSystem = False Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
                
                Case 12 ' TypeOf (File or Folder ?) ******************************************************************************
                    Select Case UneRegle.RuleCondition + 1
                        Case 1  ' Isfile ?
                            If LeFichier.FileType = True Then
                                vok = True
                            Else
                                vok = False
                            End If
                        Case 2  ' IsFolder ?
                            If LeFichier.FileType = False Then
                                vok = True
                            Else
                                vok = False
                            End If
                    End Select
            End Select
        End If
        If mvarRulesType = 0 Then   ' Rename when they satisfy conditions
            If vok = False Then
                TestRules = False
                Exit Function
            End If
        Else                        ' Rename when they don't satisfy conditions
            If vok = True Then
                TestRules = False
                Exit Function
            Else
                vok = True
            End If
        End If
    Next
    TestRules = vok
End Function
