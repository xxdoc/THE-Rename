VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAFM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private sFilename As String
Private mvarCreationDate As String
Private mvarUniqueID As String
Private mvarVMusage As String
Private mvarCopyright As String
Private mvarWeight As String
Private mvarNotice As String
Private mvarMetricsSets As String
Private mvarFullName As String
Private mvarFontVersion As String
Private mvarFontName As String
Private mvarFontMetricsVersion As String
Private mvarFamilyName As String
Private mvarEncodingScheme As String
Private mvarCharacterSet As String
Private mvarFileName As String
Public Sub FillTagsList(lv As ListView)
    Dim itmX As ListItem
    lv.ListItems.Clear
    Set itmX = lv.ListItems.Add(, , "CharacterSet")
    itmX.SubItems(1) = mvarCharacterSet
    Set itmX = lv.ListItems.Add(, , "Copyright")
    itmX.SubItems(1) = mvarCopyright
    Set itmX = lv.ListItems.Add(, , "CreationDate")
    itmX.SubItems(1) = mvarCreationDate
    Set itmX = lv.ListItems.Add(, , "EncodingScheme")
    itmX.SubItems(1) = mvarEncodingScheme
    Set itmX = lv.ListItems.Add(, , "FamilyName")
    itmX.SubItems(1) = mvarFamilyName
    Set itmX = lv.ListItems.Add(, , "FontMetricsVersion")
    itmX.SubItems(1) = mvarFontMetricsVersion
    Set itmX = lv.ListItems.Add(, , "FontName")
    itmX.SubItems(1) = mvarFontName
    Set itmX = lv.ListItems.Add(, , "FontVersion")
    itmX.SubItems(1) = mvarFontVersion
    Set itmX = lv.ListItems.Add(, , "FullName")
    itmX.SubItems(1) = mvarFullName
    Set itmX = lv.ListItems.Add(, , "MetricsSets")
    itmX.SubItems(1) = mvarMetricsSets
    Set itmX = lv.ListItems.Add(, , "Notice")
    itmX.SubItems(1) = mvarNotice
    Set itmX = lv.ListItems.Add(, , "UniqueID")
    itmX.SubItems(1) = mvarUniqueID
    Set itmX = lv.ListItems.Add(, , "VMusage")
    itmX.SubItems(1) = mvarVMusage
    Set itmX = lv.ListItems.Add(, , "Weight")
    itmX.SubItems(1) = mvarWeight
    ResizeLvMp3
End Sub
Public Sub Clear()
    mvarCreationDate = ""
    mvarUniqueID = ""
    mvarVMusage = ""
    mvarCopyright = ""
    mvarWeight = ""
    mvarNotice = ""
    mvarMetricsSets = ""
    mvarFullName = ""
    mvarFontVersion = ""
    mvarFontName = ""
    mvarFontMetricsVersion = ""
    mvarFamilyName = ""
    mvarEncodingScheme = ""
    mvarCharacterSet = ""
    mvarFileName = ""
End Sub
Public Property Get CharacterSet() As String
    CharacterSet = mvarCharacterSet
End Property
Public Property Get EncodingScheme() As String
    EncodingScheme = mvarEncodingScheme
End Property
Public Property Get FamilyName() As String
    FamilyName = mvarFamilyName
End Property
Public Property Get FontMetricsVersion() As String
    FontMetricsVersion = mvarFontMetricsVersion
End Property
Public Property Get FontName() As String
    FontName = mvarFontName
End Property
Public Property Get FontVersion() As String
    FontVersion = mvarFontVersion
End Property
Public Property Get FullName() As String
    FullName = mvarFullName
End Property
Public Property Get MetricsSets() As String
    MetricsSets = mvarMetricsSets
End Property
Public Property Get Notice() As String
    Notice = mvarNotice
End Property
Public Property Get Weight() As String
    Weight = mvarWeight
End Property
Public Property Get Copyright() As String
    Copyright = mvarCopyright
End Property
Public Property Get VMusage() As String
    VMusage = mvarVMusage
End Property
Public Property Get UniqueID() As String
    UniqueID = mvarUniqueID
End Property
Public Property Get CreationDate() As String
    CreationDate = mvarCreationDate
End Property
Public Sub GetAFMInfos(ByVal sPassFileName As String)
    On Error GoTo ErrGen
    Dim vin As Integer
    Dim vligne As String
    Dim vtmp As String
    Dim TheData As String
    Dim vok As Boolean
    Dim VnbTags As Integer
    
   If sPassFileName = sFilename Then
        Exit Sub
    End If
    sFilename = sPassFileName
    
    vin = FreeFile
    vok = True
    Clear
    Open sPassFileName For Input As #vin
    
    Line Input #vin, vligne
    While vok And Not EOF(vin)
        vligne = Trim$(vligne)
        vtmp = Trim$(UCase$(GetToken(vligne, " ", 1)))
        TheData = GetToken(vligne, " ", 2)
        Select Case vtmp
            Case "STARTFONTMETRICS"
                mvarFontMetricsVersion = LTrim$(Mid$(vligne, InStr(vligne, "StartFontMetrics") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "WEIGHT"
                mvarWeight = LTrim$(Mid$(vligne, InStr(vligne, "Weight") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "NOTICE"
                mvarNotice = LTrim$(Mid$(vligne, InStr(vligne, "Notice") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "METRICSSETS"
                mvarMetricsSets = LTrim$(Mid$(vligne, InStr(vligne, "MetricsSets") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "FULLNAME"
                mvarFullName = LTrim$(Mid$(vligne, InStr(vligne, "FullName") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "VERSION"
                mvarFontVersion = LTrim$(Mid$(vligne, InStr(vligne, "Version") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "FONTNAME"
                mvarFontName = LTrim$(Mid$(vligne, InStr(vligne, "FontName") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "FAMILYNAME"
                mvarFamilyName = LTrim$(Mid$(vligne, InStr(vligne, "FamilyName") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "ENCODINGSCHEME"
                mvarEncodingScheme = LTrim$(Mid$(vligne, InStr(vligne, "EncodingScheme") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "CHARACTERSET"
                mvarCharacterSet = LTrim$(Mid$(vligne, InStr(vligne, "CharacterSet") + Len(vtmp) + 1))
                VnbTags = VnbTags + 1
            Case "COMMENT"
                Select Case TheData
                    Case "Copyright"
                        mvarCopyright = LTrim$(Mid$(vligne, InStr(vligne, "Copyright") + 9))
                        VnbTags = VnbTags + 1
                    Case "Creation"
                        mvarCreationDate = LTrim$(Mid$(vligne, InStr(vligne, "Creation") + 8))
                        VnbTags = VnbTags + 1
                    Case "UniqueID"
                        mvarUniqueID = LTrim$(Mid$(vligne, InStr(vligne, "UniqueID") + 8))
                        VnbTags = VnbTags + 1
                    Case "VMusage"
                        mvarVMusage = LTrim$(Mid$(vligne, InStr(vligne, "VMusage") + 7))
                        VnbTags = VnbTags + 1
                End Select
        End Select
        If UCase$(Left$(vligne, 5)) = "START" And Left$(vligne, 16) <> "StartFontMetrics" Then
            vok = False
        End If
        If VnbTags = 14 Then
            vok = False
        End If
        Line Input #vin, vligne
    Wend
    Close #vin
    Exit Sub
    
ErrGen:
 ErreurGrave "GetAFMInfos"
 Exit Sub
End Sub
