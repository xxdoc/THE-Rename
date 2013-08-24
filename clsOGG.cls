VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOGG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe de gestion des fichiers Ogg"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Rule"
Option Explicit
Private Declare Sub CloseOggFile Lib "renogg.dll" ()
Private Declare Function OpenOggFile Lib "renogg.dll" (ByVal File As String, ByVal Length As String, ByVal Playtime As String) As Long
Private Declare Function GetTag Lib "renogg.dll" (ByVal TagNumber As Long, ByVal TxtTag As String) As String
Private Declare Function GetVendor Lib "renogg.dll" (ByVal Vendor As String) As String
Private Declare Function GetBitstreamChannel Lib "renogg.dll" (ByVal NbChannels As String, ByVal Bitstream As String) As String
Private Declare Function GetComplInf Lib "renogg.dll" (ByVal NumInfo As Long, ByVal Info As String) As String

Private SepTags As String       ' Séparateur à utiliser entre les tags, lu dans la base de registre ou valeur par défaut codée en dur dans le programme
Private Tags() As String            ' Les 2, Tag+valeur
Private KnowTags(18) As String
Private mvarDefaultTitleToUse As String
Private mvarDefaultAlbumToUse As String
Private mvarDefaultGenreToUse As String
Private mvarDefaultArtistToUse As String
Private mvarPlaceWhereToPut As Integer
Private mvarRule As String
Private mvarFileName As String
Private mvarLength As String
Private mvarPlayTime As String
Private mvarVendor As String
Private mvarBitstream As String
Private mvarChannels As String
Private mvarAverageBitrate As String
Private mvarNominalBitrate As String
Private mvarUpperBitrate As String
Private mvarLowerBitrate As String
Private mvarEncoderVersion As String
Private mvarSerialNumber As String
Private mvarNumberOfTags As Integer
Private sFilename As String
Public Function GetOggInfos(sPassFileName As String, Optional Tout As Boolean = True) As String
    On Error GoTo ErrGen
    Dim Inter As String
    
    If UCase$(Suffixe(sPassFileName)) <> "OGG" Then
        ClearTags
        GetOggInfos = ""
        Exit Function
    End If
    
    
    If sPassFileName <> sFilename Then
        OuvreOgg sPassFileName
    End If
    sFilename = sPassFileName
    
    If Tout = False Then    ' la méthode n'a pas été appelée par une fenêtre mais par une fonction du Free Form
        Exit Function
    End If

    Inter = mvarRule
    If Title() = "" Then
        Inter = Replace(Inter, "%Title%", mvarDefaultTitleToUse, , , vbTextCompare)
    Else
        Inter = Replace(Inter, "%Title%", Title(), , , vbTextCompare)
    End If
    

    If Artist() = "" Then
        Inter = Replace(Inter, "%Artist%", mvarDefaultArtistToUse, , , vbTextCompare)
    Else
        Inter = Replace(Inter, "%Artist%", Artist(), , , vbTextCompare)
    End If
    

    If Album() = "" Then
        Inter = Replace(Inter, "%Album%", mvarDefaultAlbumToUse, , , vbTextCompare)
    Else
        Inter = Replace(Inter, "%Album%", Album(), , , vbTextCompare)
    End If

    If Genre() = "" Then
        Inter = Replace(Inter, "%Genre%", mvarDefaultGenreToUse, , , vbTextCompare)
    Else
        Inter = Replace(Inter, "%Genre%", Genre(), , , vbTextCompare)
    End If

    Inter = Replace(Inter, "%AverageBitrate%", AverageBitrate(), , , vbTextCompare)
    Inter = Replace(Inter, "%Channels%", Channels(), , , vbTextCompare)
    Inter = Replace(Inter, "%Comment%", Comment(), , , vbTextCompare)
    Inter = Replace(Inter, "%CopyRight%", Copyright(), , , vbTextCompare)
    Inter = Replace(Inter, "%Date%", SaDate(), , , vbTextCompare)
    Inter = Replace(Inter, "%Description%", Description(), , , vbTextCompare)
    Inter = Replace(Inter, "%EncoderVersion%", EncoderVersion(), , , vbTextCompare)
    Inter = Replace(Inter, "%ISRC%", ISRC(), , , vbTextCompare)
    Inter = Replace(Inter, "%Length%", Length(), , , vbTextCompare)
    Inter = Replace(Inter, "%Location%", Location(), , , vbTextCompare)
    Inter = Replace(Inter, "%LowerBitrate%", LowerBitrate(), , , vbTextCompare)
    Inter = Replace(Inter, "%NominalBitrate%", NominalBitrate(), , , vbTextCompare)
    Inter = Replace(Inter, "%NumberOfTags%", NumberOfTags(), , , vbTextCompare)
    Inter = Replace(Inter, "%Organization%", Organization(), , , vbTextCompare)
    Inter = Replace(Inter, "%Playtime%", Playtime(), , , vbTextCompare)
    Inter = Replace(Inter, "%SampleRate%", SampleRate(), , , vbTextCompare)
    Inter = Replace(Inter, "%SerialNumber%", SerialNumber(), , , vbTextCompare)
    Inter = Replace(Inter, "%TotalTracks%", TotalTracks(), , , vbTextCompare)
    Inter = Replace(Inter, "%TrackNumber%", TrackNumber(), , , vbTextCompare)
    Inter = Replace(Inter, "%UpperBitrate%", UpperBitrate(), , , vbTextCompare)
    Inter = Replace(Inter, "%Vendor%", Vendor(), , , vbTextCompare)
    Inter = Replace(Inter, "%Version%", Version(), , , vbTextCompare)
    Inter = Replace(Inter, "%Composer%", Composer(), , , vbTextCompare)
    Inter = Replace(Inter, "%Conductor%", Conductor(), , , vbTextCompare)
    Inter = Replace(Inter, "%Ensemble%", Ensemble(), , , vbTextCompare)
    Inter = Replace(Inter, "%Performer%", Performer(), , , vbTextCompare)
    GetOggInfos = Inter
    Exit Function
    
ErrGen:
 ErreurGrave "GetOggInfos"
 Exit Function
End Function
Public Property Get NumberOfTags() As Integer
    NumberOfTags = mvarNumberOfTags
End Property
Public Property Get SerialNumber() As String
    SerialNumber = mvarSerialNumber
End Property
Public Property Get EncoderVersion() As String
    EncoderVersion = mvarEncoderVersion
End Property
Public Property Get LowerBitrate() As String
    LowerBitrate = mvarLowerBitrate
End Property
Public Property Get UpperBitrate() As String
    UpperBitrate = mvarUpperBitrate
End Property
Public Property Get NominalBitrate() As String
    NominalBitrate = mvarNominalBitrate
End Property
Public Property Get AverageBitrate() As String
    AverageBitrate = mvarAverageBitrate
End Property
Public Property Get Channels() As String
    Channels = mvarChannels
End Property
Public Property Get SampleRate() As String
    SampleRate = mvarBitstream
End Property
Public Property Get Vendor() As String
    Vendor = mvarVendor
End Property
Public Property Get Playtime() As String
    Playtime = mvarPlayTime
End Property
Public Property Get Length() As String
    Length = mvarLength
End Property
Public Property Get ISRC() As String
    ISRC = GetTagValue("ISRC")
End Property
' Renvoie la valeur d'un tag selon sa position
Public Property Get GetTagByPosition(TagPos As Integer, TagFormat As Integer, Separator As String) As String
    Dim i As Integer        ' Pour la boucle
    Dim LeTag As String     ' Nom du tag
    Dim ValTag As String    ' Valeur du tag
    Dim valret As String    ' Valeur de retour de la fonction
    
    valret = ""
    For i = 1 To mvarNumberOfTags
        LeTag = GetToken(Tags(i), "=", 1)
        If i = TagPos Then  ' c'est le tag recherché
            ValTag = GetToken(Tags(i), "=", 2)
            Select Case TagFormat
                Case 1  ' tag=valeur
                    valret = LeTag + Separator + ValTag
                Case 2  ' tag seulement
                    valret = LeTag
                Case 3  ' valeur seulement
                    valret = ValTag
            End Select
            i = mvarNumberOfTags
        End If
    Next
    GetTagByPosition = valret
End Property
' Renvoie la valeur d'un tag selon le nom du tag passé en paramètre
Public Property Get GetTagByName(TagName As String, TagFormat As Integer, Separator As String) As String
    Dim i As Integer        ' Pour la boucle
    Dim LeTag As String     ' Nom du tag
    Dim ValTag As String    ' Valeur du tag
    Dim valret As String    ' Valeur de retour de la fonction
    Dim TagComp As String   ' Tag passé par l'utilisateur et modifié pour les effectuer les comparaisons
    TagComp = UCase$(Trim$(TagName))
    
    valret = ""
    For i = 1 To mvarNumberOfTags
        LeTag = GetToken(Tags(i), "=", 1)
        If UCase$(Trim$(LeTag)) = TagComp Then  ' C'est le tag recherché
            ValTag = GetToken(Tags(i), "=", 2)
            Select Case LesOptions.OggOpt1
                Case 0  ' Use the first tag
                    i = mvarNumberOfTags
                    Select Case TagFormat
                        Case 1  ' tag=valeur
                            valret = LeTag + Separator + ValTag + SepTags
                        Case 2  ' tag seulement
                            valret = LeTag + SepTags
                        Case 3  ' valeur seulement
                            valret = ValTag + SepTags
                    End Select
                    
                Case 1  ' Use the last
                    Select Case TagFormat
                        Case 1  ' tag=valeur
                            valret = LeTag + Separator + ValTag + SepTags
                        Case 2  ' tag seulement
                            valret = LeTag + SepTags
                        Case 3  ' valeur seulement
                            valret = ValTag + SepTags
                    End Select
                    
                Case 2  ' Use the longest
                    Select Case TagFormat
                        Case 1  ' tag=valeur
                            If Len(ValTag) + Len(LeTag) > valret Then
                                valret = LeTag + Separator + ValTag + SepTags
                            End If
                        Case 2  ' tag seulement
                            If Len(LeTag) > valret Then
                                valret = LeTag + SepTags
                            End If
                            
                        Case 3  ' valeur seulement
                            If Len(ValTag) > valret Then
                                valret = ValTag + SepTags
                            End If
                    End Select
                    
                Case 3  ' Use filled
                    If Len(ValTag) > 0 Then
                        Select Case TagFormat
                            Case 1  ' tag=valeur
                                valret = valret + LeTag + Separator + ValTag + SepTags
                            Case 2  ' tag seulement
                                valret = valret + LeTag + SepTags
                                
                            Case 3  ' valeur seulement
                                valret = valret + ValTag + SepTags
                        End Select
                    End If
                    
                Case 4  ' Combine tags
                    Select Case TagFormat
                        Case 1  ' tag=valeur
                            If Len(valret) > 0 Then
                                valret = valret & LesOptions.OggOpt2 + LeTag + Separator + ValTag + SepTags
                            Else
                                valret = valret + LeTag + Separator + ValTag + SepTags
                            End If
                        Case 2  ' tag seulement
                            If Len(valret) > 0 Then
                                valret = valret & LesOptions.OggOpt2 + LeTag + SepTags
                            Else
                                valret = valret + LeTag + SepTags
                            End If
                        Case 3  ' valeur seulement
                            If Len(valret) > 0 Then
                                valret = valret & LesOptions.OggOpt2 + ValTag + SepTags
                            Else
                                valret = valret + ValTag + SepTags
                            End If
                    End Select
            End Select
        End If
    Next
    
    If Len(valret) > 1 Then
        valret = Left$(valret, Len(valret) - Len(SepTags))
    End If
    GetTagByName = valret
End Property
' Renvoie, sous la forme d'une chaine, tous les tags d'un fichier
Public Property Get GetAllTags(TagFormat As Integer, Separator As String) As String
    Dim i As Integer
    Dim LeTag As String     ' Nom du tag
    Dim ValTag As String    ' Valeur du tag
    Dim valret As String    ' Valeur de retour de la fonction
   
    valret = ""
    For i = 1 To mvarNumberOfTags
        LeTag = GetToken(Tags(i), "=", 1)
        ValTag = GetToken(Tags(i), "=", 2)
        Select Case TagFormat
            Case 1  ' tag=valeur
                valret = valret + LeTag + Separator + ValTag + SepTags
            Case 2  ' tag seulement
                valret = valret + LeTag + SepTags
            Case 3  ' valeur seulement
                valret = valret + ValTag + SepTags
        End Select
    Next
    If Len(valret) > 1 Then
        valret = Left$(valret, Len(valret) - Len(SepTags))
    End If
    GetAllTags = valret
End Property
' Renvoie, sous la forme d'une chaine, les tags non reconnus par le programme
Public Property Get GetUnknowTags(TagFormat As Integer, Separator As String) As String
    Dim i As Integer        ' Pour la boucle
    Dim LeTag As String     ' Nom du tag
    Dim ValTag As String    ' Valeur du tag
    Dim valret As String    ' Valeur de retour de la fonction
    
    valret = ""
    For i = 1 To mvarNumberOfTags
        LeTag = GetToken(Tags(i), "=", 1)
        If Not IsKnowTag(LeTag) Then    ' Le tag n'est pas reconnu par le système, on peut l'ajouter à la liste
            ValTag = GetToken(Tags(i), "=", 2)
            Select Case TagFormat
                Case 1  ' tag=valeur
                    valret = valret + LeTag + Separator + ValTag + SepTags
                Case 2  ' tag seulement
                    valret = valret + LeTag + SepTags
                Case 3  ' valeur seulement
                    valret = valret + ValTag + SepTags
            End Select
        End If
    Next
    If Len(valret) > 1 Then
        valret = Left$(valret, Len(valret) - Len(SepTags))
    End If
    GetUnknowTags = valret
End Property
Public Property Get Composer() As String
    Composer = GetTagValue("COMPOSER")
End Property
Public Property Get Conductor() As String
    Conductor = GetTagValue("CONDUCTOR")
End Property
Public Property Get Performer() As String
    Performer = GetTagValue("PERFORMER")
End Property
Public Property Get Ensemble() As String
    Ensemble = GetTagValue("ENSEMBLE")
End Property
Public Property Get SaDate() As String
    SaDate = GetTagValue("DATE")
End Property
Public Property Get Copyright() As String
    Copyright = GetTagValue("COPYRIGHT")
End Property
Public Property Get Location() As String
    Location = GetTagValue("LOCATION")
End Property
Public Property Get Description() As String
    Description = GetTagValue("DESCRIPTION")
End Property
Public Property Get Organization() As String
    Organization = GetTagValue("ORGANIZATION")
End Property
Public Property Get TotalTracks() As String
    Dim chtempo As String
    If LesOptions.Mp3VqfOpt4 <> 0 Then
        chtempo = String$(LesOptions.Mp3VqfOpt4, "0")
        TotalTracks = Format$(Val(GetTagValue("TOTALTRACKS")), chtempo)
    End If
End Property
Public Property Get TrackNumber() As String
    Dim chtempo As String
    If LesOptions.Mp3VqfOpt4 <> 0 Then
        chtempo = String$(LesOptions.Mp3VqfOpt4, "0")
        TrackNumber = Format$(Val(GetTagValue("TRACKNUMBER")), chtempo)
    End If
End Property
Public Property Get Version() As String
    Version = GetTagValue("VERSION")
End Property
Public Property Get Comment() As String
    Comment = GetTagValue("COMMENT")
End Property
Public Property Get Album() As String
    Album = GetTagValue("ALBUM")
End Property
Public Property Get Genre() As String
    Genre = GetTagValue("GENRE")
End Property
Public Property Get Artist() As String
    Artist = GetTagValue("ARTIST")
End Property
Public Property Get Title() As String
    Title = GetTagValue("TITLE")
End Property
Public Property Let Rule(ByVal vData As String)
    mvarRule = vData
End Property
Public Property Get Rule() As String
    Rule = mvarRule
End Property
Public Property Let PlaceWhereToPut(ByVal vData As Integer)
    mvarPlaceWhereToPut = vData
End Property
Public Property Get PlaceWhereToPut() As Integer
    PlaceWhereToPut = mvarPlaceWhereToPut
End Property
Public Property Let DefaultArtistToUse(ByVal vData As String)
    mvarDefaultArtistToUse = vData
End Property
Public Property Get DefaultArtistToUse() As String
    DefaultArtistToUse = mvarDefaultArtistToUse
End Property
Public Property Let DefaultGenreToUse(ByVal vData As String)
    mvarDefaultGenreToUse = vData
End Property
Public Property Get DefaultGenreToUse() As String
    DefaultGenreToUse = mvarDefaultGenreToUse
End Property
Public Property Let DefaultAlbumToUse(ByVal vData As String)
    mvarDefaultAlbumToUse = vData
End Property
Public Property Get DefaultAlbumToUse() As String
    DefaultAlbumToUse = mvarDefaultAlbumToUse
End Property
Public Property Let DefaultTitleToUse(ByVal vData As String)
    mvarDefaultTitleToUse = vData
End Property
Public Property Get DefaultTitleToUse() As String
    DefaultTitleToUse = mvarDefaultTitleToUse
End Property
' Procédure chargée d'ouvrir le fichier Ogg voulu avec les options sélectionnées
Private Sub OuvreOgg(NomFichier As String)
    On Error GoTo ErrGen
    Dim vret As Long
    Dim Length As String
    Dim Playtime As String
    Dim vtmp As String
    Dim vtmp2 As String
    Dim vtmp3 As String
    Dim i As Long
    ClearTags
    Length = String$(1024, " ") + Chr$(0)
    Playtime = String$(1024, " ") + Chr$(0)
    vret = OpenOggFile(NomFichier, Length, Playtime)
    If vret <= 0 Then
        ReDim Tags(0)
        mvarNumberOfTags = 0
        Exit Sub
    End If
    
    mvarNumberOfTags = vret
    ReDim Tags(mvarNumberOfTags)
    mvarPlayTime = Menage(Playtime)
    mvarLength = Menage(Length)
    
    For i = 1 To mvarNumberOfTags
        vtmp = String$(256, " ") + Chr$(0)
        vtmp = GetTag(i, vtmp)
        Tags(i) = Menage(vtmp)
    Next
    
    vtmp = String$(256, " ") + Chr$(0)
    vtmp = GetVendor(vtmp)
    mvarVendor = Menage(vtmp)

    vtmp = String$(256, " ") + Chr$(0)    ' Bitstream
    vtmp2 = String$(256, " ") + Chr$(0)   ' Channels
    vtmp3 = GetBitstreamChannel(vtmp2, vtmp)
    mvarBitstream = Menage(vtmp)
    mvarChannels = Menage(vtmp2)
    
    vtmp = String$(256, " ") + Chr$(0)
    vtmp = GetComplInf(5, vtmp)
    mvarAverageBitrate = Menage(vtmp)
    
    vtmp = String$(256, " ") + Chr$(0)
    vtmp = GetComplInf(3, vtmp)
    mvarNominalBitrate = Menage(vtmp)
    
    vtmp = String$(256, " ") + Chr$(0)
    vtmp = GetComplInf(1, vtmp)
    mvarEncoderVersion = Menage(vtmp)
    
    vtmp = String$(256, " ") + Chr$(0)
    vtmp = GetComplInf(2, vtmp)
    mvarUpperBitrate = Menage(vtmp)
    
    vtmp = String$(256, " ") + Chr$(0)
    vtmp = GetComplInf(4, vtmp)
    mvarLowerBitrate = Menage(vtmp)
    
    vtmp = String(256, " ") + Chr$(0)
    vtmp = GetComplInf(6, vtmp)
    mvarSerialNumber = Menage(vtmp)
    CloseOggFile
    Exit Sub
    
ErrGen:
 ErreurGrave "OuvreOgg"
 Exit Sub
End Sub

Private Function Menage(chaine As String) As String
    Dim vtmp As String
    vtmp = Trim$(Replace(chaine, Chr$(0), "", , , vbBinaryCompare))
    vtmp = Replace(vtmp, Chr$(9), "", , , vbBinaryCompare)
    Menage = vtmp
End Function

Private Function GetTagValue(TagName As String) As String
    Dim i As Integer
    Dim LeTag As String
    Dim SaValeur As String
    Dim TagRech As String
    Dim valret As String
    
    If mvarNumberOfTags <= 0 Then
        GetTagValue = ""
        Exit Function
    End If
    
    TagRech = Trim$(UCase$(TagName))
    valret = ""
    For i = 1 To mvarNumberOfTags
        LeTag = GetToken(Tags(i), "=", 1)
        If Trim$(UCase$(LeTag)) = TagRech Then
            SaValeur = Menage(GetToken(Tags(i), "=", 2))
            Select Case LesOptions.OggOpt1
                Case 0  ' Use the first tag
                    valret = SaValeur
                    i = mvarNumberOfTags
                Case 1  ' Use the last
                    valret = SaValeur
                Case 2  ' Use the longest
                    If Len(SaValeur) > Len(valret) Then
                        valret = SaValeur
                    End If
                Case 3  ' Use filled
                    If Len(SaValeur) > 0 Then
                        valret = SaValeur
                    End If
                Case 4  ' Combine tags
                    If Len(valret) > 0 Then
                        valret = valret & LesOptions.OggOpt2 & SaValeur
                    Else
                        valret = SaValeur
                    End If
            End Select
        End If
    Next
    
    If LesOptions.Mp3VqfOpt1 = 1 Then  ' Separate words
        valret = ExtractWords(valret)
    End If
    
    If LesOptions.Mp3VqfOpt2 = 1 Then  ' Remove multiple spaces
        valret = RemoveMultipleSpacing(valret)
    End If
    
    valret = ChangeMP3Case(valret)  ' Changement de casse des tags
    
    GetTagValue = valret
End Function

' Permet de remplir le treeview de prévisualisation des tags de la fenêtre principale
Public Sub FillTagsList(lv As ListView)
    Dim itmX As ListItem
    Dim i As Integer
    If mvarNumberOfTags <= 0 Then
        Exit Sub
    End If
    lv.ListItems.Clear
    
    For i = 1 To mvarNumberOfTags
        Set itmX = lv.ListItems.Add(, , Trim$(Str$(i)) + " - " + GetToken(Tags(i), "=", 1))
        itmX.SubItems(1) = Menage(GetToken(Tags(i), "=", 2))
    Next
    
    Set itmX = lv.ListItems.Add(, , " Serial Number")
    itmX.SubItems(1) = mvarSerialNumber
    Set itmX = lv.ListItems.Add(, , " Encoder Version")
    itmX.SubItems(1) = mvarEncoderVersion
    Set itmX = lv.ListItems.Add(, , " Lower Bitrate")
    itmX.SubItems(1) = mvarLowerBitrate
    Set itmX = lv.ListItems.Add(, , " Upper Bitrate")
    itmX.SubItems(1) = mvarUpperBitrate
    Set itmX = lv.ListItems.Add(, , " Nominal Bitrate")
    itmX.SubItems(1) = mvarNominalBitrate
    Set itmX = lv.ListItems.Add(, , " Average Bitrate")
    itmX.SubItems(1) = mvarAverageBitrate
    Set itmX = lv.ListItems.Add(, , " Channels")
    itmX.SubItems(1) = mvarChannels
    Set itmX = lv.ListItems.Add(, , " Sample Rate")
    itmX.SubItems(1) = mvarBitstream
    Set itmX = lv.ListItems.Add(, , " Vendor")
    itmX.SubItems(1) = mvarVendor
    Set itmX = lv.ListItems.Add(, , " Playtime")
    itmX.SubItems(1) = mvarPlayTime
    Set itmX = lv.ListItems.Add(, , " Length")
    itmX.SubItems(1) = mvarLength
    
    ResizeLvMp3
End Sub
Private Sub Class_Initialize()
    ClearTags
    ' Le séparateur à utiliser entre les tags est lu à partir de la base de registres.
    ' Si la clé n'est pas trouvée, on utilise ", "
    SepTags = GetSetting("THERename", "Param", "OggTagSeparator", ", ")
End Sub

Private Sub ClearTags()
    mvarFileName = ""
    mvarRule = ""
    mvarPlaceWhereToPut = 0
    mvarDefaultArtistToUse = ""
    mvarDefaultGenreToUse = ""
    mvarDefaultAlbumToUse = ""
    mvarDefaultTitleToUse = ""
    mvarLength = ""
    mvarPlayTime = ""
    mvarVendor = ""
    mvarBitstream = ""
    mvarChannels = ""
    mvarAverageBitrate = ""
    mvarNominalBitrate = ""
    mvarUpperBitrate = ""
    mvarLowerBitrate = ""
    mvarEncoderVersion = ""
    mvarSerialNumber = ""
    mvarNumberOfTags = 0
    KnowTags(1) = "Vendor"
    KnowTags(2) = "ISRC"
    KnowTags(3) = "Date"
    KnowTags(4) = "CopyRight"
    KnowTags(5) = "Location"
    KnowTags(6) = "Description"
    KnowTags(7) = "Organization"
    KnowTags(8) = "TrackNumber"
    KnowTags(9) = "Version"
    KnowTags(10) = "Comment"
    KnowTags(11) = "Album"
    KnowTags(12) = "Genre"
    KnowTags(13) = "Artist"
    KnowTags(14) = "Title"
    KnowTags(15) = "Composer"
    KnowTags(16) = "Conductor"
    KnowTags(17) = "Ensemble"
    KnowTags(18) = "Performer"
End Sub
' Renvoie vrai si un tag est reconnu par la classe
Private Function IsKnowTag(TagName As String) As Boolean
    Dim i As Long
    Dim vnb As Long
    Dim LeTag As String
    LeTag = Trim$(UCase$(TagName))
    vnb = UBound(KnowTags)
    
    For i = 1 To vnb
        If Trim$(UCase$(KnowTags(i))) = LeTag Then
            IsKnowTag = True
            Exit Function
        End If
    Next
    IsKnowTag = False
End Function

