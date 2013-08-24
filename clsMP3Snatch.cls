VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Rule"
Option Explicit
Private Declare Sub OpenMP3File Lib "therename.dll" (ByVal sFile As String, ByVal TypeTag As Integer)
Private Declare Function GetV2Tags Lib "therename.dll" (ByVal TypeTag As Long, ByVal TagValue As String) As String
Private Declare Function GetV1Tags Lib "therename.dll" (ByVal TypeTag As Long, ByVal TagValue As String) As String
Private Declare Sub CloseMP3File Lib "therename.dll" ()
Private sFilename As String
Private Type Info
    Album                           As String
    Band                            As String
    BPM                             As String
    Comment                         As String
    Composer                        As String
    Conductor                       As String
    ContentGroup                    As String
    ContentType                     As String
    Copyright                       As String
    EncryptionMethod                As String
    Date                            As String
    EncodedBy                       As String
    SoftwareEncodingSettings        As String
    FileOwner                       As String
    FileType                        As String
    GroupIdent                      As String
    InitialKey                      As String
    InvolvedPeopleList              As String
    ISRC                            As String
    Language                        As String
    LeadArtist                      As String
    LinkedInformation               As String
    Lyricist                        As String
    MediaType                       As String
    MixArtist                       As String
    NetRadioOwner                   As String
    NetRadioStation                 As String
    OriginalAlbum                   As String
    OriginalArtist                  As String
    OriginalFilename                As String
    OriginalLyricist                As String
    OriginalYear                    As String
    PartOfASet                      As String
    PlayListDelay                   As String
    PopulariMeter                   As String
    Publisher                       As String
    RecordingDates                  As String
    SongLength                      As String
    SubTitle                        As String
    SynchronizedLyric               As String
    TermsOfUse                      As String
    Time                            As String
    Title                           As String
    TrackNumber                     As String
    TotalTracks                     As String
    UnsynchronizedLyric             As String
    UserText                        As String
    wwwArtist                       As String
    wwwAudioFile                    As String
    wwwAudioSource                  As String
    wwwCommercialInfo               As String
    wwwCopyright                    As String
    wwwPayment                      As String
    wwwPublisher                    As String
    wwwRadioPage                    As String
    wwwUserURL                      As String
    Year                            As String
End Type

Private MP3Info As Info
Private mvarRule As String
Private mvarPlaceWhereToPut As Integer
Private mvarDefaultArtistToUse As String
Private mvarDefaultYearToUse As String
Private mvarDefaultGenreToUse As String
Private mvarDefaultAlbumToUse As String
Private mvarDefaultTitleToUse As String
Public Sub Clear()
Attribute Clear.VB_Description = "Clear all felds"
    mvarRule = ""
    mvarPlaceWhereToPut = 0
    mvarDefaultArtistToUse = ""
    mvarDefaultYearToUse = ""
    mvarDefaultGenreToUse = ""
    mvarDefaultAlbumToUse = ""
    mvarDefaultTitleToUse = ""
    sFilename = ""
End Sub
Public Property Let DefaultTitleToUse(ByVal vData As String)
    mvarDefaultTitleToUse = vData
End Property
Public Property Get DefaultTitleToUse() As String
    DefaultTitleToUse = mvarDefaultTitleToUse
End Property
Public Property Let DefaultAlbumToUse(ByVal vData As String)
    mvarDefaultAlbumToUse = vData
End Property
Public Property Get DefaultAlbumToUse() As String
    DefaultAlbumToUse = mvarDefaultAlbumToUse
End Property
Public Property Let DefaultGenreToUse(ByVal vData As String)
    mvarDefaultGenreToUse = vData
End Property
Public Property Get DefaultGenreToUse() As String
    DefaultGenreToUse = mvarDefaultGenreToUse
End Property
Public Property Let DefaultYearToUse(ByVal vData As String)
    mvarDefaultYearToUse = vData
End Property
Public Property Get DefaultYearToUse() As String
    DefaultYearToUse = mvarDefaultYearToUse
End Property
Public Property Let DefaultArtistToUse(ByVal vData As String)
    mvarDefaultArtistToUse = vData
End Property
Public Property Get DefaultArtistToUse() As String
    DefaultArtistToUse = mvarDefaultArtistToUse
End Property
Public Property Let PlaceWhereToPut(ByVal vData As Integer)
    mvarPlaceWhereToPut = vData
End Property
Public Property Get PlaceWhereToPut() As Integer
    PlaceWhereToPut = mvarPlaceWhereToPut
End Property
Public Property Let Rule(ByVal vData As String)
Attribute Rule.VB_Description = "Permet de saisir la ligne de commande servant à donner les régles"
    mvarRule = vData
End Property
Public Property Get Rule() As String
    Rule = mvarRule
End Property
Public Property Get Title() As String
    Title = MP3Info.Title
End Property
Public Property Get Artist() As String
    Artist = MP3Info.LeadArtist
End Property
Public Property Get Genre() As String
    Genre = MP3Info.ContentType
End Property
Public Property Get Album() As String
    Album = MP3Info.Album
End Property
Public Property Get Year() As String
    Year = MP3Info.Year
End Property
Public Property Get Comment() As String
    Comment = MP3Info.Comment
End Property
Public Property Get Band() As String
    Band = MP3Info.Band
End Property
Public Property Get BPM() As String
    BPM = MP3Info.BPM
End Property
Public Property Get Composer() As String
    Composer = MP3Info.Composer
End Property
Public Property Get Conductor() As String
    Conductor = MP3Info.Conductor
End Property
Public Property Get ContentGroup() As String
    ContentGroup = MP3Info.ContentGroup
End Property
Public Property Get Copyright() As String
    Copyright = MP3Info.Copyright
End Property
Public Property Get EncryptionMethod() As String
    EncryptionMethod = MP3Info.EncryptionMethod
End Property
Public Property Get mDate() As String
    mDate = MP3Info.Date
End Property
Public Property Get EncodedBy() As String
    EncodedBy = MP3Info.EncodedBy
End Property
Public Property Get SoftwareEncodingSettings() As String
    SoftwareEncodingSettings = MP3Info.SoftwareEncodingSettings
End Property
Public Property Get FileOwner() As String
    FileOwner = MP3Info.FileOwner
End Property
Public Property Get FileType() As String
    FileType = MP3Info.FileType
End Property
Public Property Get GroupIdent() As String
    GroupIdent = MP3Info.GroupIdent
End Property
Public Property Get InitialKey() As String
    InitialKey = MP3Info.InitialKey
End Property
Public Property Get InvolvedPeopleList() As String
    InvolvedPeopleList = MP3Info.InvolvedPeopleList
End Property
Public Property Get ISRC() As String
    ISRC = MP3Info.ISRC
End Property
Public Property Get Language() As String
    Language = MP3Info.Language
End Property
Public Property Get LinkedInformation() As String
    LinkedInformation = MP3Info.LinkedInformation
End Property
Public Property Get Lyricist() As String
    Lyricist = MP3Info.Lyricist
End Property
Public Property Get MediaType() As String
    MediaType = MP3Info.MediaType
End Property
Public Property Get MixArtist() As String
    MixArtist = MP3Info.MixArtist
End Property
Public Property Get NetRadioOwner() As String
    NetRadioOwner = MP3Info.NetRadioOwner
End Property
Public Property Get NetRadioStation() As String
    NetRadioStation = MP3Info.NetRadioStation
End Property
Public Property Get OriginalAlbum() As String
    OriginalAlbum = MP3Info.OriginalAlbum
End Property
Public Property Get OriginalArtist() As String
    OriginalArtist = MP3Info.OriginalArtist
End Property
Public Property Get OriginalFilename() As String
    OriginalFilename = MP3Info.OriginalFilename
End Property
Public Property Get OriginalLyricist() As String
    OriginalLyricist = MP3Info.OriginalLyricist
End Property
Public Property Get OriginalYear() As String
    OriginalYear = MP3Info.OriginalYear
End Property
Public Property Get PartOfASet() As String
    PartOfASet = MP3Info.PartOfASet
End Property
Public Property Get PlayListDelay() As String
    PlayListDelay = MP3Info.PlayListDelay
End Property
Public Property Get PopulariMeter() As String
    PopulariMeter = MP3Info.PopulariMeter
End Property
Public Property Get Publisher() As String
    Publisher = MP3Info.Publisher
End Property
Public Property Get RecordingDates() As String
    RecordingDates = MP3Info.RecordingDates
End Property
Public Property Get SongLength() As String
    SongLength = MP3Info.SongLength
End Property
Public Property Get SubTitle() As String
    SubTitle = MP3Info.SubTitle
End Property
Public Property Get SynchronizedLyric() As String
    SynchronizedLyric = MP3Info.SynchronizedLyric
End Property
Public Property Get TermsOfUse() As String
    TermsOfUse = MP3Info.TermsOfUse
End Property
Public Property Get Time() As String
    Time = MP3Info.Time
End Property
Public Property Get TrackNumber() As String
    TrackNumber = MP3Info.TrackNumber
End Property
Public Property Get TotalTracks() As String
    TotalTracks = MP3Info.TotalTracks
End Property
Public Property Get UserText() As String
    UserText = MP3Info.UserText
End Property
Public Property Get UnsynchronizedLyric() As String
    UnsynchronizedLyric = MP3Info.UnsynchronizedLyric
End Property
Public Property Get wwwArtist() As String
    wwwArtist = MP3Info.wwwArtist
End Property
Public Property Get wwwAudioFile() As String
    wwwAudioFile = MP3Info.wwwAudioFile
End Property
Public Property Get wwwAudioSource() As String
    wwwAudioSource = MP3Info.wwwAudioSource
End Property
Public Property Get wwwCommercialInfo() As String
    wwwCommercialInfo = MP3Info.wwwCommercialInfo
End Property
Public Property Get wwwCopyright() As String
    wwwCopyright = MP3Info.wwwCopyright
End Property
Public Property Get wwwPayment() As String
    wwwPayment = MP3Info.wwwPayment
End Property
Public Property Get wwwPublisher() As String
    wwwPublisher = MP3Info.wwwPublisher
End Property
Public Property Get wwwRadioPage() As String
    wwwRadioPage = MP3Info.wwwRadioPage
End Property
Public Property Get wwwUserURL() As String
    wwwUserURL = MP3Info.wwwUserURL
End Property
Public Function GetMP3Infos(ByVal sPassFileName As String, Optional Tout As Boolean = True) As String
    On Error GoTo ErrGen
    Dim Inter As String
    Dim chtempo As String
    
    If UCase$(Suffixe(sPassFileName)) <> "MP3" Then
        ClearTags
        GetMP3Infos = ""
        Exit Function
    End If
    
    If sPassFileName <> sFilename Then
        OuvreMP3 sPassFileName
    End If
    sFilename = sPassFileName
    
    If LesOptions.Mp3VqfOpt4 <> 0 Then
        chtempo = String$(LesOptions.Mp3VqfOpt4, "0")
        MP3Info.TrackNumber = Format$(Val(MP3Info.TrackNumber), chtempo)
        chtempo = String$(LesOptions.Mp3VqfOpt4, "0")
        MP3Info.TotalTracks = Format$(Val(MP3Info.TotalTracks), chtempo)
    End If
    
    If Tout = False Then    ' la méthode n'a pas été appelée par une fenêtre mais par une fonction du Free Form
        Exit Function
    End If
    
    Inter = mvarRule
    If MP3Info.Title = "" Then
        MP3Info.Title = mvarDefaultTitleToUse
    End If
    Inter = Replace(Inter, "%Title%", MP3Info.Title, , , vbTextCompare)
    
    If MP3Info.LeadArtist = "" Then
        MP3Info.LeadArtist = mvarDefaultArtistToUse
    End If
    Inter = Replace(Inter, "%Artist%", MP3Info.LeadArtist, , , vbTextCompare)

    If MP3Info.Album = "" Then
        MP3Info.Album = mvarDefaultAlbumToUse
    End If
    Inter = Replace(Inter, "%Album%", MP3Info.Album, , , vbTextCompare)
    
    If MP3Info.Year = "" Then
        MP3Info.Year = mvarDefaultYearToUse
    End If
    
    If MP3Info.ContentType = "" Then
        MP3Info.ContentType = mvarDefaultGenreToUse
    End If
    Inter = Replace(Inter, "%Genre%", MP3Info.ContentType, , , vbTextCompare)
    
    Inter = Replace(Inter, "%Year%", MP3Info.Year, , , vbTextCompare)
    Inter = Replace(Inter, "%Comment%", MP3Info.Comment, , , vbTextCompare)
    Inter = Replace(Inter, "%Band%", MP3Info.Band, , , vbTextCompare)
    Inter = Replace(Inter, "%BPM%", MP3Info.BPM, , , vbTextCompare)
    Inter = Replace(Inter, "%Composer%", MP3Info.Composer, , , vbTextCompare)
    Inter = Replace(Inter, "%Conductor%", MP3Info.Conductor, , , vbTextCompare)
    Inter = Replace(Inter, "%ContentGroup%", MP3Info.ContentGroup, , , vbTextCompare)
    Inter = Replace(Inter, "%Copyright%", MP3Info.Copyright, , , vbTextCompare)
    Inter = Replace(Inter, "%EncryptionMethod%", MP3Info.EncryptionMethod, , , vbTextCompare)
    Inter = Replace(Inter, "%Date%", MP3Info.Date, , , vbTextCompare)
    Inter = Replace(Inter, "%EncodedBy%", MP3Info.EncodedBy, , , vbTextCompare)
    Inter = Replace(Inter, "%SoftwareEncodingSettings%", MP3Info.SoftwareEncodingSettings, , , vbTextCompare)
    Inter = Replace(Inter, "%FileOwner%", MP3Info.FileOwner, , , vbTextCompare)
    Inter = Replace(Inter, "%FileType%", MP3Info.FileType, , , vbTextCompare)
    Inter = Replace(Inter, "%GroupIdent%", MP3Info.GroupIdent, , , vbTextCompare)
    Inter = Replace(Inter, "%InitialKey%", MP3Info.InitialKey, , , vbTextCompare)
    Inter = Replace(Inter, "%InvolvedPeopleList%", MP3Info.InvolvedPeopleList, , , vbTextCompare)
    Inter = Replace(Inter, "%ISRC%", MP3Info.ISRC, , , vbTextCompare)
    Inter = Replace(Inter, "%Language%", MP3Info.Language, , , vbTextCompare)
    Inter = Replace(Inter, "%LinkedInformation%", MP3Info.LinkedInformation, , , vbTextCompare)
    Inter = Replace(Inter, "%Lyricist%", MP3Info.Lyricist, , , vbTextCompare)
    Inter = Replace(Inter, "%MediaType%", MP3Info.MediaType, , , vbTextCompare)
    Inter = Replace(Inter, "%MixArtist%", MP3Info.MixArtist, , , vbTextCompare)
    Inter = Replace(Inter, "%NetRadioOwner%", MP3Info.NetRadioOwner, , , vbTextCompare)
    Inter = Replace(Inter, "%NetRadioStation%", MP3Info.NetRadioStation, , , vbTextCompare)
    Inter = Replace(Inter, "%OriginalAlbum%", MP3Info.OriginalAlbum, , , vbTextCompare)
    Inter = Replace(Inter, "%OriginalArtist%", MP3Info.OriginalArtist, , , vbTextCompare)
    Inter = Replace(Inter, "%OriginalFilename%", MP3Info.OriginalFilename, , , vbTextCompare)
    Inter = Replace(Inter, "%OriginalLyricist%", MP3Info.OriginalLyricist, , , vbTextCompare)
    Inter = Replace(Inter, "%OriginalYear%", MP3Info.OriginalYear, , , vbTextCompare)
    Inter = Replace(Inter, "%PartOfASet%", MP3Info.PartOfASet, , , vbTextCompare)
    Inter = Replace(Inter, "%PlayListDelay%", MP3Info.PlayListDelay, , , vbTextCompare)
    Inter = Replace(Inter, "%PopulariMeter%", MP3Info.PopulariMeter, , , vbTextCompare)
    Inter = Replace(Inter, "%Publisher%", MP3Info.Publisher, , , vbTextCompare)
    Inter = Replace(Inter, "%RecordingDates%", MP3Info.RecordingDates, , , vbTextCompare)
    Inter = Replace(Inter, "%SongLength%", MP3Info.SongLength, , , vbTextCompare)
    Inter = Replace(Inter, "%SubTitle%", MP3Info.SubTitle, , , vbTextCompare)
    Inter = Replace(Inter, "%SynchronizedLyric%", MP3Info.SynchronizedLyric, , , vbTextCompare)
    Inter = Replace(Inter, "%TermsOfUse%", MP3Info.TermsOfUse, , , vbTextCompare)
    Inter = Replace(Inter, "%Time%", MP3Info.Time, , , vbTextCompare)
    Inter = Replace(Inter, "%TrackNumber%", MP3Info.TrackNumber, , , vbTextCompare)
    Inter = Replace(Inter, "%TotalTracks%", MP3Info.TotalTracks, , , vbTextCompare)
    Inter = Replace(Inter, "%UnsynchronizedLyric%", MP3Info.UnsynchronizedLyric, , , vbTextCompare)
    Inter = Replace(Inter, "%UserText%", MP3Info.UserText, , , vbTextCompare)
    Inter = Replace(Inter, "%wwwArtist%", MP3Info.wwwArtist, , , vbTextCompare)
    Inter = Replace(Inter, "%wwwAudioFile%", MP3Info.wwwAudioFile, , , vbTextCompare)
    Inter = Replace(Inter, "%wwwAudioSource%", MP3Info.wwwAudioSource, , , vbTextCompare)
    Inter = Replace(Inter, "%wwwCommercialInfo%", MP3Info.wwwCommercialInfo, , , vbTextCompare)
    Inter = Replace(Inter, "%wwwCopyright%", MP3Info.wwwCopyright, , , vbTextCompare)
    Inter = Replace(Inter, "%wwwPayment%", MP3Info.wwwPayment, , , vbTextCompare)
    Inter = Replace(Inter, "%wwwPublisher%", MP3Info.wwwPublisher, , , vbTextCompare)
    Inter = Replace(Inter, "%wwwRadioPage%", MP3Info.wwwRadioPage, , , vbTextCompare)
    Inter = Replace(Inter, "%wwwUserURL%", MP3Info.wwwUserURL, , , vbTextCompare)
    GetMP3Infos = Inter
    Exit Function
    
ErrGen:
 ErreurGrave "GetMP3Infos"
 Exit Function
End Function
Private Sub Class_Initialize()
    Clear
End Sub
' Procédure chargée d'ouvrir le fichier MP3 voulu avec les options sélectionnées
Private Sub OuvreMP3(NomFichier As String)
    ClearTags
    If LesOptions.TagsVersionToUse = 0 Then ' Seulement les tag de version 1
        OpenMP3File NomFichier, 1
        RempliV1Tags
        CloseMP3File
    Else    ' Tag de version 2
        If LesOptions.TagsPriority <> 0 Then
            OpenMP3File NomFichier, 1
            RempliV1Tags
            CloseMP3File
        End If
        OpenMP3File NomFichier, 2
        RempliV2Tags
        CloseMP3File
    End If
End Sub
Private Sub RempliV2Tags()
' si on est là c'est que les tags de la version 1 ont été remplis
    If LesOptions.TagsPriority = 0 Then
        MP3Info.Album = RenvoieUnV2Tag(1)
    Else
        If RenvoieUnV2Tag(1) <> "" Then
            MP3Info.Album = RenvoieUnV2Tag(1)
        End If
    End If
    MP3Info.Band = RenvoieUnV2Tag(3)
    MP3Info.BPM = RenvoieUnV2Tag(4)
    If LesOptions.TagsPriority = 0 Then
        MP3Info.Comment = RenvoieUnV2Tag(7)
    Else
        If RenvoieUnV2Tag(7) <> "" Then
            MP3Info.Comment = RenvoieUnV2Tag(7)
        End If
    End If
    MP3Info.Composer = RenvoieUnV2Tag(9)
    MP3Info.Conductor = RenvoieUnV2Tag(10)
    MP3Info.ContentGroup = RenvoieUnV2Tag(11)
    
    If LesOptions.TagsPriority = 0 Then
        MP3Info.ContentType = RenvoieLeGenre(RenvoieUnV2Tag(12))
        'MP3Info.ContentType = RenvoieUnV2Tag(12)
    Else
        If RenvoieUnV2Tag(12) <> "" Then
            MP3Info.ContentType = RenvoieLeGenre(RenvoieUnV2Tag(12))
            'MP3Info.ContentType = RenvoieUnV2Tag(12)
        End If
    End If
    MP3Info.Copyright = RenvoieUnV2Tag(13)
    MP3Info.EncryptionMethod = RenvoieUnV2Tag(14)
    MP3Info.Date = RenvoieUnV2Tag(15)
    MP3Info.EncodedBy = RenvoieUnV2Tag(16)
    MP3Info.SoftwareEncodingSettings = RenvoieUnV2Tag(17)
    MP3Info.FileOwner = RenvoieUnV2Tag(19)
    MP3Info.FileType = RenvoieUnV2Tag(20)
    MP3Info.GroupIdent = RenvoieUnV2Tag(22)
    MP3Info.InitialKey = RenvoieUnV2Tag(23)
    MP3Info.InvolvedPeopleList = RenvoieUnV2Tag(24)
    MP3Info.ISRC = RenvoieUnV2Tag(25)
    MP3Info.Language = RenvoieUnV2Tag(26)
    
    If LesOptions.TagsPriority = 0 Then
        MP3Info.LeadArtist = RenvoieUnV2Tag(27)
    Else
        If RenvoieUnV2Tag(27) <> "" Then
            MP3Info.LeadArtist = RenvoieUnV2Tag(27)
        End If
    End If
    
    MP3Info.LinkedInformation = RenvoieUnV2Tag(28)
    MP3Info.Lyricist = RenvoieUnV2Tag(29)
    MP3Info.MediaType = RenvoieUnV2Tag(30)
    MP3Info.MixArtist = RenvoieUnV2Tag(31)
    MP3Info.NetRadioOwner = RenvoieUnV2Tag(33)
    MP3Info.NetRadioStation = RenvoieUnV2Tag(34)
    MP3Info.OriginalAlbum = RenvoieUnV2Tag(35)
    MP3Info.OriginalArtist = RenvoieUnV2Tag(36)
    MP3Info.OriginalFilename = RenvoieUnV2Tag(37)
    MP3Info.OriginalLyricist = RenvoieUnV2Tag(38)
    MP3Info.OriginalYear = RenvoieUnV2Tag(39)
    MP3Info.PartOfASet = RenvoieUnV2Tag(41)
    MP3Info.PlayListDelay = RenvoieUnV2Tag(43)
    MP3Info.PopulariMeter = RenvoieUnV2Tag(44)
    MP3Info.Publisher = RenvoieUnV2Tag(47)
    MP3Info.RecordingDates = RenvoieUnV2Tag(48)
    MP3Info.SongLength = RenvoieUnV2Tag(51)
    MP3Info.SubTitle = RenvoieUnV2Tag(52)
    MP3Info.SynchronizedLyric = RenvoieUnV2Tag(53)
    MP3Info.TermsOfUse = RenvoieUnV2Tag(55)
    MP3Info.Time = RenvoieUnV2Tag(56)
    
    If LesOptions.TagsPriority = 0 Then
        MP3Info.Title = RenvoieUnV2Tag(57)
    Else
        If RenvoieUnV2Tag(57) <> "" Then
            MP3Info.Title = RenvoieUnV2Tag(57)
        End If
    End If
    
    If LesOptions.TagsPriority = 0 Then
        MP3Info.TrackNumber = GetTrackNum(RenvoieUnV2Tag(58))
    Else
        If RenvoieUnV2Tag(58) <> "" Then
            MP3Info.TrackNumber = GetTrackNum(RenvoieUnV2Tag(58))
        End If
    End If
    
    MP3Info.TotalTracks = GetTotalTracks(RenvoieUnV2Tag(58))
    
    MP3Info.UnsynchronizedLyric = RenvoieUnV2Tag(60)
    MP3Info.UserText = RenvoieUnV2Tag(61)
    MP3Info.wwwArtist = RenvoieUnV2Tag(63)
    MP3Info.wwwAudioFile = RenvoieUnV2Tag(64)
    MP3Info.wwwAudioSource = RenvoieUnV2Tag(65)
    MP3Info.wwwCommercialInfo = RenvoieUnV2Tag(66)
    MP3Info.wwwCopyright = RenvoieUnV2Tag(67)
    MP3Info.wwwPayment = RenvoieUnV2Tag(68)
    MP3Info.wwwPublisher = RenvoieUnV2Tag(69)
    MP3Info.wwwRadioPage = RenvoieUnV2Tag(70)
    MP3Info.wwwUserURL = RenvoieUnV2Tag(71)
    
    If LesOptions.TagsPriority = 0 Then
        MP3Info.Year = RenvoieUnV2Tag(72)
    Else
        If RenvoieUnV2Tag(72) <> "" Then
            MP3Info.Year = RenvoieUnV2Tag(72)
        End If
    End If
End Sub
Private Sub RempliV1Tags()
    MP3Info.Album = RenvoieUnV1Tag(1)                           ' sAlbum
    MP3Info.LeadArtist = RenvoieUnV1Tag(27)                     ' sArtist
    MP3Info.Comment = RenvoieUnV1Tag(7)                         ' sComment
    MP3Info.Title = RenvoieUnV1Tag(57)                          ' sTitle
    MP3Info.Year = RenvoieUnV1Tag(72)                           ' sYear
    MP3Info.ContentType = RenvoieLeGenre(RenvoieUnV1Tag(12))    ' sGenre
    MP3Info.TrackNumber = RenvoieUnV1Tag(58)                    ' Inexistant
End Sub
Private Function RenvoieUnV2Tag(NumTag As Long) As String
    Dim vtmp As String
    Dim zNumTag As Long
    zNumTag = NumTag
    vtmp = String$(1024, " ")
    vtmp = GetV2Tags(zNumTag, vtmp)
    vtmp = Trim$(Replace(vtmp, Chr$(0), "", , , vbBinaryCompare))
    vtmp = Replace(vtmp, Chr$(9), "", , , vbBinaryCompare)
    If LesOptions.Mp3VqfOpt1 = 1 Then  ' Separate words
        vtmp = ExtractWords(vtmp)
    End If
    If LesOptions.Mp3VqfOpt2 = 1 Then  ' Remove multiple spaces
        vtmp = RemoveMultipleSpacing(vtmp)
    End If
    vtmp = ChangeMP3Case(vtmp)
    RenvoieUnV2Tag = vtmp
End Function
Private Function RenvoieUnV1Tag(NumTag As Long) As String
    Dim vtmp As String
    Dim zNumTag As Long
    zNumTag = NumTag
    vtmp = String$(1024, " ")
    vtmp = Trim$(Replace(GetV1Tags(zNumTag, vtmp), Chr$(0), "", , , vbBinaryCompare))
    vtmp = Replace(vtmp, Chr$(9), "", , , vbBinaryCompare)
    If LesOptions.Mp3VqfOpt1 = 1 Then  ' Separate words
        vtmp = ExtractWords(vtmp)
    End If
    If LesOptions.Mp3VqfOpt2 = 1 Then  ' Remove multiple spaces
        vtmp = RemoveMultipleSpacing(vtmp)
    End If
    vtmp = ChangeMP3Case(vtmp)
    RenvoieUnV1Tag = vtmp
End Function

Private Function RenvoieLeGenre(Genre As String) As String
    Dim vtmp As String
    Dim i As Integer
    Dim longueur As Integer
    Dim Nombres As String
    Nombres = "0123456789"
    longueur = Len(Genre)
    vtmp = ""
    For i = 1 To longueur
        If InStr(Nombres, Mid$(Genre, i, 1)) <> 0 Then
            vtmp = vtmp + Mid$(Genre, i, 1)
        End If
    Next
    Dim sGenre(148) As String
    sGenre(0) = "Blues"
    sGenre(1) = "Classic Rock"
    sGenre(2) = "Country"
    sGenre(3) = "Dance"
    sGenre(4) = "Disco"
    sGenre(5) = "Funk"
    sGenre(6) = "Grunge"
    sGenre(7) = "Hip-Hop"
    sGenre(8) = "Jazz"
    sGenre(9) = "Metal"
    sGenre(10) = "New Age"
    sGenre(11) = "Oldies"
    sGenre(12) = "Other"
    sGenre(13) = "Pop"
    sGenre(14) = "R&B"
    sGenre(15) = "Rap"
    sGenre(16) = "Reggae"
    sGenre(17) = "Rock"
    sGenre(18) = "Techno"
    sGenre(19) = "Industrial"
    sGenre(20) = "Alternative"
    sGenre(21) = "Ska"
    sGenre(22) = "Death Metal"
    sGenre(23) = "Pranks"
    sGenre(24) = "Soundtrack"
    sGenre(25) = "Euro-Techno"
    sGenre(26) = "Ambient"
    sGenre(27) = "Trip Hop"
    sGenre(28) = "Vocal"
    sGenre(29) = "Jazz+Funk"
    sGenre(30) = "Fusion"
    sGenre(31) = "Trance"
    sGenre(32) = "Classical"
    sGenre(33) = "Instrumental"
    sGenre(34) = "Acid"
    sGenre(35) = "House"
    sGenre(36) = "Game"
    sGenre(37) = "Sound Clip"
    sGenre(38) = "Gospel"
    sGenre(39) = "Noise"
    sGenre(40) = "Alt. Rock"
    sGenre(41) = "Bass"
    sGenre(42) = "Soul"
    sGenre(43) = "Punk"
    sGenre(44) = "Space"
    sGenre(45) = "Meditative"
    sGenre(46) = "Instrumental Pop"
    sGenre(47) = "Instrumental Rock"
    sGenre(48) = "Ethnic"
    sGenre(49) = "Gothic"
    sGenre(50) = "Darkwave"
    sGenre(51) = "Techno-Industrial"
    sGenre(52) = "Electronic"
    sGenre(53) = "Pop-Folk"
    sGenre(54) = "Eurodance"
    sGenre(55) = "Dream"
    sGenre(56) = "Southern Rock"
    sGenre(57) = "Comedy"
    sGenre(58) = "Cult"
    sGenre(59) = "Gangsta Rap"
    sGenre(60) = "Top 40"
    sGenre(61) = "Christian Rap"
    sGenre(62) = "Pop/Punk"
    sGenre(63) = "Jungle"
    sGenre(64) = "Native American"
    sGenre(65) = "Cabaret"
    sGenre(66) = "New Wave"
    sGenre(67) = "Phychedelic"
    sGenre(68) = "Rave"
    sGenre(69) = "Showtunes"
    sGenre(70) = "Trailer"
    sGenre(71) = "Lo-Fi"
    sGenre(72) = "Tribal"
    sGenre(73) = "Acid Punk"
    sGenre(74) = "Acid Jazz"
    sGenre(75) = "Polka"
    sGenre(76) = "Retro"
    sGenre(77) = "Musical"
    sGenre(78) = "Rock & Roll"
    sGenre(79) = "Hard Rock"
    sGenre(80) = "Folk"
    sGenre(81) = "Folk/Rock"
    sGenre(82) = "National Folk"
    sGenre(83) = "Swing"
    sGenre(84) = "Fast-Fusion"
    sGenre(85) = "Bebob"
    sGenre(86) = "Latin"
    sGenre(87) = "Revival"
    sGenre(88) = "Celtic"
    sGenre(89) = "Blue Grass"
    sGenre(90) = "Avantegarde"
    sGenre(91) = "Gothic Rock"
    sGenre(92) = "Progressive Rock"
    sGenre(93) = "Psychedelic Rock"
    sGenre(94) = "Symphonic Rock"
    sGenre(95) = "Slow Rock"
    sGenre(96) = "Big Band"
    sGenre(97) = "Chorus"
    sGenre(98) = "Easy Listening"
    sGenre(99) = "Acoustic"
    sGenre(100) = "Humour"
    sGenre(101) = "Speech"
    sGenre(102) = "Chanson"
    sGenre(103) = "Opera"
    sGenre(104) = "Chamber Music"
    sGenre(105) = "Sonata"
    sGenre(106) = "Symphony"
    sGenre(107) = "Booty Bass"
    sGenre(108) = "Primus"
    sGenre(109) = "Porn Groove"
    sGenre(110) = "Satire"
    sGenre(111) = "Slow Jam"
    sGenre(112) = "Club"
    sGenre(113) = "Tango"
    sGenre(114) = "Samba"
    sGenre(115) = "Folklore"
    sGenre(116) = "Ballad"
    sGenre(117) = "power Ballad"
    sGenre(118) = "Rhythmic Soul"
    sGenre(119) = "Freestyle"
    sGenre(120) = "Duet"
    sGenre(121) = "Punk Rock"
    sGenre(122) = "Drum Solo"
    sGenre(123) = "A Capella"
    sGenre(124) = "Euro-House"
    sGenre(125) = "Dance Hall"
    sGenre(126) = "Goa"
    sGenre(127) = "Drum & Bass"
    sGenre(128) = "Club-House"
    sGenre(129) = "Hardcore"
    sGenre(130) = "Terror"
    sGenre(131) = "indie"
    sGenre(132) = "Brit Pop"
    sGenre(133) = "Negerpunk"
    sGenre(134) = "Polsk Punk"
    sGenre(135) = "Beat"
    sGenre(136) = "Christian Gangsta Rap"
    sGenre(137) = "Heavy Metal"
    sGenre(138) = "Black Metal"
    sGenre(139) = "Crossover"
    sGenre(140) = "Comteporary Christian"
    sGenre(141) = "Christian Rock"
    sGenre(142) = "Merengue"
    sGenre(143) = "Salsa"
    sGenre(144) = "Trash Metal"
    sGenre(145) = "Anime"
    sGenre(146) = "JPop"
    sGenre(147) = "Synth Pop"
    If Val(vtmp) > 0 And Val(vtmp) <= UBound(sGenre) Then
        RenvoieLeGenre = Trim$(sGenre(Val(vtmp)))
    Else
        RenvoieLeGenre = ""
    End If
End Function
' renvoie le numéro de piste
Private Function GetTrackNum(lachaine As String) As String
Dim i As Integer
Dim vnb As Integer
Dim vtmp As String
vtmp = ""
lachaine = Trim$(lachaine)
vnb = Len(lachaine)
If vnb = 0 Then
    GetTrackNum = ""
    Exit Function
End If
For i = 1 To vnb
    If IsNumeric(Mid$(lachaine, i, 1)) Then
        vtmp = vtmp + Mid$(lachaine, i, 1)
    Else
        i = vnb
    End If
Next
GetTrackNum = vtmp
End Function

Private Function GetTotalTracks(lachaine As String) As String
Dim vnb As Integer
lachaine = Trim$(lachaine)
vnb = InStr(lachaine, "/")
If vnb <> 0 Then
    GetTotalTracks = Trim$(Mid$(lachaine, vnb + 1))
Else
    GetTotalTracks = ""
End If
End Function

Private Sub ClearTags()
    MP3Info.Album = ""
    MP3Info.Band = ""
    MP3Info.BPM = ""
    MP3Info.Comment = ""
    MP3Info.Composer = ""
    MP3Info.Conductor = ""
    MP3Info.ContentGroup = ""
    MP3Info.ContentType = ""
    MP3Info.Copyright = ""
    MP3Info.EncryptionMethod = ""
    MP3Info.Date = ""
    MP3Info.EncodedBy = ""
    MP3Info.SoftwareEncodingSettings = ""
    MP3Info.FileOwner = ""
    MP3Info.FileType = ""
    MP3Info.GroupIdent = ""
    MP3Info.InitialKey = ""
    MP3Info.InvolvedPeopleList = ""
    MP3Info.ISRC = ""
    MP3Info.Language = ""
    MP3Info.LeadArtist = ""
    MP3Info.LinkedInformation = ""
    MP3Info.Lyricist = ""
    MP3Info.MediaType = ""
    MP3Info.MixArtist = ""
    MP3Info.NetRadioOwner = ""
    MP3Info.NetRadioStation = ""
    MP3Info.OriginalAlbum = ""
    MP3Info.OriginalArtist = ""
    MP3Info.OriginalFilename = ""
    MP3Info.OriginalLyricist = ""
    MP3Info.OriginalYear = ""
    MP3Info.PartOfASet = ""
    MP3Info.PlayListDelay = ""
    MP3Info.PopulariMeter = ""
    MP3Info.Publisher = ""
    MP3Info.RecordingDates = ""
    MP3Info.SongLength = ""
    MP3Info.SubTitle = ""
    MP3Info.SynchronizedLyric = ""
    MP3Info.TermsOfUse = ""
    MP3Info.Time = ""
    MP3Info.Title = ""
    MP3Info.TrackNumber = ""
    MP3Info.TotalTracks = ""
    MP3Info.UnsynchronizedLyric = ""
    MP3Info.UserText = ""
    MP3Info.wwwArtist = ""
    MP3Info.wwwAudioFile = ""
    MP3Info.wwwAudioSource = ""
    MP3Info.wwwCommercialInfo = ""
    MP3Info.wwwCopyright = ""
    MP3Info.wwwPayment = ""
    MP3Info.wwwPublisher = ""
    MP3Info.wwwRadioPage = ""
    MP3Info.wwwUserURL = ""
    MP3Info.Year = ""
End Sub
