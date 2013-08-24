VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Rule"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Function WMAInfo Lib "renmm.dll" (ByVal File As String, ByVal ChannelMode As String, ByVal SampleRate As String, ByVal Duration As String, ByVal BitRate As String, ByVal Track As String, ByVal Title As String, ByVal Artist As String, ByVal Album As String, ByVal Year As String, ByVal Genre As String, ByVal Comment As String) As Long
Private mvarRule As String
Private mvarPlaceWhereToPut As Integer
Private mvarChannelMode As String
Private mvarSampleRate As String
Private mvarDuration As String
Private mvarBitRate As String
Private mvarTrack As String
Private mvarTitle As String
Private mvarDefaultTitleToUse As String
Private mvarArtist As String
Private mvarDefaultArtistToUse As String
Private mvarAlbum As String
Private mvarDefaultAlbumToUse As String
Private mvarGenre As String
Private mvarDefaultGenreToUse As String
Private mvarComment As String
Private mvarYear As String
Private mvarDefaultYearToUse As String
Private sFilename As String
' Fonction qui récupère les infos d'un fichier WMA(si elles sont présentes...)
Public Function GetWMAInfos(sPassFileName As String, Optional Tout As Boolean = True) As String
    On Error GoTo ErrGen
    Dim Inter As String
   
    If UCase$(Suffixe(sPassFileName)) <> "WMA" Then
        ClearTags
        GetWMAInfos = ""
        Exit Function
    End If
    
    If sPassFileName <> sFilename Then
        OuvreWMA sPassFileName
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
    If Year() = "" Then
        Inter = Replace(Inter, "%Year%", mvarDefaultYearToUse, , , vbTextCompare)
    Else
        Inter = Replace(Inter, "%Year%", Genre(), , , vbTextCompare)
    End If
    Inter = Replace(Inter, "%ChannelMode%", ChannelMode(), , , vbTextCompare)
    Inter = Replace(Inter, "%SampleRate%", SampleRate(), , , vbTextCompare)
    Inter = Replace(Inter, "%Duration%", Duration(), , , vbTextCompare)
    Inter = Replace(Inter, "%BitRate%", BitRate(), , , vbTextCompare)
    Inter = Replace(Inter, "%Track%", Track(), , , vbTextCompare)
    Inter = Replace(Inter, "%Comment%", Comment(), , , vbTextCompare)
    
    GetWMAInfos = Inter
    Exit Function
    
ErrGen:
    ErreurGrave "GetWMAInfos"
    Exit Function
End Function
Private Sub OuvreWMA(NomFichier As String)
    On Error GoTo ErrGen
    Dim valret As Long
    Dim fichier As String, BitRate As String, Track As String, Title As String, Artist As String, Album As String, Year As String, Genre As String, Comment As String, ChannelMode As String, SampleRate As String, Duration As String
    Dim chtempo As String
    ClearTags
    fichier = NomFichier
    BitRate = String$(256, " ")
    Track = String$(256, " ")
    Title = String$(256, " ")
    Artist = String$(256, " ")
    Album = String$(256, " ")
    Year = String$(256, " ")
    Genre = String$(256, " ")
    Comment = String$(256, " ")
    SampleRate = String$(256, " ")
    Duration = String$(256, " ")
    ChannelMode = String$(256, " ")
    valret = WMAInfo(fichier, ChannelMode, SampleRate, Duration, BitRate, Track, Title, Artist, Album, Year, Genre, Comment)
    If valret = 1 Then
        mvarChannelMode = ReformatTag(Menage(ChannelMode))
        mvarSampleRate = ReformatTag(Menage(SampleRate))
        mvarDuration = ReformatTag(Menage(Duration))
        mvarBitRate = ReformatTag(Menage(BitRate))
        mvarTrack = ReformatTag(Menage(Track))
        If LesOptions.Mp3VqfOpt4 <> 0 Then
            chtempo = String$(LesOptions.Mp3VqfOpt4, "0")
            mvarTrack = Format$(Val(mvarTrack), chtempo)
        End If
        mvarTitle = ReformatTag(Menage(Title))
        mvarArtist = ReformatTag(Menage(Artist))
        mvarAlbum = ReformatTag(Menage(Album))
        mvarGenre = ReformatTag(Menage(Genre))
        mvarComment = ReformatTag(Menage(Comment))
        mvarYear = ReformatTag(Menage(Year))
    End If
    Exit Sub
    
ErrGen:
    ErreurGrave "OuvreWMA"
    Exit Sub
End Sub
Public Sub FillTagsList(lv As ListView)
    lv.ListItems.Clear
    AddOneTag "BitRate", BitRate(), lv
    AddOneTag "Track", Track(), lv
    AddOneTag "Title", Title(), lv
    AddOneTag "Artist", Artist(), lv
    AddOneTag "Album", Album(), lv
    AddOneTag "Year", Year(), lv
    AddOneTag "Genre", Genre(), lv
    AddOneTag "Comment", Comment(), lv
    AddOneTag "ChannelMode", ChannelMode(), lv
    AddOneTag "SampleRate", SampleRate(), lv
    AddOneTag "Duration", Duration(), lv
    ResizeLvMp3
End Sub
Private Sub AddOneTag(txt As String, Tag As String, lv As ListView)
Dim Aff As Boolean
Dim itmX As ListItem
Aff = True
If LesOptions.RemoveEmptyTags = 1 And Trim$(Tag) = "" Then
    Aff = False
End If
If Aff Then
    Set itmX = lv.ListItems.Add(, , txt)
    itmX.SubItems(1) = Tag
End If

End Sub
Private Function ReformatTag(Tag As String) As String
    Dim tTag As String
    tTag = Tag
    If LesOptions.Mp3VqfOpt1 = 1 Then ' Separate Words
        tTag = ExtractWords(tTag)
    End If
    If LesOptions.Mp3VqfOpt2 = 1 Then ' Remove multiple spaces
        tTag = RemoveMultipleSpacing(tTag)
    End If
    ReformatTag = ChangeMP3Case(tTag)
End Function
Private Function Menage(chaine As String) As String
    Dim vtmp As String
    vtmp = Trim$(Replace(chaine, Chr$(0), "", , , vbBinaryCompare))
    vtmp = Replace(vtmp, Chr$(9), "", , , vbBinaryCompare)
    Menage = vtmp
End Function

Private Sub Class_Initialize()
    ClearTags
End Sub
Private Sub ClearTags()
    mvarChannelMode = ""
    mvarSampleRate = ""
    mvarDuration = ""
    mvarBitRate = ""
    mvarTrack = ""
    mvarTitle = ""
    mvarArtist = ""
    mvarAlbum = ""
    mvarGenre = ""
    mvarComment = ""
    mvarYear = ""
End Sub
Public Property Get Year() As String
    Year = mvarYear
End Property
Public Property Get Comment() As String
    Comment = mvarComment
End Property
Public Property Get Genre() As String
    Genre = mvarGenre
End Property
Public Property Get Album() As String
    Album = mvarAlbum
End Property
Public Property Get Artist() As String
    Artist = mvarArtist
End Property
Public Property Get Title() As String
    Title = mvarTitle
End Property
Public Property Get Track() As String
    Track = mvarTrack
End Property
Public Property Get BitRate() As String
    BitRate = mvarBitRate
End Property
Public Property Get Duration() As String
    Duration = mvarDuration
End Property
Public Property Get SampleRate() As String
    SampleRate = mvarSampleRate
End Property
Public Property Get ChannelMode() As String
    ChannelMode = mvarChannelMode
End Property
Public Property Let PlaceWhereToPut(ByVal vData As Integer)
    mvarPlaceWhereToPut = vData
End Property
Public Property Get PlaceWhereToPut() As Integer
    PlaceWhereToPut = mvarPlaceWhereToPut
End Property
Public Property Let Rule(ByVal vData As String)
    mvarRule = vData
End Property
Public Property Get Rule() As String
    Rule = mvarRule
End Property
Public Property Let DefaultTitleToUse(ByVal vData As String)
    mvarDefaultTitleToUse = vData
End Property
Public Property Get DefaultTitleToUse() As String
    DefaultTitleToUse = mvarDefaultTitleToUse
End Property
Public Property Let DefaultArtistToUse(ByVal vData As String)
    mvarDefaultArtistToUse = vData
End Property
Public Property Get DefaultArtistToUse() As String
    DefaultArtistToUse = mvarDefaultArtistToUse
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

