VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVQF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A Visual Basic class to get information from VQF files"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Rule"
Option Explicit
'Private sFilename As String
Private mvarQuality As String
Private mvarMono_Stereo As String
Private mvarBitRate As String
Private mvarSampleRate As String
Private mvarTitle As String
Private mvarComment As String
Private mvarAuthor As String
Private mvarCopyright As String
Private mvarSaveAsFilename As String
Private mvarDefaultArtistToUse As String
Private mvarDefaultTitle As String
Private mvarFileName As String
Private mvarPlaceWhereToPut As Integer
Private mvarRule As String
' Fonction qui récupère les infos d'un fichier VQF (si elles sont présentes...)
Public Function GetVQFInfos(ByVal sPassFileName As String, Optional Tout As Boolean = True) As String
Dim vin As Integer, Buffer As String, TailleEntete As Integer
Dim position As Integer, tmp As String, tmp2 As Integer, Inter As String
Dim rechVQF As New CSearch
On Error GoTo ErrGen
 Inter = ""
 mvarQuality = ""
 mvarMono_Stereo = ""
 mvarBitRate = ""
 mvarSampleRate = ""
 mvarTitle = ""
 mvarComment = ""
 mvarAuthor = ""
 mvarCopyright = ""
 mvarSaveAsFilename = ""
 mvarFileName = sPassFileName

    If UCase$(Suffixe(sPassFileName)) <> "VQF" Then
        GetVQFInfos = ""
        Exit Function
    End If

 vin = FreeFile
 Open sPassFileName For Binary As #vin
 Buffer = Space$(4)
 Get #vin, 1, Buffer
 ' Vérification, est-ce un fichier VQF ?
 If UCase$(Buffer) <> "TWIN" Then
     Close #vin
     Exit Function
 End If
 Buffer = Space$(1)
 Get #vin, 16, Buffer
 TailleEntete = Asc(Buffer)
 ' Taille de l'entête
 If TailleEntete = 0 Then
    ' Entête à zéro, pas la peine de continuer
     Close #vin
    Exit Function
 End If
 ' Rendu là, on connait la longueur de l'entête et on sait où le lire
 Buffer = Space$(TailleEntete)
 Get #vin, 17, Buffer
 Close #vin
 ' Il ne reste "plus" qu'à analyzer le contenu de cet entête.
 position = 1
 While position <= TailleEntete
    tmp = UCase$(Mid$(Buffer, position, 4))
    Select Case tmp
        Case "COMM" ' Informations sur le fichier
            mvarQuality = Trim$(Str$(Asc(Mid$(Buffer, position + 7, 1))))
            If Asc(Mid$(Buffer, position + 11, 1)) = 0 Then
                mvarMono_Stereo = "Mono"
            Else
                mvarMono_Stereo = "Stereo"
            End If
            mvarBitRate = Trim$(Str$(Asc(Mid$(Buffer, position + 15, 1))))
            mvarSampleRate = Trim$(Str$(Asc(Mid$(Buffer, position + 19, 1))))
            ' Il ne reste plus qu'à sauter ce TAG
            position = position + 24
            
        Case "NAME" ' Titre de la chanson
            ' Il faut extraire la longueur de la chaine contenant le titre
            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
            If tmp2 <> 0 Then
                mvarTitle = Trim$(Mid$(Buffer, position + 8, tmp2))
                If LesOptions.Mp3VqfOpt1 = 1 Then ' Separate Words
                    mvarTitle = ExtractWords(mvarTitle)
                End If
                If LesOptions.Mp3VqfOpt2 = 1 Then ' Remove multiple spaces
                    mvarTitle = RemoveMultipleSpacing(mvarTitle)
                End If
                mvarTitle = ChangeMP3Case(mvarTitle)
            End If
            ' Il ne reste plus qu'à sauter ce TAG
            position = position + 8 + tmp2
            
        Case "COMT" ' Commentaire
            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
            If tmp2 <> 0 Then
                mvarComment = Trim$(Mid$(Buffer, position + 8, tmp2))
            End If
            If LesOptions.Mp3VqfOpt1 = 1 Then ' Separate Words
                mvarComment = ExtractWords(mvarComment)
            End If
            If LesOptions.Mp3VqfOpt2 = 1 Then ' Remove multiple spaces
                mvarComment = RemoveMultipleSpacing(mvarComment)
            End If
            mvarComment = ChangeMP3Case(mvarComment)
            ' Il ne reste plus qu'à sauter ce TAG
            position = position + 8 + tmp2
            
        Case "AUTH" ' Nom de l'auteur
            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
            If tmp2 <> 0 Then
                mvarAuthor = Trim$(Mid$(Buffer, position + 8, tmp2))
            End If
            If LesOptions.Mp3VqfOpt1 = 1 Then ' Separate Words
                mvarAuthor = ExtractWords(mvarAuthor)
            End If
            If LesOptions.Mp3VqfOpt2 = 1 Then ' Remove multiple spaces
                mvarAuthor = RemoveMultipleSpacing(mvarAuthor)
            End If
            mvarAuthor = ChangeMP3Case(mvarAuthor)
            ' Il ne reste plus qu'à sauter ce TAG
            position = position + 8 + tmp2
            
        Case "(C) " ' copyright
            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
            If tmp2 <> 0 Then
                mvarCopyright = Trim$(Mid$(Buffer, position + 8, tmp2))
            End If
            If LesOptions.Mp3VqfOpt1 = 1 Then ' Separate Words
                mvarCopyright = ExtractWords(mvarCopyright)
            End If
            If LesOptions.Mp3VqfOpt2 = 1 Then ' Remove multiple spaces
                mvarCopyright = RemoveMultipleSpacing(mvarCopyright)
            End If
            mvarCopyright = ChangeMP3Case(mvarCopyright)
            ' Il ne reste plus qu'à sauter ce TAG
            position = position + 8 + tmp2
            
        Case "FILE" ' Nom par défaut à donner au fichier lors de sa sauvegarde
            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
            If tmp2 <> 0 Then
                mvarSaveAsFilename = Trim$(Mid$(Buffer, position + 8, tmp2))
            End If
            If LesOptions.Mp3VqfOpt1 = 1 Then ' Separate Words
                mvarSaveAsFilename = ExtractWords(mvarSaveAsFilename)
            End If
            If LesOptions.Mp3VqfOpt2 = 1 Then ' Remove multiple spaces
                mvarSaveAsFilename = RemoveMultipleSpacing(mvarSaveAsFilename)
            End If
            mvarSaveAsFilename = ChangeMP3Case(mvarSaveAsFilename)
            ' Il ne reste plus qu'à sauter ce TAG
            position = position + 8 + tmp2
            
        Case "DSIZ" ' Inutilisé dans le programme
            tmp2 = Asc(Mid$(Buffer, position + 7, 1))
            ' On saute directement ce TAG puisqu'il n'est pas utilisé
            position = position + 8 + tmp2
        Case Else
            Exit Function ' On n'a pas du bien interpreter les données
    End Select
Wend
    If Tout = False Then
        Exit Function
    End If
    rechVQF.SourceString = mvarRule
    rechVQF.SearchString = "%Title%"
    rechVQF.MatchCase = False
    If mvarTitle <> "" Then
     rechVQF.ReplaceString = mvarTitle
    Else
     rechVQF.ReplaceString = mvarDefaultTitle
    End If
    Inter = rechVQF.BeginSearchAndReplace
    
    rechVQF.SourceString = Inter
    rechVQF.SearchString = "%Artist%"
    If mvarAuthor <> "" Then
     rechVQF.ReplaceString = mvarAuthor
    Else
     rechVQF.ReplaceString = mvarDefaultArtistToUse
    End If
    Inter = rechVQF.BeginSearchAndReplace
    
    rechVQF.SourceString = Inter
    rechVQF.SearchString = "%Quality%"
    rechVQF.ReplaceString = mvarQuality
    Inter = rechVQF.BeginSearchAndReplace
        
    rechVQF.SourceString = Inter
    rechVQF.SearchString = "%Mono_Stereo%"
    rechVQF.ReplaceString = mvarMono_Stereo
    Inter = rechVQF.BeginSearchAndReplace
       
    rechVQF.SourceString = Inter
    rechVQF.SearchString = "%Comment%"
    rechVQF.ReplaceString = mvarComment
    Inter = rechVQF.BeginSearchAndReplace
     
    rechVQF.SourceString = Inter
    rechVQF.SearchString = "%Bitrate%"
    rechVQF.ReplaceString = mvarBitRate
    Inter = rechVQF.BeginSearchAndReplace
    
    rechVQF.SourceString = Inter
    rechVQF.SearchString = "%SampleRate%"
    rechVQF.ReplaceString = mvarSampleRate
    Inter = rechVQF.BeginSearchAndReplace
    
    rechVQF.SourceString = Inter
    rechVQF.SearchString = "%Copyright%"
    rechVQF.ReplaceString = mvarCopyright
    Inter = rechVQF.BeginSearchAndReplace
    
    rechVQF.SourceString = Inter
    rechVQF.SearchString = "%FileSaveAs%"
    rechVQF.ReplaceString = mvarSaveAsFilename
    Inter = rechVQF.BeginSearchAndReplace
    
    GetVQFInfos = Inter

Exit Function

ErrGen:
 ErreurGrave "GetVQFInfos"
 Exit Function

End Function
Public Sub Clear()
Attribute Clear.VB_Description = "Clear all felds"
    mvarRule = ""
    mvarPlaceWhereToPut = 0
    mvarDefaultArtistToUse = ""
    mvarDefaultTitle = ""
End Sub
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
Public Property Let DefaultTitle(ByVal vData As String)
    mvarDefaultTitle = vData
End Property
Public Property Get DefaultTitle() As String
    DefaultTitle = mvarDefaultTitle
End Property
Public Property Let DefaultArtistToUse(ByVal vData As String)
    mvarDefaultArtistToUse = vData
End Property
Public Property Get DefaultArtistToUse() As String
    DefaultArtistToUse = mvarDefaultArtistToUse
End Property
Public Property Let SaveAsFilename(ByVal vData As String)
    mvarSaveAsFilename = vData
End Property
Public Property Get SaveAsFilename() As String
    SaveAsFilename = mvarSaveAsFilename
End Property
Public Property Let Copyright(ByVal vData As String)
    mvarCopyright = vData
End Property
Public Property Get Copyright() As String
    Copyright = mvarCopyright
End Property
Public Property Let Author(ByVal vData As String)
    mvarAuthor = vData
End Property
Public Property Get Author() As String
    Author = mvarAuthor
End Property
Public Property Let Comment(ByVal vData As String)
    mvarComment = vData
End Property
Public Property Get Comment() As String
    Comment = mvarComment
End Property
Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
End Property
Public Property Get Title() As String
    Title = mvarTitle
End Property
Public Property Let SampleRate(ByVal vData As String)
    mvarSampleRate = vData
End Property
Public Property Get SampleRate() As String
    SampleRate = mvarSampleRate
End Property
Public Property Let BitRate(ByVal vData As String)
    mvarBitRate = vData
End Property
Public Property Get BitRate() As String
    BitRate = mvarBitRate
End Property
Public Property Let Mono_Stereo(ByVal vData As String)
    mvarMono_Stereo = vData
End Property
Public Property Get Mono_Stereo() As String
    Mono_Stereo = mvarMono_Stereo
End Property
Public Property Let Quality(ByVal vData As String)
    mvarQuality = vData
End Property
Public Property Get Quality() As String
    Quality = mvarQuality
End Property

Private Sub Class_Initialize()
    Clear
End Sub

Public Sub FillTagsList(lv As ListView)
    Dim itmX As ListItem
    lv.ListItems.Clear
    Set itmX = lv.ListItems.Add(, , "Quality")
    itmX.SubItems(1) = mvarQuality
    Set itmX = lv.ListItems.Add(, , "Mono/Stereo ?")
    itmX.SubItems(1) = mvarMono_Stereo
    Set itmX = lv.ListItems.Add(, , "Bitrate")
    itmX.SubItems(1) = mvarBitRate
    Set itmX = lv.ListItems.Add(, , "Sample Rate")
    itmX.SubItems(1) = mvarSampleRate
    Set itmX = lv.ListItems.Add(, , "Title")
    itmX.SubItems(1) = mvarTitle
    Set itmX = lv.ListItems.Add(, , "Comment")
    itmX.SubItems(1) = mvarComment
    Set itmX = lv.ListItems.Add(, , "Author")
    itmX.SubItems(1) = mvarAuthor
    Set itmX = lv.ListItems.Add(, , "Copyright")
    itmX.SubItems(1) = mvarCopyright
    ResizeLvMp3
End Sub

