VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe de gestion de fichier"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarFileName As String          ' Nom complet sans le chemin
Private mvarFilePath As String          ' Le chemin uniquement
Private mvarCompleteFileName As String  ' chemin + nom
Private mvarFilePrefix As String        ' prefix uniquement
Private mvarFileExtension As String     ' extension uniquement
Private mvarFileSize As Long            ' Taille du fichier en octets
Private mvarCreationDate As Date        ' Date de création
Private mvarLastAccessDate As Date      ' Date de dernier accès
Private mvarLastUpdateDate As Date      ' Date de dernière modif
Private mvarAttrReadOnly As Boolean     ' Attribut Lecture seule
Private mvarAttrHidden As Boolean       ' Attribut Caché
Private mvarAttrSystem As Boolean       ' Attribut System
Private mvarAttrArchive As Boolean      ' Attribut Archive
Private mvarShortName As String         ' Nom court du fichier
Private mvarChAttrib As String          ' Chaine contenant les attributs sous la forme HSRA
Private mvarFileType As Boolean         ' Indique si c'est un fichier ou un répertoire
Public Property Let FileType(ByVal vData As Boolean)
Attribute FileType.VB_Description = "Indique si c'est un fichier ou un répertoire"
    mvarFileType = vData
End Property
Public Property Get FileType() As Boolean
    FileType = mvarFileType
End Property
Public Sub Clear()
mvarFileType = True
mvarShortName = ""
mvarChAttrib = ""
mvarFileName = ""
mvarFilePath = ""
mvarCompleteFileName = ""
mvarFilePrefix = ""
mvarFileExtension = ""
mvarFileSize = 0
mvarAttrReadOnly = False
mvarAttrHidden = False
mvarAttrSystem = False
mvarAttrArchive = False
End Sub
' Renvoie une chaine sous la forme "HSRA" contenant les attributs du fichier
Public Property Get GetStringAttrib() As String
    GetStringAttrib = mvarChAttrib
End Property
' Renvoie le nom court d'un fichier
Public Property Get ShortName() As String
    ShortName = mvarShortName
End Property
Public Sub SetFileName(Filename As String, SonType As Boolean)
    Dim vtmp As Integer
    Dim FicTmp As New clsFindFile
    
    FicTmp.Find Filename
    Clear
    
    mvarFileType = SonType
    mvarShortName = FicTmp.GetShortName
    mvarFileName = Prefixe(Filename) & "." & Suffixe(Filename)
    mvarFilePath = ExtractPath(Filename)
    mvarCompleteFileName = Filename
    mvarFilePrefix = Prefixe(Filename)
    mvarFileExtension = Suffixe(Filename)
    mvarFileSize = FileLen(Filename)
    FicTmp.GetCreationDate mvarCreationDate
    FicTmp.GetLastAccessDate mvarLastAccessDate
    FicTmp.GetLastWriteDate mvarLastUpdateDate
    vtmp = GetAttr(Filename)
    If vtmp And vbReadOnly Then
        mvarAttrReadOnly = True
        mvarChAttrib = mvarChAttrib + "R"
    End If
    If vtmp And vbHidden Then
        mvarAttrHidden = True
        mvarChAttrib = mvarChAttrib + "H"
    End If
    If vtmp And vbSystem Then
        mvarAttrSystem = True
        mvarChAttrib = mvarChAttrib + "S"
    End If
    If vtmp And vbArchive Then
        mvarAttrArchive = True
        mvarChAttrib = mvarChAttrib + "A"
    End If
End Sub
Public Property Get AttrArchive() As Boolean
    AttrArchive = mvarAttrArchive
End Property
Public Property Get AttrSystem() As Boolean
    AttrSystem = mvarAttrSystem
End Property
Public Property Get AttrHidden() As Boolean
    AttrHidden = mvarAttrHidden
End Property
Public Property Get AttrReadOnly() As Boolean
    AttrReadOnly = mvarAttrReadOnly
End Property
Public Property Get LastUpdateDate() As Date
Attribute LastUpdateDate.VB_Description = "Date de dernière modification"
    LastUpdateDate = mvarLastUpdateDate
End Property
Public Property Get LastAccessDate() As Date
Attribute LastAccessDate.VB_Description = "Date de dernier accès"
    LastAccessDate = mvarLastAccessDate
End Property
Public Property Get CreationDate() As Date
Attribute CreationDate.VB_Description = "Date de création"
    CreationDate = mvarCreationDate
End Property
Public Property Get FileSize() As Long
Attribute FileSize.VB_Description = "Taille du fichier en octets"
    FileSize = mvarFileSize
End Property
Public Property Get FileExtension() As String
    FileExtension = mvarFileExtension
End Property
Public Property Get FilePrefix() As String
    FilePrefix = mvarFilePrefix
End Property
Public Property Get CompleteFileName() As String
    CompleteFileName = mvarCompleteFileName
End Property
Public Property Get FilePath() As String
    FilePath = mvarFilePath
End Property
Public Property Get Filename() As String
Attribute Filename.VB_Description = "Nom du fichier"
    Filename = mvarFileName
End Property
Private Sub Class_Initialize()
    Clear
End Sub

