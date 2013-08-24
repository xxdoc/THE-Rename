VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SharedClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarCountFiles As Long 'local copy
Private mvarSelectedCount As Long 'local copy
Private mvarFileTraited As Long 'local copy
Private mvarOrCompleteFileName As String 'local copy
Private mvarOrFileName As String 'local copy
Private mvarOrPath As String 'local copy
Private mvarOrPrefix As String 'local copy
Private mvarOrExtension As String 'local copy
Private mvarFileSize As Long 'local copy
Private mvarFileDate As Date 'local copy
Private mvarFileAttrib As String 'local copy
Private mvarFileType As Integer 'local copy
Private mvarRunMode As Integer 'local copy
Private mvarNwCompleteFileName As String 'local copy
Private mvarNwFileName As String 'local copy
Private mvarNwPath As String 'local copy
Private mvarNwPrefix As String 'local copy
Private mvarNwExtension As String 'local copy
Private mvarOkRename As Boolean 'local copy
Private mvarForceNewName As String 'local copy
Private mvarTheRenColl As Collection 'local copy
Private mvarCommandLine As String 'local copy
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent BeforeRename[(arg1, arg2, ... , argn)]
Public Event BeforeRename(ByVal OldName As String, ByVal NewName As String)
Attribute BeforeRename.VB_Description = "Occurs before the program renames the file"
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent AfterRename[(arg1, arg2, ... , argn)]
Public Event AfterRename(ByVal OldName As String, ByVal NewName As String)
'To fire this event, use RaiseEvent with the following syntax:
'RaiseEvent ErrorOccured[(arg1, arg2, ... , argn)]
Public Event ErrorOccured(ByVal ErrText As String)
Public Property Get CommandLine() As String
Attribute CommandLine.VB_Description = "Returns the command's line when THE Rename was launched"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommandLine
   CommandLine = Command$
End Property
Public Function EXBefore(ByVal Expr As String) As String
    EXBefore = "EXBefore"
End Function
Public Function EXAfter(ByVal Expr As String) As String
    EXAfter = "CA MARCHE !"
End Function
Public Function CyclicSelection() As String
    CyclicSelection = "CyclicSelection"
End Function
Public Function curprefix() As String
Attribute curprefix.VB_Description = "Take current file's prefix"
    curprefix = "curprefix"
End Function
Public Function curext() As String
Attribute curext.VB_Description = "Take current file's extension"
    curext = "curext"
End Function
Public Property Set TheRenColl(ByVal vData As Collection)
Attribute TheRenColl.VB_Description = "Collection libre d'utilisation"
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.TheRenColl = Form1
    Set mvarTheRenColl = vData
End Property
Public Property Get TheRenColl() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TheRenColl
    Set TheRenColl = mvarTheRenColl
End Property
Public Property Let ForceNewName(ByVal vData As String)
Attribute ForceNewName.VB_Description = "Nouveau nom qu'il faut donner quand c'est l'utilisateur qui veut renommer le fichier"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForceNewName = 5
    mvarForceNewName = vData
End Property
Public Property Get ForceNewName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForceNewName
    ForceNewName = mvarForceNewName
End Property
Public Property Let OkRename(ByVal vData As Boolean)
Attribute OkRename.VB_Description = "Faut'il renommer le fichier ?"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OkRename = 5
    mvarOkRename = vData
End Property
Public Property Get OkRename() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OkRename
    OkRename = mvarOkRename
End Property
Public Property Let NwExtension(ByVal vData As String)
Attribute NwExtension.VB_Description = "Extension du nouveau fichier"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NwExtension = 5
    mvarNwExtension = vData
End Property
Public Property Get NwExtension() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NwExtension
    NwExtension = mvarNwExtension
End Property
Public Property Let NwPrefix(ByVal vData As String)
Attribute NwPrefix.VB_Description = "Préfixe du nouveau fichier"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NwPrefix = 5
    mvarNwPrefix = vData
End Property
Public Property Get NwPrefix() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NwPrefix
    NwPrefix = mvarNwPrefix
End Property
Public Property Let NwPath(ByVal vData As String)
Attribute NwPath.VB_Description = "Chemin du nouveau fichier"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NwPath = 5
    mvarNwPath = vData
End Property
Public Property Get NwPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NwPath
    NwPath = mvarNwPath
End Property
Public Property Let NwFileName(ByVal vData As String)
Attribute NwFileName.VB_Description = "Nom du nouveau fichier  (prefixe+extension)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NwFileName = 5
    mvarNwFileName = vData
End Property
Public Property Get NwFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NwFileName
    NwFileName = mvarNwFileName
End Property
Public Property Let NwCompleteFileName(ByVal vData As String)
Attribute NwCompleteFileName.VB_Description = "Nom complet du nouveau fichier (chemin+prefixe+extension)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NwCompleteFileName = 5
    mvarNwCompleteFileName = vData
End Property
Public Property Get NwCompleteFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NwCompleteFileName
    NwCompleteFileName = mvarNwCompleteFileName
End Property
Public Property Let RunMode(ByVal vData As Integer)
Attribute RunMode.VB_Description = "Mode d'utilisation, normal ou récursif"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RunMode = 5
    mvarRunMode = vData
End Property
Public Property Get RunMode() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RunMode
    RunMode = mvarRunMode
End Property
Public Property Let FileType(ByVal vData As Integer)
Attribute FileType.VB_Description = "Type, Fichier ou Répertoire ?"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileType = 5
    mvarFileType = vData
End Property
Public Property Get FileType() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileType
    FileType = mvarFileType
End Property
Public Property Let FileAttrib(ByVal vData As String)
Attribute FileAttrib.VB_Description = "Les attributs du fichier"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileAttrib = 5
    mvarFileAttrib = vData
End Property
Public Property Get FileAttrib() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileAttrib
    FileAttrib = mvarFileAttrib
End Property
Public Property Let FileDate(ByVal vData As Date)
Attribute FileDate.VB_Description = "Date du fichier en cours de traitement"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileDate = 5
    mvarFileDate = vData
End Property
Public Property Get FileDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileDate
    FileDate = mvarFileDate
End Property
Public Property Let FileSize(ByVal vData As Long)
Attribute FileSize.VB_Description = "Taille du fichier en cours de traitement"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileSize = 5
    mvarFileSize = vData
End Property
Public Property Get FileSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSize
    FileSize = mvarFileSize
End Property
Public Property Let OrExtension(ByVal vData As String)
Attribute OrExtension.VB_Description = "Extension du fichier en cours de traitement"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrExtension = 5
    mvarOrExtension = vData
End Property
Public Property Get OrExtension() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrExtension
    OrExtension = mvarOrExtension
End Property
Public Property Let OrPrefix(ByVal vData As String)
Attribute OrPrefix.VB_Description = "Préfixe du fichier en cours de traitement"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrPrefix = 5
    mvarOrPrefix = vData
End Property
Public Property Get OrPrefix() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrPrefix
    OrPrefix = mvarOrPrefix
End Property
Public Property Let OrPath(ByVal vData As String)
Attribute OrPath.VB_Description = "Chemin du fichier en cours de traitement"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrPath = 5
    mvarOrPath = vData
End Property
Public Property Get OrPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrPath
    OrPath = mvarOrPath
End Property
Public Property Let OrFileName(ByVal vData As String)
Attribute OrFileName.VB_Description = "Nom du fichier en cours de traitement (prefixe+extension)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrFileName = 5
    mvarOrFileName = vData
End Property
Public Property Get OrFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrFileName
    OrFileName = mvarOrFileName
End Property
Public Property Let OrCompleteFileName(ByVal vData As String)
Attribute OrCompleteFileName.VB_Description = "Nom complet du fichier en cours de traitement (chemin+prefixe+extension)"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.OrCompleteFileName = 5
    mvarOrCompleteFileName = vData
End Property
Public Property Get OrCompleteFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OrCompleteFileName
    OrCompleteFileName = mvarOrCompleteFileName
End Property
Public Property Let FileTraited(ByVal vData As Long)
Attribute FileTraited.VB_Description = "Indice du fichier en cours de traitement"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileTraited = 5
    mvarFileTraited = vData
End Property
Public Property Get FileTraited() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileTraited
    FileTraited = mvarFileTraited
End Property
Public Property Get SelectedCount() As Long
Attribute SelectedCount.VB_Description = "Nombre de fichiers sélectionnés"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedCount
    SelectedCount = LVGetCountSelected(RENAME.ListView1)
End Property
Public Property Get CountFiles() As Long
Attribute CountFiles.VB_Description = "Nombre total de fichiers"
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CountFiles
    CountFiles = Val(RENAME.état.Panels(3).Text)
End Property

