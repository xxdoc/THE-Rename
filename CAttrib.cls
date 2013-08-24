VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAttrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
'Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private mvarArchive As Boolean
Private mvarSystem As Boolean
Private mvarReadOnly As Boolean
Private mvarHidden As Boolean
Private mvarSETArchive As Boolean
Private mvarSETReadOnly As Boolean
Private mvarSETHidden As Boolean
Private mvarSETSystem As Boolean
Private mvarAtrOk As Boolean
Public Sub ChangeAttr(Filename As String)
 Dim AttrEncours As Long
 Dim retour As Long
 If mvarAtrOk = True Then
    AttrEncours = GetFileAttributes(Filename)
    If AttrEncours = -1 Then
     Exit Sub
    End If
    If mvarSETArchive Then
        AttrEncours = AttrEncours And (Not FILE_ATTRIBUTE_ARCHIVE)
    Else
        If mvarArchive Then
            AttrEncours = AttrEncours Or FILE_ATTRIBUTE_ARCHIVE
        End If
    End If
     
    If mvarSETReadOnly Then
        AttrEncours = AttrEncours And (Not FILE_ATTRIBUTE_READONLY)
    Else
        If mvarReadOnly Then
            AttrEncours = AttrEncours Or FILE_ATTRIBUTE_READONLY
        End If
    End If
 
    If mvarSETHidden Then
        AttrEncours = AttrEncours And (Not FILE_ATTRIBUTE_HIDDEN)
    Else
        If mvarHidden Then
            AttrEncours = AttrEncours Or FILE_ATTRIBUTE_HIDDEN
        End If
    End If
 
    If mvarSETSystem Then
        AttrEncours = AttrEncours And (Not FILE_ATTRIBUTE_SYSTEM)
    Else
        If mvarSystem Then
            AttrEncours = AttrEncours Or FILE_ATTRIBUTE_SYSTEM
        End If
    End If
    retour = SetFileAttributes(Filename, AttrEncours)
 End If
End Sub

Public Property Let SETSystem(ByVal vData As Boolean)
    mvarSETSystem = vData
End Property

Public Property Get SETSystem() As Boolean
    SETSystem = mvarSETSystem
End Property

Public Property Let SETHidden(ByVal vData As Boolean)
    mvarSETHidden = vData
End Property

Public Property Get SETHidden() As Boolean
    SETHidden = mvarSETHidden
End Property

Public Property Let SETReadOnly(ByVal vData As Boolean)
    mvarSETReadOnly = vData
End Property

Public Property Get SETReadOnly() As Boolean
    SETReadOnly = mvarSETReadOnly
End Property

Public Property Let SETArchive(ByVal vData As Boolean)
    mvarSETArchive = vData
End Property

Public Property Get SETArchive() As Boolean
    SETArchive = mvarSETArchive
End Property

Public Property Let Hidden(ByVal vData As Boolean)
    mvarHidden = vData
End Property

Public Property Get Hidden() As Boolean
    Hidden = mvarHidden
End Property

Public Property Let ReadOnly(ByVal vData As Boolean)
    mvarReadOnly = vData
End Property

Public Property Get ReadOnly() As Boolean
    ReadOnly = mvarReadOnly
End Property

Public Property Let System(ByVal vData As Boolean)
    mvarSystem = vData
End Property

Public Property Get System() As Boolean
    System = mvarSystem
End Property

Public Property Let Archive(ByVal vData As Boolean)
    mvarArchive = vData
End Property

Public Property Get Archive() As Boolean
    Archive = mvarArchive
End Property

Public Property Let AtrOk(ByVal vData As Boolean)
    mvarAtrOk = vData
End Property

Public Property Get AtrOk() As Boolean
    AtrOk = mvarAtrOk
End Property

Public Sub ResetAttrib()
mvarArchive = False
mvarSystem = False
mvarReadOnly = False
mvarHidden = False
mvarSETArchive = False
mvarSETReadOnly = False
mvarSETHidden = False
mvarSETSystem = False
mvarAtrOk = False
End Sub

Private Sub Class_Initialize()
mvarArchive = False
mvarSystem = False
mvarReadOnly = False
mvarHidden = False
mvarSETArchive = False
mvarSETReadOnly = False
mvarSETHidden = False
mvarSETSystem = False
mvarAtrOk = False
End Sub
