VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDateTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Modify date and time of files"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Option Explicit
Private Const INVALID_HANDLE_VALUE = -1
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpcreation As FILETIME, lpLecture As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Dim wfd As WIN32_FIND_DATA
Private m_tmCreation As Double
Private m_tmAccess As Double
Private m_tmWrite As Double
Private mvarOk As Boolean
Private mvarCreateDOption As Integer
Private mvarCreateDFixed As Date
Private mvarCreateDExif As Integer
Private mvarCreateTExif As Integer
Private mvarCreateDInc1 As Integer
Private mvarCreateDInc2 As Integer
Private mvarCreateDInc3 As Integer
Private mvarCreateTOption As Integer
Private mvarCreateTFixed As String
Private mvarAccessDOption As Integer
Private mvarAccessDFixed As Date
Private mvarAccessDExif As Integer
Private mvarAccessTExif As Integer
Private mvarAccessDInc1 As Integer
Private mvarAccessDInc2 As Integer
Private mvarAccessDInc3 As Integer
Private mvarAccessTOption As Integer
Private mvarAccessTFixed As String
Private mvarWriteDOption As Integer
Private mvarWriteDFixed As Date
Private mvarWriteDExif As Integer
Private mvarWriteTExif As Integer
Private mvarWriteDInc1 As Integer
Private mvarWriteDInc2 As Integer
Private mvarWriteDInc3 As Integer
Private mvarWriteTOption As Integer
Private mvarWriteTFixed As String
Private mvarFileName As String

Private Function GetFT(QDoption As Integer, QDDateFixe As Date, QDInc1 As Integer, QDInc2 As Integer, QDInc3 As Integer, QTOption As Integer, QTTimeFixe As String, QDate As Integer) As FILETIME
Dim anc As SYSTEMTIME
Dim tqtempo As Integer, SonInfo As String
Dim udtSysTime As SYSTEMTIME, udtLocalTime As FILETIME
Dim RetVal As Long, date22 As Date, DateTempo1 As Date, DateTmp2 As Date
Dim ExifDOpt As Integer, ExifTOpt As Integer

Select Case QDate
 Case 1
    ExifDOpt = mvarCreateDExif
    ExifTOpt = mvarCreateTExif
    anc.wYear = Year(m_tmCreation)
    anc.wMonth = Month(m_tmCreation)
    anc.wDay = Day(m_tmCreation)
    anc.wDayOfWeek = Weekday(m_tmCreation) - 1
    anc.wHour = Hour(m_tmCreation)
    anc.wMinute = Minute(m_tmCreation)
    anc.wSecond = Second(m_tmCreation)
 Case 2
    ExifDOpt = mvarAccessDExif
    ExifTOpt = mvarAccessTExif
    anc.wYear = Year(m_tmAccess)
    anc.wMonth = Month(m_tmAccess)
    anc.wDay = Day(m_tmAccess)
    anc.wDayOfWeek = Weekday(m_tmAccess) - 1
    anc.wHour = Hour(m_tmAccess)
    anc.wMinute = Minute(m_tmAccess)
    anc.wSecond = Second(m_tmAccess)
 Case 3
    ExifDOpt = mvarWriteDExif
    ExifTOpt = mvarWriteTExif
    anc.wYear = Year(m_tmWrite)
    anc.wMonth = Month(m_tmWrite)
    anc.wDay = Day(m_tmWrite)
    anc.wDayOfWeek = Weekday(m_tmWrite) - 1
    anc.wHour = Hour(m_tmWrite)
    anc.wMinute = Minute(m_tmWrite)
    anc.wSecond = Second(m_tmWrite)
End Select

' Cas de la date
Select Case QDoption
    Case 0 ' "Keep" date
        With udtSysTime
            .wYear = anc.wYear
            .wMonth = anc.wMonth
            .wDay = anc.wDay
            .wDayOfWeek = anc.wDayOfWeek
        End With
 
    Case 1 ' "Set to", date fixe
        With udtSysTime
            .wYear = Year(QDDateFixe)
            .wMonth = Month(QDDateFixe)
            .wDay = Day(QDDateFixe)
            .wDayOfWeek = Weekday(QDDateFixe) - 1
        End With
 
    Case 2 ' Increase/Decrease
        DateTempo1 = CDate(Str$(anc.wYear) + "/" + Str$(anc.wMonth) + "/" + Str$(anc.wDay))
        tqtempo = QDInc2
        If QDInc1 = 1 Then ' Decrease
            tqtempo = tqtempo * (-1)
        End If
        Select Case QDInc3
            Case 0 ' Day
                DateTempo1 = DateAdd("d", tqtempo, DateTempo1)
            Case 1 ' Month
                DateTempo1 = DateAdd("m", tqtempo, DateTempo1)
            Case 2 ' Year
                DateTempo1 = DateAdd("yyyy", tqtempo, DateTempo1)
        End Select
        With udtSysTime
            .wYear = Year(DateTempo1)
            .wMonth = Month(DateTempo1)
            .wDay = Day(DateTempo1)
            .wDayOfWeek = Weekday(DateTempo1) - 1
        End With
    
    Case 3 ' Use EXIF date
        SonInfo = PicEXIF.GetEXIFInfos(mvarFileName, False)
        If IsDate(PicEXIF.GetDate(ExifDOpt, 1)) Then
            DateTmp2 = PicEXIF.GetDate(ExifDOpt, 1)
            With udtSysTime
                .wYear = Year(DateTmp2)
                .wMonth = Month(DateTmp2)
                .wDay = Day(DateTmp2)
                .wDayOfWeek = Weekday(DateTmp2) - 1
            End With
            
        Else    ' Date invalide, on garde la date actuelle
            With udtSysTime
                .wYear = anc.wYear
                .wMonth = anc.wMonth
                .wDay = anc.wDay
                .wDayOfWeek = anc.wDayOfWeek
            End With
        End If
                
End Select

' Cas de l'heure
Select Case QTOption
    Case 0 ' Keep time
        With udtSysTime
            .wHour = anc.wHour
            .wMinute = anc.wMinute
            .wSecond = anc.wSecond
        End With
 
    Case 1 ' Set to
        date22 = QTTimeFixe
        With udtSysTime
            .wHour = Hour(date22)
            .wMinute = Minute(date22)
            .wSecond = Second(date22)
        End With
        
    Case 2 ' Use EXIF time
        SonInfo = PicEXIF.GetEXIFInfos(mvarFileName, False)
        If IsDate(PicEXIF.GetDate(ExifTOpt, 2)) Then
            date22 = PicEXIF.GetDate(ExifTOpt, 2)
            With udtSysTime
                .wHour = Hour(date22)
                .wMinute = Minute(date22)
                .wSecond = Second(date22)
            End With
        Else    ' Heure invalide, on garde l'heure actuelle
            With udtSysTime
                .wHour = anc.wHour
                .wMinute = anc.wMinute
                .wSecond = anc.wSecond
            End With
        End If
    
End Select
RetVal = SystemTimeToFileTime(udtSysTime, udtLocalTime)
RetVal = LocalFileTimeToFileTime(udtLocalTime, GetFT)
End Function
' Modification des dates et heure d'un fichier selon les options.
' Note : Le fichier est sensé exister
Public Sub SetFileDateTime(Filename As String)
Attribute SetFileDateTime.VB_Description = "Modifie les date et heures du fichier passé en paramètres"
Dim hSearch As Long
Dim hFile As Long, R As Long
Dim Ftc As FILETIME, Fta As FILETIME, Ftw As FILETIME
If mvarOk = False Then
 Exit Sub
End If

mvarFileName = Filename
hSearch = FindFirstFile(Filename, wfd)
If hSearch <> INVALID_HANDLE_VALUE Then
      Call FindClose(hSearch)
Else
 Exit Sub
End If

m_tmCreation = FileTimeToDouble(wfd.ftCreationTime, True)
m_tmAccess = FileTimeToDouble(wfd.ftLastAccessTime, True)
m_tmWrite = FileTimeToDouble(wfd.ftLastWriteTime, True)

hFile = CreateFile(Filename, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
If hFile = -1 Then
 Exit Sub
End If

Ftc = GetFT(mvarCreateDOption, mvarCreateDFixed, mvarCreateDInc1, mvarCreateDInc2, mvarCreateDInc3, mvarCreateTOption, mvarCreateTFixed, 1)
Fta = GetFT(mvarAccessDOption, mvarAccessDFixed, mvarAccessDInc1, mvarAccessDInc2, mvarAccessDInc3, mvarAccessTOption, mvarAccessTFixed, 2)
Ftw = GetFT(mvarWriteDOption, mvarWriteDFixed, mvarWriteDInc1, mvarWriteDInc2, mvarWriteDInc3, mvarWriteTOption, mvarWriteTFixed, 3)
' La modif en elle même des dates et des heures
R = SetFileTime(hFile, Ftc, Fta, Ftw)
Call CloseHandle(hFile)
End Sub

Public Property Let WriteTFixed(ByVal vData As String)
    mvarWriteTFixed = vData
End Property

Public Property Get WriteTFixed() As String
    WriteTFixed = mvarWriteTFixed
End Property

Public Property Let WriteTOption(ByVal vData As Integer)
    mvarWriteTOption = vData
End Property

Public Property Get WriteTOption() As Integer
    WriteTOption = mvarWriteTOption
End Property
Public Property Let WriteDInc3(ByVal vData As Integer)
    mvarWriteDInc3 = vData
End Property
Public Property Get WriteDInc3() As Integer
    WriteDInc3 = mvarWriteDInc3
End Property
Public Property Let WriteDInc2(ByVal vData As Integer)
    mvarWriteDInc2 = vData
End Property
Public Property Get WriteDInc2() As Integer
    WriteDInc2 = mvarWriteDInc2
End Property
Public Property Let WriteDInc1(ByVal vData As Integer)
    mvarWriteDInc1 = vData
End Property
Public Property Get WriteDInc1() As Integer
    WriteDInc1 = mvarWriteDInc1
End Property
Public Property Let WriteDFixed(ByVal vData As Date)
    mvarWriteDFixed = vData
End Property
Public Property Get WriteDFixed() As Date
    WriteDFixed = mvarWriteDFixed
End Property
Public Property Let WriteDOption(ByVal vData As Integer)
    mvarWriteDOption = vData
End Property
Public Property Get WriteDOption() As Integer
    WriteDOption = mvarWriteDOption
End Property
Public Property Let AccessTFixed(ByVal vData As String)
    mvarAccessTFixed = vData
End Property
Public Property Get AccessTFixed() As String
    AccessTFixed = mvarAccessTFixed
End Property
Public Property Let AccessTOption(ByVal vData As Integer)
    mvarAccessTOption = vData
End Property
Public Property Get AccessTOption() As Integer
    AccessTOption = mvarAccessTOption
End Property
Public Property Let AccessDInc3(ByVal vData As Integer)
    mvarAccessDInc3 = vData
End Property
Public Property Get AccessDInc3() As Integer
    AccessDInc3 = mvarAccessDInc3
End Property
Public Property Let AccessDInc2(ByVal vData As Integer)
    mvarAccessDInc2 = vData
End Property
Public Property Get AccessDInc2() As Integer
    AccessDInc2 = mvarAccessDInc2
End Property
Public Property Let AccessDInc1(ByVal vData As Integer)
    mvarAccessDInc1 = vData
End Property
Public Property Get AccessDInc1() As Integer
    AccessDInc1 = mvarAccessDInc1
End Property
Public Property Let AccessDFixed(ByVal vData As Date)
    mvarAccessDFixed = vData
End Property
Public Property Get AccessDFixed() As Date
    AccessDFixed = mvarAccessDFixed
End Property
Public Property Let AccessDOption(ByVal vData As Integer)
    mvarAccessDOption = vData
End Property
Public Property Get AccessDOption() As Integer
    AccessDOption = mvarAccessDOption
End Property
Public Property Let CreateTFixed(ByVal vData As String)
Attribute CreateTFixed.VB_Description = "Heure Fixe"
    mvarCreateTFixed = vData
End Property
Public Property Get CreateTFixed() As String
    CreateTFixed = mvarCreateTFixed
End Property
Public Property Let CreateTOption(ByVal vData As Integer)
Attribute CreateTOption.VB_Description = "Heure\r\n1=Keep\r\n2=Set to"
    mvarCreateTOption = vData
End Property
Public Property Get CreateTOption() As Integer
    CreateTOption = mvarCreateTOption
End Property
Public Property Let CreateDInc3(ByVal vData As Integer)
Attribute CreateDInc3.VB_Description = "Date incrément/décrément\r\n1=Day\r\n2=Month\r\n3=Year"
    mvarCreateDInc3 = vData
End Property
Public Property Get CreateDInc3() As Integer
    CreateDInc3 = mvarCreateDInc3
End Property
Public Property Let CreateDInc2(ByVal vData As Integer)
Attribute CreateDInc2.VB_Description = "Date, nb de jours, de mois ou d'années"
    mvarCreateDInc2 = vData
End Property
Public Property Get CreateDInc2() As Integer
    CreateDInc2 = mvarCreateDInc2
End Property
Public Property Let CreateDInc1(ByVal vData As Integer)
Attribute CreateDInc1.VB_Description = "Date, Increase/Decrease ?"
    mvarCreateDInc1 = vData
End Property
Public Property Get CreateDInc1() As Integer
    CreateDInc1 = mvarCreateDInc1
End Property
Public Property Let CreateDFixed(ByVal vData As Date)
Attribute CreateDFixed.VB_Description = "Date Fixe"
    mvarCreateDFixed = vData
End Property
Public Property Get CreateDFixed() As Date
    CreateDFixed = mvarCreateDFixed
End Property
Public Property Let CreateDExif(ByVal vData As Integer)
    mvarCreateDExif = vData
End Property
Public Property Get CreateDExif() As Integer
    CreateDExif = mvarCreateDExif
End Property
Public Property Let AccessDExif(ByVal vData As Integer)
    mvarAccessDExif = vData
End Property
Public Property Get AccessDExif() As Integer
    AccessDExif = mvarAccessDExif
End Property
Public Property Let WriteDExif(ByVal vData As Integer)
    mvarWriteDExif = vData
End Property
Public Property Get WriteDExif() As Integer
    WriteDExif = mvarWriteDExif
End Property
Public Property Let CreateTExif(ByVal vData As Integer)
    mvarCreateTExif = vData
End Property
Public Property Get CreateTExif() As Integer
    CreateTExif = mvarCreateTExif
End Property
Public Property Let AccessTExif(ByVal vData As Integer)
    mvarAccessTExif = vData
End Property
Public Property Get AccessTExif() As Integer
    AccessTExif = mvarAccessTExif
End Property
Public Property Let WriteTExif(ByVal vData As Integer)
    mvarWriteTExif = vData
End Property
Public Property Get WriteTExif() As Integer
    WriteTExif = mvarWriteTExif
End Property

Public Property Let CreateDOption(ByVal vData As Integer)
Attribute CreateDOption.VB_Description = "Date\r\n1=Keep\r\n2=Set to\r\n3=Increase/Decrease"
    mvarCreateDOption = vData
End Property
Public Property Get CreateDOption() As Integer
    CreateDOption = mvarCreateDOption
End Property
Private Sub Class_Initialize()
    mvarOk = False
    mvarCreateDOption = 0
    mvarCreateDFixed = Date
    mvarCreateDInc1 = 0
    mvarCreateDInc2 = 1
    mvarCreateDInc3 = 0
    mvarCreateTOption = 0
    mvarCreateTFixed = Time
    mvarAccessDOption = 0
    mvarAccessDFixed = Date
    mvarAccessDInc1 = 0
    mvarAccessDInc2 = 1
    mvarAccessDInc3 = 0
    mvarAccessTOption = 0
    mvarAccessTFixed = Time
    mvarWriteDOption = 0
    mvarWriteDFixed = Date
    mvarWriteDInc1 = 0
    mvarWriteDInc2 = 1
    mvarWriteDInc3 = 0
    mvarWriteTOption = 0
    mvarWriteTFixed = Time
End Sub
Public Property Let DTOk(ByVal vData As Boolean)
    mvarOk = vData
End Property
Public Property Get DTOk() As Boolean
    DTOk = mvarOk
End Property
Private Function FileTimeToDouble(ftUTC As FILETIME, Localize As Boolean) As Double
   Dim ft As FILETIME
   Dim st As SYSTEMTIME
   If Localize Then
      Call FileTimeToLocalFileTime(ftUTC, ft)
   Else
      ft = ftUTC
   End If
   Call FileTimeToSystemTime(ft, st)
   FileTimeToDouble = DateSerial(st.wYear, st.wMonth, st.wDay) + TimeSerial(st.wHour, st.wMinute, st.wSecond)
End Function

