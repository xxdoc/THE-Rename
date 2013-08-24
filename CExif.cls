VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe pour obtenir les tags EXIF d'une image"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Rule"
Option Explicit
Private sFilename As String
Private mvarExifVersion As String
Private mvarAperture As String
Private mvarBrightness As String
Private mvarCompressedBitsPerPixel As String
Private mvarCopyright As String
Private mvarDateTime As String
Private mvarDateTimeOrig As String
Private mvarDateTimeDigitized As String
Private mvarDateTimeDigitizedOrig As String
Private mvarDateTimeOriginal As String
Private mvarDateTimeOriginalOrig As String
Private mvarVersion As String
Private mvarExposureBias As String
Private mvarExposureProgram As String
Private mvarExposureTime As String
Private mvarFirmwareVersion As String
Private mvarFlash As String
Private mvarFNumber As String
Private mvarFocalLength As String
Private mvarFocalPlaneResolutionUnit As String
Private mvarFocalPlaneXResolution As String
Private mvarFocalPlaneYResolution As String
Private mvarImageDescription As String
Private mvarImageHeight As String
Private mvarImageWidth As String
Private mvarISOSpeedRatings As String
Private mvarMake As String
Private mvarMaxAperture As String
Private mvarMeteringMode As String
Private mvarModel As String
Private mvarOrientation As String
Private mvarRelatedSoundFile As String
Private mvarResolutionUnit As String
Private mvarShutterSpeed As String
Private mvarSubjectDistance As String
Private mvarWhiteBalance As String
Private mvarXResolution As String
Private mvarYResolution As String
Private mvarRule As String
Private mvarPlaceWhereToPut As Integer
Private mvarUseEXIF As Boolean
' Renvoie soit la date soit l'heure et on précise de quel champs il s'agit
Public Property Get GetDate(QDate As Integer, DorT As Integer) As String
    Dim vtmp As String
    Dim LaDate As Date
    
    If QDate < 0 Or QDate > 3 Then
        GetDate = ""
        Exit Property
    End If
    
    If DorT < 1 Or DorT > 2 Then
        GetDate = ""
        Exit Property
    End If
    
    Select Case QDate
        Case 0  ' Date time
            vtmp = mvarDateTimeOrig
        Case 1  ' DateTime Digitized
            vtmp = mvarDateTimeDigitizedOrig
        Case 2  ' DateTime Original
            vtmp = mvarDateTimeOriginalOrig
    End Select
    If Trim$(vtmp) = "" Then
        GetDate = ""
        Exit Property
    End If
    
    If DorT = 1 Then    ' On veut la date
        If Trim$(Left$(vtmp, 10)) <> "" Then
            vtmp = Replace(Left$(vtmp, 10), ":", "-")
            LaDate = CDate(vtmp)
            If IsDate(LaDate) Then
                GetDate = LaDate
            Else
                GetDate = ""
            End If
        Else
            GetDate = ""
        End If
    Else                ' on veut l'heure
        If Trim$(Right$(vtmp, 8)) <> "" Then
            GetDate = CDate(Right$(vtmp, 8))
        Else
            GetDate = ""
        End If
    End If
End Property
Public Property Let UseEXIF(ByVal vData As Boolean)
    mvarUseEXIF = vData
End Property
Public Property Get UseEXIF() As Boolean
    UseEXIF = mvarUseEXIF
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
Public Property Get YResolution() As String
    YResolution = mvarYResolution
End Property
Public Property Get XResolution() As String
    XResolution = mvarXResolution
End Property
Public Property Get WhiteBalance() As String
    WhiteBalance = mvarWhiteBalance
End Property
Public Property Get SubjectDistance() As String
    SubjectDistance = mvarSubjectDistance
End Property
Public Property Get ShutterSpeed() As String
    ShutterSpeed = mvarShutterSpeed
End Property
Public Property Get ResolutionUnit() As String
    ResolutionUnit = mvarResolutionUnit
End Property
Public Property Get RelatedSoundFile() As String
    RelatedSoundFile = mvarRelatedSoundFile
End Property
Public Property Get Orientation() As String
    Orientation = mvarOrientation
End Property
Public Property Get Model() As String
    Model = mvarModel
End Property
Public Property Get MeteringMode() As String
    MeteringMode = mvarMeteringMode
End Property
Public Property Get MaxAperture() As String
    MaxAperture = mvarMaxAperture
End Property
Public Property Get Make() As String
    Make = mvarMake
End Property
Public Property Get ISOSpeedRatings() As String
    ISOSpeedRatings = mvarISOSpeedRatings
End Property
Public Property Get ImageWidth() As String
    ImageWidth = mvarImageWidth
End Property
Public Property Get ImageHeight() As String
    ImageHeight = mvarImageHeight
End Property
Public Property Get ImageDescription() As String
    ImageDescription = mvarImageDescription
End Property
Public Property Get FocalPlaneYResolution() As String
    FocalPlaneYResolution = mvarFocalPlaneYResolution
End Property
Public Property Get FocalPlaneXResolution() As String
    FocalPlaneXResolution = mvarFocalPlaneXResolution
End Property
Public Property Get FocalPlaneResolutionUnit() As String
    FocalPlaneResolutionUnit = mvarFocalPlaneResolutionUnit
End Property
Public Property Get FocalLength() As String
    FocalLength = mvarFocalLength
End Property
Public Property Get FNumber() As String
    FNumber = mvarFNumber
End Property
Public Property Get Flash() As String
    Flash = mvarFlash
End Property
Public Property Get FirmwareVersion() As String
    FirmwareVersion = mvarFirmwareVersion
End Property
Public Property Get ExposureTime() As String
    ExposureTime = mvarExposureTime
End Property
Public Property Get ExposureProgram() As String
    ExposureProgram = mvarExposureProgram
End Property
Public Property Get ExposureBias() As String
    ExposureBias = mvarExposureBias
End Property
Public Property Get Version() As String
    Version = mvarVersion
End Property
Public Property Get DateTimeOriginal() As String
    DateTimeOriginal = mvarDateTimeOriginal
End Property
Public Property Get DateTimeDigitized() As String
    DateTimeDigitized = mvarDateTimeDigitized
End Property
Public Property Get DateTime() As String
    DateTime = mvarDateTime
End Property
Public Property Get Copyright() As String
    Copyright = mvarCopyright
End Property
Public Property Get CompressedBitsPerPixel() As String
    CompressedBitsPerPixel = mvarCompressedBitsPerPixel
End Property
Public Property Get Brightness() As String
    Brightness = mvarBrightness
End Property
Public Property Get Aperture() As String
    Aperture = mvarAperture
End Property
Public Property Get ExifVersion() As String
    ExifVersion = mvarExifVersion
End Property
Public Function GetEXIFInfos(ByVal sPassFileName As String, Optional Tout As Boolean = True) As String
On Error Resume Next
Dim Inter As String
Dim col As ExifTags
Dim tg As ExifTag
Dim exobj As ExifPage

If sPassFileName <> sFilename Then
    ClearTags
    Set exobj = New ExifPage
    Set col = exobj.ExtractExifInfo(sPassFileName)
        For Each tg In col
            Select Case Trim$(RInternalSpaces(RemoveMultipleSpacing(tg.Name)))
                Case "ExifVersion"
                    mvarExifVersion = tg.Value
                Case "Aperture"
                    mvarAperture = tg.Value
                Case "Brightness"
                    mvarBrightness = tg.Value
                Case "CompressedBitsPerPixel"
                    mvarCompressedBitsPerPixel = tg.Value
                Case "Copyright"
                    mvarCopyright = tg.Value
                Case "DateTime"
                    mvarDateTime = DatesDelimitors(tg.Value)
                    mvarDateTimeOrig = tg.Value
                Case "DateTimeDigitized"
                    mvarDateTimeDigitized = DatesDelimitors(tg.Value)
                    mvarDateTimeDigitizedOrig = tg.Value
                Case "DateTimeOriginal"
                    mvarDateTimeOriginal = DatesDelimitors(tg.Value)
                    mvarDateTimeOriginalOrig = tg.Value
                Case "Version"
                    mvarVersion = tg.Value
                Case "ExposureBias"
                    mvarExposureBias = tg.Value
                Case "ExposureProgram"
                    mvarExposureProgram = tg.Value
                Case "ExposureTime"
                    mvarExposureTime = tg.Value
                Case "FirmwareVersion"
                    mvarFirmwareVersion = tg.Value
                Case "Flash"
                    mvarFlash = tg.Value
                Case "FNumber"
                    mvarFNumber = tg.Value
                Case "FocalLength"
                    mvarFocalLength = tg.Value
                Case "FocalPlaneResolutionUnit"
                    mvarFocalPlaneResolutionUnit = tg.Value
                Case "FocalPlaneXResolution"
                    mvarFocalPlaneXResolution = tg.Value
                Case "FocalPlaneYResolution"
                    mvarFocalPlaneYResolution = tg.Value
                Case "ImageDescription"
                    mvarImageDescription = tg.Value
                Case "ImageHeight"
                    mvarImageHeight = tg.Value
                Case "ImageWidth"
                    mvarImageWidth = tg.Value
                Case "ISOSpeedRatings"
                    mvarISOSpeedRatings = tg.Value
                Case "Make"
                    mvarMake = tg.Value
                Case "MaxAperture"
                    mvarMaxAperture = tg.Value
                Case "MeteringMode"
                    mvarMeteringMode = tg.Value
                Case "Model"
                    mvarModel = tg.Value
                Case "Orientation"
                    mvarOrientation = tg.Value
                Case "RelatedSoundFile"
                    mvarRelatedSoundFile = tg.Value
                Case "ResolutionUnit"
                    mvarResolutionUnit = tg.Value
                Case "ShutterSpeed"
                    mvarShutterSpeed = tg.Value
                Case "SubjectDistance"
                    mvarSubjectDistance = tg.Value
                Case "WhiteBalance"
                    mvarWhiteBalance = tg.Value
                Case "XResolution"
                    mvarXResolution = tg.Value
                Case "YResolution"
                    mvarYResolution = tg.Value
            End Select
        Next
End If
On Error GoTo ErrGen
sFilename = sPassFileName

If Tout = False Then    ' la méthode n'a pas été appelée par une fenêtre mais par une fonction du Free Form
    Exit Function
End If

Inter = mvarRule
Inter = Replace(Inter, "%Aperture%", mvarAperture, , , vbTextCompare)
Inter = Replace(Inter, "%Brightness%", mvarBrightness, , , vbTextCompare)
Inter = Replace(Inter, "%CompressedBitsPerPixel%", mvarCompressedBitsPerPixel, , , vbTextCompare)
Inter = Replace(Inter, "%Copyright%", mvarCopyright, , , vbTextCompare)
Inter = Replace(Inter, "%DateTime%", mvarDateTime, , , vbTextCompare)
Inter = Replace(Inter, "%DateTimeDigitized%", mvarDateTimeDigitized, , , vbTextCompare)
Inter = Replace(Inter, "%DateTimeOriginal%", mvarDateTimeOriginal, , , vbTextCompare)
Inter = Replace(Inter, "%ExifVersion%", mvarExifVersion, , , vbTextCompare)
Inter = Replace(Inter, "%ExposureBias%", mvarExposureBias, , , vbTextCompare)
Inter = Replace(Inter, "%ExposureProgram%", mvarExposureProgram, , , vbTextCompare)
Inter = Replace(Inter, "%ExposureTime%", mvarExposureTime, , , vbTextCompare)
Inter = Replace(Inter, "%FirmwareVersion%", mvarFirmwareVersion, , , vbTextCompare)
Inter = Replace(Inter, "%Flash%", mvarFlash, , , vbTextCompare)
Inter = Replace(Inter, "%FNumber%", mvarFNumber, , , vbTextCompare)
Inter = Replace(Inter, "%FocalLength%", mvarFocalLength, , , vbTextCompare)
Inter = Replace(Inter, "%FocalPlaneResolutionUnit%", mvarFocalPlaneResolutionUnit, , , vbTextCompare)
Inter = Replace(Inter, "%FocalPlaneXResolution%", mvarFocalPlaneXResolution, , , vbTextCompare)
Inter = Replace(Inter, "%FocalPlaneYResolution%", mvarFocalPlaneYResolution, , , vbTextCompare)
Inter = Replace(Inter, "%ImageDescription%", mvarImageDescription, , , vbTextCompare)
Inter = Replace(Inter, "%ImageHeight%", mvarImageHeight, , , vbTextCompare)
Inter = Replace(Inter, "%ImageWidth%", mvarImageWidth, , , vbTextCompare)
Inter = Replace(Inter, "%ISOSpeedRatings%", mvarISOSpeedRatings, , , vbTextCompare)
Inter = Replace(Inter, "%Make%", mvarMake, , , vbTextCompare)
Inter = Replace(Inter, "%MaxAperture%", mvarMaxAperture, , , vbTextCompare)
Inter = Replace(Inter, "%MeteringMode%", mvarMeteringMode, , , vbTextCompare)
Inter = Replace(Inter, "%Model%", mvarModel, , , vbTextCompare)
Inter = Replace(Inter, "%Orientation%", mvarOrientation, , , vbTextCompare)
Inter = Replace(Inter, "%RelatedSoundFile%", mvarRelatedSoundFile, , , vbTextCompare)
Inter = Replace(Inter, "%ResolutionUnit%", mvarResolutionUnit, , , vbTextCompare)
Inter = Replace(Inter, "%ShutterSpeed%", mvarShutterSpeed, , , vbTextCompare)
Inter = Replace(Inter, "%SubjectDistance%", mvarSubjectDistance, , , vbTextCompare)
Inter = Replace(Inter, "%Version%", mvarVersion, , , vbTextCompare)
Inter = Replace(Inter, "%WhiteBalance%", mvarWhiteBalance, , , vbTextCompare)
Inter = Replace(Inter, "%XResolution%", mvarXResolution, , , vbTextCompare)
Inter = Replace(Inter, "%YResolution%", mvarYResolution, , , vbTextCompare)
GetEXIFInfos = Inter
Exit Function
    
ErrGen:
 ErreurGrave "GetEXIFInfos"
 Exit Function

End Function

Private Sub ClearTags()
    mvarExifVersion = ""
    mvarAperture = ""
    mvarBrightness = ""
    mvarCompressedBitsPerPixel = ""
    mvarCopyright = ""
    mvarDateTime = ""
    mvarDateTimeOrig = ""
    mvarDateTimeDigitized = ""
    mvarDateTimeDigitizedOrig = ""
    mvarDateTimeOriginal = ""
    mvarDateTimeOriginalOrig = ""
    mvarVersion = ""
    mvarExposureBias = ""
    mvarExposureProgram = ""
    mvarExposureTime = ""
    mvarFirmwareVersion = ""
    mvarFlash = ""
    mvarFNumber = ""
    mvarFocalLength = ""
    mvarFocalPlaneResolutionUnit = ""
    mvarFocalPlaneXResolution = ""
    mvarFocalPlaneYResolution = ""
    mvarImageDescription = ""
    mvarImageHeight = ""
    mvarImageWidth = ""
    mvarISOSpeedRatings = ""
    mvarMake = ""
    mvarMaxAperture = ""
    mvarMeteringMode = ""
    mvarModel = ""
    mvarOrientation = ""
    mvarRelatedSoundFile = ""
    mvarResolutionUnit = ""
    mvarShutterSpeed = ""
    mvarSubjectDistance = ""
    mvarWhiteBalance = ""
    mvarXResolution = ""
    mvarYResolution = ""
End Sub
Private Sub Class_Initialize()
    ClearTags
End Sub
Private Function DatesDelimitors(LaDate As String) As String
    If Trim$(LaDate) = "" Then
        DatesDelimitors = ""
    Else
        LaDate = Replace(LaDate, ":", LesOptions.ExifDelimD, 1, 2)
        LaDate = Replace(LaDate, ":", LesOptions.ExifDelimT, 1, 2)
        DatesDelimitors = LaDate
    End If
End Function
