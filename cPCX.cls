VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' PCX image file reader
' Copyright (c) 2000 by Warren Galyen [wgalyen@earthlink.net]
' Currently only supports ver 3.0+ (RLE compressed)

Private HasLoaded As Boolean
Private BmpInfo As BITMAPINFO
Private BmpData() As Byte
Private PCXHdr As PCXHeader

Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    SizeImage As Long
    XPelsPerMeter As Long
    YPelsPerMeter As Long
    ClrUsed As Long
    ClrImportant As Long
End Type
Private Type RGBQUAD
    Blue As Byte
    Green As Byte
    Red As Byte
    Reserved As Byte
End Type
Private Type PALENTRY
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

' PCX Header structure
Private Type PCXHeader
    Signature As Byte    ' (10 = ZSoft)
    Version As Byte      ' (0 = 2.5, 2 = 2.8 w/ palette, 3 = 2.8 w/o palette, 3.0+ w/ 24 bit encoding)
    Compression As Byte  ' (0 = None, 1 = Run Length Encoded)
    BitsPerPixel As Byte ' (1, 2, 4, 8)
    XMIN As Integer
    YMIN As Integer
    XMAX As Integer
    YMAX As Integer
    DpiX As Integer
    DpiY As Integer
    ColorTable(0 To 15) As PALENTRY
    Reserved As Byte
    Planes As Byte
    BytesPerRow As Integer
    PaletteFlags As Integer
    Reserved2(1 To 58) As Byte
End Type

Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Colors() As RGBQUAD
End Type

Private Type BITMAPFILEHEADER
    Signature As Integer
    Size As Long
    Reserved1 As Integer
    Reserved2 As Integer
    OffBits As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long
' Decodes the image data if it uses Run Length encoding
Private Function Decompress(RLEStream() As Byte) As Byte()
    Dim m As Long, n As Long, c As Long, Cnt As Long
    Dim UStream As Long, InitSize As Long, TheByte As Byte
    Dim Temp() As Byte
    
    UStream = UBound(RLEStream)
    InitSize = 3 * UStream
    ReDim Temp(1 To InitSize)
    
    n = LBound(RLEStream) - 1
    Do
        n = n + 1
        If n > UStream Then Exit Do
        If RLEStream(n) >= 192 Then
            Cnt = RLEStream(n) - 192
            m = m + Cnt
            If m > InitSize Then ReDim Preserve Temp(1 To m)
            TheByte = RLEStream(n + 1)
            For c = Cnt - 1 To 0 Step -1
                Temp(m - c) = TheByte
            Next c
            n = n + 1
        Else
            m = m + 1
            If m > InitSize Then ReDim Preserve Temp(1 To m)
            Temp(m) = RLEStream(n)
        End If
    Loop
    
    ReDim Preserve Temp(1 To m)
    Decompress = Temp
End Function
' Takes the loaded function and draws it to a device context
Public Sub DrawToDC(hdc As Long, X As Long, Y As Long)
    If Not HasLoaded Then Exit Sub
    
    Dim BitsInfo() As Byte

    BitsInfo = GetBmpInfo(BmpInfo)
    SetDIBitsToDevice hdc, X, Y, BmpInfo.Header.Width, _
        BmpInfo.Header.Height, 0, 0, 0, BmpInfo.Header.Height, _
        BmpData(0), BitsInfo(0), 0
End Sub
Private Function GetBmpInfo(MyBmpInfo As BITMAPINFO) As Byte()
    Dim Temp() As Byte, ClrTableSize As Long
    
    If MyBmpInfo.Header.BitCount <= 8 Then
        If MyBmpInfo.Header.ClrUsed > 0 Then
            ClrTableSize = MyBmpInfo.Header.ClrUsed
        Else
            ClrTableSize = 2 ^ MyBmpInfo.Header.BitCount
        End If
    End If
    
    ReDim Temp(0 To Len(MyBmpInfo.Header) + ClrTableSize * 4 - 1)
    
    CopyMemory Temp(0), MyBmpInfo.Header, Len(MyBmpInfo.Header)
    
    If ClrTableSize > 0 Then
        CopyMemory Temp(Len(MyBmpInfo.Header)), _
            MyBmpInfo.Colors(0), ClrTableSize * 4
    End If
    
    GetBmpInfo = Temp
End Function
' Check if file has a correct ZSoft header
Private Function IsValidHeader(Header As PCXHeader) As Boolean
    With Header
        If .Signature <> &HA Then Exit Function
        Select Case .Version
            Case 0, 2, 3, 5
            Case Else: Exit Function
        End Select
        Select Case .Compression
            Case 0, 1
            Case Else: Exit Function
        End Select
        If .XMAX - .XMIN + 1 < 1 Then Exit Function
        If .YMAX - .YMIN + 1 < 1 Then Exit Function
        If .Reserved <> 0 Then Exit Function
        If (.Planes < 1) Or (.Planes > 4) Then Exit Function
        Select Case .PaletteFlags
            Case 1, 2
            Case Else: Exit Function
        End Select
    End With
        
    IsValidHeader = True
End Function
Public Function LoadFile(FileName As String) As Boolean
    Dim FileNum As Integer, pHdr As PCXHeader
    Dim DataSize As Long, pData() As Byte
    Dim pBmpHdr As BITMAPINFOHEADER, pClrTbl() As RGBQUAD
    Dim n As Long, pBmpData() As Byte
    
    FileNum = FreeFile
    
    On Error GoTo FileErr
    Open FileName For Binary Access Read As FileNum
    On Error GoTo 0
    
    Get FileNum, , pHdr

    If Not IsValidHeader(pHdr) Then Exit Function
    If pHdr.Version <> 5 Then Exit Function  ' Only supporting 3.0+

    ' Fill in bitmap header
    With pBmpHdr
        .Size = Len(pBmpHdr)
        .Width = pHdr.XMAX - pHdr.XMIN + 1
        .Height = pHdr.YMAX - pHdr.YMIN + 1
        .Planes = 1
        If (pHdr.BitsPerPixel = 8) And (pHdr.Planes = 3) Then
            .BitCount = 24
        ElseIf (pHdr.BitsPerPixel = 8) And (pHdr.Planes = 1) Then
            .BitCount = 8
        ElseIf (pHdr.BitsPerPixel = 1) And (pHdr.Planes = 4) Then
            .BitCount = 4
        ElseIf (pHdr.BitsPerPixel = 1) And (pHdr.Planes = 1) Then
            .BitCount = 1
        End If
        .SizeImage = BmpSize(pBmpHdr)
        .XPelsPerMeter = 2953
        .YPelsPerMeter = 2953
    End With

    ' Get color table, if needed
    Select Case pBmpHdr.BitCount
        Case 1
            ReDim pClrTbl(0 To 1)
            For n = 14 To 15
                With pClrTbl(n - 14)
                    .Red = pHdr.ColorTable(n).Red
                    .Green = pHdr.ColorTable(n).Green
                    .Blue = pHdr.ColorTable(n).Blue
                End With
            Next n
        Case 4
            ReDim pClrTbl(0 To 15)
            For n = 0 To 15
                With pClrTbl(n)
                    .Red = pHdr.ColorTable(n).Red
                    .Green = pHdr.ColorTable(n).Green
                    .Blue = pHdr.ColorTable(n).Blue
                End With
            Next n
        Case 8
            Dim ClrSig As Byte
            Get FileNum, LOF(FileNum) - 768, ClrSig
            If ClrSig = &HC Then
                Dim PCXTbl(0 To 255) As PALENTRY
                Get FileNum, , PCXTbl
                ReDim pClrTbl(0 To 255)
                For n = 0 To 255
                    With pClrTbl(n)
                        .Red = PCXTbl(n).Red
                        .Green = PCXTbl(n).Green
                        .Blue = PCXTbl(n).Blue
                    End With
                Next n
                Seek FileNum, Len(pHdr) + 1
                DataSize = LOF(FileNum) - Len(pHdr) - 769
            Else
                ReDim pClrTbl(0 To 15)
                For n = 0 To 15
                    With pClrTbl(n)
                        .Red = pHdr.ColorTable(n).Red
                        .Green = pHdr.ColorTable(n).Green
                        .Blue = pHdr.ColorTable(n).Blue
                    End With
                Next n
            End If
    End Select

    If DataSize = 0 Then DataSize = LOF(FileNum) - Len(pHdr)
    ReDim pData(0 To DataSize - 1)
    
    Get FileNum, , pData
    Close FileNum
    
    'Decompress
    If pHdr.Compression = 1 Then pData = Decompress(pData)
    
    ' Get bitmap data
    pBmpData = MakeBMPData(pData, pHdr, pBmpHdr)
    
    BmpInfo.Header = pBmpHdr
    BmpInfo.Colors = pClrTbl
    BmpData = pBmpData
    PCXHdr = pHdr

    HasLoaded = True
    LoadFile = True
    Exit Function
FileErr:
    Kill FileName
End Function
Private Function BmpSize(BmpHdr As BITMAPINFOHEADER) As Long
    With BmpHdr
        BmpSize = LineBytes(.Width, .BitCount) * .Height
    End With
End Function

Private Function LineBytes(Width As Long, BitCount As Integer) As Long
    LineBytes = ((Width * BitCount + 31) \ 32) * 4
End Function
Private Function MakeBMPData(Data() As Byte, PCXHdr As PCXHeader, BmpHdr As BITMAPINFOHEADER) As Byte()
    Dim Temp() As Byte, Row As Long, Plane As Long, X As Long
    Dim StartByte As Long, DestByte As Long, BPerRow As Long
    
    ReDim Temp(0 To BmpHdr.SizeImage - 1)
    
    BPerRow = LineBytes(BmpHdr.Width, BmpHdr.BitCount)
    Select Case BmpHdr.BitCount
        Case 8
            For Row = 0 To BmpHdr.Height - 1
                StartByte = Row * PCXHdr.BytesPerRow + 1
                DestByte = BPerRow * (BmpHdr.Height - Row - 1)
                CopyMemory Temp(DestByte), Data(StartByte), PCXHdr.BytesPerRow
            Next Row
        Case 24
            For Row = 0 To BmpHdr.Height - 1
                DestByte = BPerRow * (BmpHdr.Height - Row - 1)
                For Plane = 0 To 2
                    StartByte = (Row * 3 + Plane) * PCXHdr.BytesPerRow + 1
                    For X = 0 To BmpHdr.Width - 1
                        Temp(DestByte + X * 3 + (2 - Plane)) = _
                            Data(StartByte + X)
                    Next X
                Next Plane
            Next Row
    End Select

    MakeBMPData = Temp
End Function
Public Property Get Width() As Long
    Width = BmpInfo.Header.Width
End Property

Public Property Get MinX() As Long
    MinX = PCXHdr.XMIN
End Property

Public Property Get MinY() As Long
    MinY = PCXHdr.YMIN
End Property
Public Property Get MaxX() As Long
    MaxX = PCXHdr.XMAX
End Property
Public Property Get MaxY() As Long
    MaxY = PCXHdr.YMAX
End Property
Public Property Get BitCount() As Long
    BitCount = BmpInfo.Header.BitCount
End Property
Public Property Get Planes() As Long
    Planes = PCXHdr.Planes
End Property
Public Property Get IsGrayscale() As Boolean
    IsGrayscale = (PCXHdr.PaletteFlags = 2)
End Property
Public Property Get DpiX() As Long
    DpiX = PCXHdr.DpiX
End Property
Public Property Get DpiY() As Long
    DpiY = PCXHdr.DpiY
End Property
Public Property Get IsCompressed() As Boolean
    IsCompressed = (PCXHdr.Compression = 1)
End Property
Public Property Get Height() As Long
    Height = BmpInfo.Header.Height
End Property
Private Sub Class_Initialize()
    HasLoaded = False
End Sub


