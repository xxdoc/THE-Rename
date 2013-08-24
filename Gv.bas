Attribute VB_Name = "GV1"
' Besoin d'images au format suivant :
' AVS, extension .x
' ICC from kodak XL7700, extension ICC ?
' LUMENA files from Vista, extensions : PIX or BPX
' MacDraw, extension .mac ?
' MIFF image file, extension ?
' Kodak PCD-format, extension PCD ?
' pgm pictures, extension pgm ?
' PIX from IRIX, extension ????
' Fax images, extension QFX
' ColoRIX pictures from RIX SoftWorks, extensions : SCD,SCE,SCR,SCP,SCG,SCU,SCI,SCK,SCQ,SCL,SCF,SCN,SCO,SCZ
' SGI from Silicon Graphics Computer Systems, extension SGI ????
' VICAR (VAX/VMS), extension ????
' XBM - X BitMap Format, extension ????
' XWD - X Window Dump, extension xwd ?
Option Explicit
DefInt A-Z

Dim Fi%
Dim IntMot%
Dim Tags$(254 To 532)
Dim Typs$(4)
Dim BT As String * 1
Dim Canc%
Dim Found_BMP%
' Graphics Viewer
' Written by: Joe C. Oliphant
' CompuServe 71742, 1451
' E-Mail joe_oliphant@csufresno.edu


Private Type PNGHeader
   png As String * 16
End Type
  
Private Type PSDHeader
  skip As String * 14
End Type
 
Private Type BITMAPINFOHEADER '40 bytes
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type
  
Private Type BITMAPFILEHEADER
    bfType            As Integer
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOffBits         As Long
End Type

' Type defs for CUT Files
Private Type CUTHEAD
    width As Integer
    height As Integer
    Key As Integer
End Type
          
' Type defs for GIF files
Private Type GIFHEADER
    GIF As String * 6
    width As Integer
    height As Integer
    flags As String * 1
    Background As String * 1
    Aspect As String * 1
End Type

Private Type IMAGEBLOCK
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
    flags As String * 1
End Type

Private Type PLAINTEXT
    BlockSize As String * 1
    Left As Integer
    Top As Integer
    GridWidth As Integer
    GridHeight As Integer
    CellWidth As String * 1
    CellHeight As String * 1
    ForeColor As String * 1
    BackColor As String * 1
End Type

Private Type CONTROLBLOCK
    BlockSize As String * 1
    flags As String * 1
    Delay As Integer
    TransParent_Color As String * 1
    Terminator As String * 1
End Type

Private Type APPLICATION
    BlockSize As String * 1
    ApplString As String * 8
    Authentication As String * 3
End Type

' Type def for IFF/LBM files
Private Type IFFHEAD
    Ftype As String * 4
    Size As String * 4
    SubType As String * 4
End Type

Private Type BMHD
    W As String * 2
    H As String * 2
    X As String * 2
    Y As String * 2
    nPlanes As String * 1
    Masking As String * 1
    Compression As String * 1
    Pad1 As String * 1
    TransparentColor As String * 2
    XAspect As String * 1
    YAspect As String * 1
    PageW As String * 2
    PageH As String * 2
End Type

' Type defs for MAC files
'Type MACHEAD
'    ZeroByte As String * 1
'    Name As String * 64
'    Type As String * 4
'    Creator As String * 4
'    Filler As String * 10
'    DataFork_Size As String * 4
'    RsrcFork_Size As String * 4
'    Creation_Date As String * 4
'    Modif_Date As String * 4
'    Filler2 As String * 29
'End Type

' Type def for MSP files
Private Type MSPHEAD
    Key1 As Integer
    Key2 As Integer
    width As Integer
    height As Integer
    ScrAspX As Integer
    ScrAspY As Integer
    PrnAspX As Integer
    PrnAspY As Integer
    PrndX As Integer
    PrndY As Integer
    Wcheck As Integer
    Res1 As Integer
    Res2 As Integer
    Res3 As Integer
End Type

' Type def for PCX files
Private Type PCXHEAD
    Manufacturer As String * 1
    Version As String * 1
    Encoding As String * 1
    Bits_Per_Pixel As String * 1
    XMin As Integer
    YMin As Integer
    XMax As Integer
    YMax As Integer
    HRes As Integer
    VRes As Integer
    Palette As String * 48
    Reserved As String * 1
    Color_Planes As String * 1
    Bytes_Per_Line As Integer
    Palette_Type As Integer
    Filler As String * 58
End Type

' Type def for PIC files
Private Type PICHEAD
    Mark As Integer
    XSize As Integer
    YSize As Integer
    XOff As Integer
    YOff As Integer
    BitsInf As String * 1
    Emark As String * 1
    EVideo As String * 1
    EDesc As Integer
    ESize As Integer
End Type

' Type def for TGA files
Private Type TGAHEAD
    IdentSize As String * 1
    ColorMapType As String * 1
    ImageType As String * 1
    ColorMapStart As Integer
    ColorMapLength As Integer
    ColorMapBits As String * 1
    XStart As Integer
    YStart As Integer
    width As Integer
    height As Integer
    Bits As String * 1
    Descriptor As String * 1
End Type

' Type def for TIF files
Private Type TIFFTAG
    Tag As Long
    Type As Long
    Length As Long
    Offset As Long
End Type

' Type defs for WMF files
Private Type RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Private Type METAFILEHEADER
    Key As Long
    hMF As Integer
    bbox As RECT
    inch As Integer
    Reserved As Long
    checksum As Integer
End Type

Private Type METAHEADER
    mtType As Integer
    mtHeaderSize As Integer
    mtVersion As Integer
    mtSize As Long
    mtNoObjects As Integer
    mtMaxRecord As Long
    mtNoParameters As Integer
End Type

' Type defs for WPG files
Private Type WPGHEAD
    ID As String * 4
    start As Long
    Product As String * 1
    FileType As String * 1
    MajorVersion As String * 1
    Minorversion As String * 1
    Encrypt As Integer
    Reserved As Integer
End Type

'Type STARTRECORD
'    Version As String * 1
'    flags As String * 1
'    ScreenWidth As Integer
'    ScreenHeight As Integer
'End Type

'Type COLORMAP
'    StartIndex As Integer
'    PalleteSize As Integer
'End Type

'Type BITMAP
'    width As Integer
'    height As Integer
'    Bits As Integer
'    XResolution As Integer
'    YResolution As Integer
'End Type
          
' Type defs for displaying logical fonts
'Type LOGFONT
'    lfHeight As Integer
'    lfWidth As Integer
'    lfEscapement As Integer
'    lfOrientation As Integer
'    lfWeight As Integer
'    lfItalic As String * 1
'    lfUnderline As String * 1
'    lfStrikeOut As String * 1
'    lfCharSet As String * 1
'    lfOutPrecision As String * 1
'    lfClipPrecision As String * 1
'    lfQuality As String * 1
'    lfPitchAndFamily As String * 1
'    lfFaceName As String * 32
'End Type

'Type TEXTMETRIC
'    tmHeight As Integer
'    tmAscent As Integer
'    tmDescent As Integer
'    tmInternalLeading As Integer
'    tmExternalLeading As Integer
'    tmAveCharWidth As Integer
'    tmMaxCharWidth As Integer
'    tmWeight As Integer
'    tmItalic As String * 1
'    tmUnderlined As String * 1
'    tmStruckOut As String * 1
'    tmFirstChar As String * 1
'    tmLastChar As String * 1
'    tmDefaultChar As String * 1
'    tmBreakChar As String * 1
'    tmPitchAndFamily As String * 1
'    tmCharSet As String * 1
'    tmOverhang As Integer
'    tmDigitizedAspectX As Integer
'    tmDigitizedAspectY As Integer
'End Type

'Global Const BITSPIXEL = 12
'Global Const GMEM_MOVEABLE = &H2
'Global Const HORZRES = 8
'Global Const LB_SETTABSTOPS = &H413
'Global Const MM_ANISOTROPIC = 8
'Global Const PLANES = 14
'Global Const RASTERCAPS = 38
'Global Const SRCCOPY = &HCC0020
'Global Const VERTRES = 10
Private Const M_SOF0 = &HC0
Private Const M_SOF1 = &HC1
Private Const M_SOF2 = &HC2
Private Const M_SOF3 = &HC3
Private Const M_SOF5 = &HC5
Private Const M_SOF6 = &HC6
Private Const M_SOF7 = &HC7
Private Const M_SOF9 = &HC9
Private Const M_SOF10 = &HCA
Private Const M_SOF11 = &HCB
Private Const M_SOF13 = &HCD
Private Const M_SOF14 = &HCE
Private Const M_SOF15 = &HCF
Private Const M_SOI = &HD8
Private Const M_EOI = &HD9
Private Const M_SOS = &HDA
Private Const M_COM = &HFE
Public Function ImgInfo(Filename As String) As String
 On Error Resume Next
 Dim sonextension As String
 Dim lWidth As Long
 Dim lHeight As Long
 Dim vretour As String
 lWidth = -1
 lHeight = -1
 sonextension = UCase$(Suffixe(Filename))
    Select Case sonextension
    Case "ART"
        Info_ART Filename, lWidth, lHeight
    Case "BMP"
        Info_BMP Filename, lWidth, lHeight
    Case "CUT"
        Info_CUT Filename, lWidth, lHeight
    Case "DIB"
        Info_BMP Filename, lWidth, lHeight
    Case "GEM"
        Info_IMG Filename, lWidth, lHeight
    Case "GIF"
        Info_GIF Filename, lWidth, lHeight
    Case "HRZ"
        Info_HRZ Filename, lWidth, lHeight
    Case "IFF"
        Info_IFF Filename, lWidth, lHeight
    Case "IMG"
        Info_IMG Filename, lWidth, lHeight
    Case "JPG"
        Info_JPG Filename, lWidth, lHeight
    Case "JPE"
        Info_JPG Filename, lWidth, lHeight
    Case "JPEG"
        Info_JPG Filename, lWidth, lHeight
    Case "LBM"
        Info_IFF Filename, lWidth, lHeight
    Case "MAC"
        Info_MAC lWidth, lHeight
    Case "MSP"
        Info_MSP Filename, lWidth, lHeight
    Case "PCX"
        Info_PCX Filename, lWidth, lHeight
    Case "PIC"
        Info_PIC Filename, lWidth, lHeight
    Case "PNG"
        Info_PNG Filename, lWidth, lHeight
    Case "PSD"
        Info_PSD Filename, lWidth, lHeight
    Case "PSP"
        Info_PSP Filename, lWidth, lHeight
    Case "RAS"
        Info_RAS Filename, lWidth, lHeight
    Case "RLE"
        Info_BMP Filename, lWidth, lHeight
    Case "TGA"
        Info_TGA Filename, lWidth, lHeight
    Case "TIF"
        Info_TIF Filename, lWidth, lHeight
    Case "WMF"
        Info_WMF Filename, lWidth, lHeight
    Case "WPG"
        Info_WPG Filename, lWidth, lHeight
    Case "ICO"
        lWidth = 32
        lHeight = 32
    Case Else
        lWidth = -1
        lHeight = -1
    End Select
    If lWidth = -1 Or lHeight = -1 Then
     ImgInfo = ""
     Exit Function
    End If
    vretour = LesOptions.PicturesFormat
    vretour = Replace(vretour, "%w%", Trim$(Str$(lWidth)), , , vbTextCompare)
    vretour = Replace(vretour, "%h%", Trim$(Str$(lHeight)), , , vbTextCompare)
    ImgInfo = vretour
End Function
Private Sub Info_ART(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim c&
    IntMot = True
    Fi = FreeFile
    Open File$ For Binary As Fi
    c = GetInt()
    c = GetInt()
    lWidth = c
    c = GetInt()
    c = GetInt()
    lHeight = c
    Close Fi
End Sub
' ******************************************************************************
' Retourne des infos sur les BMP
' ******************************************************************************
Private Sub Info_BMP(fichier$, ByRef lWidth As Long, ByRef lHeight As Long)
  Dim ff As Integer
  Dim FileHeader As BITMAPFILEHEADER
  Dim InfoHeader As BITMAPINFOHEADER

  On Error GoTo cmdSelect_FileErrorHandler
  
 'read the file header info
  ff = FreeFile
  Open fichier$ For Binary Access Read As #ff
  Get #ff, , FileHeader
  Get #ff, , InfoHeader
  Close #ff


  lWidth = InfoHeader.biWidth
  lHeight = InfoHeader.biHeight
Exit Sub

'handle file errors or the user choosing cancel
cmdSelect_FileErrorHandler:
  lWidth = -1
  lHeight = -1

End Sub
Private Sub Info_PSD(File$, ByRef lWidth As Long, ByRef lHeight As Long)
 Dim psd As PSDHeader
 Fi = FreeFile
 Dim i&
 Open File$ For Binary As Fi
 Get #Fi, , psd
 IntMot = False
 i = GetLng()
 lHeight = i
 i = GetLng()
 lWidth = i
 Close Fi
End Sub
Private Sub Info_PNG(File$, ByRef lWidth As Long, ByRef lHeight As Long)
 Dim png As PNGHeader
 Dim i&
 Fi = FreeFile
 Open File$ For Binary As Fi
 Get #Fi, , png
 IntMot = False
 i = GetLng()
 lWidth = i
 i = GetLng()
 lHeight = i
 Close Fi
End Sub
Private Sub Info_CUT(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim CUT As CUTHEAD
    Fi = FreeFile
    Open File$ For Binary As Fi
    Get #Fi, , CUT
    Close Fi
    lWidth = CUT.width
    lHeight = CUT.height
End Sub
Private Sub Info_GIF(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim GIF As GIFHEADER
    Dim Image As IMAGEBLOCK
    Dim a$, B$, i%, clr%
    Dim Flag%, NumClrs%, NumClrBits%
    lWidth = -1
    lHeight = -1

    Fi = FreeFile
    Open File$ For Binary As Fi
    Get #Fi, , GIF
    Flag = Asc(GIF.flags)
    B$ = "No"
    If (Flag And &H80) Then
        NumClrBits = (Flag And &H7) + 1
        NumClrs = 2 ^ NumClrBits
        a$ = String$(NumClrs * 3, 0)
        Get #Fi, , a$
        B$ = "Yes"
    End If
    Do
        Get #Fi, , BT$
        Select Case BT$
        Case ","
            Get #Fi, , Image
            lWidth = Image.width
            lHeight = Image.height
            Flag = Asc(Image.flags)
            B$ = "No"
            If (Flag And &H80) Then
                NumClrBits = (Flag And &H7) + 1
                NumClrs = 2 ^ NumClrBits
                a$ = String$(NumClrs * 3, 0)
                Get #Fi, , a$
                B$ = "Yes"
            End If
            B$ = "No"
            If (Flag And &H40) Then B$ = "Yes"
            i = GetC()
            i = 1
            Do Until i = 0
            i = GetC()
            Seek #Fi, Seek(Fi) + i
            Loop

        Case "!"
            Get #Fi, , BT$

            Select Case Asc(BT$)          ' Plain Text Extension
            Case 1
                Dim PlnTxt As PLAINTEXT
                Get #Fi, , PlnTxt
                clr = Asc(PlnTxt.ForeColor)
                clr = Asc(PlnTxt.BackColor)
                a$ = ""
                Do
                For i = 1 To GetC()
                Get #Fi, , BT$
                If BT$ = Chr$(0) Or EOF(Fi) Then Exit Do
                a$ = a$ + B$
                Next
                Loop

            Case 249                      'Control Block Extension
                Dim Cntrlblk As CONTROLBLOCK
                Get #Fi, , Cntrlblk
                Flag = Asc(Cntrlblk.flags)
            
            Case 254                      'Comment Extension
                a$ = ""
                Do
                For i = 1 To GetC()
                Get #Fi, , BT$
                If BT$ = Chr$(0) Or EOF(Fi) Then Exit Do
                a$ = a$ & BT$
                Next
                Loop
             
            Case 255                      'Application Extension
                Dim Appl As APPLICATION
                Get #Fi, , Appl
                Do
                For i = 1 To GetC()
                Get #Fi, , BT$
                If BT$ = Chr$(0) Or EOF(Fi) Then Exit Do
                Next
                Loop

            Case Else
                Do
                For i = 1 To GetC()
                Get #Fi, , BT$
                If BT$ = Chr$(0) Or EOF(Fi) Then Exit Do
                Next
                Loop
            
            End Select

        Case Chr$(0)
            If EOF(Fi) Then Exit Do
        
        Case Else
            Exit Do

        End Select
    Loop
    Close Fi

If lWidth = -1 And lHeight = -1 Then
    lWidth = GIF.width
    lHeight = GIF.height
End If
    
End Sub

Private Sub Info_HRZ(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    If FileLen(File$) <> 184320 Then
        lWidth = -1
        lHeight = -1
    End If
 lWidth = 256
 lHeight = 240

End Sub

Private Sub Info_IFF(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim IFF As IFFHEAD, BMHEAD As BMHD
    Dim B$, Lng As String * 4
    Dim Chnk As String * 4, Pos&, Size&
    
    Fi = FreeFile
    Open File$ For Binary As Fi
    Get #Fi, , IFF
    Do
    Get #Fi, , Chnk$
    Get #Fi, , Lng$
    Pos = Seek(Fi)
    Size = CnvtLng(Lng$)
    If Size And 1 Then Size = Size + 1
    Select Case Chnk$
    Case "BMHD"
        Get #Fi, , BMHEAD
        lWidth = CnvtInt(BMHEAD.W)
        lHeight = CnvtInt(BMHEAD.H)

    Case "TEXT"
        If Size <= 40 Then
            B$ = Space$(Size)
            Get #Fi, , B$
        Else
        End If
        
    End Select
    Seek #Fi, Pos + Size
    Loop Until Chnk$ = "BODY" Or EOF(Fi)
    Close Fi
End Sub

Private Sub Info_IMG(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim i&, H&, n&

    IntMot = False
    Fi = FreeFile
    Open File$ For Binary As Fi
    i = GetInt()
    H = GetInt()
    n = GetInt()
    i = GetInt()
    i = GetInt()
    i = GetInt()
    i = GetInt()
    lWidth = i
    i = GetInt()
    lHeight = i
    Close Fi
End Sub

Public Sub Info_JPG(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim Marker
    
    IntMot = False
    Fi = FreeFile
    Open File$ For Binary As Fi
    
    If First_Marker() <> M_SOI Then
        lWidth = -1
        lHeight = -1
        Close Fi
        Exit Sub
    End If
    Do
        Marker = Next_Marker()
        Select Case Marker
        Case -1
            Close Fi
            Exit Sub
            
        Case M_SOF0, M_SOF1, M_SOF2, M_SOF3, M_SOF5, M_SOF6, M_SOF7, M_SOF9, M_SOF10, M_SOF11, M_SOF13, M_SOF14, M_SOF15
            Process_SOFn Marker, lWidth, lHeight
           
        Case M_SOS
            Close Fi
            Exit Sub
        
        Case M_SOI
            Close Fi
            Exit Sub
        
        Case M_EOI
            Close Fi
            Exit Sub
        
        Case M_COM
            Process_COM
            If Canc Then
                Close Fi
                Exit Sub
            End If
        
        Case Else
            Skip_Variable
        
        End Select
        
    Loop
    Close Fi
End Sub

Private Sub Info_MAC(ByRef lWidth As Long, ByRef lHeight As Long)
    lWidth = 576
    lHeight = 720
End Sub

Private Sub Info_MSP(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim MSP As MSPHEAD
    Fi = FreeFile
    Open File$ For Binary As Fi
    Get #Fi, , MSP
    Close Fi
    lWidth = MSP.width
    lHeight = MSP.height
End Sub

Private Sub Info_PCX(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim PCX As PCXHEAD
    Fi = FreeFile
    Open File$ For Binary As Fi
    Get #Fi, , PCX
    Close Fi
    lWidth = PCX.XMax - PCX.XMin + 1
    lHeight = PCX.YMax - PCX.YMin + 1
End Sub
Private Sub Info_PIC(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim PIC As PICHEAD
    Fi = FreeFile
    Open File$ For Binary As Fi
    Get #Fi, , PIC
    Close Fi
    lWidth = PIC.XSize
    lHeight = PIC.YSize
End Sub
Private Sub Info_RAS(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim i&
    IntMot = False
    Fi = FreeFile
    Open File$ For Binary As Fi
    i = GetLng()
    i = GetLng()
    lWidth = i
    i = GetLng()
    lHeight = i
    Close Fi
End Sub
Private Sub Info_TGA(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim TGA As TGAHEAD
    Fi = FreeFile
    Open File$ For Binary As Fi
    Get #Fi, , TGA
    Close Fi
    lWidth = TGA.width
    lHeight = TGA.height
End Sub
Private Sub Info_TIF(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim a$, i%
    Dim Offset&, Tag&, Typ&, Length&, NumTags%

    Fi = FreeFile
    Open File$ For Binary As Fi
    a$ = Space$(2)
    Get #Fi, , a$
    If a$ = "II" Then IntMot = True Else IntMot = False
    If IntMot Then a$ = "Intel" Else a$ = "Motorola"
    a$ = Space$(2)
    Get #Fi, , a$
    If IntMot Then a$ = Left$(a$, 1) Else a$ = Right$(a$, 1)
    Offset = GetLng()
    Seek #Fi, Offset + 1
    NumTags = GetInt()
    ReDim TagsInfo(NumTags) As TIFFTAG
    For i = 1 To NumTags
        Tag = GetInt()
        Typ = GetInt()
        Length = GetLng()
        Offset = GetLng()
        TagsInfo(i).Tag = Tag
        TagsInfo(i).Type = Typ
        TagsInfo(i).Length = Length
        TagsInfo(i).Offset = Offset
    Next
    For i = 1 To NumTags
    If TagsInfo(i).Tag <= 532 Then
     
    End If
    a$ = ""
    Select Case TagsInfo(i).Type
    Case 1
        If TagsInfo(i).Length <= 1 Then
            a$ = Format$(TagsInfo(i).Offset And &HF)
        Else
            a$ = "Offset = " & Format$(TagsInfo(i).Offset) & "  Length = " & Format$(TagsInfo(i).Length)
        End If
    Case 2
        Seek #Fi, TagsInfo(i).Offset + 1
        Do
        Get #Fi, , BT$
        If BT$ <> "" Then a$ = a$ & BT$
        Loop Until Asc(BT$) = 0
    Case 3
        If TagsInfo(i).Length <= 1 Then
            a$ = Format$(TagsInfo(i).Offset And &HFFF)
        Else
            a$ = "Offset = " & Format$(TagsInfo(i).Offset) & "  Length = " & Format$(TagsInfo(i).Length)
        End If
    Case 4
        If TagsInfo(i).Length <= 1 Then
            a$ = Format$(TagsInfo(i).Offset)
        Else
            a$ = "Offset = " & Format$(TagsInfo(i).Offset) & "  Length = " & Format$(TagsInfo(i).Length)
        End If
    Case 5
        Seek #Fi, TagsInfo(i).Offset + 1
        a$ = Str$(GetLng() / GetLng())
    End Select
    If i = 256 Then
     lWidth = Val(a$)
    End If
    If i = 257 Then
     lHeight = Val(a$)
    End If
    Next
    Close Fi

End Sub

Private Sub Info_WMF(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim WMFH As METAFILEHEADER
    Dim WMF As METAHEADER
    Fi = FreeFile
    Open File$ For Binary As Fi
    Get #Fi, , WMFH
    If WMFH.Key <> &H9AC6CDD7 Then Seek #Fi, 1
    Get #Fi, , WMF
    Close Fi
    If WMFH.Key = &H9AC6CDD7 Then
        lWidth = WMFH.bbox.Right
        lHeight = WMFH.bbox.Bottom
    End If
End Sub

Private Sub Info_PSP(File$, ByRef lWidth As Long, ByRef lHeight As Long)
 Dim header As String * 50
 Dim i&
 IntMot = True
 Fi = FreeFile
 Open File$ For Binary As Fi
 Get #Fi, , header
 i = GetLng()
 lWidth = i
 i = GetInt()
 lHeight = i
 Close #Fi
End Sub

Private Sub Info_WPG(File$, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim WPG As WPGHEAD
    Dim Typ, t&, i&, l&

    Fi = FreeFile
    IntMot = True
    Open File$ For Binary As Fi
    Get #Fi, , WPG
    Seek #Fi, WPG.start + 1
    Do
        Typ = GetC()
        t = Seek(Fi)
        i = GetC()
        If i = 255 Then
            i = GetInt()
            If i And &H8000 Then
                l = (i And &H7FFF) * 2 ^ 16
                i = GetInt()
                l = l + i + 4
            Else
                l = i + 2
            End If
        Else
            l = i
        End If
        
        Select Case Typ
        Case 11
            lWidth = GetInt()
            lHeight = GetInt()
            Found_BMP = True
        
        
        End Select
        
        Seek #Fi, t + l + 1
    Loop While Seek(Fi) < LOF(Fi)
    Close Fi
End Sub

Private Function GetC%()
    Get #Fi, , BT$
    GetC = Asc(BT$)
End Function

Private Function GetInt&()
    Dim c&, n&
    c = GetC()
    If IntMot Then n = c Else n = c * 256
    c = GetC()
    If IntMot Then n = n + c * 256 Else n = n + c
    GetInt = n
End Function

Private Function GetLng&()
    Dim c&, n&
    c = GetC()
    If IntMot Then n = c Else n = c * 16777216
    c = GetC()
    If IntMot Then n = n + c * 256 Else n = n + c * 65536
    c = GetC()
    If IntMot Then n = n + c * 65536 Else n = n + c * 256
    c = GetC()
    If IntMot Then n = n + c * 16777216 Else n = n + c
    GetLng = n
End Function
' Procédure gérant les données EXIF d'une image JPEG
'Private Sub Process_EXIF()
'    Dim Length
'    Dim PosCour As Long
'    Dim ExifBuf() As Byte
'    Dim i As Integer
'    Dim vtmp As String
'    Dim Intel As Boolean
'
'    PosCour = Seek(Fi)
'    Length = GetInt()
'    If Length < 2 Then
'        Close Fi
'        Exit Sub
'    End If
'    Length = Length - 2
'    ReDim ExifBuf(Length)
'    ' on redimensionne le buffer à la taille des infos EXIF
'    ' Puis on lit ces informations
'    Get #Fi, , ExifBuf
'    Clipboard.SetText ExifBuf
'    vtmp = Chr$(ExifBuf(0)) & Chr$(ExifBuf(1)) & Chr$(ExifBuf(2)) & Chr$(ExifBuf(3)) & Chr$(ExifBuf(4)) & Chr$(ExifBuf(5))
'    ' Si le buffer ne commence pas par Exif00 on se barre
'    If left(vtmp, 4) <> "Exif" And Asc(Mid(vtmp, 5, 1)) <> 0 And Asc(Mid(vtmp, 6, 1)) <> 0 Then
'        GoTo fin
'    End If
'    ' Si on est encore là c'est que le buffer est reconnu comme des données EXIF
'    vtmp = Chr$(ExifBuf(6)) & Chr$(ExifBuf(7))
'    If vtmp = "II" Then ' On est au format Intel
'        Intel = True
'    Else                ' on est au format Motorola
'        Intel = False
'    End If
'    vtmp = Chr$(ExifBuf(8)) & Chr$(ExifBuf(9))
'fin:
'    ' retour à la position initiale
'    Seek #Fi, PosCour
'    ' et on dégage
'    Skip_Variable
'End Sub
Private Sub Process_SOFn(Marker, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim Length, Image_Height, Image_Width, Data_Precision, Num_Components
    Dim Ci, C1, C2, C3
    Dim Process$

    Length = GetInt()
    Data_Precision = GetC()
    Image_Height = GetInt()
    Image_Width = GetInt()
    Num_Components = GetC()
    
    Select Case Marker
    Case M_SOF0
        Process = "Baseline"
    
    Case M_SOF1
        Process = "Extended sequential"
    
    Case M_SOF2
        Process = "Progressive"
    
    Case M_SOF3
        Process = "Lossless"
    
    Case M_SOF5
        Process = "Differential sequential"
    
    Case M_SOF6
        Process = "Differential progressive"
    
    Case M_SOF7
        Process = "Differential lossless"
    
    Case M_SOF9
        Process = "Extended sequential, arithmetic coding"
    
    Case M_SOF10
        Process = "Progressive, arithmetic coding"
    
    Case M_SOF11
        Process = "Lossless, arithmetic coding"
    
    Case M_SOF13
        Process = "Differential sequential, arithmetic coding"
    
    Case M_SOF14
        Process = "Differential progressive, arithmetic coding"
    
    Case M_SOF15
        Process = "Differential lossless, arithmetic coding"
    
    Case Else
        Process = "Unknown"
    
    End Select
    lWidth = Image_Width
    lHeight = Image_Height
    If Length <> 8 + Num_Components * 3 Then
        Close Fi
        Canc = True
        Exit Sub
    End If
    For Ci = 0 To Num_Components - 1
    C1 = GetC()
    C2 = GetC()
    C3 = GetC()
    Next
End Sub

Private Sub Skip_Variable()
    Dim Length
    
    Length = GetInt()
    If Length < 2 Then
        Close Fi
        Exit Sub
    End If
    Length = Length - 2
    Seek #Fi, Seek(Fi) + Length
End Sub

Private Sub Process_COM()
    Dim cH, Length, a$
    
    Length = GetInt()
    If Length < 2 Then
        Close Fi
        Exit Sub
    End If
    Length = Length - 2
    While Length > 0
        cH = GetC()
        a$ = a$ & Chr$(cH)
        Length = Length - 1
    Wend
End Sub

Public Sub LoadTags()
Dim i As Integer
On Error GoTo ErrGen
    For i = 254 To 532
        Tags(i) = "Unknown"
    Next
    Typs(0) = "Byte"
    Typs(1) = "ASCII"
    Typs(2) = "Unsigned Int"
    Typs(3) = "Unsigned Long"
    Typs(4) = "Rational"
    Tags(254) = "NewSubFileType"
    Tags(255) = "SubFileType"
    Tags(256) = "ImageWidth"
    Tags(257) = "ImageHeight"
    Tags(258) = "BitsPerSample"
    Tags(259) = "Compression"
    Tags(262) = "PhotometricInterpretation"
    Tags(263) = "Threshholding"
    Tags(264) = "CellWidth"
    Tags(265) = "CellLength"
    Tags(266) = "FillOrder"
    Tags(269) = "DocumentName"
    Tags(270) = "ImageDescription"
    Tags(271) = "Make"
    Tags(272) = "Model"
    Tags(273) = "StripOffsets"
    Tags(274) = "Orientation"
    Tags(277) = "SamplesPerPixel"
    Tags(278) = "RowsPerStrip"
    Tags(279) = "StripByteCounts"
    Tags(280) = "MinSampleValue"
    Tags(281) = "MaxSampleValue"
    Tags(282) = "XResolution"
    Tags(283) = "YResolution"
    Tags(284) = "PlanarConfiguration"
    Tags(285) = "PageName"
    Tags(286) = "XPosition"
    Tags(287) = "YPosition"
    Tags(288) = "FreeOffsets"
    Tags(289) = "FreeByteCounts"
    Tags(290) = "GrayResponseUnit"
    Tags(291) = "GrayResponseCurve"
    Tags(292) = "Group3Options"
    Tags(293) = "Group4Options"
    Tags(296) = "ResolutionUnit"
    Tags(297) = "PageNumber"
    Tags(300) = "ColorResponseUnit"
    Tags(301) = "ColorResponseCurves"
    Tags(305) = "Software"
    Tags(306) = "DateTime"
    Tags(315) = "Artist"
    Tags(316) = "HostComputer"
    Tags(317) = "Predictor"
    Tags(318) = "WhitePoint"
    Tags(319) = "PrimaryChromaticities"
    Tags(320) = "ColorMap"
    Tags(321) = "HalfToneHints"
    Tags(322) = "TileWidth"
    Tags(323) = "TileLength"
    Tags(324) = "TileOffsets"
    Tags(325) = "TileByteCounts"
    Tags(326) = "BadFaxLines"
    Tags(327) = "CleanFaxData"
    Tags(328) = "ConsecutiveBadFaxLines"
    Tags(332) = "InkSet"
    Tags(333) = "InkNames"
    Tags(334) = "NumberofInks"
    Tags(336) = "DotRange"
    Tags(337) = "TargetPrinter"
    Tags(338) = "ExtraSamples"
    Tags(339) = "SampleFormat"
    Tags(340) = "SMinSampleValue"
    Tags(341) = "SMaxSampleValue"
    Tags(342) = "TransferRange"
    Tags(512) = "JPEGProc"
    Tags(513) = "JPEGInterchangeFormat"
    Tags(514) = "JPEGInterchangeFormatLength"
    Tags(515) = "JPEGRestartInterval"
    Tags(517) = "JPEGLosslessPredictors"
    Tags(518) = "JPEGPointTransforms"
    Tags(519) = "JPEGQTables"
    Tags(520) = "JPEGDCTTables"
    Tags(521) = "JPEGACCTTables"
    Tags(529) = "YCbCrCoefficients"
    Tags(530) = "YCbCrSubSampling"
    Tags(531) = "YCbCrPositioning"
    Tags(532) = "ReferenceBlackWhite"
    Exit Sub
ErrGen:
ErreurGrave "LoadTags"
End Sub

Private Function CnvtLng#(Lng$)
    Dim c#, i#
    For i = 3 To 0 Step -1
    c = c + Asc(Mid$(Lng$, 4 - i, 1)) * 256 ^ i
    Next
    CnvtLng = c
End Function
Private Function CnvtInt(zin As String) As Long
    Dim c&
    c = Asc(Left$(zin, 1))
    CnvtInt = c * 256 + Asc(Right$(zin, 1))
End Function
Private Function First_Marker()
    Dim C1, C2
    C1 = GetC()
    C2 = GetC()
    If C1 <> &HFF Or C2 <> M_SOI Then
        Close Fi
        First_Marker = -1
        Exit Function
    End If
    First_Marker = C2
End Function
Private Function Next_Marker()
    Dim c, Discarded_Bytes
    
    c = GetC()
    While c <> &HFF
    Discarded_Bytes = Discarded_Bytes + 1
    c = GetC()
    Wend
    Do
    c = GetC()
    Loop While c = &HFF
    If Discarded_Bytes <> 0 Then
        MsgBox "Garbage found in JPEG file", vbOKOnly
        Close Fi
        Next_Marker = -1
        Exit Function
    End If
    Next_Marker = c
End Function
