Attribute VB_Name = "mod_Fade"
Public Const FADE_T_TO_B = 0
Public Const FADE_B_TO_T = 1
Public Const FADE_L_TO_R = 2
Public Const FADE_R_TO_L = 3
Public Const FADE_RANDOM = 4
Public Const FADE_OUTWARD = 5
Public Type RECT
        Left As Long
    Top As Long
    Right As Long
        Bottom As Long
End Type

Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub Fade(Pic As PictureBox, Style As Integer, Blocks As Integer)
    
    Dim width_section_size As Integer
    Dim height_section_size As Integer
    Dim i As Integer, j As Integer
    Dim save_color As Long
    
    'Saves the picbox's current forecolor
    save_color = Pic.ForeColor

    'Set Pics forecolor to its backcolor
    Pic.ForeColor = Pic.BackColor

    'Corrects the Blocks if needed
    If Blocks < 5 Then Blocks = 5
    If Blocks > 100 Then Blocks = 100

    'Sets the size of each width section
    width_section_size = Pic.ScaleWidth / Blocks

    'Sets the size of each height section
    height_section_size = Pic.ScaleHeight / Blocks


    Select Case Style
       '-------------------------------------------------------------------------------------
       Case 0  'Fading top to bottom
          
          For i = 0 To Blocks
             For j = 0 To Blocks
                Pic.Line ((j * width_section_size), (i * height_section_size))-((j + 1) * width_section_size, (i + 1) * height_section_size), , BF
                DoEvents
             Next
             DoEvents
          Next
       '-------------------------------------------------------------------------------------
       Case 1  'Fading bottom to top
          
          For i = Blocks To 0 Step -1
             For j = 0 To Blocks
                Pic.Line (((j - 1) * width_section_size), ((i - 1) * height_section_size))-(j * width_section_size, i * height_section_size), , BF
                DoEvents
             Next
             DoEvents
          Next
       '-------------------------------------------------------------------------------------
       Case 2  'Fading left to right
          
          For i = 0 To Blocks
             For j = 0 To Blocks
                Pic.Line ((i * width_section_size), (j * height_section_size))-((i + 1) * width_section_size, (j + 1) * height_section_size), , BF
                DoEvents
             Next
             DoEvents
          Next
       '-------------------------------------------------------------------------------------
       Case 3  'Fading right to left
          
          For i = Blocks To 0 Step -1
             For j = 0 To Blocks
                Pic.Line (((i - 1) * width_section_size), (j * height_section_size))-(i * width_section_size, (j + 1) * height_section_size), , BF
                DoEvents
             Next
             DoEvents
          Next
       '-------------------------------------------------------------------------------------
       Case 4  'Fading Random
       
          Dim bit_array() As Byte
          ReDim bit_array(Blocks, Blocks)
              
          Dim counter As Integer
       
          Do
             Do
                width_next_block = Int(Blocks * Rnd) 'Generate the random numbers
                height_next_block = Int(Blocks * Rnd) 'Generate the random numbers
                'MsgBox bit_array(width_next_block, height_next_block)
                If bit_array(width_next_block, height_next_block) = 0 Then
                  Exit Do
                End If
                counter = counter + 1
                If counter = Blocks * 10 Then Exit Do
             Loop
             
             If counter = Blocks * 10 Then Exit Do
             counter = 0
          
             'Update the bit_array
             bit_array(width_next_block, height_next_block) = 1
          
    
              
             Pic.Line ((width_next_block * width_section_size), (height_next_block * height_section_size))-((width_next_block + 1) * width_section_size, (height_next_block + 1) * height_section_size), , BF
          
             DoEvents
          Loop
          
          Pic.Line (0, 0)-(Pic.ScaleWidth, Pic.ScaleHeight), , BF
  
       '-------------------------------------------------------------------------------------
       Case 5 'Fading Outward
       
          For i = (Blocks / 2) To 0 Step -1
             Sleep (20)
             Pic.Line (i * width_section_size, i * height_section_size)-(((Blocks - i) + 1) * width_section_size, ((Blocks - i) + 1) * height_section_size), , BF
          Next
          
       '-------------------------------------------------------------------------------------
    End Select

    'Restores the picbox's original forecolor
    Pic.ForeColor = save_color
        
End Sub

Public Sub TransparentBlt(OutDstDC As Long, DstDC As Long, SrcDC As Long, SrcRect As RECT, DstX As Integer, DstY As Integer, TransColor As Long)
'DstDC- Device context into which image must be drawn transparently
'OutDstDC- Device context into image is actually drawn, even though
'it is made transparent in terms of DstDC
'Src- Device context of source to be made transparent in color TransColor
'SrcRect- Rectangular region within SrcDC to be made transparent in terms of
'DstDC, and drawn to OutDstDC
'DstX, DstY - Coordinates in OutDstDC (and DstDC) where the transparent bitmap must go
'In most cases, OutDstDC and DstDC will be the same
Dim nRet As Long, W As Integer, H As Integer
Dim MonoMaskDC As Long, hMonoMask As Long
Dim MonoInvDC As Long, hMonoInv As Long
Dim ResultDstDC As Long, hResultDst As Long
Dim ResultSrcDC As Long, hResultSrc As Long
Dim hPrevMask As Long, hPrevInv As Long
Dim hPrevSrc As Long, hPrevDst As Long
W = SrcRect.Right - SrcRect.Left + 1
H = SrcRect.Bottom - SrcRect.Top + 1
'create monochrome mask and inverse masks
MonoMaskDC = CreateCompatibleDC(DstDC)
MonoInvDC = CreateCompatibleDC(DstDC)
hMonoMask = CreateBitmap(W, H, 1, 1, ByVal 0&)
hMonoInv = CreateBitmap(W, H, 1, 1, ByVal 0&)
hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
hPrevInv = SelectObject(MonoInvDC, hMonoInv)
'create keeper DCs and bitmaps
ResultDstDC = CreateCompatibleDC(DstDC)
ResultSrcDC = CreateCompatibleDC(DstDC)
hResultDst = CreateCompatibleBitmap(DstDC, W, H)
hResultSrc = CreateCompatibleBitmap(DstDC, W, H)
hPrevDst = SelectObject(ResultDstDC, hResultDst)
hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
'copy src to monochrome mask
Dim OldBC As Long
OldBC = SetBkColor(SrcDC, TransColor)
nRet = BitBlt(MonoMaskDC, 0, 0, W, H, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
TransColor = SetBkColor(SrcDC, OldBC)
'create inverse of mask
nRet = BitBlt(MonoInvDC, 0, 0, W, H, MonoMaskDC, 0, 0, vbNotSrcCopy)
'get background
nRet = BitBlt(ResultDstDC, 0, 0, W, H, DstDC, DstX, DstY, vbSrcCopy)
'AND with Monochrome mask
nRet = BitBlt(ResultDstDC, 0, 0, W, H, MonoMaskDC, 0, 0, vbSrcAnd)
'get overlapper
nRet = BitBlt(ResultSrcDC, 0, 0, W, H, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
'AND with inverse monochrome mask
nRet = BitBlt(ResultSrcDC, 0, 0, W, H, MonoInvDC, 0, 0, vbSrcAnd)
'XOR these two
nRet = BitBlt(ResultDstDC, 0, 0, W, H, ResultSrcDC, 0, 0, vbSrcInvert)
'output results
nRet = BitBlt(OutDstDC, DstX, DstY, W, H, ResultDstDC, 0, 0, vbSrcCopy)
'clean up
hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
DeleteObject hMonoMask
hMonoInv = SelectObject(MonoInvDC, hPrevInv)
DeleteObject hMonoInv
hResultDst = SelectObject(ResultDstDC, hPrevDst)
DeleteObject hResultDst
hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
DeleteObject hResultSrc
DeleteDC MonoMaskDC
DeleteDC MonoInvDC
DeleteDC ResultDstDC
DeleteDC ResultSrcDC
End Sub

'*********************************************************************
' Paints a bitmap on a given surface using the surface backcolor
' everywhere lngMaskColor appears on the picSource bitmap
'*********************************************************************
Sub TransparentPaint(objDest As Object, picSource As StdPicture, _
    lngX As Long, lngY As Long, ByVal lngMaskColor As Long)
    '*****************************************************************
    ' This sub uses a bunch of variables, so let's declare and explain
    ' them in advance...
    '*****************************************************************
    Dim lngSrcDC As Long     'ource bitmap
    Dim lngSaveDC As Long    'Copy of Source bitmap
    Dim lngMaskDC As Long    'Monochrome Mask bitmap
    Dim lngInvDC As Long     'Monochrome Inverse of Mask bitmap
    Dim lngNewPicDC As Long  'Combination of Source & Background bmps
    
    Dim bmpSource As BITMAP  'Description of the Source bitmap
    
    Dim hResultBmp As Long   'Combination of source & background
    Dim hSaveBmp As Long     'Copy of Source bitmap
    Dim hMaskBmp As Long     'Monochrome Mask bitmap
    Dim hInvBmp As Long      'Monochrome Inverse of Mask bitmap
    
    Dim hSrcPrevBmp As Long  'Holds prev bitmap in source DC
    Dim hSavePrevBmp As Long 'Holds prev bitmap in saved DC
    Dim hDestPrevBmp As Long 'Holds prev bitmap in destination DC
    Dim hMaskPrevBmp As Long 'Holds prev bitmap in the mask DC
    Dim hInvPrevBmp As Long  'Holds prev bitmap in inverted mask DC
    
    Dim lngOrigScaleMode&    'Holds the original ScaleMode
    Dim lngOrigColor&        'Holds original backcolor from source DC
    '*****************************************************************
    ' Set ScaleMode to pixels for Windows GDI
    '*****************************************************************
    lngOrigScaleMode = objDest.ScaleMode
    objDest.ScaleMode = vbPixels
    '*****************************************************************
    ' Load the source bitmap to get its width (bmpSource.bmWidth)
    ' and height (bmpSource.bmHeight)
    '*****************************************************************
    GetObject picSource, Len(bmpSource), bmpSource
    '*****************************************************************
    ' Create compatible device contexts (DC's) to hold the temporary
    ' bitmaps used by this sub
    '*****************************************************************
    lngSrcDC = CreateCompatibleDC(objDest.hdc)
    lngSaveDC = CreateCompatibleDC(objDest.hdc)
    lngMaskDC = CreateCompatibleDC(objDest.hdc)
    lngInvDC = CreateCompatibleDC(objDest.hdc)
    lngNewPicDC = CreateCompatibleDC(objDest.hdc)
    '*****************************************************************
    ' Create monochrome bitmaps for the mask-related bitmaps
    '*****************************************************************
    hMaskBmp = CreateBitmap(bmpSource.bmWidth, bmpSource.bmHeight, _
        1, 1, ByVal 0&)
    hInvBmp = CreateBitmap(bmpSource.bmWidth, bmpSource.bmHeight, _
        1, 1, ByVal 0&)
    '*****************************************************************
    ' Create color bitmaps for the final result and the backup copy
    ' of the source bitmap
    '*****************************************************************
    hResultBmp = CreateCompatibleBitmap(objDest.hdc, _
        bmpSource.bmWidth, bmpSource.bmHeight)
    hSaveBmp = CreateCompatibleBitmap(objDest.hdc, _
        bmpSource.bmWidth, bmpSource.bmHeight)
    '*****************************************************************
    ' Select bitmap into the device context (DC)
    '*****************************************************************
    hSrcPrevBmp = SelectObject(lngSrcDC, picSource)
    hSavePrevBmp = SelectObject(lngSaveDC, hSaveBmp)
    hMaskPrevBmp = SelectObject(lngMaskDC, hMaskBmp)
    hInvPrevBmp = SelectObject(lngInvDC, hInvBmp)
    hDestPrevBmp = SelectObject(lngNewPicDC, hResultBmp)
    '*****************************************************************
    ' Make a backup of source bitmap to restore later
    '*****************************************************************
    BitBlt lngSaveDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, _
        lngSrcDC, 0, 0, vbSrcCopy
    '*****************************************************************
    ' Create the mask by setting the background color of source to
    ' transparent color, then BitBlt'ing that bitmap into the mask
    ' device context
    '*****************************************************************
    lngOrigColor = SetBkColor(lngSrcDC, lngMaskColor)
    BitBlt lngMaskDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, _
        lngSrcDC, 0, 0, vbSrcCopy
    '*****************************************************************
    ' Restore the original backcolor in the device context
    '*****************************************************************
    SetBkColor lngSrcDC, lngOrigColor
    '*****************************************************************
    ' Create an inverse of the mask to AND with the source and combine
    ' it with the background
    '*****************************************************************
    BitBlt lngInvDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, _
        lngMaskDC, 0, 0, vbNotSrcCopy
    '*****************************************************************
    ' Copy the background bitmap to the new picture device context
    ' to begin creating the final transparent bitmap
    '*****************************************************************
    BitBlt lngNewPicDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, _
        objDest.hdc, lngX, lngY, vbSrcCopy
    '*****************************************************************
    ' AND the mask bitmap with the result device context to create
    ' a cookie cutter effect in the background by painting the black
    ' area for the non-transparent portion of the source bitmap
    '*****************************************************************
    BitBlt lngNewPicDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, _
        lngMaskDC, 0, 0, vbSrcAnd
    '*****************************************************************
    ' AND the inverse mask with the source bitmap to turn off the bits
    ' associated with transparent area of source bitmap by making it
    ' black
    '*****************************************************************
    BitBlt lngSrcDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, _
        lngInvDC, 0, 0, vbSrcAnd
    '*****************************************************************
    ' XOR the result with the source bitmap to replace the mask color
    ' with the background color
    '*****************************************************************
    BitBlt lngNewPicDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, _
        lngSrcDC, 0, 0, vbSrcPaint
    '*****************************************************************
    ' Paint the transparent bitmap on source surface
    '*****************************************************************
    BitBlt objDest.hdc, lngX, lngY, bmpSource.bmWidth, _
        bmpSource.bmHeight, lngNewPicDC, 0, 0, vbSrcCopy
    '*****************************************************************
    ' Restore backup of bitmap
    '*****************************************************************
    BitBlt lngSrcDC, 0, 0, bmpSource.bmWidth, bmpSource.bmHeight, _
        lngSaveDC, 0, 0, vbSrcCopy
    '*****************************************************************
    ' Restore the original objects by selecting their original values
    '*****************************************************************
    SelectObject lngSrcDC, hSrcPrevBmp
    SelectObject lngSaveDC, hSavePrevBmp
    SelectObject lngNewPicDC, hDestPrevBmp
    SelectObject lngMaskDC, hMaskPrevBmp
    SelectObject lngInvDC, hInvPrevBmp
    '*****************************************************************
    ' Free system resources created by this sub
    '*****************************************************************
    DeleteObject hSaveBmp
    DeleteObject hMaskBmp
    DeleteObject hInvBmp
    DeleteObject hResultBmp
    DeleteDC lngSrcDC
    DeleteDC lngSaveDC
    DeleteDC lngInvDC
    DeleteDC lngMaskDC
    DeleteDC lngNewPicDC
    '*****************************************************************
    ' Restores the ScaleMode to its original value
    '*****************************************************************
    objDest.ScaleMode = lngOrigScaleMode
End Sub
