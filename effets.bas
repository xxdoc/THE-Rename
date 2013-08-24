Attribute VB_Name = "effets"
'****************************************************************
' Routine pour mettre de l'ombre sur des objets
'****************************************************************
'****************************************************************
'Windows API/Global Declarations for :FormControlShadow
'****************************************************************

'     ' Label and Shape Styles
Global Const GFM_STANDARD = 0
Global Const GFM_RAISED = 1
Global Const GFM_SUNKEN = 2
'     ' Control Shadow Styles
Global Const GFM_BACKSHADOW = 1
Global Const GFM_DROPSHADOW = 2
'     ' Color constants
Global Const BOX_WHITE& = &HFFFFFF
Global Const BOX_LIGHTGRAY& = &HC0C0C0
Global Const BOX_DARKGRAY& = &H808080
Global Const BOX_BLACK& = &H0&


'****************************************************************
' Name: FormControlShadow
' Description:This routine is used to create a Back or Drop
'     shadow effect on any controls which are placed on a form. Si
'     mply place the control as normal and invoke the 'shadow with
'      the code below.
' By: Gord's VB Code Snippets
'
' Inputs:' Parameters Type Comment
'     'fForm the form containing the control
'     'CControlthe control to shadow
'     'shadow_effectintegerGFM_DROPSHADOW or GFM_BACKSHADOW
'     'shadow_widthintegerwidth of the shadow in pixels
'     'shadow_colorlong color of the shadow
' Returns:None
' Assumes:' Example:
'     ' In the Form_Paint event:
' FormControlShadow Me, Text1, GFM_DROPSHADOW, 2, QBColor(8)
' Side Effects:None
'
'Code provided by Planet Source Code(tm) 'as is', without
'     warranties as to performance, fitness, merchantability,
'     and any other warranty (whether expressed or implied).
'****************************************************************


Sub FormControlShadow(f As Form, C As Control, shadow_effect As Integer, shadow_width As Integer, shadow_color As Long)
       Dim shColor As Long
       Dim shWidth As Integer
       Dim oldWidth As Integer
       Dim oldScale As Integer
       shWidth = shadow_width
       shColor = shadow_color
       oldWidth = f.DrawWidth
       oldScale = f.ScaleMode
       f.ScaleMode = 3 'Pixels
       f.DrawWidth = 1
        
        Select Case shadow_effect
        Case GFM_DROPSHADOW
       f.Line (C.Left + shWidth, C.Top + shWidth)-Step(C.Width - 1, C.Height - 1), shColor, BF
        Case GFM_BACKSHADOW
       f.Line (C.Left - shWidth, C.Top - shWidth)-Step(C.Width - 1, C.Height - 1), shColor, BF
End Select

f.DrawWidth = oldWidth
f.ScaleMode = oldScale
End Sub


'****************************************************************
' Name: How to draw 3D offset bevels around control
' Description:Here's a routine for 3D offset bevels on contr
'     ols
' By: Gord's VB Code Snippets
'
' Inputs:
'     'Parameters:
'     ' Ctrl= apply 3D look to control name
'     ' nBevel% = bevel width (pixels)
'     ' nSpace% = surround distance from control (pixels)
'     ' bInset% = True is 3D inset border
'     'False is 3D outset border

' Returns:None
' Assumes:Looks best when background of form or container is light gray.
'Example of calling this routine:
'In the form's Paint event:
'MakeIt3D Text1, 1, 0, True
' Side Effects:None
'
'Code provided by Planet Source Code(tm) 'as is', without
'     warranties as to performance, fitness, merchantability,
'     and any other warranty (whether expressed or implied).
'****************************************************************


Sub MakeIt3D(Ctrl As Control, nBevel%, nSpace%, bInset%)

        
       '     'Makes the passed control appear 3D.
       PixX% = Screen.TwipsPerPixelX
       PixY% = Screen.TwipsPerPixelY
       CTop% = Ctrl.Top - PixX%
       CLft% = Ctrl.Left - PixY%
       CRgt% = Ctrl.Left + Ctrl.Width
       CBtm% = Ctrl.Top + Ctrl.Height
       '     ' Color used below:
       '     ' dark gray = &H808080
       '     ' white = &HFFFFFF

              If bInset% Then 'recessed border

                            For i% = nSpace% To (nBevel% + nSpace% - 1)
                                   AddX% = i% * PixX%
                                   AddY% = i% * PixY%
                                   Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CRgt% + AddX%, CTop% - AddY%), 8421504
                                   Ctrl.Parent.Line (CLft% - AddX%, CTop% - AddY%)-(CLft% - AddX%, CBtm% + AddY%), 8421504
                                   Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CRgt% + AddX% + PixX%, CBtm% + AddY%), &HFFFFFF
                                   Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CRgt% + AddX%, CBtm% + AddY%), 16777215
                            Next

              Else 'raised border

                            For i% = nSpace% To (nBevel% + nSpace% - 1)
                                   AddX% = i% * PixX%
                                   AddY% = i% * PixY%
                                   Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CRgt% + AddX%, CTop% - AddY%), 8421504
                                   Ctrl.Parent.Line (CRgt% + AddX%, CBtm% + AddY%)-(CLft% - AddX%, CBtm% + AddY%), 8421504
                                   Ctrl.Parent.Line (CRgt% + AddX%, CTop% - AddY%)-(CLft% - AddX% - PixX%, CTop% - AddY%), &HFFFFFF
                                   Ctrl.Parent.Line (CLft% - AddX%, CBtm% + AddY%)-(CLft% - AddX%, CTop% - AddY%), 16777215
                            Next

              End If

End Sub

