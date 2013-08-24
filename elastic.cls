'****************************************************************
' Name: Elasic
' Description:This class can change size and location of controls on your form if you
' 1. Resize form
' 2. Change screen resolution
' By: Mikhail Shmukler
'
' Inputs:None
' Returns:None
' Assumes:1. Add Elastic.cls 
' 2. Add declaration 'Dim El as New Elastic'
' 3. Insert string like 'El.init Me' (formload event)
' 4. Insert string like 'El.FormResize Me' (Resize event)
' 5. Press 'F5' and resize form ....
' Side Effects:None
'****************************************************************

VERSION 1.0 CLASS
BEGIN
 MultiUse = -1 'True
END

Attribute VB_Name = "Elastic"
Attribute VB_Creatable = True
Attribute VB_Exposed = False
Option Explicit
Dim iFormHeight As Integer, iFormWidth As Integer, iNumOfControls As Integer
Dim iTop() As Integer, iLeft() As Integer, iHeight() As Integer, iWidth() As Integer, iFontSize() As Integer, iRightMargin() As Integer
Dim bFirstTime As Boolean

Sub Init(FormName As Form, Optional WindState)

       Dim I As Integer
       Dim WinMax As Boolean
        WinMax = Not IsMissing(WindState)
        
        iFormHeight = FormName.Height
        iFormWidth = FormName.Width
        iNumOfControls = FormName.Controls.Count - 1
        bFirstTime = True
        ReDim iTop(iNumOfControls)
        ReDim iLeft(iNumOfControls)
        ReDim iHeight(iNumOfControls)
        ReDim iWidth(iNumOfControls)
        ReDim iFontSize(iNumOfControls)
        ReDim iRightMargin(iNumOfControls)
       On Error Resume Next
        For I = 0 To iNumOfControls

              If TypeOf FormName.Controls(I) Is Line Then
                      iTop(I) = FormName.Controls(I).Y1
                      iLeft(I) = FormName.Controls(I).X1
                      iHeight(I) = FormName.Controls(I).Y2
                      iWidth(I) = FormName.Controls(I).X2
              Else
                      iTop(I) = FormName.Controls(I).Top
                      iLeft(I) = FormName.Controls(I).Left
                      iHeight(I) = FormName.Controls(I).Height
                      iWidth(I) = FormName.Controls(I).Width
                      iFontSize(I) = FormName.FontSize
                      iRightMargin(I) = FormName.Controls(I).RightMargin
              End If

        Next
        
        If WinMax Or FormName.WindowState = 2 Then ' maxim
        FormName.Height = Screen.Height
        FormName.Width = Screen.Width
        Else
        FormName.Height = FormName.Height * Screen.Height / 7290
        FormName.Width = FormName.Width * Screen.Width / 9690
        End If
        
        bFirstTime = True
        
End Sub


Sub FormResize(FormName As Form)

       Dim I As Integer, Inc As Integer, CaptionSize As Integer
       Dim RatioX As Double, RatioY As Double
       Dim SaveRedraw%
       On Error Resume Next
       SaveRedraw% = FormName.AutoRedraw

              FormName.AutoRedraw = True

                            If bFirstTime Then
                                   bFirstTime = False
                                   Exit Sub
                            End If


                            If FormName.Height < iFormHeight / 2 Then FormName.Height = iFormHeight / 2

                                          If FormName.Width < iFormWidth / 2 Then FormName.Width = iFormWidth / 2
                                                 CaptionSize = 400
                                                 RatioY = 1# * (iFormHeight - CaptionSize) / (FormName.Height - CaptionSize)
                                                 RatioX = 1# * iFormWidth / FormName.Width
                                                 On Error Resume Next ' for comboboxes, timeres and other nonsizible controls

                                                        For I = 0 To iNumOfControls

                                                                      If TypeOf FormName.Controls(I) Is Line Then
                                                                              FormName.Controls(I).Y1 = Int(iTop(I) / RatioY)
                                                                              FormName.Controls(I).X1 = Int(iLeft(I) / RatioX)
                                                                              FormName.Controls(I).Y2 = Int(iHeight(I) / RatioY)
                                                                              FormName.Controls(I).X2 = Int(iWidth(I) / RatioX)
                                                                      Else
                                                                              FormName.Controls(I).Top = Int(iTop(I) / RatioY)
                                                                              FormName.Controls(I).Left = Int(iLeft(I) / RatioX)
                                                                              FormName.Controls(I).Height = Int(iHeight(I) / RatioY)
                                                                              FormName.Controls(I).Width = Int(iWidth(I) / RatioX)
                                                                              FormName.Controls(I).FontSize = Int(iFontSize(I) / RatioX) + Int(iFontSize(I) / RatioX) Mod 2
                                                                              FormName.Controls(I).RightMargin = Int(iRightMargin(I) / RatioY)
                                                                      End If

                                                        Next


                                                        FormName.AutoRedraw = SaveRedraw%
                                                        End Sub

