VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cObjectSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************
' Name: cResize
' Description:Automatically resizes all controls on a form t
'     o provide resolution independence.
' By: Michael D. Long
'
' Inputs:Requires that the class be declared on each form, and the .Resize method to be called from each forms Resize event.
' Returns:None
' Assumes:All controls on a managed form will be proportionally scale when the form is resized. Limitations: no font sizing support, and the ComboBox Height property does not resize.
' Side Effects:None
'
'Code provided by Planet Source Code(tm) 'as is', without
'     warranties as to performance, fitness, merchantability,
'     and any other warranty (whether expressed or implied).
'****************************************************************

'The cResize class proportionally scales all controls on a f
'     orm.
'This is used to allow automatic adjustment for different sc
'     reen
'     'resolutions. Also, implementation in code reduces resource
'     'consumption and distribution requirements.
'     '
'To implement, simply add a declaration to each form that yo
'     u
'want resized and a call to Resize method in each Form_Resiz
'     e
'     'event.
'     '
'     '------------------------------------
'     'Code to add to each form
Private m_cResize As New cResize

Private Sub Form_Resize()

        m_cResize.Resize Me
End Sub

'     '------------------------------------
'     '
'     'Class cObjectSize - member variable declarations only
'     ' Object size properties
Public Height As Long
Public Left As Long
Public Top As Long
Public Width As Long
'     '------------------------------------
'     '
'     'Class cResize - declarations and control resizing code
Option Explicit
Private m_cSize As New Collection ' Collection of size properties
Private m_iNbrControls As Integer ' Previous number of controls on form
Private m_nFormHeight As Long ' Original form height
Private m_nFormWidth As Long ' Original form width

Private Sub Class_Initialize()

        m_nFormHeight = 0
        m_nFormWidth = 0
        m_iNbrControls = 0
End Sub

' Expose the method to perform scaling of the controls on th
'     e form

Public Sub Resize(ByRef frmToResize As Form)

        Dim cControls As cObjectSize
        Dim dblScaleH As Double
        Dim dblScaleW As Double
        Dim iCnt As Integer
        Dim iItm As Integer
        Dim iNbrControls As Integer
        On Error Resume Next
        iNbrControls = frmToResize.Controls.Count
        If iNbrControls <> m_iNbrControls Then
       '     ' Save off the original controls size properties

              For iCnt = m_iNbrControls To iNbrControls - 1
                      Set cControls = New cObjectSize
                      cControls.Height = frmToResize.Controls(iCnt).Height
                      cControls.Left = frmToResize.Controls(iCnt).Left
                      cControls.Top = frmToResize.Controls(iCnt).Top
                      cControls.Width = frmToResize.Controls(iCnt).Width
                      m_cSize.Add cControls
              Next iCnt

       m_iNbrControls = iNbrControls
        End If
       '      ' Check to see if the form has just been loaded; we only
       '      ' need to resize the controls on subsequent events.
        If (m_nFormHeight <> 0 And m_nFormWidth <> 0) Then
       '     ' Calculate the scaling factor; doubles used for precision
       dblScaleH = frmToResize.ScaleHeight / m_nFormHeight
       dblScaleW = frmToResize.ScaleWidth / m_nFormWidth
       '     ' Loop through the controls and resize

              For iCnt = 0 To iNbrControls - 1
                      iItm = iCnt + 1 ' Compensate for collection being 1 based
                      If TypeOf frmToResize.Controls(iCnt) Is Timer Then
                     '     ' do nothing
                      ElseIf TypeOf frmToResize.Controls(iCnt) Is ComboBox Then
                     '     ' attempting to resize height results in error
                     frmToResize.Controls(iCnt).Move _
                      m_cSize.Item(iItm).Left * dblScaleW, _
                      m_cSize.Item(iItm).Top * dblScaleH, _
                      m_cSize.Item(iItm).Width * dblScaleW
                      Else
                     frmToResize.Controls(iCnt).Move _
                      m_cSize.Item(iItm).Left * dblScaleW, _
                      m_cSize.Item(iItm).Top * dblScaleH, _
                      m_cSize.Item(iItm).Width * dblScaleW, _
                      m_cSize.Item(iItm).Height * dblScaleH
                      End If
              Next iCnt

        Else
       ' Store original form height and width; used for scale calcs
       '
       m_nFormHeight = frmToResize.Height
       m_nFormWidth = frmToResize.Width
        End If
End Sub


