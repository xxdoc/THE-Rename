VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "CCRPFTV6.OCX"
Object = "{06D5A045-D511-11D3-9875-BB56A32B4523}#1.0#0"; "PROPBRWS.OCX"
Begin VB.Form Form4 
   Caption         =   "Files & Folders"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   7680
   Begin TabDlg.SSTab Tab1 
      Height          =   6495
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11456
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Files && Fodlers"
      TabPicture(0)   =   "Form4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FTV1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LV1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Splitter"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Properties"
      TabPicture(1)   =   "Form4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Pb2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Viewer"
      TabPicture(2)   =   "Form4.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin PB.PropertyBrowser Pb2 
         Height          =   6090
         Left            =   -74955
         TabIndex        =   6
         Top             =   45
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   10742
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty CatFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NameWidth       =   75
      End
      Begin VB.PictureBox Splitter 
         AutoRedraw      =   -1  'True
         Height          =   5775
         Left            =   3000
         ScaleHeight     =   5775
         ScaleWidth      =   15
         TabIndex        =   5
         Top             =   120
         Visible         =   0   'False
         Width           =   15
      End
      Begin MSComctlLib.ListView LV1 
         Height          =   6090
         Left            =   3000
         TabIndex        =   4
         Top             =   45
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   10742
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin CCRPFolderTV6.FolderTreeview FTV1 
         Height          =   6090
         Left            =   45
         TabIndex        =   3
         Top             =   45
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   10742
         IntegralHeight  =   0   'False
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9120
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   13421772
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0054
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":05A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":0AFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":1050
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":1AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":204C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":25A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":2AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":3048
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":359C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":3AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":4044
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":4598
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":4AEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":5040
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":5594
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":5AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":603C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form4.frx":6590
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   174
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select all"
            Description     =   "Select all files"
            Object.ToolTipText     =   "Select all files in current directory"
            Object.Tag             =   "Select all files in current directory"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "unselect"
            Description     =   "Unselect files from current directory"
            Object.ToolTipText     =   "Unselect files from current directory"
            Object.Tag             =   "Unselect files from current directory"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "invert selection"
            Description     =   "Invert selection"
            Object.ToolTipText     =   "Invert selection"
            Object.Tag             =   "Invert selection"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step"
            Description     =   "Enter a step value for selection"
            Object.ToolTipText     =   "Enter a step value for selection"
            Object.Tag             =   "Enter a step value for selection"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DropFiles"
            Description     =   "Drop files to copy them, click to configure"
            Object.ToolTipText     =   "Drop files to copy them, click to configure"
            Object.Tag             =   "Drop files to copy them, click to configure"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Recursive mode"
            Description     =   "Recursive mode"
            Object.ToolTipText     =   "Recursive mode"
            Object.Tag             =   "Recursive mode"
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Up one level"
            Description     =   "Up one level"
            Object.ToolTipText     =   "Up one level"
            Object.Tag             =   "Up one level"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Root directory"
            Description     =   "Root directory"
            Object.ToolTipText     =   "Root directory"
            Object.Tag             =   "Root directory"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Show history of your moves"
            Description     =   "Show history of your moves"
            Object.ToolTipText     =   "Show history of your moves"
            Object.Tag             =   "Show history of your moves"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add to your favorites"
            Object.ToolTipText     =   "Add to your favorites"
            Object.Tag             =   "Add to your favorites"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "organize your favorites"
            Object.ToolTipText     =   "Organize your favorites"
            Object.Tag             =   "organize your favorites"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "First favorite"
            Object.ToolTipText     =   "First favorite"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previous favorite"
            Object.ToolTipText     =   "previous favorite"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next favorite"
            Object.ToolTipText     =   "Next favorite"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Last favorite"
            Object.ToolTipText     =   "Last favorite"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
      Begin VB.ComboBox Combo5 
         Height          =   315
         IntegralHeight  =   0   'False
         ItemData        =   "Form4.frx":6AE4
         Left            =   6075
         List            =   "Form4.frx":6AE6
         TabIndex        =   1
         ToolTipText     =   "Type a file filter and press enter or select a filter from the list"
         Top             =   20
         WhatsThisHelpID =   182
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'variable to hold the width of the splitter bar
 Private Const SPLT_WDTH As Integer = 3
'variable to hold the last-sized position
 Private currSplitPosX As Long
'variable to hold the horizontal
'and vertical offsets of the 2 controls
 Dim CTRL_OFFSET As Integer
'variable to hold the Splitter bar colour
 Dim SPLT_COLOUR As Long


Private Sub Form_Load()
  AddPropertiesPB2
 
 'set the startup variables
  CTRL_OFFSET = 5
  SPLT_COLOUR = &H808080
 ' set the current splitter bar position to an
 'arbitrary value that will always be outside
 'the possible range.  This allows us to check
 'for movement of the splitter bar in subsequent
 'mousexxx subs.
  currSplitPosX = &H7FFFFFFF
  
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Tab1.width = Form4.ScaleWidth
    Tab1.height = Form4.ScaleHeight - Toolbar1.height - 50
    FTV1.height = Tab1.height - 405
    LV1.height = FTV1.height
    FTV1.width = (Tab1.width * 30) / 100
    LV1.left = FTV1.left + FTV1.width + 60
    LV1.width = Tab1.width - FTV1.left - FTV1.width - 135
    Pb2.width = Tab1.width - 120
    Pb2.height = LV1.height

'  Dim x1 As Integer
'  Dim x2 As Integer
'  Dim height1 As Integer
'  Dim width1 As Integer
'  Dim width2 As Integer
'
'  On Error Resume Next
'
' 'set the height of the controls
'  height1 = ScaleHeight - (CTRL_OFFSET * 2)
'
'  x1 = CTRL_OFFSET
'  width1 = ListLeft.width
'
'  x2 = x1 + ListLeft.width + SPLT_WDTH - 1
'  width2 = ScaleWidth - x2 - CTRL_OFFSET
'
' 'move the left list
'  ListLeft.Move x1 - 1, CTRL_OFFSET, width1, height1
'
' 'move the right list
'  TextRight.Move x2, CTRL_OFFSET, width2 + 1, height1
'
' 'reposition the splitter bar
'  Splitter.Move x1 + ListLeft.width - 1, _
'               CTRL_OFFSET, SPLT_WDTH, height1
  

End Sub

Private Sub FTV1_Click()
    Form4.Caption = FTV1.SelectedFolder
End Sub

Private Sub Splitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
    
       'change the splitter colour
        Splitter.BackColor = SPLT_COLOUR
       
       'set the current position to x
        currSplitPosX = CLng(X)
    
    Else
    
       'not the left button, so...
       'if the current position <> default, cause a MouseUp
        If currSplitPosX <> &H7FFFFFFF Then
           Splitter_MouseUp Button, Shift, X, Y
        End If
       
       'set the current position to the default value
        currSplitPosX = &H7FFFFFFF
    
    End If
End Sub

Private Sub Splitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if the splitter has been moved...
  If currSplitPosX <> &H7FFFFFFF Then
  
   'if the current position <> default, reposition
   'the splitter and set this as the current value
    
    If CLng(X) <> currSplitPosX Then
        Splitter.Move Splitter.left + X, _
                      CTRL_OFFSET, SPLT_WDTH, _
                      ScaleHeight - (CTRL_OFFSET * 2)
        currSplitPosX = CLng(X)
    End If
End If

End Sub

Private Sub Splitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 'if the splitter has been moved...
  If currSplitPosX <> &H7FFFFFFF Then
      
     'if the current position <> the last
     'position do a final move of the splitter
      If CLng(X) <> currSplitPosX Then
        Splitter.Move Splitter.left + X, _
                     CTRL_OFFSET, SPLT_WDTH, _
                     ScaleHeight - (CTRL_OFFSET * 2)
      End If
      
     'call this the default position
      currSplitPosX = &H7FFFFFFF
      
     'restore the normal splitter colour
      Splitter.BackColor = &H8000000F
     
     'and check for valid sizing.
     'Either enforce the default minimum & maximum widths for
     'the left list, or, if within range, set the width
     
      If Splitter.left > 60 And Splitter.left < (ScaleWidth - 60) Then
            'the pane is within range
             ListLeft.width = Splitter.left - ListLeft.left
      
      ElseIf Splitter.left < 60 Then
            'the pane is too small
             ListLeft.width = 60
      
      Else: 'the pane is too wide
             ListLeft.width = ScaleWidth - 60
      End If
      
     'reposition both lists, and the splitter bar
      Form_Resize
  
  End If
End Sub

Sub AddPropertiesPB2()

    With Pb2

        With .Categories.Add("Numeric Types", pbImgOpenFolder).Properties
            .Add "Byte", 128, pbByte
            
            With .Add("Currency", 12300, pbCurrency, , , "@pbCurrency properties have a" & vbCrLf & "default format of ""$ #,##0.00""")
               .UpDownIncrement = 0.05
            End With
            
            .Add "Integer", 1, pbInteger
            
            With .Add("Long", 200, pbLong, , , "This property has maximum and minimum values and an UpDown control.")
               .SetRange 100, 1000
               .UpDownIncrement = 10
            End With
            
            .Add "Decimal", 32312223.21, pbDecimal
            .Add "Double", 1639043.324, pbDouble
            .Add "Single", 123 / 3, pbSingle
        End With
        
        With .Categories.Add("Date/Time Types", pbImgOpenFolder).Properties
            .Add "Time", Now(), pbTime, , pbImgClock, "PropertyBrowser supports Time and Date properties."
            .Add "Date", #10/22/1932#, pbDate, , pbImgCalendar2, "Date properties shows a calendar" & vbCrLf & "control to select a valid date."
        End With
        
        With .Categories.Add("String Types", pbImgOpenFolder).Properties
            .Add "String", "Hello World!", pbString
            .Add("String * 8", "12345678", pbString).SetRange , 8
            .Add("Password", "pwd", pbString).Format = "Password"
            .Add "File", "C:\AUTOEXEC.BAT", pbFile, , pbImgPaperClip
            .Add "Combo", "Text", pbCombo
            
            With .Item("Combo").ListValues

                .Add "Sample Item"
                .Add "Combo item", "New Combo Type"

            End With
        
        End With
        
        With .Categories.Add("Object Types", pbImgOpenFolder).Properties
            .Add "Font", Me.Font, , , pbImgFont
            .Add "Object", Empty
            .Add("Picture", Nothing, pbPicture, , pbImgPicture1).Format = "CustomDisplay"
        End With
        
        With .Categories.Add("Other Types", pbImgOpenFolder).Properties
            
            .Add "Array", Array(1, 2, 3)
            .Add "Boolean", False, pbBoolean
            .Add("Color", vbBlue, pbColor, , , "This property uses the CustomDisplay format.").Format = "CustomDisplay"
            
            With .Add("DropDown List", 0, pbDropDownList, , , _
                  "@This kind of property" & vbCrLf & _
                  "is not limited only to" & vbCrLf & _
                  "Long values. You can use" & vbCrLf & _
                  "anything that can be" & vbCrLf & _
                  "stored in a Variant.").ListValues
                .Add "Item A"
                .Add "Item B"
            End With
            
        End With
        
        With .Categories.Add("Empty Category").Properties
        End With

    End With
End Sub
