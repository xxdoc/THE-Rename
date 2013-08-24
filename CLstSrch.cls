VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CListSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' *********************************************************************
'  Copyright ©1996-99 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  Warning: This computer program is protected by copyright law and
'  international treaties. Unauthorized reproduction or distribution
'  of this program, or any portion of it, may result in severe civil
'  and criminal penalties, and will be prosecuted to the maximum
'  extent possible under the law.
' *********************************************************************
Option Explicit
'
' API Declarations and constants
'
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_FINDSTRING = &H18F
Private Const LB_ERR = (-1)
'
' Set aside storage for private member variables.
'
Private m_ExtendSearch As Boolean
Private m_LastKey As Double
Private m_TimeLimit As Long
Private m_Beep As Boolean
'
' Use WithEvents for VB5 and later.
'
Private WithEvents m_List As ListBox
Attribute m_List.VB_VarHelpID = -1

' ********************************************
'  Initialize
' ********************************************
Private Sub Class_Initialize()
   '
   ' Set default values for class properties.
   '
   Set m_List = Nothing
   m_ExtendSearch = False
   m_TimeLimit = 2000      '2 seconds
   m_LastKey = 0
   m_Beep = True           'Beep if search fails
End Sub

' ********************************************
'  Public Properties
' ********************************************
Public Property Set Client(NewObj As Object)
   '
   ' Set new ListBox as Client property.
   '
   If TypeName(NewObj) = "ListBox" Then
      If NewObj.Sorted Then
         Set m_List = NewObj
         m_LastKey = Now
      Else
         Err.Raise Number:=vbObjectError + 2, _
                   Source:="CListSearch.Client", _
                   Description:="Client ListBox must have Sorted = True."
      End If
   Else
      Err.Raise Number:=vbObjectError + 1, _
                Source:="CListSearch.Client", _
                Description:="Client property must be of type ListBox."
   End If
End Property

Public Property Get Client() As Object
   '
   ' Return ListBox as Client property.
   '
   Set Client = m_List
End Property

Public Property Let TimeLimit(NewVal As Long)
   '
   ' Set new value for number of milliseconds
   ' to wait between keystrokes when continuing
   ' a search.
   '
   If NewVal > 0 Then
      m_TimeLimit = NewVal
   Else
      m_TimeLimit = 0
   End If
End Property

Public Property Get TimeLimit() As Long
   '
   ' Return current value for TimeLimit property.
   '
   TimeLimit = m_TimeLimit
End Property

Public Property Let AudibleError(NewVal As Boolean)
   '
   ' Store whether or not to beep when search fails.
   '
   m_Beep = NewVal
End Property

Public Property Get AudibleError() As Boolean
   '
   ' Return current value for AudibleError property.
   '
   AudibleError = m_Beep
End Property

' ********************************************
'  Public Methods
' ********************************************
Public Function KeyPress(KeyAscii As Integer)
   '
   ' Leave public function in place, so that
   ' VB4 users will still have an entry point
   ' if they use this class.
   '
   Static Search As String
   Dim Index As Long
   Dim DoSearch As Boolean
   Dim Elapsed As Double
   Const SecsPerDay = 86400
   '
   ' Check if more than allowed time has elapsed.
   '
   If m_ExtendSearch Then
      Elapsed = Now - m_LastKey
      If (Elapsed * SecsPerDay) > (m_TimeLimit / 1000) Then
         m_ExtendSearch = False
      End If
   End If
   '
   ' Start over if delay was too long.
   '
   If Not m_ExtendSearch Then
      Search = ""
      m_ExtendSearch = True
      Index = m_List.ListIndex
   Else
      Index = -1
   End If
   '
   ' Check for valid keystrokes.
   '
   If KeyAscii = vbKeyBack Then
      '
      ' Allow user to take back last key.
      '
      If Len(Search) Then
         Search = left$(Search, Len(Search) - 1)
         DoSearch = True
      End If
   ElseIf KeyAscii >= vbKeySpace Then
      '
      ' Append latest key.
      '
      Search = Search & Chr$(KeyAscii)
      DoSearch = True
   End If
   '
   ' Perform search after valid keystrokes.
   '
   If (DoSearch = True) And (Len(Search) > 0) Then
      Index = SendMessage(m_List.hWnd, LB_FINDSTRING, Index, ByVal Search)
      If Index <> LB_ERR Then 'Found a match!
         m_List.ListIndex = Index
      Else 'No match
         Search = left$(Search, Len(Search) - 1)
         If m_Beep Then Beep
      End If
      '
      ' Record when key was pressed, and consume
      ' keystroke (by returning 0) so VB doesn't
      ' automatically move list to entry that
      ' starts with last key.
      '
      m_LastKey = Now
      KeyPress = 0
   Else
      '
      ' Return passed KeyAscii value so original
      ' KeyPress routine can continue processing.
      '
      KeyPress = KeyAscii
   End If
End Function

' ********************************************
'  Private Methods
' ********************************************
Private Sub m_List_KeyPress(KeyAscii As Integer)
   '
   ' Call searching code on each keystroke.
   '
   KeyAscii = KeyPress(KeyAscii)
End Sub
