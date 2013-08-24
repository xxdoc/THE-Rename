Attribute VB_Name = "zlistview"
Option Explicit
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public objFind As LV_FINDINFO
Public objItem As LV_ITEM
Dim zItemX As ListItem
'variable to hold the sort order (ascending or descending)
Public sOrder As Boolean
     
'API declarations
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Function CompareDates(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal hWnd As Long) As Long
  'CompareDates: This is the sorting routine that gets passed to the
  'ListView control to provide the comparison test for date values.
  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
   Dim dDate1 As Date
   Dim dDate2 As Date

  'Obtain the item names and dates corresponding to the
  'input parameters
   dDate1 = ListView_GetItemDate(lParam1, hWnd)
   dDate2 = ListView_GetItemDate(lParam2, hWnd)
     
  'based on the Public variable sOrder set in the
  'columnheader click sub, sort the dates appropriately:
   Select Case sOrder
      Case True:    'sort descending
            If dDate1 < dDate2 Then
                  CompareDates = 0
            ElseIf dDate1 = dDate2 Then
                  CompareDates = 1
            Else: CompareDates = 2
            End If
      
      Case Else: 'sort ascending
            If dDate1 > dDate2 Then
                  CompareDates = 0
            ElseIf dDate1 = dDate2 Then
                  CompareDates = 1
            Else: CompareDates = 2
            End If
   
   End Select

End Function

Public Function CompareValues(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal hWnd As Long) As Long
  'CompareValues: This is the sorting routine that gets passed to the
  'ListView control to provide the comparison test for numeric values.
  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
   Dim val1 As Long
   Dim val2 As Long
     
  'Obtain the item names and values corresponding
  'to the input parameters
   val1 = ListView_GetItemValuestr(hWnd, lParam1)
   val2 = ListView_GetItemValuestr(hWnd, lParam2)
  'based on the Public variable sOrder set in the
  'columnheader click sub, sort the values appropriately:
   Select Case sOrder
      Case True:    'sort descending
            If val1 < val2 Then
                  CompareValues = 0
            ElseIf val1 = val2 Then
                  CompareValues = 1
            Else: CompareValues = 2
            End If
      
      Case Else: 'sort ascending
            If val1 > val2 Then
                  CompareValues = 0
            ElseIf val1 = val2 Then
                  CompareValues = 1
            Else: CompareValues = 2
            End If
   End Select
End Function

Public Function ListView_GetItemDate(lParam As Long, hWnd As Long) As Date
   Dim R As Long
   Dim hIndex As Long
  'Convert the input parameter to an index in the list view
   objFind.flags = LVFI_PARAM
   objFind.lParam = lParam
   hIndex = SendMessageAny(hWnd, LVM_FINDITEM, -1, objFind)
  'Obtain the value of the specified list view item.
  'The objItem.iSubItem member is set to the index
  'of the column that is being retrieved.
   objItem.Mask = LVIF_TEXT
Rem vérifier
   objItem.iSubItem = 2
   objItem.pszText = Space$(32)
   objItem.cchTextMax = Len(objItem.pszText)
  'get the string at subitem 1
   R = SendMessageAny(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
  'and convert it into a date and exit
   If R > 0 Then
      ListView_GetItemDate = CDate(Left$(objItem.pszText, R))
   End If
End Function

Public Function ListView_GetItemValuestr(hWnd As Long, lParam As Long) As Long
   Dim R As Long
   Dim hIndex As Long
  'Convert the input parameter to an index in the list view
   objFind.flags = LVFI_PARAM
   objFind.lParam = lParam
   hIndex = SendMessageAny(hWnd, LVM_FINDITEM, -1, objFind)
  'Obtain the value of the specified list view item.
  'The objItem.iSubItem member is set to the index
  'of the column that is being retrieved.
   objItem.Mask = LVIF_TEXT
   objItem.iSubItem = 1
   objItem.pszText = Space$(32)
   objItem.cchTextMax = Len(objItem.pszText)
     
  'get the string at subitem 2
   R = SendMessageAny(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
     
  'and convert it into a long
   If R > 0 Then
      ListView_GetItemValuestr = CLng(Left$(objItem.pszText, R))
   End If
End Function
' Return label of an item in a listview
Public Function LVGetName(lv As ListView, lindex As Long) As String
 Dim R As Long, sItem As String
 R = 0
 objItem.Mask = LVIF_TEXT
 objItem.iSubItem = 0
 objItem.pszText = Space$(256)
 objItem.cchTextMax = 256
 R = SendMessageAny(lv.hWnd, LVM_GETITEMTEXT, lindex, objItem)
 sItem = Left$(objItem.pszText, R)
 LVGetName = sItem
End Function
' Return True if an item is selected, or false
Public Function LVIsSelected(lv As ListView, Index As Long) As Boolean
 Dim R As Long
 R = 0
 R = SendMessageLong(lv.hWnd, LVM_GETITEMSTATE, Index, LVIS_SELECTED)
 If R And LVIS_SELECTED Then
  LVIsSelected = True
 Else
  LVIsSelected = False
 End If
End Function
' Delete an item from a listview
Public Function LVRemoveItem(lv As ListView, Index As Long) As Boolean
 Dim R As Long
 R = 0
 R = SendMessageAny(lv.hWnd, LVM_DELETEITEM, Index, 0&)
 LVRemoveItem = R
End Function
' Return # of selected items in a listview
Public Function LVGetCountSelected(lv As ListView) As Long
 Dim R As Long
 R = 0
 R = SendMessageLong(lv.hWnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
 LVGetCountSelected = R
End Function
' Search for a text and return its index or -1 if it fails
' Modif du 06/08/99, le paramètre de retour était un INT, je l'ai passé en long
Public Function LVSearch(lv As ListView, texte As String) As Long
 Dim R As Long
 R = 0
 objFind.flags = LVFI_STRING
 objFind.pSz = texte
 objFind.lParam = 0
 R = SendMessageAny(lv.hWnd, LVM_FINDITEM, -1, objFind)
 LVSearch = R
End Function
' Select an item in a listview
Public Sub LVSetItemSelected(lv As ListView, Index As Long)
 Dim R As Long
 objItem.Mask = LVIF_STATE
 objItem.iSubItem = 0
 objItem.State = LVIS_SELECTED
 objItem.StateMask = LVIS_SELECTED
 R = SendMessageAny(lv.hWnd, LVM_SETITEMSTATE, Index, objItem)
End Sub
' Unselect an item in a listview
Public Sub LVSetItemNotSelected(lv As ListView, Index As Long)
 Dim R As Long
 On Error Resume Next
 R = 0
 objItem.Mask = LVIF_STATE
 objItem.iSubItem = 0
 objItem.State = Not LVIS_SELECTED
 objItem.StateMask = LVIS_SELECTED
 R = SendMessageAny(lv.hWnd, LVM_SETITEMSTATE, Index, objItem)
End Sub
' Return indexes of selected items
' Modif du 06/08/99, le paramètre de retour était un INT, je l'ai passé en long
Public Function LVGetItemSelected(lv As ListView, Index As Long) As Long
 Dim lRet As Long
 lRet = SendMessageLong(lv.hWnd, LVM_GETNEXTITEM, Index, LVNI_SELECTED)
 LVGetItemSelected = lRet
End Function
' Return the label of a sub item in a listview
Public Function LVGetItemName(lv As ListView, Index As Long, Item As Long) As String
 Dim R As Long, sItem As String
 R = 0
 objItem.Mask = LVIF_TEXT
 objItem.iSubItem = Item
 objItem.pszText = Space$(256)
 objItem.cchTextMax = 256
 R = SendMessageAny(lv.hWnd, LVM_GETITEMTEXT, Index, objItem)
 sItem = Left$(objItem.pszText, R)
 LVGetItemName = sItem
End Function
Public Sub AutoSizeColumnHeader(LView As ListView, Column As ColumnHeader, ByVal SizeToHeader As Boolean)
Dim l As Long
If SizeToHeader Then
 l = -2
Else
 l = -1
End If
Call SendMessage(LView.hWnd, LVM_FIRST + 30, Column.Index - 1, l)
End Sub
Public Function ListView_GetItemValueString(lParam As Long, hWnd As Long) As String
   Dim R As Long
   Dim hIndex As Long
   objFind.flags = LVFI_PARAM
   objFind.lParam = lParam
   hIndex = SendMessageAny(hWnd, LVM_FINDITEM, -1, objFind)
   objItem.Mask = LVIF_TEXT
   objItem.iSubItem = 0
   objItem.pszText = Space$(256)
   objItem.cchTextMax = Len(objItem.pszText)
   R = SendMessageAny(hWnd, LVM_GETITEMTEXT, hIndex, objItem)
   If R > 0 Then
      ListView_GetItemValueString = Left$(objItem.pszText, R)
   End If
End Function

Public Function CompareNatural(ByVal lParam1 As Long, ByVal lParam2 As Long, ByVal hWnd As Long) As Long
  'CompareDates: This is the sorting routine that gets passed to the
  'ListView control to provide the comparison test for date values.
  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
   Dim dDate1 As String
   Dim dDate2 As String
   Dim vretour As Integer

  'Obtain the item names and dates corresponding to the input parameters
   dDate1 = ListView_GetItemValueString(lParam1, hWnd)
   dDate2 = ListView_GetItemValueString(lParam2, hWnd)
   'based on the Public variable sOrder set in the
   Select Case sOrder
      Case True:    'sort descending
            vretour = strnatcasecmp(dDate1, dDate2)
            Select Case vretour
                Case 0  ' égal
                    CompareNatural = 1
                Case 1  ' Supérieur
                    CompareNatural = 2
                Case -1 ' Inférieur
                 CompareNatural = 0
            End Select
      Case Else: 'sort ascending
            vretour = strnatcasecmp(dDate1, dDate2)
            Select Case vretour
                Case 0  ' égal
                    CompareNatural = 1
                Case 1  ' Supérieur
                    CompareNatural = 0
                Case -1
                 CompareNatural = 2
            End Select
   End Select
End Function

Public Function MoveRow(lv As ListView, from As Long, tto As Long) As Boolean
    On Error Resume Next
    If tto > lv.ListItems.Count Then
        MoveRow = False
        Exit Function
    End If
    If from = tto Or from < 1 Or tto < 1 Then
        MoveRow = False
        Exit Function
    End If
    
    If from < tto Then
        tto = tto + 1
    End If
    
    If (CopyRow(lv, from, tto)) Then
        If (from > tto) Then
            'Call LVRemoveItem(lv, from + 1)
            lv.ListItems.Remove (from + 1)
        Else
            'Call LVRemoveItem(lv, from)
            lv.ListItems.Remove (from)
        End If
        MoveRow = True
        Exit Function
    Else
        MoveRow = False
        Exit Function
    End If
End Function

Public Function CopyRow(lv As ListView, from As Long, tto As Long) As Boolean
    Dim i As Long
    Dim zvnb As Integer
    On Error Resume Next
    If from = tto Or from < 0 Or tto < 0 Then
        CopyRow = False
        Exit Function
    End If
    Call InsertItem(lv, tto, lv.ListItems(from).Text)  ' LVGetName(lv, from))
    zvnb = lv.ColumnHeaders.Count
    For i = 1 To zvnb
        If from < tto Then
            zItemX.SubItems(i) = lv.ListItems(from).SubItems(i)
        Else
            zItemX.SubItems(i) = lv.ListItems(from + 1).SubItems(i)
        End If
    Next
    CopyRow = True
End Function

Public Function InsertItem(lv As ListView, tto As Long, zTexte As String) As Boolean
    On Error Resume Next
    Set zItemX = lv.ListItems.Add(tto, , zTexte)
    zItemX.Text = zTexte
    InsertItem = True
End Function
