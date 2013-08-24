VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Public lMax As Long
Public sAppName As String
Public sSection As String
Public sKey As String
Public Sub Items(zCombo As ComboBox)
    Dim i As Long
    Dim sCurItem As String
    Dim vnb As Integer
    vnb = 0
    For i = 1 To lMax
        sCurItem = GetSetting(sAppName, sSection, sKey & i)
        If sCurItem = Empty Then Exit For
        vnb = vnb + 1
        zCombo.AddItem sCurItem
    Next
End Sub
Public Sub AddNewItem(ByVal sNewItem As String)
    '// this sub adds a new file
    '// to the list
    Dim i As Integer
    Dim sCurItem As String
    If sNewItem = "" Then Exit Sub
    If OnList(sNewItem) Then
        MoveItemToTop sNewItem
    Else
        '// Copy RecentFile1 to RecentFile2, and so on.
        For i = lMax - 1 To 1 Step -1
            sCurItem = GetSetting(sAppName, sSection, sKey & i)
            If sCurItem <> Empty Then
                'Key = strRecent & (i + 1)
                SaveSetting sAppName, sSection, sKey & (i + 1), sCurItem ' sKey & i + 1
            End If
        Next
        '// Write the open file to first recent file.
        SaveSetting sAppName, sSection, sKey & "1", sNewItem
    End If
End Sub

Public Sub MoveItemToTop(ByVal sExistingItem As String)
    '// This sub moves the specified file
    '// to the top of the list
    '// from wherever it is
    Dim i As Long
    Dim sCurItem As String
    Dim lItemPos As Integer

    '// Get location of specified file
    For i = 1 To lMax
        sCurItem = GetSetting(sAppName, sSection, sKey & i)
        If sCurItem = sExistingItem Then
            '// Save the value
            lItemPos = i
            Exit For
        End If
    Next
    '// Copy RecentFile1 to RecentFile2, and so on.
    For i = lItemPos - 1 To 1 Step -1
        sCurItem = GetSetting(sAppName, sSection, sKey & i)
        If sCurItem <> Empty Then
            'Key = txtRecent & (i + 1)
            If Not i = lMax Then
                '// save it to the next one down
                SaveSetting sAppName, sSection, sKey & i + 1, sCurItem
            End If
        End If
    Next
    SaveSetting sAppName, sSection, sKey & "1", sExistingItem
End Sub

Public Function OnList(ByVal sItem As String) As Boolean
    Dim i As Long
    Dim sCurItem As String
    For i = 1 To lMax
        sCurItem = GetSetting(sAppName, sSection, sKey & i)
        If sCurItem = sItem Then
            '// Save the value
            OnList = True
            Exit For
        End If
    Next
End Function
Private Sub Class_Initialize()
    lMax = 10
    sAppName = "THERename"
    sSection = "History"
End Sub
