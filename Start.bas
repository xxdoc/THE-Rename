Attribute VB_Name = "Start"
Sub Main()
On Error GoTo ErrGen
RENAME.Show 1
End

ErrGen:
MsgBox "Error, description : " & Err.Description & " Number : " & Err.Number & " Source : " & Err.Source & "ErrDll=" & Err.LastDllError
End Sub
