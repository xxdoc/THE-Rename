Attribute VB_Name = "associat"
Option Explicit
Public Sub CreateAssociation()
   Dim sPath As String
   Exit Sub
   sPath = AppPath
   sPath = sPath & App.EXEName & ".exe"
   Dim cR As New cRegistry
   ' Create an open association, and set the file icon to be the icon with resource id 24 within
   ' the executable (note that resource id 1 is the exe's icon):
   cR.CreateEXEAssociation sPath, "THE Rename.Document", "THE Rename", "REN", lDefaultIconIndex:=1
End Sub
