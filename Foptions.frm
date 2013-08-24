VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Foptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   5760
      Width           =   1815
   End
   Begin MSComctlLib.TreeView tv1 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   11668
      _Version        =   393217
      Indentation     =   499
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Foptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub LoadOptions()
   Dim nodX As Node  ' Declare Node variable.
   Set nodX = tv1.Nodes.Add(, , "General", "General")
   Set nodX = tv1.Nodes.Add("General", tvwChild, "GStartupDirectory", "Startup Directory")
   Set nodX = tv1.Nodes.Add("General", tvwChild, "GOther", "Other")
   Set nodX = tv1.Nodes.Add("General", tvwChild, "GBatchFile", "Batch File")
   Set nodX = tv1.Nodes.Add("General", tvwChild, "GLaunchProgram", "Launch a program")
   Set nodX = tv1.Nodes.Add("General", tvwChild, "GDateTimeFormat", "Date and Time format")
   Set nodX = tv1.Nodes.Add("General", tvwChild, "GFindLongFileNames", "Find long file names")
   Set nodX = tv1.Nodes.Add(, , "FilesFilters", "Files Filters")
   Set nodX = tv1.Nodes.Add(, , "Other", "Other")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OCharactersDelimiters", "Characters to delimit words")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OHtmlReport", "Directory & html report")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OCountersFormat", "Counters format")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OCountersRecursive", "Counters and recursive mode")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OFreeFom", "Free From")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OStartingSpaces", "Starting Spaces")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OSearchAbbrev", "Search & Replace and Abbreviations")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "ORules", "Rules")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OHistory", "History")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OInvalidCharacters", "Invalid characters")
   Set nodX = tv1.Nodes.Add("Other", tvwChild, "OSep", "Separator for lists")
   Set nodX = tv1.Nodes.Add(, , "Display", "Display")
   Set nodX = tv1.Nodes.Add("Display", tvwChild, "DWindPos", "Window's position")
   Set nodX = tv1.Nodes.Add("Display", tvwChild, "DStartOpt", "Startup options")
   Set nodX = tv1.Nodes.Add("Display", tvwChild, "DOther", "Other")
   Set nodX = tv1.Nodes.Add("Display", tvwChild, "DCurDir", "Current directory")
   Set nodX = tv1.Nodes.Add("Display", tvwChild, "DPreview", "Preview window")
   Set nodX = tv1.Nodes.Add(, , "Bin", "Bin")
   Set nodX = tv1.Nodes.Add(, , "Settings", "Settings")
   Set nodX = tv1.Nodes.Add("Settings", tvwChild, "SSaveSet", "Save settings")
   Set nodX = tv1.Nodes.Add("Settings", tvwChild, "SAutoSave", "Autosave")
   Set nodX = tv1.Nodes.Add("Settings", tvwChild, "SDefaultFiles", "Default Files")
   Set nodX = tv1.Nodes.Add(, , "Multimedia", "Multimedia")
   Set nodX = tv1.Nodes.Add("Multimedia", tvwChild, "MPictures", "Pictures width & height")
   Set nodX = tv1.Nodes.Add("Multimedia", tvwChild, "MMP3Tags", "Mp3 tags")
   Set nodX = tv1.Nodes.Add("Multimedia", tvwChild, "Mmp3VqfOggOpt", "Mp3, Vqf, Ogg options")
   Set nodX = tv1.Nodes.Add(, , "Lists", "Lists")
End Sub

Private Sub Form_Load()
    LoadOptions
End Sub
