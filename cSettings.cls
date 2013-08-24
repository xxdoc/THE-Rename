VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarWordsDelimiters As String
Private mvarRxSyntax As Integer
Private mvarLogFile As String
Private mvarTagsVersionToUse As Integer
Private mvarMp3VqfOpt1 As Integer
Private mvarMp3VqfOpt2 As Integer
Private mvarMp3VqfOpt3 As Integer
Private mvarMp3VqfOpt4 As Integer
Private mvarTagsPriority As Integer
Private mvarPicturesPreview As Integer
Private mvarLastFilter As String
Private mvarBackRefNotation As Integer
Private mvarRegExpEngine As Integer
Private mvarRegExOpt1 As Integer
Private mvarChineseKorean As Integer
Private mvarRemoveEmptyTags As Integer
Private mvarPCRE1 As Integer
Private mvarPCRE2 As Integer
Private mvarPCRE3 As Integer
Private mvarPCRE4 As Integer
Private mvarPCRE5 As Integer
Private mvarOggOpt1 As Integer
Private mvarOggOpt2 As String
Private mvarUseDefaultCyclicFile As Integer
Private mvarDefaultCyclicFile As String
Private mvarUseDefaultAbbrevFile As Integer
Private mvarDefaultAbbrevFile As String
Private mvarRememberWSize As Integer
Private mvarShowMP3Tab As Integer
Private mvarShowMusicTab As Integer
Private mvarShowTextTab As Integer
Private mvarSelectAllFiles As Integer
Private mvarAskQuestion As Integer
Private mvarUseNaturalSort As Integer
Private mvarRulesOpt1 As Integer
Private mvarwWindowState As Integer
Private mvarTextToView As String
Private mvarDefaultCommand As String
Private mvarDefaultSearchPr As String
Private mvarDefaultReplacePr As String
Private mvarDefaultSearchExt As String
Private mvarDefaultReplaceExt As String
Private mvarRemoveIllegals As Integer
Private mvarRemoveStartingSpaces As Integer
Private mvarCharTokens As String
Private mvarIncludePictInfo As Integer
Private mvarCheckLongFileNameOption As Integer
Private mvarCheckLongFileNameSize As Integer
Private mvarToolbarButtons As Integer
Private mvarCompleCounters As Integer
Private mvarShowPathInCaption As Integer
Private mvarFilesToInclude As Integer
Private mvarHtmlIncFolder As Integer
Private mvarHtmlIncSize As Integer
Private mvarHtmlIncDate As Integer
Private mvarHtmlIncAttr As Integer
Private mvarHtmlIncMusic As Integer
Private mvarIncHiddenFolders As Integer
Private mvarPreviewGridLines As Integer
Private mvarIncVirtualFolders As Integer
Private mvarSettingsDirectory As String
Private mvarListDelimiter As Integer
Private mvarRemoveGuill As Integer
Private mvarRememberLastCommand As Integer
Private mvarLastCommand As String
Private mvarRememberLastFolder As Integer
Private mvarLastFolder As String
Private mvarIncludeLinks As Integer
Private mvarPicturesFormat As String
Private mvarShowWhenFileNameChange As Integer
Private mvarExifDelimD As String
Private mvarExifDelimT As String
Private mvarDefOption1 As Integer
Private mvarDefOption2 As Integer
Private mvarColumnsWiths As Boolean
Private mvarUseAutoSave As Integer
Private mvarCol1 As Long
Private mvarCol2 As Long
Private mvarCol3 As Long
Private mvarCol4 As Long
Private mvarCol5 As Long
Private mvarConfirmOperation As Boolean
Private mvarRenameOnCollision As Boolean
Private mvarSilentMode As Boolean
Private mvarAllowUndo As Boolean
Private mvarConfirmMakeDir As Boolean
Private mvarShutDown As Boolean
Private mvarFullRow As Boolean
Private mvarGridLines As Boolean
Private mvarPersonnalDate As String
Private mvarDisplayDate As String
Private mvarUseHistory As Boolean
Private mvarDirectoryReport As Integer
Private mvarDateformat As Integer
Private mvarSearchAndReplace As Integer
Private mvarStartupDir As String
Private mvarAutoArrange As Boolean
Private mvarRestartCounter As Integer
Private mvarLevelRestart As Integer
Private mvarUseLowerInLetterCounters As Integer
Private mvarprog1 As String
Private mvarprog2 As String
Private mvarprog3 As String
Private mvarFormatDate As Integer
Private mvarFormatTime As Integer
Private mvarActDblClick As Integer
Private mvarCenter0rSave As Integer
Private mvarCopyRename As Boolean
Private mvarFirstDateUse As String
Private mvarLastDirectory As String
Private mvarLastUseDate As String
Private mvarLastUseTime As String
Private mvarNumberOFiles As Long
Private mvarNumberOfRuns As Long
Private mvarUndoFile As String
Private mvarbatch As String
Private mvarHidden As Boolean
Private mvarlLeft As Long
Private mvarlTOp As Long
Private mvarwWidth As Long
Private mvarwHeight As Long
Private mvarReadOnly As Boolean
Private mvarSystem As Boolean
Private mvarMisc1 As Integer
Private mvarMisc2 As Integer
Private mvarMisc3 As Integer
Private mvarMisc4 As Integer
Private mvarMisc5 As Integer
Private mvarMisc6 As Integer
Private mvarMisc7 As Integer
Private mvarMisc8 As Integer
Private mvarMisc9 As Integer
Private mvarMisc10 As Integer
Private mvarMisc11 As String
Private mvarMisc12 As String
Private mvarMisc13 As String
Public Property Let wWindowState(ByVal vData As Integer)
    mvarwWindowState = vData
End Property
Public Property Get wWindowState() As Integer
    wWindowState = mvarwWindowState
End Property
Public Property Let RulesOpt1(ByVal vData As Integer)
    mvarRulesOpt1 = vData
End Property
Public Property Get RulesOpt1() As Integer
    RulesOpt1 = mvarRulesOpt1
End Property
Public Property Let UseNaturalSort(ByVal vData As Integer)
    mvarUseNaturalSort = vData
End Property
Public Property Get UseNaturalSort() As Integer
    UseNaturalSort = mvarUseNaturalSort
End Property
Public Property Let AskQuestion(ByVal vData As Integer)
    mvarAskQuestion = vData
End Property
Public Property Get AskQuestion() As Integer
    AskQuestion = mvarAskQuestion
End Property
Public Property Let SelectAllFiles(ByVal vData As Integer)
    mvarSelectAllFiles = vData
End Property
Public Property Get SelectAllFiles() As Integer
    SelectAllFiles = mvarSelectAllFiles
End Property
Public Property Let ShowTextTab(ByVal vData As Integer)
    mvarShowTextTab = vData
End Property
Public Property Get ShowTextTab() As Integer
    ShowTextTab = mvarShowTextTab
End Property
Public Property Let ShowMusicTab(ByVal vData As Integer)
    mvarShowMusicTab = vData
End Property
Public Property Get ShowMusicTab() As Integer
    ShowMusicTab = mvarShowMusicTab
End Property
Public Property Let ShowMP3Tab(ByVal vData As Integer)
    mvarShowMP3Tab = vData
End Property
Public Property Get ShowMP3Tab() As Integer
    ShowMP3Tab = mvarShowMP3Tab
End Property
Public Property Let RememberWSize(ByVal vData As Integer)
    mvarRememberWSize = vData
End Property
Public Property Get RememberWSize() As Integer
    RememberWSize = mvarRememberWSize
End Property
Public Property Let DefaultAbbrevFile(ByVal vData As String)
    mvarDefaultAbbrevFile = vData
End Property
Public Property Get DefaultAbbrevFile() As String
    DefaultAbbrevFile = mvarDefaultAbbrevFile
End Property
Public Property Let UseDefaultAbbrevFile(ByVal vData As Integer)
    mvarUseDefaultAbbrevFile = vData
End Property
Public Property Get UseDefaultAbbrevFile() As Integer
    UseDefaultAbbrevFile = mvarUseDefaultAbbrevFile
End Property
Public Property Let DefaultCyclicFile(ByVal vData As String)
    mvarDefaultCyclicFile = vData
End Property
Public Property Get DefaultCyclicFile() As String
    DefaultCyclicFile = mvarDefaultCyclicFile
End Property
Public Property Let UseDefaultCyclicFile(ByVal vData As Integer)
    mvarUseDefaultCyclicFile = vData
End Property
Public Property Get UseDefaultCyclicFile() As Integer
    UseDefaultCyclicFile = mvarUseDefaultCyclicFile
End Property
Public Property Let OggOpt2(ByVal vData As String)
    mvarOggOpt2 = vData
End Property
Public Property Get OggOpt2() As String
    OggOpt2 = mvarOggOpt2
End Property
Public Property Let OggOpt1(ByVal vData As Integer)
    mvarOggOpt1 = vData
End Property
Public Property Get OggOpt1() As Integer
    OggOpt1 = mvarOggOpt1
End Property
Public Property Let PCRE5(ByVal vData As Integer)
    mvarPCRE5 = vData
End Property
Public Property Get PCRE5() As Integer
    PCRE5 = mvarPCRE5
End Property
Public Property Let PCRE4(ByVal vData As Integer)
    mvarPCRE4 = vData
End Property
Public Property Get PCRE4() As Integer
    PCRE4 = mvarPCRE4
End Property
Public Property Let PCRE3(ByVal vData As Integer)
    mvarPCRE3 = vData
End Property
Public Property Get PCRE3() As Integer
    PCRE3 = mvarPCRE3
End Property
Public Property Let PCRE2(ByVal vData As Integer)
    mvarPCRE2 = vData
End Property
Public Property Get PCRE2() As Integer
    PCRE2 = mvarPCRE2
End Property
Public Property Let PCRE1(ByVal vData As Integer)
    mvarPCRE1 = vData
End Property
Public Property Get PCRE1() As Integer
    PCRE1 = mvarPCRE1
End Property
Public Property Let RemoveEmptyTags(ByVal vData As Integer)
    mvarRemoveEmptyTags = vData
End Property
Public Property Get RemoveEmptyTags() As Integer
    RemoveEmptyTags = mvarRemoveEmptyTags
End Property
Public Property Let ChineseKorean(ByVal vData As Integer)
    mvarChineseKorean = vData
End Property
Public Property Get ChineseKorean() As Integer
    ChineseKorean = mvarChineseKorean
End Property
Public Property Let RegExOpt1(ByVal vData As Integer)
    mvarRegExOpt1 = vData
End Property
Public Property Get RegExOpt1() As Integer
    RegExOpt1 = mvarRegExOpt1
End Property
Public Property Let RegExpEngine(ByVal vData As Integer)
    mvarRegExpEngine = vData
End Property
Public Property Get RegExpEngine() As Integer
    RegExpEngine = mvarRegExpEngine
End Property
Public Property Let BackRefNotation(ByVal vData As Integer)
    mvarBackRefNotation = vData
End Property
Public Property Get BackRefNotation() As Integer
    BackRefNotation = mvarBackRefNotation
End Property
Public Property Let LastFilter(ByVal vData As String)
    mvarLastFilter = vData
End Property
Public Property Get LastFilter() As String
    LastFilter = mvarLastFilter
End Property
Public Property Let PicturesPreview(ByVal vData As Integer)
    mvarPicturesPreview = vData
End Property
Public Property Get PicturesPreview() As Integer
    PicturesPreview = mvarPicturesPreview
End Property
Public Property Let TagsPriority(ByVal vData As Integer)
    mvarTagsPriority = vData
End Property
Public Property Get TagsPriority() As Integer
    TagsPriority = mvarTagsPriority
End Property
Public Property Let Mp3VqfOpt4(ByVal vData As Integer)
    mvarMp3VqfOpt4 = vData
End Property
Public Property Get Mp3VqfOpt4() As Integer
    Mp3VqfOpt4 = mvarMp3VqfOpt4
End Property
Public Property Let Mp3VqfOpt3(ByVal vData As Integer)
    mvarMp3VqfOpt3 = vData
End Property
Public Property Get Mp3VqfOpt3() As Integer
    Mp3VqfOpt3 = mvarMp3VqfOpt3
End Property
Public Property Let Mp3VqfOpt2(ByVal vData As Integer)
    mvarMp3VqfOpt2 = vData
End Property
Public Property Get Mp3VqfOpt2() As Integer
    Mp3VqfOpt2 = mvarMp3VqfOpt2
End Property
Public Property Let Mp3VqfOpt1(ByVal vData As Integer)
    mvarMp3VqfOpt1 = vData
End Property
Public Property Get Mp3VqfOpt1() As Integer
    Mp3VqfOpt1 = mvarMp3VqfOpt1
End Property
Public Property Let TagsVersionToUse(ByVal vData As Integer)
    mvarTagsVersionToUse = vData
End Property
Public Property Get TagsVersionToUse() As Integer
    TagsVersionToUse = mvarTagsVersionToUse
End Property
Public Property Let LogFile(ByVal vData As String)
    mvarLogFile = vData
End Property
Public Property Get LogFile() As String
    LogFile = mvarLogFile
End Property
Public Property Let RxSyntax(ByVal vData As Integer)
    mvarRxSyntax = vData
End Property
Public Property Get RxSyntax() As Integer
    RxSyntax = mvarRxSyntax
End Property
Public Property Let TextToView(ByVal vData As String)
    mvarTextToView = vData
End Property
Public Property Get TextToView() As String
    TextToView = mvarTextToView
End Property
Public Property Let DefaultCommand(ByVal vData As String)
    mvarDefaultCommand = vData
End Property
Public Property Get DefaultCommand() As String
    DefaultCommand = mvarDefaultCommand
End Property
Public Property Let DefaultSearchPr(ByVal vData As String)
    mvarDefaultSearchPr = vData
End Property
Public Property Get DefaultSearchPr() As String
    DefaultSearchPr = mvarDefaultSearchPr
End Property
Public Property Let DefaultReplacePr(ByVal vData As String)
    mvarDefaultReplacePr = vData
End Property
Public Property Get DefaultReplacePr() As String
    DefaultReplacePr = mvarDefaultReplacePr
End Property
Public Property Let DefaultSearchExt(ByVal vData As String)
    mvarDefaultSearchExt = vData
End Property
Public Property Get DefaultSearchExt() As String
    DefaultSearchExt = mvarDefaultSearchExt
End Property
Public Property Let DefaultReplaceExt(ByVal vData As String)
    mvarDefaultReplaceExt = vData
End Property
Public Property Get DefaultReplaceExt() As String
    DefaultReplaceExt = mvarDefaultReplaceExt
End Property
Public Property Let RemoveIllegals(ByVal vData As Integer)
    mvarRemoveIllegals = vData
End Property
Public Property Get RemoveIllegals() As Integer
    RemoveIllegals = mvarRemoveIllegals
End Property
Public Property Let RemoveStartingSpaces(ByVal vData As Integer)
    mvarRemoveStartingSpaces = vData
End Property
Public Property Get RemoveStartingSpaces() As Integer
    RemoveStartingSpaces = mvarRemoveStartingSpaces
End Property
Public Property Let CharTokens(ByVal vData As String)
    mvarCharTokens = vData
End Property
Public Property Get CharTokens() As String
    CharTokens = mvarCharTokens
End Property
Public Property Let IncludePictInfo(ByVal vData As Integer)
    mvarIncludePictInfo = vData
End Property
Public Property Get IncludePictInfo() As Integer
    IncludePictInfo = mvarIncludePictInfo
End Property
Public Property Let CheckLongFileNameOption(ByVal vData As Integer)
    mvarCheckLongFileNameOption = vData
End Property
Public Property Get CheckLongFileNameOption() As Integer
    CheckLongFileNameOption = mvarCheckLongFileNameOption
End Property
Public Property Let CheckLongFileNameSize(ByVal vData As Integer)
    mvarCheckLongFileNameSize = vData
End Property
Public Property Get CheckLongFileNameSize() As Integer
    CheckLongFileNameSize = mvarCheckLongFileNameSize
End Property
Public Property Let ToolbarButtons(ByVal vData As Integer)
    mvarToolbarButtons = vData
End Property
Public Property Get ToolbarButtons() As Integer
    ToolbarButtons = mvarToolbarButtons
End Property
Public Property Let CompleCounters(ByVal vData As Integer)
    mvarCompleCounters = vData
End Property
Public Property Get CompleCounters() As Integer
    CompleCounters = mvarCompleCounters
End Property
Public Property Let ShowPathInCaption(ByVal vData As Integer)
    mvarShowPathInCaption = vData
End Property
Public Property Get ShowPathInCaption() As Integer
    ShowPathInCaption = mvarShowPathInCaption
End Property
Public Property Let FilesToInclude(ByVal vData As Integer)
    mvarFilesToInclude = vData
End Property
Public Property Get FilesToInclude() As Integer
    FilesToInclude = mvarFilesToInclude
End Property
Public Property Let HtmlIncFolder(ByVal vData As Integer)
    mvarHtmlIncFolder = vData
End Property
Public Property Get HtmlIncFolder() As Integer
    HtmlIncFolder = mvarHtmlIncFolder
End Property
Public Property Let HtmlIncSize(ByVal vData As Integer)
    mvarHtmlIncSize = vData
End Property
Public Property Get HtmlIncSize() As Integer
    HtmlIncSize = mvarHtmlIncSize
End Property
Public Property Let HtmlIncDate(ByVal vData As Integer)
    mvarHtmlIncDate = vData
End Property
Public Property Get HtmlIncDate() As Integer
    HtmlIncDate = mvarHtmlIncDate
End Property
Public Property Let HtmlIncAttr(ByVal vData As Integer)
    mvarHtmlIncAttr = vData
End Property
Public Property Get HtmlIncAttr() As Integer
    HtmlIncAttr = mvarHtmlIncAttr
End Property
Public Property Let HtmlIncMusic(ByVal vData As Integer)
    mvarHtmlIncMusic = vData
End Property
Public Property Get HtmlIncMusic() As Integer
    HtmlIncMusic = mvarHtmlIncMusic
End Property
Public Property Let IncHiddenFolders(ByVal vData As Integer)
    mvarIncHiddenFolders = vData
End Property
Public Property Get IncHiddenFolders() As Integer
    IncHiddenFolders = mvarIncHiddenFolders
End Property
Public Property Let PreviewGridLines(ByVal vData As Integer)
    mvarPreviewGridLines = vData
End Property
Public Property Get PreviewGridLines() As Integer
    PreviewGridLines = mvarPreviewGridLines
End Property
Public Property Let IncVirtualFolders(ByVal vData As Integer)
    mvarIncVirtualFolders = vData
End Property
Public Property Get IncVirtualFolders() As Integer
    IncVirtualFolders = mvarIncVirtualFolders
End Property
Public Property Let WordsDelimiters(ByVal vData As String)
    mvarWordsDelimiters = vData
End Property
Public Property Get WordsDelimiters() As String
    WordsDelimiters = mvarWordsDelimiters
End Property
Public Property Let SettingsDirectory(ByVal vData As String)
    mvarSettingsDirectory = vData
End Property
Public Property Get SettingsDirectory() As String
    SettingsDirectory = mvarSettingsDirectory
End Property
Public Property Let ListDelimiter(ByVal vData As Integer)
    mvarListDelimiter = vData
End Property
Public Property Get ListDelimiter() As Integer
    ListDelimiter = mvarListDelimiter
End Property
Public Property Let RemoveGuill(ByVal vData As Integer)
    mvarRemoveGuill = vData
End Property
Public Property Get RemoveGuill() As Integer
    RemoveGuill = mvarRemoveGuill
End Property
Public Property Let RememberLastCommand(ByVal vData As Integer)
    mvarRememberLastCommand = vData
End Property
Public Property Get RememberLastCommand() As Integer
    RememberLastCommand = mvarRememberLastCommand
End Property
Public Property Let LastCommand(ByVal vData As String)
    mvarLastCommand = vData
End Property
Public Property Get LastCommand() As String
    LastCommand = mvarLastCommand
End Property
Public Property Let RememberLastFolder(ByVal vData As Integer)
    mvarRememberLastFolder = vData
End Property
Public Property Get RememberLastFolder() As Integer
    RememberLastFolder = mvarRememberLastFolder
End Property
Public Property Let LastFolder(ByVal vData As String)
    mvarLastFolder = vData
End Property
Public Property Get LastFolder() As String
    LastFolder = mvarLastFolder
End Property
Public Property Let IncludeLinks(ByVal vData As Integer)
    mvarIncludeLinks = vData
End Property
Public Property Get IncludeLinks() As Integer
    IncludeLinks = mvarIncludeLinks
End Property
Public Property Let PicturesFormat(ByVal vData As String)
    mvarPicturesFormat = vData
End Property
Public Property Get PicturesFormat() As String
    PicturesFormat = mvarPicturesFormat
End Property
Public Property Let ShowWhenFileNameChange(ByVal vData As Integer)
    mvarShowWhenFileNameChange = vData
End Property
Public Property Get ShowWhenFileNameChange() As Integer
    ShowWhenFileNameChange = mvarShowWhenFileNameChange
End Property
Public Property Let ExifDelimD(ByVal vData As String)
    mvarExifDelimD = vData
End Property
Public Property Get ExifDelimD() As String
    ExifDelimD = mvarExifDelimD
End Property
Public Property Let ExifDelimT(ByVal vData As String)
    mvarExifDelimT = vData
End Property
Public Property Get ExifDelimT() As String
    ExifDelimT = mvarExifDelimT
End Property
Public Property Let DefOption1(ByVal vData As Integer)
    mvarDefOption1 = vData
End Property
Public Property Get DefOption1() As Integer
    DefOption1 = mvarDefOption1
End Property
Public Property Let DefOption2(ByVal vData As Integer)
    mvarDefOption2 = vData
End Property
Public Property Get DefOption2() As Integer
    DefOption2 = mvarDefOption2
End Property
Public Property Let ColumnsWiths(ByVal vData As Boolean)
    mvarColumnsWiths = vData
End Property
Public Property Get ColumnsWiths() As Boolean
    ColumnsWiths = mvarColumnsWiths
End Property
Public Property Let UseAutoSave(ByVal vData As Integer)
    mvarUseAutoSave = vData
End Property
Public Property Get UseAutoSave() As Integer
    UseAutoSave = mvarUseAutoSave
End Property
Public Property Let WCol1(ByVal vData As Long)
    mvarCol1 = vData
End Property
Public Property Get WCol1() As Long
    WCol1 = mvarCol1
End Property
Public Property Let WCol2(ByVal vData As Long)
    mvarCol2 = vData
End Property
Public Property Get WCol2() As Long
    WCol2 = mvarCol2
End Property
Public Property Let WCol3(ByVal vData As Long)
    mvarCol3 = vData
End Property
Public Property Get WCol3() As Long
    WCol3 = mvarCol3
End Property
Public Property Let WCol4(ByVal vData As Long)
    mvarCol4 = vData
End Property
Public Property Get WCol4() As Long
    WCol4 = mvarCol4
End Property
Public Property Let WCol5(ByVal vData As Long)
    mvarCol5 = vData
End Property
Public Property Get WCol5() As Long
    WCol5 = mvarCol5
End Property
Public Property Let ConfirmOperation(ByVal vData As Boolean)
    mvarConfirmOperation = vData
End Property
Public Property Get ConfirmOperation() As Boolean
    ConfirmOperation = mvarConfirmOperation
End Property
Public Property Let RenameOnCollision(ByVal vData As Boolean)
    mvarRenameOnCollision = vData
End Property
Public Property Get RenameOnCollision() As Boolean
    RenameOnCollision = mvarRenameOnCollision
End Property
Public Property Let SilentMode(ByVal vData As Boolean)
    mvarSilentMode = vData
End Property
Public Property Get SilentMode() As Boolean
    SilentMode = mvarSilentMode
End Property
Public Property Let AllowUndo(ByVal vData As Boolean)
    mvarAllowUndo = vData
End Property
Public Property Get AllowUndo() As Boolean
    AllowUndo = mvarAllowUndo
End Property
Public Property Let ConfirmMakeDir(ByVal vData As Boolean)
    mvarConfirmMakeDir = vData
End Property
Public Property Get ConfirmMakeDir() As Boolean
    ConfirmMakeDir = mvarConfirmMakeDir
End Property
Public Property Let ShutDown(ByVal vData As Boolean)
    mvarShutDown = vData
End Property
Public Property Get ShutDown() As Boolean
    ShutDown = mvarShutDown
End Property
Public Property Let FullRow(ByVal vData As Boolean)
    mvarFullRow = vData
End Property
Public Property Get FullRow() As Boolean
    FullRow = mvarFullRow
End Property
Public Property Let GridLines(ByVal vData As Boolean)
    mvarGridLines = vData
End Property
Public Property Get GridLines() As Boolean
    GridLines = mvarGridLines
End Property
Public Property Let PersonnalDate(ByVal vData As String)
    mvarPersonnalDate = vData
End Property
Public Property Get PersonnalDate() As String
    PersonnalDate = mvarPersonnalDate
End Property
Public Property Let DisplayDate(ByVal vData As String)
    mvarDisplayDate = vData
End Property
Public Property Get DisplayDate() As String
    DisplayDate = mvarDisplayDate
End Property
Public Property Let UseHistory(ByVal vData As Boolean)
    mvarUseHistory = vData
End Property
Public Property Get UseHistory() As Boolean
    UseHistory = mvarUseHistory
End Property
Public Property Let DirectoryReport(ByVal vData As Integer)
    mvarDirectoryReport = vData
End Property
Public Property Get DirectoryReport() As Integer
    DirectoryReport = mvarDirectoryReport
End Property
Public Property Let Dateformat(ByVal vData As Integer)
    mvarDateformat = vData
End Property
Public Property Get Dateformat() As Integer
    Dateformat = mvarDateformat
End Property
Public Property Let SearchAndReplace(ByVal vData As Integer)
    mvarSearchAndReplace = vData
End Property
Public Property Get SearchAndReplace() As Integer
    SearchAndReplace = mvarSearchAndReplace
End Property
Public Property Let StartupDir(ByVal vData As String)
    mvarStartupDir = vData
End Property
Public Property Get StartupDir() As String
    StartupDir = mvarStartupDir
End Property
Public Property Let AutoArrange(ByVal vData As Boolean)
    mvarAutoArrange = vData
End Property
Public Property Get AutoArrange() As Boolean
    AutoArrange = mvarAutoArrange
End Property
Public Property Let RestartCounter(ByVal vData As Integer)
    mvarRestartCounter = vData
End Property
Public Property Get RestartCounter() As Integer
    RestartCounter = mvarRestartCounter
End Property
Public Property Let LevelRestart(ByVal vData As Integer)
    mvarLevelRestart = vData
End Property
Public Property Get LevelRestart() As Integer
    LevelRestart = mvarLevelRestart
End Property
Public Property Let UseLowerInLetterCounters(ByVal vData As Integer)
    mvarUseLowerInLetterCounters = vData
End Property
Public Property Get UseLowerInLetterCounters() As Integer
    UseLowerInLetterCounters = mvarUseLowerInLetterCounters
End Property
Public Property Let prog1(ByVal vData As String)
    mvarprog1 = vData
End Property
Public Property Get prog1() As String
    prog1 = mvarprog1
End Property
Public Property Let prog2(ByVal vData As String)
    mvarprog2 = vData
End Property
Public Property Get prog2() As String
    prog2 = mvarprog2
End Property
Public Property Let prog3(ByVal vData As String)
    mvarprog3 = vData
End Property
Public Property Get prog3() As String
    prog3 = mvarprog3
End Property
Public Property Let FormatDate(ByVal vData As Integer)
    mvarFormatDate = vData
End Property
Public Property Get FormatDate() As Integer
    FormatDate = mvarFormatDate
End Property
Public Property Let FormatTime(ByVal vData As Integer)
    mvarFormatTime = vData
End Property
Public Property Get FormatTime() As Integer
    FormatTime = mvarFormatTime
End Property
Public Property Let ActDblClick(ByVal vData As Integer)
    mvarActDblClick = vData
End Property
Public Property Get ActDblClick() As Integer
    ActDblClick = mvarActDblClick
End Property
Public Property Let Center0rSave(ByVal vData As Integer)
    mvarCenter0rSave = vData
End Property
Public Property Get Center0rSave() As Integer
    Center0rSave = mvarCenter0rSave
End Property
Public Property Let CopyRename(ByVal vData As Boolean)
    mvarCopyRename = vData
End Property
Public Property Get CopyRename() As Boolean
    CopyRename = mvarCopyRename
End Property
Public Property Let FirstDateUse(ByVal vData As String)
    mvarFirstDateUse = vData
End Property
Public Property Get FirstDateUse() As String
    FirstDateUse = mvarFirstDateUse
End Property
Public Property Let LastDirectory(ByVal vData As String)
    mvarLastDirectory = vData
End Property
Public Property Get LastDirectory() As String
    LastDirectory = mvarLastDirectory
End Property
Public Property Let LastUseDate(ByVal vData As String)
    mvarLastUseDate = vData
End Property
Public Property Get LastUseDate() As String
    LastUseDate = mvarLastUseDate
End Property
Public Property Let LastUseTime(ByVal vData As String)
    mvarLastUseTime = vData
End Property
Public Property Get LastUseTime() As String
    LastUseTime = mvarLastUseTime
End Property
Public Property Let NumberOFiles(ByVal vData As Long)
    mvarNumberOFiles = vData
End Property
Public Property Get NumberOFiles() As Long
    NumberOFiles = mvarNumberOFiles
End Property
Public Property Let NumberOfRuns(ByVal vData As Long)
    mvarNumberOfRuns = vData
End Property
Public Property Get NumberOfRuns() As Long
    NumberOfRuns = mvarNumberOfRuns
End Property
Public Property Let UndoFile(ByVal vData As String)
    mvarUndoFile = vData
End Property
Public Property Get UndoFile() As String
    UndoFile = mvarUndoFile
End Property
Public Property Let batch(ByVal vData As String)
    mvarbatch = vData
End Property
Public Property Get batch() As String
    batch = mvarbatch
End Property
Public Property Let Hidden(ByVal vData As Boolean)
    mvarHidden = vData
End Property
Public Property Get Hidden() As Boolean
    Hidden = mvarHidden
End Property
Public Property Let lLeft(ByVal vData As Long)
    mvarlLeft = vData
End Property
Public Property Get lLeft() As Long
    lLeft = mvarlLeft
End Property
Public Property Let lTOp(ByVal vData As Long)
    mvarlTOp = vData
End Property
Public Property Get lTOp() As Long
    lTOp = mvarlTOp
End Property
Public Property Let wWidth(ByVal vData As Long)
    mvarwWidth = vData
End Property
Public Property Get wWidth() As Long
    wWidth = mvarwWidth
End Property
Public Property Let wHeight(ByVal vData As Long)
    mvarwHeight = vData
End Property
Public Property Get wHeight() As Long
    wHeight = mvarwHeight
End Property
Public Property Let ReadOnly(ByVal vData As Boolean)
    mvarReadOnly = vData
End Property
Public Property Get ReadOnly() As Boolean
    ReadOnly = mvarReadOnly
End Property
Public Property Let System(ByVal vData As Boolean)
    mvarSystem = vData
End Property
Public Property Get System() As Boolean
    System = mvarSystem
End Property
Public Property Let Misc1(ByVal vData As Integer)
    mvarMisc1 = vData
End Property
Public Property Get Misc1() As Integer
    Misc1 = mvarMisc1
End Property
Public Property Let Misc2(ByVal vData As Integer)
    mvarMisc2 = vData
End Property
Public Property Get Misc2() As Integer
    Misc2 = mvarMisc2
End Property
Public Property Let Misc3(ByVal vData As Integer)
    mvarMisc3 = vData
End Property
Public Property Get Misc3() As Integer
    Misc3 = mvarMisc3
End Property
Public Property Let Misc4(ByVal vData As Integer)
    mvarMisc4 = vData
End Property
Public Property Get Misc4() As Integer
    Misc4 = mvarMisc4
End Property
Public Property Let Misc5(ByVal vData As Integer)
    mvarMisc5 = vData
End Property
Public Property Get Misc5() As Integer
    Misc5 = mvarMisc5
End Property
Public Property Let Misc6(ByVal vData As Integer)
    mvarMisc6 = vData
End Property
Public Property Get Misc6() As Integer
    Misc6 = mvarMisc6
End Property
Public Property Let Misc7(ByVal vData As Integer)
    mvarMisc7 = vData
End Property
Public Property Get Misc7() As Integer
    Misc7 = mvarMisc7
End Property
Public Property Let Misc8(ByVal vData As Integer)
    mvarMisc8 = vData
End Property
Public Property Get Misc8() As Integer
    Misc8 = mvarMisc8
End Property
Public Property Let Misc9(ByVal vData As Integer)
    mvarMisc9 = vData
End Property
Public Property Get Misc9() As Integer
    Misc9 = mvarMisc9
End Property
Public Property Let Misc10(ByVal vData As Integer)
    mvarMisc10 = vData
End Property
Public Property Get Misc10() As Integer
    Misc10 = mvarMisc10
End Property
Public Property Let Misc11(ByVal vData As String)
    mvarMisc11 = vData
End Property
Public Property Get Misc11() As String
    Misc11 = mvarMisc11
End Property
Public Property Let Misc12(ByVal vData As String)
    mvarMisc12 = vData
End Property
Public Property Get Misc12() As String
    Misc12 = mvarMisc12
End Property
Public Property Let Misc13(ByVal vData As String)
    mvarMisc13 = vData
End Property
Public Property Get Misc13() As String
    Misc13 = mvarMisc13
End Property
Public Function SaveSettings() As Boolean
    SaveSetting "THERename", "Param", "system", CInt(mvarSystem)
    SaveSetting "THERename", "Param", "readonly", CInt(mvarReadOnly)
    If RENAME.WindowState <> vbMaximized Then
        If wWidth > 0 Then
            SaveSetting "THERename", "Param", "wWidth", mvarwWidth
        End If
        If wHeight > 0 Then
            SaveSetting "THERename", "Param", "wHeight", mvarwHeight
        End If
    Else
        SaveSetting "THERename", "Param", "wWidth", -1
        SaveSetting "THERename", "Param", "wHeight", -1
    End If
    If lLeft > 0 Then
        SaveSetting "THERename", "Param", "left", mvarlLeft
    End If
    If lTOp > 0 Then
        SaveSetting "THERename", "Param", "top", mvarlTOp
    End If
    SaveSetting "THERename", "Param", "hidden", CInt(mvarHidden)
    SaveSetting "THERename", "Param", "batch", mvarbatch
    SaveSetting "THERename", "Param", "UndoFile", mvarUndoFile
    SaveSetting "THERename", "Param", "NumberOfRuns", mvarNumberOfRuns
    SaveSetting "THERename", "Param", "NumberOFiles", mvarNumberOFiles
    SaveSetting "THERename", "Param", "LastUseTime", mvarLastUseTime
    SaveSetting "THERename", "Param", "LastUseDate", mvarLastUseDate
    SaveSetting "THERename", "Param", "LastDirectory", mvarLastDirectory
    SaveSetting "THERename", "Param", "FirstDateUse", mvarFirstDateUse
    SaveSetting "THERename", "Param", "CopyRename", CInt(mvarCopyRename)
    SaveSetting "THERename", "Param", "CenterOrSave", mvarCenter0rSave
    SaveSetting "THERename", "Param", "ActDblClick", mvarActDblClick
    SaveSetting "THERename", "Param", "FormatTime", mvarFormatTime
    SaveSetting "THERename", "Param", "FormatDate", mvarFormatDate
    SaveSetting "THERename", "Param", "Prog1", mvarprog1
    SaveSetting "THERename", "Param", "Prog2", mvarprog2
    SaveSetting "THERename", "Param", "Prog3", mvarprog3
    SaveSetting "THERename", "Param", "UseLowerInLetterCounters", mvarUseLowerInLetterCounters
    SaveSetting "THERename", "Param", "LevelRestart", mvarLevelRestart
    SaveSetting "THERename", "Param", "RestartCounter", mvarRestartCounter
    SaveSetting "THERename", "Param", "AutoArrange", CInt(mvarAutoArrange)
    SaveSetting "THERename", "Param", "StartupDir", mvarStartupDir
    SaveSetting "THERename", "Param", "SearchAndReplace", mvarSearchAndReplace
    SaveSetting "THERename", "Param", "DateFormat", mvarDateformat
    SaveSetting "THERename", "Param", "DirectoryReport", mvarDirectoryReport
    SaveSetting "THERename", "Param", "UseHistory", CInt(mvarUseHistory)
    SaveSetting "THERename", "Param", "DisplayDate", mvarDisplayDate
    SaveSetting "THERename", "Param", "PersonnalDate", mvarPersonnalDate
    SaveSetting "THERename", "Param", "GridLines", CInt(mvarGridLines)
    SaveSetting "THERename", "Param", "FullRow", CInt(mvarFullRow)
    SaveSetting "THERename", "Param", "ShutDown", CInt(mvarShutDown)
    SaveSetting "THERename", "Param", "ConfirmMakeDir", CInt(mvarConfirmMakeDir)
    SaveSetting "THERename", "Param", "AllowUndo", CInt(mvarAllowUndo)
    SaveSetting "THERename", "Param", "SilentMode", CInt(mvarSilentMode)
    SaveSetting "THERename", "Param", "RenameOnCollision", CInt(mvarRenameOnCollision)
    SaveSetting "THERename", "Param", "ConfirmOperation", CInt(mvarConfirmOperation)
    SaveSetting "THERename", "Param", "WCol1", RENAME.ListView1.ColumnHeaders(1).width
    SaveSetting "THERename", "Param", "WCol2", RENAME.ListView1.ColumnHeaders(2).width
    SaveSetting "THERename", "Param", "WCol3", RENAME.ListView1.ColumnHeaders(3).width
    SaveSetting "THERename", "Param", "WCol4", RENAME.ListView1.ColumnHeaders(4).width
    SaveSetting "THERename", "Param", "WCol5", RENAME.ListView1.ColumnHeaders(5).width
    SaveSetting "THERename", "Param", "UseAutoSave", mvarUseAutoSave
    SaveSetting "THERename", "Param", "ColumnsWiths", CInt(mvarColumnsWiths)
    SaveSetting "THERename", "Param", "DefOption1", mvarDefOption1
    SaveSetting "THERename", "Param", "DefOption2", mvarDefOption2
    SaveSetting "THERename", "Param", "ExifDelimT", mvarExifDelimT
    SaveSetting "THERename", "Param", "ExifDelimD", mvarExifDelimD
    SaveSetting "THERename", "Param", "ShowWhenFileNameChange", mvarShowWhenFileNameChange
    SaveSetting "THERename", "Param", "PicturesFormat", mvarPicturesFormat
    SaveSetting "THERename", "Param", "IncludeLinks", mvarIncludeLinks
    SaveSetting "THERename", "Param", "LastFolder", mvarLastFolder
    SaveSetting "THERename", "Param", "RememberLastFolder", mvarRememberLastFolder
    SaveSetting "THERename", "Param", "LastCommand", mvarLastCommand
    SaveSetting "THERename", "Param", "RememberLastCommand", mvarRememberLastCommand
    SaveSetting "THERename", "Param", "RemoveGuill", mvarRemoveGuill
    SaveSetting "THERename", "Param", "ListDelimiter", mvarListDelimiter
    SaveSetting "THERename", "Param", "SettingsDirectory", mvarSettingsDirectory
    SaveSetting "THERename", "Param", "IncVirtualFolders", mvarIncVirtualFolders
    SaveSetting "THERename", "Param", "PreviewGridLines", mvarPreviewGridLines
    SaveSetting "THERename", "Param", "IncHiddenFolders", mvarIncHiddenFolders
    SaveSetting "THERename", "Param", "HtmlIncMusic", mvarHtmlIncMusic
    SaveSetting "THERename", "Param", "HtmlIncAttr", mvarHtmlIncAttr
    SaveSetting "THERename", "Param", "HtmlIncDate", mvarHtmlIncDate
    SaveSetting "THERename", "Param", "HtmlIncSize", mvarHtmlIncSize
    SaveSetting "THERename", "Param", "HtmlIncFolder", mvarHtmlIncFolder
    SaveSetting "THERename", "Param", "FilesToInclude", mvarFilesToInclude
    SaveSetting "THERename", "Param", "ShowPathInCaption", mvarShowPathInCaption
    SaveSetting "THERename", "Param", "CompleCounters", mvarCompleCounters
    SaveSetting "THERename", "Param", "ToolbarButtons", mvarToolbarButtons
    SaveSetting "THERename", "Param", "CheckLongFileNameSize", mvarCheckLongFileNameSize
    SaveSetting "THERename", "Param", "CheckLongFileNameOption", mvarCheckLongFileNameOption
    SaveSetting "THERename", "Param", "IncludePictInfo", mvarIncludePictInfo
    SaveSetting "THERename", "Param", "CharTokens", mvarCharTokens
    SaveSetting "THERename", "Param", "RemoveStartingSpaces", mvarRemoveStartingSpaces
    SaveSetting "THERename", "Param", "RemoveIllegals", mvarRemoveIllegals
    SaveSetting "THERename", "Param", "DefaultReplaceExt", mvarDefaultReplaceExt
    SaveSetting "THERename", "Param", "DefaultSearchExt", mvarDefaultSearchExt
    SaveSetting "THERename", "Param", "DefaultReplacePr", mvarDefaultReplacePr
    SaveSetting "THERename", "Param", "DefaultSearchPr", mvarDefaultSearchPr
    SaveSetting "THERename", "Param", "DefaultCommand", mvarDefaultCommand
    SaveSetting "THERename", "Param", "TextToView", UCase$(mvarTextToView)
    SaveSetting "THERename", "Param", "WindowState", mvarwWindowState
    SaveSetting "THERename", "Param", "RulesOpt1", mvarRulesOpt1
    SaveSetting "THERename", "Param", "UseNaturalSort", mvarUseNaturalSort
    SaveSetting "THERename", "Param", "SelectAllFiles", mvarSelectAllFiles
    SaveSetting "THERename", "Param", "ShowTextTab", mvarShowTextTab
    SaveSetting "THERename", "Param", "ShowMusicTab", mvarShowMusicTab
    SaveSetting "THERename", "Param", "ShowMP3Tab", mvarShowMP3Tab
    SaveSetting "THERename", "Param", "RememberWSize", mvarRememberWSize
    SaveSetting "THERename", "Param", "DefaultAbbrevFile", mvarDefaultAbbrevFile
    SaveSetting "THERename", "Param", "UseDefaultAbbrevFile", mvarUseDefaultAbbrevFile
    SaveSetting "THERename", "Param", "DefaultCyclicFile", mvarDefaultCyclicFile
    SaveSetting "THERename", "Param", "UseDefaultCyclicFile", mvarUseDefaultCyclicFile
    SaveSetting "THERename", "Param", "OggOpt1", mvarOggOpt1
    SaveSetting "THERename", "Param", "OggOpt2", mvarOggOpt2
    SaveSetting "THERename", "Param", "PCRE2", mvarPCRE2
    SaveSetting "THERename", "Param", "PCRE3", mvarPCRE3
    SaveSetting "THERename", "Param", "PCRE4", mvarPCRE4
    SaveSetting "THERename", "Param", "PCRE5", mvarPCRE5
    SaveSetting "THERename", "Param", "PCRE1", mvarPCRE1
    SaveSetting "THERename", "Param", "RemoveEmptyTags", mvarRemoveEmptyTags
    SaveSetting "THERename", "Param", "ChineseKorean", mvarChineseKorean
    SaveSetting "THERename", "Param", "RegExOpt1", mvarRegExOpt1
    SaveSetting "THERename", "Param", "RegExpEngine", mvarRegExpEngine
    SaveSetting "THERename", "Param", "BackRefNotation", mvarBackRefNotation
    SaveSetting "THERename", "Param", "LastFilter", mvarLastFilter
    SaveSetting "THERename", "Param", "PicturesPreview", mvarPicturesPreview
    SaveSetting "THERename", "Param", "TagsPriority", mvarTagsPriority
    SaveSetting "THERename", "Param", "WordsDelimiters", mvarWordsDelimiters
    SaveSetting "THERename", "Param", "RxSyntax", mvarRxSyntax
    SaveSetting "THERename", "Param", "LogFile", mvarLogFile
    SaveSetting "THERename", "Param", "TagsVersionToUse", mvarTagsVersionToUse
    SaveSetting "THERename", "Param", "Mp3VqfOpt1", mvarMp3VqfOpt1
    SaveSetting "THERename", "Param", "Mp3VqfOpt2", mvarMp3VqfOpt2
    SaveSetting "THERename", "Param", "Mp3VqfOpt3", mvarMp3VqfOpt3
    SaveSetting "THERename", "Param", "Mp3VqfOpt4", mvarMp3VqfOpt4
    SaveSetting "THERename", "Param", "Misc1", mvarMisc1
    SaveSetting "THERename", "Param", "Misc2", mvarMisc2
    SaveSetting "THERename", "Param", "Misc3", mvarMisc3
    SaveSetting "THERename", "Param", "Misc4", mvarMisc4
    SaveSetting "THERename", "Param", "Misc5", mvarMisc5
    SaveSetting "THERename", "Param", "Misc6", mvarMisc6
    SaveSetting "THERename", "Param", "Misc7", mvarMisc7
    SaveSetting "THERename", "Param", "Misc8", mvarMisc8
    SaveSetting "THERename", "Param", "Misc9", mvarMisc9
    SaveSetting "THERename", "Param", "Misc10", mvarMisc10
    SaveSetting "THERename", "Param", "Misc11", mvarMisc11
    SaveSetting "THERename", "Param", "Misc12", mvarMisc12
    SaveSetting "THERename", "Param", "Misc13", mvarMisc13
End Function
Public Function LoadSettings() As Boolean
 On Error Resume Next
 Dim vret As Boolean
 mvarLastUseDate = GetSetting("THERename", "Param", "LastUseDate", "")
 If mvarLastUseDate = "" Then   ' C'est la première utilisation du logiciel ***********************************************
    InitDefValues
    vret = SaveSettings()
 Else   ' Ce n'est pas la première utilisation ************************************************************************
    InitDefValues
    mvarRxSyntax = GetSetting("THERename", "Param", "RxSyntax", 1)
    mvarLogFile = GetSetting("THERename", "Param", "LogFile", "")
    mvarTagsVersionToUse = GetSetting("THERename", "Param", "TagsVersionToUse", 1)
    mvarMp3VqfOpt1 = GetSetting("THERename", "Param", "Mp3VqfOpt1", 0)
    mvarMp3VqfOpt2 = GetSetting("THERename", "Param", "Mp3VqfOpt2", 0)
    mvarMp3VqfOpt3 = GetSetting("THERename", "Param", "Mp3VqfOpt3", 0)
    mvarMp3VqfOpt4 = GetSetting("THERename", "Param", "Mp3VqfOpt4", 0)
    mvarTagsPriority = GetSetting("THERename", "Param", "TagsPriority", 1)
    mvarPicturesPreview = GetSetting("THERename", "Param", "PicturesPreview", 0)
    mvarLastFilter = GetSetting("THERename", "Param", "LastFilter", "*.*")
    mvarBackRefNotation = GetSetting("THERename", "Param", "BackRefNotation", 0)
    mvarRegExpEngine = GetSetting("THERename", "Param", "RegExpEngine", 1)
    mvarRegExOpt1 = GetSetting("THERename", "Param", "RegExOpt1", 1)
    mvarChineseKorean = GetSetting("THERename", "Param", "ChineseKorean", 0)
    mvarRemoveEmptyTags = GetSetting("THERename", "Param", "RemoveEmptyTags", 1)
    mvarPCRE1 = GetSetting("THERename", "Param", "PCRE1", 0)
    mvarPCRE2 = GetSetting("THERename", "Param", "PCRE2", 0)
    mvarPCRE3 = GetSetting("THERename", "Param", "PCRE3", 0)
    mvarPCRE4 = GetSetting("THERename", "Param", "PCRE4", 0)
    mvarPCRE5 = GetSetting("THERename", "Param", "PCRE5", 0)
    mvarOggOpt1 = GetSetting("THERename", "Param", "OggOpt1", 0)
    mvarOggOpt2 = GetSetting("THERename", "Param", "OggOpt2", " - ")
    mvarUseDefaultCyclicFile = GetSetting("THERename", "Param", "UseDefaultCyclicFile", mvarUseDefaultCyclicFile)
    mvarDefaultCyclicFile = GetSetting("THERename", "Param", "DefaultCyclicFile", mvarDefaultCyclicFile)
    mvarUseDefaultAbbrevFile = GetSetting("THERename", "Param", "UseDefaultAbbrevFile", mvarUseDefaultAbbrevFile)
    mvarDefaultAbbrevFile = GetSetting("THERename", "Param", "DefaultAbbrevFile", mvarDefaultAbbrevFile)
    mvarRememberWSize = GetSetting("THERename", "Param", "RememberWSize", mvarRememberWSize)
    mvarShowMP3Tab = GetSetting("THERename", "Param", "ShowMP3Tab", mvarShowMP3Tab)
    mvarShowMusicTab = GetSetting("THERename", "Param", "ShowMusicTab", mvarShowMusicTab)
    mvarShowTextTab = GetSetting("THERename", "Param", "ShowTextTab", mvarShowTextTab)
    mvarSelectAllFiles = GetSetting("THERename", "Param", "SelectAllFiles", 0)
    mvarAskQuestion = GetSetting("THERename", "Param", "AskQuestion", 0)
    mvarUseNaturalSort = GetSetting("THERename", "Param", "UseNaturalSort", mvarUseNaturalSort)
    mvarRulesOpt1 = GetSetting("THERename", "Param", "RulesOpt1", mvarRulesOpt1)
    mvarwWindowState = GetSetting("THERename", "Param", "WindowState", mvarwWindowState)
    mvarTextToView = UCase$(GetSetting("THERename", "Param", "TextToView", "TXT|DIZ|INI|NFO|LOG"))
    mvarDefaultCommand = GetSetting("THERename", "Param", "DefaultCommand", "")
    mvarDefaultSearchPr = GetSetting("THERename", "Param", "DefaultSearchPr", "")
    mvarDefaultReplacePr = GetSetting("THERename", "Param", "DefaultReplacePr", "")
    mvarDefaultSearchExt = GetSetting("THERename", "Param", "DefaultSearchExt", "")
    mvarDefaultReplaceExt = GetSetting("THERename", "Param", "DefaultReplaceExt", "")
    mvarRemoveIllegals = GetSetting("THERename", "Param", "RemoveIllegals", mvarRemoveIllegals)
    mvarRemoveStartingSpaces = GetSetting("THERename", "Param", "RemoveStartingSpaces", mvarRemoveStartingSpaces)
    mvarCharTokens = GetSetting("THERename", "Param", "CharTokens", " ;")
    mvarIncludePictInfo = GetSetting("THERename", "Param", "IncludePictInfo", 0)
    mvarCheckLongFileNameOption = GetSetting("THERename", "Param", "CheckLongFileNameOption", 0)
    mvarCheckLongFileNameSize = GetSetting("THERename", "Param", "CheckLongFileNameSize", 64)
    mvarToolbarButtons = GetSetting("THERename", "Param", "ToolbarButtons", 1)
    mvarCompleCounters = GetSetting("THERename", "Param", "CompleCounters", 1)
    mvarShowPathInCaption = GetSetting("THERename", "Param", "ShowPathInCaption", mvarShowPathInCaption)
    mvarFilesToInclude = GetSetting("THERename", "Param", "FilesToInclude", mvarFilesToInclude)
    mvarHtmlIncFolder = GetSetting("THERename", "Param", "HtmlIncFolder", 1)
    mvarHtmlIncSize = GetSetting("THERename", "Param", "HtmlIncSize", 1)
    mvarHtmlIncDate = GetSetting("THERename", "Param", "HtmlIncDate", 1)
    mvarHtmlIncAttr = GetSetting("THERename", "Param", "HtmlIncAttr", 1)
    mvarHtmlIncMusic = GetSetting("THERename", "Param", "HtmlIncMusic", 0)
    mvarIncHiddenFolders = GetSetting("THERename", "Param", "IncHiddenFolders", mvarIncHiddenFolders)
    mvarPreviewGridLines = GetSetting("THERename", "Param", "PreviewGridLines", mvarPreviewGridLines)
    mvarIncVirtualFolders = GetSetting("THERename", "Param", "IncVirtualFolders", mvarIncVirtualFolders)
    mvarWordsDelimiters = GetSetting("THERename", "Param", "WordsDelimiters", " ")
    mvarSettingsDirectory = GetSetting("THERename", "Param", "SettingsDirectory", "")
    mvarListDelimiter = GetSetting("THERename", "Param", "ListDelimiter", 9)
    mvarRemoveGuill = GetSetting("THERename", "Param", "RemoveGuill", 0)
    mvarRememberLastCommand = GetSetting("THERename", "Param", "RememberLastCommand", 0)
    mvarLastCommand = GetSetting("THERename", "Param", "LastCommand", "")
    mvarRememberLastFolder = GetSetting("THERename", "Param", "RememberLastFolder", 0)
    mvarLastFolder = GetSetting("THERename", "Param", "LastFolder", "")
    mvarIncludeLinks = GetSetting("THERename", "Param", "IncludeLinks", 0)
    mvarPicturesFormat = GetSetting("THERename", "Param", "PicturesFormat", "%w%x%h%")
    mvarShowWhenFileNameChange = GetSetting("THERename", "Param", "ShowWhenFileNameChange", 0)
    mvarExifDelimD = GetSetting("THERename", "Param", "ExifDelimD", "-")
    mvarExifDelimT = GetSetting("THERename", "Param", "ExifDelimT", "-")
    mvarDefOption1 = GetSetting("THERename", "Param", "DefOption1", 7)
    mvarDefOption2 = GetSetting("THERename", "Param", "DefOption2", 6)
    mvarColumnsWiths = CBool(GetSetting("THERename", "Param", "ColumnsWiths", mvarColumnsWiths))
    mvarUseAutoSave = GetSetting("THERename", "Param", "UseAutoSave", mvarUseAutoSave)
    mvarCol1 = GetSetting("THERename", "Param", "WCol1", mvarCol1)
    mvarCol2 = GetSetting("THERename", "Param", "WCol2", mvarCol2)
    mvarCol3 = GetSetting("THERename", "Param", "WCol3", mvarCol3)
    mvarCol4 = GetSetting("THERename", "Param", "WCol4", mvarCol4)
    mvarCol5 = GetSetting("THERename", "Param", "WCol5", mvarCol5)
    mvarConfirmOperation = CBool(GetSetting("THERename", "Param", "ConfirmOperation", mvarConfirmOperation))
    mvarRenameOnCollision = CBool(GetSetting("THERename", "Param", "RenameOnCollision", mvarRenameOnCollision))
    mvarSilentMode = CBool(GetSetting("THERename", "Param", "SilentMode", mvarSilentMode))
    mvarAllowUndo = CBool(GetSetting("THERename", "Param", "AllowUndo", mvarAllowUndo))
    mvarConfirmMakeDir = CBool(GetSetting("THERename", "Param", "ConfirmMakeDir", mvarConfirmMakeDir))
    mvarShutDown = CBool(GetSetting("THERename", "Param", "ShutDown", mvarShutDown))
    mvarFullRow = CBool(GetSetting("THERename", "Param", "FullRow", mvarFullRow))
    mvarGridLines = CBool(GetSetting("THERename", "Param", "GridLines", mvarGridLines))
    mvarPersonnalDate = GetSetting("THERename", "Param", "PersonnalDate", mvarPersonnalDate)
    mvarUseHistory = CBool(GetSetting("THERename", "Param", "UseHistory", mvarUseHistory))
    mvarDirectoryReport = GetSetting("THERename", "Param", "DirectoryReport", 1)
    mvarDateformat = GetSetting("THERename", "Param", "DateFormat", 0)
    mvarSearchAndReplace = GetSetting("THERename", "Param", "SearchAndReplace", 0)
    mvarStartupDir = GetSetting("THERename", "Param", "StartupDir", mvarStartupDir)
    mvarAutoArrange = CBool(GetSetting("THERename", "Param", "AutoArrange", True))
    mvarRestartCounter = GetSetting("THERename", "Param", "RestartCounter", 0)
    mvarLevelRestart = GetSetting("THERename", "Param", "LevelRestart", 0)
    mvarUseLowerInLetterCounters = GetSetting("THERename", "Param", "UseLowerInLetterCounters", mvarUseLowerInLetterCounters)
    mvarprog1 = GetSetting("THERename", "Param", "Prog1", "")
    mvarprog2 = GetSetting("THERename", "Param", "Prog2", "")
    mvarprog3 = GetSetting("THERename", "Param", "Prog3", "")
    mvarFormatDate = GetSetting("THERename", "Param", "FormatDate", 2)
    mvarFormatTime = GetSetting("THERename", "Param", "FormatTime", 0)
    mvarActDblClick = GetSetting("THERename", "Param", "ActDblClick", 2)
    mvarCenter0rSave = GetSetting("THERename", "Param", "CenterOrSave", 1)
    mvarCopyRename = CBool(GetSetting("THERename", "Param", "CopyRename", True))
    mvarFirstDateUse = GetSetting("THERename", "Param", "FirstDateUse", Date)
    mvarLastDirectory = GetSetting("THERename", "Param", "LastDirectory", "c:\windows")
    mvarLastUseTime = GetSetting("THERename", "Param", "LastUseTime", Time)
    mvarNumberOFiles = GetSetting("THERename", "Param", "NumberOFiles", 0)
    mvarNumberOfRuns = GetSetting("THERename", "Param", "NumberOfRuns", 1)
    mvarUndoFile = GetSetting("THERename", "Param", "UndoFile", "")
    mvarbatch = GetSetting("THERename", "Param", "batch", "")
    mvarHidden = CBool(GetSetting("THERename", "Param", "hidden", True))
    mvarlLeft = GetSetting("THERename", "Param", "left", RENAME.Left)
    mvarlTOp = GetSetting("THERename", "Param", "top", RENAME.Top)
    mvarwWidth = GetSetting("THERename", "Param", "wWidth", RENAME.width)
    mvarwHeight = GetSetting("THERename", "Param", "wHeight", RENAME.height)
    mvarReadOnly = CBool(GetSetting("THERename", "Param", "readonly", True))
    mvarSystem = CBool(GetSetting("THERename", "Param", "system", True))
    mvarMisc1 = GetSetting("THERename", "Param", "Misc1", mvarMisc1)
    mvarMisc2 = GetSetting("THERename", "Param", "Misc2", mvarMisc2)
    mvarMisc3 = GetSetting("THERename", "Param", "Misc3", mvarMisc3)
    mvarMisc4 = GetSetting("THERename", "Param", "Misc4", mvarMisc4)
    mvarMisc5 = GetSetting("THERename", "Param", "Misc5", mvarMisc5)
    mvarMisc6 = GetSetting("THERename", "Param", "Misc6", mvarMisc6)
    mvarMisc7 = GetSetting("THERename", "Param", "Misc7", mvarMisc7)
    mvarMisc8 = GetSetting("THERename", "Param", "Misc8", mvarMisc8)
    mvarMisc9 = GetSetting("THERename", "Param", "Misc9", mvarMisc9)
    mvarMisc10 = GetSetting("THERename", "Param", "Misc10", mvarMisc10)
    mvarMisc11 = GetSetting("THERename", "Param", "Misc11", "1")
    mvarMisc12 = GetSetting("THERename", "Param", "Misc12", "2")
    mvarMisc13 = GetSetting("THERename", "Param", "Misc13", "2")
 End If
 LoadSettings = True
End Function
Private Sub InitDefValues()
  mvarRestartCounter = 0
  mvarLevelRestart = 0
  mvarAutoArrange = True
  mvarSearchAndReplace = 0
  mvarDateformat = 0
  mvarStartupDir = ""
  mvarUseHistory = False
  mvarConfirmOperation = True
  mvarDefOption1 = 7
  mvarDefOption2 = 6
  mvarColumnsWiths = False
  mvarUseAutoSave = False
  mvarUseLowerInLetterCounters = 0
  mvarCompleCounters = 1
  mvarToolbarButtons = 1
  mvarCharTokens = " -;"
  mvarExifDelimD = "-"
  mvarExifDelimT = "-"
  mvarShowPathInCaption = 0
  mvarIncludePictInfo = 0
  mvarCheckLongFileNameOption = 0
  mvarRememberLastCommand = 0
  mvarShowWhenFileNameChange = 0
  mvarRememberLastFolder = 0
  mvarIncludeLinks = 0
  mvarAskQuestion = 0
  mvarRememberWSize = 1
  mvarShowMP3Tab = 1
  mvarShowMusicTab = 1
  mvarShowTextTab = 1
  mvarSelectAllFiles = 0
  mvarTagsVersionToUse = 1
  mvarMp3VqfOpt1 = 0
  mvarOggOpt1 = 0
  mvarOggOpt2 = " - "
  mvarLastFilter = "*.*"
  mvarMp3VqfOpt2 = 0
  mvarMp3VqfOpt3 = 0
  mvarMp3VqfOpt4 = 0
  mvarTagsPriority = 1
  mvarPicturesPreview = 0
  mvarTextToView = "TXT|DIZ|INI|NFO|LOG"
  mvarRegExpEngine = 1
  mvarBackRefNotation = 0
  mvarRxSyntax = 1
  mvarRegExOpt1 = 1
  mvarChineseKorean = 0
  mvarRemoveEmptyTags = 1
  mvarPCRE1 = 0
  mvarPCRE2 = 0
  mvarPCRE3 = 0
  mvarPCRE4 = 0
  mvarPCRE5 = 0
  mvarUseNaturalSort = 0
  mvarUseDefaultCyclicFile = 0
  mvarUseDefaultAbbrevFile = 0
  mvarDefaultCyclicFile = ""
  mvarDefaultAbbrevFile = ""
  mvarRulesOpt1 = 1
  mvarwWindowState = 0
  mvarLastFolder = ""
  mvarLastCommand = ""
  mvarCheckLongFileNameSize = 64
  mvarFilesToInclude = 0
  mvarPreviewGridLines = 1
  mvarIncHiddenFolders = 1
  mvarHtmlIncFolder = 1
  mvarHtmlIncSize = 1
  mvarHtmlIncDate = 1
  mvarHtmlIncAttr = 1
  mvarHtmlIncMusic = 0
  mvarIncVirtualFolders = 1
  mvarRemoveIllegals = 0
  mvarRemoveStartingSpaces = 1
  mvarDefaultCommand = ""
  mvarDefaultSearchPr = ""
  mvarDefaultReplacePr = ""
  mvarDefaultSearchExt = ""
  mvarDefaultReplaceExt = ""
  mvarRenameOnCollision = True
  mvarSilentMode = True
  mvarSettingsDirectory = ""
  mvarPicturesFormat = "%w%x%h%"
  mvarWordsDelimiters = " "
  mvarListDelimiter = 9
  mvarRemoveGuill = 0
  mvarAllowUndo = True
  mvarConfirmMakeDir = True
  mvarShutDown = False
  mvarFullRow = False
  mvarGridLines = False
  mvarPersonnalDate = ""
  mvarDisplayDate = "short Date"
  mvarLastUseDate = Date
  mvarLastUseTime = Time
  mvarNumberOfRuns = 1
  mvarFirstDateUse = Date
  mvarHidden = True
  mvarSystem = True
  mvarReadOnly = True
  mvarCenter0rSave = 1 ' 1=Centrer, 2=Rapel de la position
  mvarlLeft = RENAME.Left
  mvarlTOp = RENAME.Top
  If RENAME.WindowState <> vbMaximized Then
    mvarwWidth = RENAME.width
    mvarwHeight = RENAME.height
  Else
    mvarwWidth = -1
    mvarwHeight = -1
  End If
  mvarFormatDate = 2
  mvarFormatTime = 0
  mvarprog1 = ""
  mvarprog2 = ""
  mvarprog3 = ""
  mvarbatch = ""
  mvarLogFile = ""
  mvarUndoFile = ""
  mvarDirectoryReport = 1
  mvarNumberOFiles = 0
  mvarActDblClick = 2
  mvarLastDirectory = CurDir$()
  mvarCopyRename = True 'True=Rename, False=Copy
  mvarMisc1 = 2
  mvarMisc2 = 1
  mvarMisc3 = 0
  mvarMisc4 = 1
  mvarMisc5 = 1
  mvarMisc6 = 0
  mvarMisc7 = 0
  mvarMisc8 = 1
  mvarMisc9 = 0
  mvarMisc10 = 0
  mvarMisc11 = "1"
  mvarMisc12 = "2"
  mvarMisc13 = "2"
  
End Sub
