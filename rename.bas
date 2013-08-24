Attribute VB_Name = "StrDelete"
Option Explicit
Global LesOptions As New cSettings
' ***************** ANCIENNES VARIABLES GLOBALES *************************
'Global ReadOnly As Boolean
'Global System As Boolean
'Global wWidth As Long
'Global wHeight As Long
'Global lTOp As Long
'Global lLeft As Long
'Global Hidden As Boolean
'Global batch As String
'Global UndoFile As String
'Global NumberOfRuns As Long
'Global NumberOFiles As Long
'Global LastUseTime As String
'Global LastUseDate As String
'Global LastDirectory  As String
'Global FirstDateUse As String
'Global CopyRename As Boolean
'Global Center0rSave As Integer
'Global ActDblClick As Integer
'Global FormatTime As Integer
'Global FormatDate As Integer
'Global prog1 As String
'Global prog2 As String
'Global prog3 As String
'Global UseLowerInLetterCounters As Integer
'Global LevelRestart As Integer
'Global RestartCounter As Integer
'Global AutoArrange As Boolean
'Global StartupDir As String
'Global SearchAndReplace As Integer
'Global Dateformat As Integer
'Global DirectoryReport As Integer
'Global UseHistory As Boolean
'Global DisplayDate As String
'Global PersonnalDate As String
'Global GridLines As Boolean
'Global FullRow As Boolean
'Global ShutDown As Boolean
'Global ConfirmMakeDir As Boolean
'Global AllowUndo As Boolean
'Global SilentMode As Boolean
'Global RenameOnCollision As Boolean
'Global ConfirmOperation As Boolean
'Global WCol1 As Long                ' Largeur colonne 1
'Global WCol2 As Long                ' Largeur colonne 2
'Global WCol3 As Long                ' Largeur colonne 3
'Global WCol4 As Long                ' Largeur colonne 4
'Global WCol5 As Long                ' Largeur colonne 5
'Global UseAutoSave As Integer       ' Activer l'option auto-save ?
'Global ColumnsWiths As Boolean      ' Faut il ou pas sauvegarder la largeur des colonnes ?
'Global DefOption1 As Integer        ' Option par défaut pour rename.combo1
'Global DefOption2 As Integer        ' Option par défaut pour rename.combo2
'Global ExifDelimT As String         ' Délimiteur pour l'heure
'Global ExifDelimD As String         ' Délimiteur pour les dates
'Global ShowWhenFileNameChange As Integer ' Faut'il seulement afficher les noms des fichiers lorsqu'ils changent ?
'Global PicturesFormat As String
'Global IncludeLinks As Integer      ' Faut il mettre des liens sur les images pour le rapport en HTML ?
'Global LastFolder As String         ' Le nom du dernier répertoire vu
'Global RememberLastFolder As Integer ' Faut'il se rappeler du dernier répertoire  VU ?
'Global LastCommand As String        ' La dernière commande à se rappeler
'Global RememberLastCommand As Integer ' Faut il se rappeler de la dernière commande ?
'Global RemoveGuill As Integer       ' Faut il supprimer les guillemets des listes ?
'Global ListDelimiter As Integer     ' Caractère servant à délimiter les listes
'Global SettingsDirectory As String
'Global IncVirtualFolders As Integer ' Faut'il inclure les dossiers virtuels ?
'Global PreviewGridLines As Integer  ' Faut'il montrer une grille dans la fenêtre de preview ?
'Global IncHiddenFolders As Integer  ' Faut'il inclure les dossiers cachés ?
'Global HtmlIncMusic As Integer      ' Faut il inclure les tags des fichiers lorsqu'on génère un listing en HTML ?
'Global HtmlIncAttr As Integer       ' Faut il inclure les attributs des fichiers lorsqu'on génére un listing du dossier en HTML ?
'Global HtmlIncDate As Integer       ' Faut il inclure la date des fichiers lorsqu'on génére un listing du dossier en HTML ?
'Global HtmlIncSize As Integer       ' Faut il inclure la taille des fichiers lorsqu'on génére un listing du dossier en HTML ?
'Global HtmlIncFolder As Integer     ' Faut il inclure le nom du répertoire lorsqu'on génére un listing du dossier en HTML ?
'Global FilesToInclude As Integer    ' Qu'est-ce qu'il faut inclure, les fichiers, les répertoires, les 2 ?
'Global ShowPathInCaption As Integer ' Faut'il afficher le chemin dans la barre de titres ?
'Global CompleCounters As Integer    ' Faut'il compléter les compteurs avec des zéros ?
'Global ToolbarButtons As Integer    ' Style des boutons de la barre d'outils
'Global CheckLongFileNameSize As Integer ' Sur quelle taille faut il tester la longeur des fichiers ? Paramètre caché
'Global CheckLongFileNameOption As Integer ' Que faut il scanner lorsqu'on recherche les noms longs ?
'Global IncludePictInfo As Integer   ' Faut il inclure les dimensions des images dans les éditions ?
'Global CharTokens As String         ' Caractères à utiliser pour délimiter les tokens
'Global RemoveStartingSpaces As Integer ' Faut il supprimer les blancs en début de fichier ?
'Global RemoveIllegals As Integer    ' Doit on automatiquement supprimer les caractères illégaux ?
'Global DefaultReplaceExt As String  ' Remplacement par défaut dans le prefix
'Global DefaultSearchExt As String   ' Recherche par défaut dans l'extension
'Global DefaultReplacePr As String   ' Remplacement par défaut dans le préfix
'Global DefaultSearchPr As String    ' Recherche par défaut dans le préfix
'Global DefaultCommand As String     ' Commande du free form à prendre comme commande par défaut
'Global TextToView As String         ' Extensions des fichiers qu'on veut visualiser dans la fenêtre de preview des fichiers
'Global wWindowState As Integer       ' comme le nom l'indique...
'Global RulesOpt1 As Integer         ' Faut il utiliser des comparaisons insensibles à la casse pour les règles ?
'Global UseNaturalSort As Integer    ' Faut il utiliser le tri "naturel" ?
'Global AskQuestion As Integer       ' Option cachée dans la base de registres pour savoir s'il faut poser la question (chiante) pour les compteurs
'Global SelectAllFiles As Integer    ' Faut il sélectionner tous les fichiers des que l'on rentre dans un répertoire ?
'Global ShowTextTab As Integer       ' Doit on afficher l'onglet permettant de voir les textes ?
'Global ShowMusicTab As Integer      ' Doit on afficher l'onglet sur les images ?
'Global ShowMP3Tab As Integer        ' Doit on afficher l'onglet MP3 ?
'Global RememberWSize As Integer     ' Faut il se rappeler de la taille de la fenêtre ?
'Global DefaultAbbrevFile As String  ' Nom du fichier d'abbréviations à utiliser par défaut
'Global UseDefaultAbbrevFile As Integer  ' Faut il utliser un fichier d'abréviations par défaut ?
'Global DefaultCyclicFile As String  ' Nom du fichier de sélections cycliques à utiliser par défaut
'Global UseDefaultCyclicFile As Integer ' Faut il utiliser un fichier de sélections cycliques par défaut ?
'Global OggOpt1 As Integer           ' Indique comment traiter les tags lorsqu'il y en a plusieurs
'Global OggOpt2 As String            ' Indique le séparateur de tag à utiliser lorsqu'il y en a plusieurs et qu'on les combine ensembles
'Global PCRE2 As Integer
'Global PCRE3 As Integer
'Global PCRE4 As Integer
'Global PCRE5 As Integer
'Global PCRE1 As Integer
'Global RemoveEmptyTags As Integer ' Faut il supprimer les tags vides (pour les MP3) de la liste de prévisualisation des tags ?
'Global ChineseKorean As Integer ' Changement de méthode de renommage pour les utilisateurs de windows en chinoix ou en koréen.
'Global RegExOpt1 As Integer
'Global RegExpEngine As Integer
'Global BackRefNotation As Integer
'Global LastFilter As String         ' Indique le dernier filtre utilisé sur les fichiers
'Global PicturesPreview As Integer ' Indique la façon dont on veut prévisualiser les images (Taille réelle, Stretch ou Best Fit)
'Global TagsPriority As Integer  ' Priorité des tags
'Global WordsDelimiters As String
'Global RxSyntax As Integer
'Global LogFile As String
'Global TagsVersionToUse As Integer  ' Quelle version des tags faut il utiliser ?
'Global Mp3VqfOpt1 As Integer    ' Separate Words ?
'Global Mp3VqfOpt2 As Integer    ' Remove multiple spaces ?
'Global Mp3VqfOpt3 As Integer    ' Action à effectuer sur la casse des lettres
'Global Mp3VqfOpt4 As Integer    ' Nombre de zéros avec lesquels il faut compléter les numéros de pistes
'Global Misc1 As Integer
'Global Misc2 As Integer
'Global Misc3 As Integer
'Global Misc4 As Integer
'Global Misc5 As Integer
'Global Misc6 As Integer
'Global Misc7 As Integer
'Global Misc8 As Integer
'Global Misc9 As Integer
'Global Misc10 As Integer
'Global Misc11 As String
'Global Misc12 As String
'Global Misc13 As String
' ************************************************************************
' ***** Variables utilisée pour la copie des noms de fichiers ************
Global LOption1 As Integer
Global LOption2 As Integer
Global LOption3 As Integer
Global LChaine1 As String
Global LChaine2 As String
Global LOk As Boolean
' ************************************************************************
Global ExeCmd As Boolean
Global LesRegles As New Rules
Global cHist14 As New cHistory
Global AppPath As String
Global CollAbrev As New Collection  ' Contient toutes les abréviations
Global OkUseAbbrev As Boolean       ' Doit on utiliser les abrévitions ?
Global UseCylcic As Boolean         ' On utilise les cylclics ?
Global LesCyclic() As String        ' Sauvegarde des items Cyclic
Global VnbCyclic As Integer         ' nombre d'éléments cycliques
Global OptionsCyclic As Boolean     ' Faut'il afficher les options des cyclics ?
Global PlacementCyclic As Integer   ' Option de placement choisie
Global CompteurCyclic As Integer
Global AncTitre As String           ' Sauvegarde de l'ancien titre de le fenêtre principale
Global UseMP3 As Boolean            ' Faut'il utiliser les infos de musique sur les MP3 ?
Global UseVQF As Boolean            ' Faut'il utiliser les infos de musique sur les VQF ?
Global UseOGG As Boolean            ' Faut'il utiliser les infos de musique sur les ogg ?
Global UseWMA As Boolean            ' Faut'il utiliser les infos de musique sur les WMA ?
Global FilterOk As Boolean
Global FilterRegular As Boolean
Global FilterExpr As String
Global MusMP3 As New clsMP3         ' La classe permettant de gérer les MP3
Global PicEXIF As New CExif         ' La classe permettant de gérer les tags des images au format EXIF
Global MusVQF As New clsVQF         ' La classe permettant de gérer les VQF
Global MusOgg As New clsOGG         ' La classe permettant de gérer les Ogg
Global MusWMA As New clsWMA         ' La classe permettant de gérer les WMA
Global AFM As New clsAFM            ' La classe permettant de gérer les fichiers Afm
Global DT1 As New CDateTime         ' 1=Renommer
Global DT2 As New CDateTime         ' 2=Bin
Global DT3 As New CDateTime         ' 3=Bouton copies multiples
Global DT4 As New CDateTime         ' 4=Changement immédiat
Global Attr1 As New CAttrib         ' 1=Renommer
Global Attr2 As New CAttrib         ' 2=Bin
Global Attr3 As New CAttrib         ' 3=Bouton copies multiples
Global Attr4 As New CAttrib         ' 4=changement immédiat
Global rech1 As New CSearch ' Recherche et remplacement sur le préfixe
Global rech2 As New CSearch ' Recherche et remplacement sur le suffixe
Global rech3 As New CSearch ' Recherche et remplacement dans le préfixe ET l'extension
Global DTEnCours As Integer         ' 1= Renommer, 2=bin, 3=bouton copies multiples
Global AttrEncours As Integer       ' 1= Renommer, 2=bin, 3=bouton copies multiples
Global QDateTravail As Integer      ' Sur quelle format de date faut il travailler
Global RechEnCours As Integer ' 1=Préfixe, 2=Suffixe
Global RechPref As Boolean ' Indique s'il faut faire de la recherche et du remplacement dans le préfixe
Global RechSuff As Boolean ' Indique s'il faut faire de la recherche et du remplacement dans le suffixe
Global RechGlob As Boolean ' Indique s'il faut faire de la recherche et du remplacement dans le préfixe ET l'extension
Global Folder1 As Integer
Global Folder2 As Integer
Global Folder3 As Integer
Global Folder4 As Integer
Global Folder5 As String
Global Folder6 As String
Global FolderOk As Boolean
Global VnbHistory As Integer
Global LesRepertoires(100) As String
Global VnbRep As Integer
Global TemMove As Boolean
Global Dir1Path As String
Global VancRep As String
Global Lechemin As String
Global Recursive As Boolean ' Indique si on est en recherche récursive ou pas
Global Rafraichir As Boolean
Global TemDelete As Boolean
Global Filtre As String ' Contient le filtre de sélection des fichiers
Global Annuler As Boolean
Global fav(20) As String

Public Const INVALID_HANDLE_VALUE = -1
Public Const VK_MENU = &H12
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CreateScalableFontResource Lib "gdi32" Alias "CreateScalableFontResourceA" (ByVal fHidden As Long, ByVal lpszResourceFile As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hWnd As Long, ByVal dwType As Long) As Long
Public Declare Function WNetDisconnectDialog Lib "mpr.dll" (ByVal hWnd As Long, ByVal dwType As Long) As Long
Public Const RESOURCETYPE_DISK = &H1
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const SW_NORMAL = 1
Public Const ERROR_BAD_FORMAT = 11&
Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const SE_ERR_ACCESSDENIED = 5
Public Const SE_ERR_ASSOCINCOMPLETE = 27
Public Const SE_ERR_DDEBUSY = 30
Public Const SE_ERR_DDEFAIL = 29
Public Const SE_ERR_DDETIMEOUT = 28
Public Const SE_ERR_FNF = 2
Public Const SE_ERR_NOASSOC = 31
Public Const SE_ERR_OOM = 8
Public Const SE_ERR_PNF = 3
Public Const SE_ERR_SHARE = 26
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3
Public Const NOERROR = 0
Public Const CSIDL_FAVORITES = &H6
' Ouvre le browser par défaut avec une URL
Public Function BrowseTo(ByVal sURL As String) As Boolean
Dim xRet As Long
    xRet = ShellExecute(GetDesktopWindow(), vbNullString, sURL, vbNullString, App.Path, SW_NORMAL)
    If xRet > 32 Then BrowseTo = True Else BrowseTo = False
End Function
' Renvoie le nombre d'occurence d'un caratère dans une chaine
Public Function CharOccurs(chaine As String, cherche As String) As Integer
 Dim vnb As Integer
 Dim i As Integer
 Dim tot As Integer
 If Len(chaine) = 0 Then
  CharOccurs = 0
 Else
  vnb = Len(chaine)
  tot = 0
  For i = 1 To vnb
   If Mid$(chaine, i, 1) = cherche Then
    tot = tot + 1
   End If
  Next
  CharOccurs = tot
 End If
End Function
' Renvoie la position du nième caractère "cherche" dans "chaine"
Public Function At(chaine As String, cherche As String, occurence As Integer) As Integer
 Dim vnb As Integer
 Dim i As Integer
 Dim position As Integer
 Dim loccurence As Integer
 If occurence = 0 Then
  At = 0
 Else
  If Len(chaine) = 0 Then
   At = 0
  Else
   loccurence = 0
   position = 0
   vnb = Len(chaine)
   For i = 1 To vnb
    If Mid$(chaine, i, 1) = cherche Then
     loccurence = loccurence + 1
     If loccurence = occurence Then
      position = i
      Exit For
     End If
    End If
   Next
   At = position
  End If
 End If
End Function
' Renvoie le préfixe d'un fichier en tenant compte du chemin
Public Function Prefixe(fichier As String) As String
 Dim position As Integer
 Dim prtmp As String
 If Len(fichier) = 0 Then
    Prefixe = ""
 Else
    If InStr(1, fichier, "\") = 0 Then  ' fichier sous la forme "autoexec.bat" par exemple, donc sans chemin
        position = At(fichier, ".", CharOccurs(fichier, "."))
        If position <> 0 Then   ' Il y a une extension
            Prefixe = Left$(fichier, position - 1)
        Else    ' Il n'y a pas d'extension
            Prefixe = fichier
        End If
    Else    ' Il doit y avoir un chemin dans le nom
        prtmp = Mid$(fichier, At(fichier, "\", CharOccurs(fichier, "\")) + 1)
        position = At(prtmp, ".", CharOccurs(prtmp, "."))
        If position <> 0 Then ' Il y a une extension
            Prefixe = Left$(prtmp, position - 1)
        Else                  ' Il n'y a pas d'extension
            Prefixe = prtmp
        End If
    End If
 End If
 End Function
' Renvoie le suffixe d'un fichier en tenant compte du chemin
Public Function Suffixe(fichier As String) As String
 Dim position As Integer
 Dim prtmp As String
 If Len(fichier) = 0 Then
    Suffixe = ""
 Else
    If InStr(1, fichier, "\") = 0 Then  ' fichier sans chemin "autoexec.bat" par exemple
        position = InStr(1, fichier, ".")
        If position <> 0 Then   ' Il y a une extension
            Suffixe = Mid$(fichier, At(fichier, ".", CharOccurs(fichier, ".")) + 1)
        Else    ' Il n'y a pas d'extension
            Suffixe = ""
        End If
    Else    ' Il y a un chemin dans le nom
        prtmp = Mid$(fichier, At(fichier, "\", CharOccurs(fichier, "\")) + 1) ' On ne récupère que le nom complet sans le chemin
        position = InStr(1, prtmp, ".")
        If position <> 0 Then ' Il y a une extension
            Suffixe = Mid$(prtmp, At(prtmp, ".", CharOccurs(prtmp, ".")) + 1)
        Else    ' Il n'y a pas d'extension
            Suffixe = ""
        End If
    End If
 End If
End Function
' Inverse les majuscules minuscules d'une chaine
Public Function ToggleCase(chaine As String) As String
Dim vnb As Integer
Dim i As Integer
Dim resultat As String
Dim extrait As String
resultat = ""
vnb = Len(chaine)
For i = 1 To vnb
 extrait = Mid$(chaine, i, 1)
 If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", extrait) <> 0 Then  'Passage en minuscules
  resultat = resultat + LCase$(extrait)
 Else
  If InStr("abcdefghijklmnopqrstuvwxyz", extrait) <> 0 Then  'Passage en majuscules
   resultat = resultat + UCase$(extrait)
  Else ' Pas de changement
   resultat = resultat + extrait
  End If
 End If
Next
ToggleCase = resultat
End Function
Public Function Menage(chaine As String) As String
 Dim vnb As Integer
 Dim i As Integer
 Dim resultat As String
 Dim extrait As String
 Dim vinterdits As String
 vinterdits = "\/:*?" + Chr$(34) + "<>|" + Chr$(10) + Chr$(13)
 vnb = Len(chaine)
 resultat = ""
 For i = 1 To vnb
  extrait = Mid$(chaine, i, 1)
  If InStr(vinterdits, extrait) = 0 Then
   If IsCharAlpha(Asc(extrait)) Or InStr("1234567890- ", extrait) <> 0 Then
    resultat = resultat + extrait
   End If
  End If
 Next
 Menage = resultat
End Function
' Création d'un compteur selon une base sur une longueur donnée avec complétion par des zéros
Public Function Compteur(valeur As Long, nbdiggits As Integer, formatrep As Integer) As String
Dim chaine As String
Dim longueur As Integer
Dim NewValeur As String
Select Case formatrep
    Case 0 'Decimal
        longueur = Len(Trim$(Str$(valeur)))
        NewValeur = Trim$(Str$(valeur))
    Case 1  ' Hexa
        longueur = Len(Trim$(Hex$(valeur)))
        NewValeur = Trim$(Hex$(valeur))
    Case 2 ' Octal
        longueur = Len(Trim$(Oct$(valeur)))
        NewValeur = Trim$(Oct$(valeur))
    Case 3 ' Binaire
        longueur = Len(Trim$(bin$(CDbl(valeur))))
        NewValeur = Trim$(bin$(CDbl(valeur)))
    Case 4 ' Lettres
        NewValeur = FBase26(valeur)
        If LesOptions.UseLowerInLetterCounters = 1 Then ' passer en minuscules
            NewValeur = LCase$(NewValeur)
        End If
        Compteur = Trim$(NewValeur)
        Exit Function
    Case 5 ' Romain
        NewValeur = Roman(valeur)
        If LesOptions.UseLowerInLetterCounters = 1 Then ' passer en minuscules
            NewValeur = LCase$(NewValeur)
        End If
        Compteur = Trim$(NewValeur)
        Exit Function
End Select
 
If LesOptions.CompleCounters = 1 Then ' Faut'il compléter avec des zéros ?
 If longueur < nbdiggits Then
  chaine = String$(nbdiggits - longueur, "0") + NewValeur
 Else
  chaine = NewValeur
 End If
Else
 chaine = NewValeur
End If
Compteur = Trim$(chaine)
End Function
' Lance l'éxecution d'un fichier (ou programme)
Public Sub FileExecutor(lhWnd As Long, Path As String, action As String, Optional cParms As Variant, Optional nShowCmd As Variant)
Dim lRtn As Long 'declare the needed variables
lRtn = ShellExecute(lhWnd, action, Path, 0&, Path, SW_NORMAL) 'execute or print the file or folder
If lRtn <= 32 Then 'if an error is found then call the FileError function
   FileError (lRtn)
End If
End Sub
' Fonction interne
Private Sub FileError(lRtn As Long)
Dim Msg As String
Select Case lRtn 'if any errors occur then display them to the user
       Case 0
       Msg = "Memory Error"
       Case ERROR_BAD_FORMAT
       Msg = "Bad Executeable Format"
       Case ERROR_FILE_NOT_FOUND
       Msg = "File not found"
       Case ERROR_PATH_NOT_FOUND
       Msg = "Path not found"
       Case SE_ERR_ACCESSDENIED
       Msg = "Access Denied"
       Case SE_ERR_ASSOCINCOMPLETE
       Msg = "Association incomplete"
       Case SE_ERR_DDEBUSY
       Msg = "DDE Busy error"
       Case SE_ERR_DDEFAIL
       Msg = "DDE failed"
       Case SE_ERR_DDETIMEOUT
       Msg = "DEE time out"
       Case SE_ERR_FNF
       Msg = "File not found"
       Case SE_ERR_NOASSOC
       Msg = "No association for this file"
       Case SE_ERR_OOM
       Msg = "Out of Memory"
       Case SE_ERR_PNF
       Msg = "Path could not be found"
       Case SE_ERR_SHARE
       Msg = "Sharing violation"
       Case Else
       Msg = "Unknown Error!, Please try again..."
End Select
MsgBox Msg, vbCritical
End Sub
' Chargement des préférences utilisateur
Public Function LoadPref() As Integer
 On Error Resume Next
 Dim i As Integer
 Dim response As Integer
 Dim MySettings As Variant, intSettings As Integer
 Dim laversion As String
    Dim LastUseDate As String
 LastUseDate = GetSetting("THERename", "Param", "LastUseDate", "")
 LesOptions.LoadSettings
 If LastUseDate = "" Then   ' C'est la première utilisation du logiciel
  LoadPref = 1
  MsgBox "That's the first time you launch THE Rename. I'm going to save default parameters, add your Windows directory as your first favorite and add THE Rename to the Explorer context menu. Enjoy my program. Report bugs and suggestions to beug@herve-thouzard.com."
  RENAME.MousePointer = 11
  fav(1) = GetWinPath()
  For i = 2 To 20
   fav(i) = ""
  Next
  For i = 1 To 20
   SaveSetting "THERename", "Param", "fav" + Trim$(Str$(i)), fav(i)
  Next
  MoreParameter
  response = MsgBox("Would you like to add an internet shortcut in your favorites to THE Rename home page ?", vbYesNo, "THE Rename's home page")
  If response = vbOK Then
   CreateInternetLink
  End If
  RENAME.MousePointer = 0
 Else ' Ca n'est pas la première utilisation, donc on charge les paramètres ******************************************************************************
  LoadPref = 0
  laversion = GetSetting("THERename", "Param", "Version", "Aucune")
  If laversion = "Aucune" Then
   MoreParameter
  End If
  
  If laversion <> "2.0" Then
   SaveSetting "THERename", "Param", "Version", "2.0"
  End If
  
  MySettings = GetAllSettings(appname:="THERename", Section:="ext")
  RENAME.Combo5.Clear
  For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
   RENAME.Combo5.AddItem MySettings(intSettings, 1)
  Next
  RENAME.Combo5.Text = RENAME.Combo5.List(0)
  Filtre = RENAME.Combo5.List(0)
  For i = 1 To 20
   fav(i) = GetSetting("THERename", "Param", "fav" + Trim$(Str$(i)), "")
  Next
  If Trim$(LesOptions.LastFilter) <> "" Then
    RENAME.Combo5.Text = LesOptions.LastFilter
    Filtre = LesOptions.LastFilter
  End If
 End If
 For i = 1 To 19
  Load RENAME.menufav(i)
  Load RENAME.mnufav(i)
 Next
 For i = 0 To 19
  RENAME.menufav(i).Caption = "&" + Chr$(65 + i) + " " + fav(i + 1)
  RENAME.mnufav(i).Caption = "&" + Chr$(65 + i) + " " + fav(i + 1)
 Next
 LoadPref = 1
End Function

Private Function GetWinPath() As String
Dim strFolder As String
Dim lngResult As Long
strFolder = String$(MAX_PATH, 0)
lngResult = GetWindowsDirectory(strFolder, MAX_PATH)
If lngResult <> 0 Then
    GetWinPath = Left$(strFolder, InStr(strFolder, Chr$(0)) - 1)
Else
    GetWinPath = ""
End If
End Function
' Sauvegarde des préférences utilisateur
Public Function SaveSettings() As Integer
Dim i As Integer
Dim vnb As Long
On Error GoTo ErrGen
RENAME.MousePointer = 11
LesOptions.SaveSettings
For i = 1 To 20
 SaveSetting "THERename", "Param", "fav" + Trim$(Str$(i)), fav(i)
Next
DeleteSetting "THERename", "Ext"
vnb = RENAME.Combo5.ListCount - 1
For i = 0 To vnb
 SaveSetting "THERename", "Ext", "ext" + Trim$(Str$(i)), RENAME.Combo5.List(i)
Next
RENAME.MousePointer = 0
Exit Function

ErrGen:
ErreurGrave "SaveSettings"
End Function
' **********************************************************
' Fonction Interne pour créer un raccourci Internet
' **********************************************************
Private Function GetSpecialPath(CSIDL As Long) As String
  Dim R As Long
  Dim Path As String
  Dim IDL As ITEMIDLIST
  'fill the idl structure with the specified folder item
   R = SHGetSpecialFolderLocation(RENAME.hWnd, CSIDL, IDL)
  If R = NOERROR Then
   Path$ = Space$(512)
   R = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
   GetSpecialPath = Left$(Path, InStr(Path, Chr$(0)) - 1)
   Exit Function
  End If
 GetSpecialPath = ""
End Function
' **********************************************************
' Fonction pour créer un raccourci Internet
' **********************************************************
Private Sub CreateInternetLink()
  Dim URLpath As String
  Dim CSIDLpath As String
  Dim nameofLink As String
  Dim ff As Integer
  URLpath = "http://www.herve-thouzard.com/therename.phtml"
  CSIDLpath = GetSpecialPath(CSIDL_FAVORITES) & "\"
  nameofLink = "THE Rename Home Page.url"
  ff = FreeFile
  Open CSIDLpath & nameofLink For Output As #ff
  Print #ff, "[InternetShortcut]"
  Print #ff, "URL=" & URLpath
  Close #ff
End Sub
' Retourne le nom complet d'une police truetype
Public Function GetFontName(FileNameTTF As String) As String
Dim hFile As Integer
Dim Buffer As String
Dim FontName As String
Dim TempName As String
Dim iPos As Integer
'Build name for new resource file in a temporary file, and call API.
 TempName = AppPath & "~TEMPORAIRE.FOT"
 If CreateScalableFontResource(1, TempName, FileNameTTF, vbNullString) Then
   'The name sits behind the text "FONTRES:"
    hFile = FreeFile
    Open TempName For Binary Access Read As hFile
       Buffer = Space$(LOF(hFile))
       Get hFile, , Buffer
       iPos = InStr(Buffer, "FONTRES:") + 8
       FontName = Mid$(Buffer, iPos, InStr(iPos, Buffer, vbNullChar) - iPos)
    Close hFile
    Kill TempName
  Else
   FontName = Prefixe(FileNameTTF)
  End If
'Return the font name
 GetFontName = FontName
End Function
' Retourne le "nom" d'un document HTML qui doit se trouver dans la balise <TITLE>
Public Function GetHtmlName(Filename As String) As String
 Dim ligne As String, copie As String, vnb1 As Integer
 Dim vnb2 As Integer, letitre As String, Buffer As String, vrai As Boolean
 Dim i As Integer
 Dim Accents(21, 2) As String
 Dim ff As Integer
 Accents(1, 1) = "&"
 Accents(2, 1) = "é"
 Accents(3, 1) = "è"
 Accents(4, 1) = "ç"
 Accents(5, 1) = "à"
 Accents(6, 1) = "ù"
 Accents(7, 1) = "î"
 Accents(8, 1) = "ô"
 Accents(9, 1) = "û"
 Accents(10, 1) = "â"
 Accents(11, 1) = "ö"
 Accents(12, 1) = "ë"
 Accents(13, 1) = "ù"
 Accents(14, 1) = "<"
 Accents(15, 1) = ">"
 Accents(16, 1) = "£"
 Accents(17, 1) = "â"
 Accents(18, 1) = "ê"
 Accents(19, 1) = "û"
 Accents(20, 1) = "î"
 Accents(21, 1) = "ô"
 Accents(1, 2) = "&amp;"
 Accents(2, 2) = "&eacute;"
 Accents(3, 2) = "&egrave;"
 Accents(4, 2) = "&ccedil;"
 Accents(5, 2) = "&agrave;"
 Accents(6, 2) = "&ugrave;"
 Accents(7, 2) = "&icirc;"
 Accents(8, 2) = "&ocirc;"
 Accents(9, 2) = "&ucirc;"
 Accents(10, 2) = "&acirc;"
 Accents(11, 2) = "&ouml;"
 Accents(12, 2) = "&euml;"
 Accents(13, 2) = "&ugrave;"
 Accents(14, 2) = "&lt;"
 Accents(15, 2) = "&gt;"
 Accents(16, 2) = "&pound;"
 Accents(17, 2) = "&acirc;"
 Accents(18, 2) = "&ecirc;"
 Accents(19, 2) = "&ucirc;"
 Accents(20, 2) = "&icirc;"
 Accents(21, 2) = "&ocirc;"
 
 vrai = False
 letitre = Filename
 ff = FreeFile
 Open Filename For Input As #ff
 Line Input #ff, ligne
 
 While Not EOF(1)
  copie = UCase$(ligne)
  vnb1 = InStr(copie, "<TITLE>")
  If vnb1 <> 0 Then ' le début du titre est trouvé
   vnb2 = InStr(copie, "</TITLE>")
   If vnb2 <> 0 Then
    letitre = Mid$(ligne, vnb1 + 7, vnb2 - (vnb1 + 7))
    GoTo suite
   Else
    Buffer = ligne
    While vrai = False And Not EOF(1)
     Line Input #ff, ligne
     Buffer = Buffer + ligne
     copie = UCase$(ligne)
     vnb2 = InStr(copie, "</TITLE>")
     If vnb2 <> 0 Then
      vrai = True
     End If
    Wend
    If vrai = True Then
     letitre = Mid$(Buffer, vnb1 + 7, vnb2 - (vnb1 + 7))
     GoTo suite
    End If
   End If
  End If
  Line Input #ff, ligne
 Wend
 
suite:
 Close #ff
 For i = 1 To 21
  letitre = Replace(letitre, Accents(i, 2), Accents(i, 1), , 1, vbTextCompare)
 Next
 letitre = Menage(letitre)
 GetHtmlName = letitre
End Function
' Renvoie les tokens d'une chaine
Public Function GetToken(s As String, Token As String, ByVal Nth As Integer) As String
   Dim i As Integer
   Dim P As Integer
   Dim R As Integer

   If Nth < 1 Then
      GetToken = ""
      Exit Function
   End If

   R = 0

   For i = 1 To Nth
      P = R
      R = InStr(P + 1, s, Token)
      If R = 0 Then
         If i = Nth Then
            GetToken = Mid$(s, P + 1, Len(s) - P)
         Else
            GetToken = ""
         End If
         Exit Function
      End If
   Next

   GetToken = Mid$(s, P + 1, R - P - 1)
End Function
' **************************************************************************************************
' Supprime d'une chaine tout ce qui n'est pas numérique
' **************************************************************************************************
Public Function Menage2(chaine As String) As String
 Dim vnb As Integer
 Dim i As Integer
 Dim resultat As String
 Dim extrait As String
 Dim vinterdits As String
 vinterdits = "0123456789"
 vnb = Len(chaine)
 resultat = ""
 For i = 1 To vnb
  extrait = Mid$(chaine, i, 1)
  If InStr(vinterdits, extrait) <> 0 Then
   resultat = resultat + extrait
  End If
 Next
 Menage2 = resultat
End Function
' Sauvegarde de paramètres supplémentaires
Private Sub MoreParameter()
On Error GoTo ErrGen
   SaveSetting "THERename", "Ext", "ext1", "*.*"
   SaveSetting "THERename", "Ext", "ext2", "*.bmp"
   SaveSetting "THERename", "Ext", "ext3", "*.ogg"
   SaveSetting "THERename", "Ext", "ext4", "*.doc"
   SaveSetting "THERename", "Ext", "ext5", "*.wma"
   SaveSetting "THERename", "Ext", "ext6", "*.gif"
   SaveSetting "THERename", "Ext", "ext7", "*.htm"
   SaveSetting "THERename", "Ext", "ext8", "*.html"
   SaveSetting "THERename", "Ext", "ext9", "*.jpg"
   SaveSetting "THERename", "Ext", "ext10", "*.rar"
   SaveSetting "THERename", "Ext", "ext11", "*.ttf"
   SaveSetting "THERename", "Ext", "ext12", "*.txt"
   SaveSetting "THERename", "Ext", "ext13", "*.zip"
   SaveSetting "THERename", "Ext", "ext14", "*.mp3"
   SaveSetting "THERename", "Ext", "ext15", "*.vqf"
   SaveSetting "THERename", "Param", "Version", "1.8"
   With RENAME.Combo5
        .Clear
        .AddItem "*.*"
        .AddItem "*.bmp"
        .AddItem "*.ogg"
        .AddItem "*.doc"
        .AddItem "*.wma"
        .AddItem "*.gif"
        .AddItem "*.htm"
        .AddItem "*.html"
        .AddItem "*.jpg"
        .AddItem "*.rar"
        .AddItem "*.ttf"
        .AddItem "*.txt"
        .AddItem "*.zip"
        .AddItem "*.mp3"
        .AddItem "*.vqf"
        .ListIndex = 0
        .Text = RENAME.Combo5.List(RENAME.Combo5.ListIndex)
    End With
   Exit Sub

ErrGen:
ErreurGrave "MoreParameter"
End Sub
' **************************************************************************************************
' Renvoie le nom court d'un fichier
' **************************************************************************************************
Public Function ShortName(Path$) As String
 Dim short As String * 255
 short = Space$(255)
 Dim dummy As Long
 dummy = GetShortPathName(Path$, short, Len(short))
 ShortName = Left$(short, dummy)
End Function
' **************************************************************************************************
' Fonction pour supprimer les espace internes à une chaine
' **************************************************************************************************
Public Function RInternalSpaces(chaine As String) As String
    RInternalSpaces = Replace(chaine, " ", "")
End Function
' **************************************************************************************************
' Formatage des dates selon les paramètres utilisateur
' **************************************************************************************************
Public Function FmtDate(dateaf As Date) As String
 On Error GoTo erreur
 
 Select Case LesOptions.FormatDate
  Case 0 ' JJ MM AAAA
   FmtDate = Trim$(Format$(Day(dateaf), "00") + Format$(Month(dateaf), "00") + Format$(Year(dateaf), "0000"))
  Case 1 ' JJ AAAA MM
   FmtDate = Trim$(Format$(Day(dateaf), "00") + Format$(Year(dateaf), "0000") + Format$(Month(dateaf), "00"))
  Case 2 ' AAAA MM JJ
   FmtDate = Trim$(Format$(Year(dateaf), "0000") + Format$(Month(dateaf), "00") + Format$(Day(dateaf), "00"))
  Case 3 ' AAAA JJ MM
   FmtDate = Trim$(Format$(Year(dateaf), "0000") + Format$(Day(dateaf), "00") + Format$(Month(dateaf), "00"))
  Case 4 ' MM JJ AAAA
   FmtDate = Trim$(Format$(Month(dateaf), "00") + Format$(Day(dateaf), "00") + Format$(Year(dateaf), "0000"))
  Case 5 ' MM AAAA JJ
   FmtDate = Trim$(Format$(Month(dateaf), "00") + Format$(Year(dateaf), "0000") + Format$(Day(dateaf), "00"))
  Case 6 ' long date
   FmtDate = Format$(dateaf, "long date")
  Case 7 ' Other (format personnel)
   FmtDate = Format$(dateaf, LesOptions.PersonnalDate)
 End Select
 Exit Function
erreur:
  FmtDate = "<Invalid command in personal date format>"
End Function
' **************************************************************************************************
' Formatage de l'heure selon les paramètres utilisateur
' **************************************************************************************************
Public Function FmtHeure(heureaf As String) As String
 Select Case LesOptions.FormatTime
  Case 0
   FmtHeure = Trim$(Format$(heureaf, "Long Time"))
  Case 1
   FmtHeure = Trim$(Format$(heureaf, "Medium Time"))
  Case 2
   FmtHeure = Trim$(Format$(heureaf, "Short Time"))
 End Select
End Function
'*********************************************************************
' A complex routine that finds all of the files in a directory (and its
' subdirectories), loads the results in a collection, and returns the
' number of subdirectories that were searched.
'*********************************************************************
Public Function FindAllFiles(ByVal strSearchPath$, strPattern As String, Optional colFiles As Collection, Optional colDirs As Collection, Optional blnDirsOnly As Boolean, Optional blnBoth As Boolean) As Integer
' Create a new FindFile object every time this function is called
Dim clsFind As New clsFindFile
Dim strFile As String
Dim intDirsFound As Integer
' *** Mes variables ****************
 Dim attributs As Long
 Dim chaine As String
 Dim afficher As Boolean
 Dim itmX As ListItem
 clsFind.Dateformat = "short Date"
 
' **********************************
' Make sure strSearchPath always has a trailing backslash
strSearchPath$ = AddBackSlash(strSearchPath$)
Lechemin = strSearchPath$

strFile = clsFind.Find(strSearchPath$ & strPattern) ' Get the first file
Do While Len(strFile)   ' Loop while files are being returned
    If clsFind.FileAttributes And vbDirectory Then  ' If the current file found is a directory...
        If Left$(strFile, 1) <> "." Then ' Ignore . and ..
            If LesOptions.FilesToInclude = 1 Or LesOptions.FilesToInclude = 2 Then ' Il faut inclure les répertoires
                Lechemin = strSearchPath$
                attributs = clsFind.FileAttributes
                afficher = True
                If (attributs And FILE_ATTRIBUTE_READONLY) And LesOptions.ReadOnly = False Then
                    afficher = False
                End If
                If (attributs And FILE_ATTRIBUTE_HIDDEN) And LesOptions.Hidden = False Then
                    afficher = False
                End If
                If (attributs And FILE_ATTRIBUTE_SYSTEM) And LesOptions.System = False Then
                    afficher = False
                End If
                chaine = ""
                If afficher = True Then
                    If attributs And FILE_ATTRIBUTE_READONLY Then
                        chaine = "R"
                    End If
                    If attributs And FILE_ATTRIBUTE_HIDDEN Then
                        chaine = chaine + "H"
                    End If
                    If attributs And FILE_ATTRIBUTE_SYSTEM Then
                        chaine = chaine + "S"
                    End If
                    If attributs And FILE_ATTRIBUTE_ARCHIVE Then
                        chaine = chaine + "A"
                    End If
                    If chaine = "" Then
                        chaine = " "
                    End If
                    Set itmX = RENAME.ListView1.ListItems.Add(, , strSearchPath & strFile)
                    itmX.Text = strSearchPath & strFile
                    itmX.Bold = True ' *****************************************************************************
                    itmX.SubItems(4) = "Dir"  ' Type répertoire
                    itmX.SubItems(1) = clsFind.FileSize
                    Select Case LesOptions.Dateformat
                        Case 0
                            itmX.SubItems(2) = clsFind.GetCreationDate
                        Case 1
                            itmX.SubItems(2) = clsFind.GetLastWriteDate
                        Case 2
                            itmX.SubItems(2) = clsFind.GetLastAccessDate
                    End Select
                    itmX.SubItems(3) = chaine
                End If
            
            End If
            If blnDirsOnly Or blnBoth Then ' If either bln optional arg is true, then add this directory to the optional colDirs collection
                colDirs.Add strSearchPath$ & strFile & "\"
            End If
            intDirsFound = intDirsFound + 1 ' Increment the number of directories found by one
            ' Recursively call this function to search for matches in subdirectories.  When the recursed function
            ' completes, intDirsFound must be incremented.
            intDirsFound = intDirsFound + FindAllFiles(strSearchPath$ & strFile & "\", strPattern, colFiles, colDirs, blnDirsOnly)
        End If
        strFile = clsFind.FindNext()    ' Find the next file or directory
    Else    ' ... otherwise it must be a file.
        If Not blnDirsOnly Or blnBoth Then  ' If the caller wants files, then add them to the colFiles collection
         colFiles.Add strSearchPath$ & strFile
         If LesOptions.FilesToInclude = 0 Or LesOptions.FilesToInclude = 1 Then
            Lechemin = strSearchPath$
            attributs = clsFind.FileAttributes
            afficher = True
            If (attributs And FILE_ATTRIBUTE_READONLY) And LesOptions.ReadOnly = False Then
                afficher = False
            End If
            If (attributs And FILE_ATTRIBUTE_HIDDEN) And LesOptions.Hidden = False Then
                afficher = False
            End If
            If (attributs And FILE_ATTRIBUTE_SYSTEM) And LesOptions.System = False Then
                afficher = False
            End If
            chaine = ""
            If afficher = True Then
                If attributs And FILE_ATTRIBUTE_READONLY Then
                    chaine = "R"
                End If
                If attributs And FILE_ATTRIBUTE_HIDDEN Then
                    chaine = chaine + "H"
                End If
                If attributs And FILE_ATTRIBUTE_SYSTEM Then
                    chaine = chaine + "S"
                End If
                If attributs And FILE_ATTRIBUTE_ARCHIVE Then
                    chaine = chaine + "A"
                End If
                If chaine = "" Then
                    chaine = " "
                End If
                Set itmX = RENAME.ListView1.ListItems.Add(, , strSearchPath & strFile)
                itmX.Text = strSearchPath & strFile
                itmX.SubItems(1) = clsFind.FileSize
                itmX.SubItems(4) = "File"  ' Type fichier
                Select Case LesOptions.Dateformat
                    Case 0
                        itmX.SubItems(2) = clsFind.GetCreationDate
                    Case 1
                        itmX.SubItems(2) = clsFind.GetLastWriteDate
                    Case 2
                        itmX.SubItems(2) = clsFind.GetLastAccessDate
                End Select
                itmX.SubItems(3) = chaine
            End If
         End If
        End If
        strFile = clsFind.FindNext()    ' Find the next file or directory
    End If
Loop
FindAllFiles = intDirsFound ' Return the number of directories found
End Function
'*******************************************************************
'  PURPOSE: This returns just a path name from a full/partial path.
'  INPUTS:  sFileName - String Data to remove file from.
'  OUTPUTS: N/A
'  RETURNS: This function returns all the characters from left to the last
'           first \.  Does NOT check validity of the filename/Path....
'*******************************************************************
Function ExtractPath(sFilename As String) As String
Dim nIdx As Integer
Dim debut As Integer
debut = Len(sFilename)
    For nIdx = debut To 1 Step -1
       If Mid$(sFilename, nIdx, 1) = "\" Then
          ExtractPath = Mid$(sFilename, 1, nIdx)
          Exit Function
       End If
    Next
    ExtractPath = sFilename
End Function
Public Function FolderPart(lchemin As String) As String
 Dim zchemin As String
 zchemin = Trim$(ExtractPath(lchemin))
 If Len(zchemin) = 0 Then
  FolderPart = ""
  Exit Function
 End If
 
 zchemin = Replace(zchemin, ":", "")
 
 If Folder1 = 1 Then ' Use # of levels
  If Folder2 = 0 Then  ' From left
   zchemin = Left$(zchemin, At(zchemin, "\", Val(Folder5)))
  Else ' From right
   zchemin = StrReverse(zchemin)
   zchemin = Left$(zchemin, At(zchemin, "\", Val(Folder5) + 1))
   zchemin = StrReverse(zchemin)
  End If
 End If
 
 If Folder3 = 0 Then ' Delete all "\"
  zchemin = Replace(zchemin, "\", "")
 Else ' Replace all "\" with ...
  zchemin = Replace(zchemin, "\", Folder6)
 End If
 
 If Folder1 = 1 And Folder2 <> 0 And Folder3 <> 0 Then
    If Right$(zchemin, 1) = Folder6 Then
     zchemin = Mid$(zchemin, 2)
    End If
 End If
 
 FolderPart = Trim$(zchemin)
End Function
' Convertion d'une valeur double en binaire
Private Function bin$(valeur As Double)
 valeur = Abs(valeur)
 Dim X&
 Dim dummy#
 Dim max&
 Dim Rest#
 Dim Y&
 Dim rés$
 Dim resultat$
 resultat$ = ""
 
 X& = -1
 While dummy# <= valeur
  X& = X& + 1
  dummy# = 1 * 2 ^ X&
 Wend
 
 max& = X& - 1
 Rest# = valeur
 For Y& = max& To 0 Step -1
  Rest# = Rest# - (1 * 2 ^ Y&)
  If Rest# >= 0 Then rés$ = "1"
  If Rest# < 0 Then
   Rest# = Rest# + (1 * 2 ^ Y&)
   rés$ = "0"
  End If
  resultat$ = resultat$ + rés$
 Next
 If resultat$ = "" Then resultat$ = "0"
 bin$ = resultat$
End Function
Public Sub AjoutHistorique(repertoire As String)
 If VnbHistory < 51 Then
  If VnbHistory <> 0 Then
   Load RENAME.mnuhistory(VnbHistory)
  End If
  If Len(Trim$(repertoire)) > 0 Then
   RENAME.mnuhistory(VnbHistory).Caption = repertoire
   VnbHistory = VnbHistory + 1
  End If
 End If
End Sub
Public Sub CharInterdits(txt As String)
 Dim interdits As String
 Dim i As Integer
 Dim Lng As Integer
 Lng = Len(txt)
 interdits = "\/:*?<>|" + Chr$(34)
 For i = 1 To Lng
  If InStr(interdits, Mid$(txt, i, 1)) > 0 Then
   MsgBox "Warning, the following characters are not legal in a filename :" + vbCrLf + interdits, , "Warning !"
   Exit Sub
  End If
 Next
End Sub
' Renvoie vrai si la chaine contient des caractères interdits
Public Function ChInterdits(txt As String) As Boolean
 Dim interdits As String
 Dim i As Integer
 Dim Lng As Integer
 Dim retour As Boolean
 retour = False
 Lng = Len(txt)
 interdits = "\/:*?<>|" + Chr$(34)
 For i = 1 To Lng
  If InStr(interdits, Mid$(txt, i, 1)) > 0 Then
   ChInterdits = True
   Exit Function
  End If
 Next
ChInterdits = retour
End Function
' Fonction pour "Capitalizer"
Public Function MyStrConv(chaine As String) As String
Dim i As Integer
Dim vnb As Long
Dim vretour As String
Dim temoin As Boolean
Dim extrait As String
temoin = False
vretour = ""
vretour = UCase$(Mid$(chaine, 1, 1))
vnb = Len(chaine)
For i = 2 To vnb
 extrait = Mid$(chaine, i, 1)
 If temoin = True Then ' Il faut passer en majuscules, on est sur un nouveau mot
  extrait = UCase$(extrait)
  temoin = False
 Else ' il faut passer en minuscules
  extrait = LCase$(extrait)
 End If
 If InStr(LesOptions.WordsDelimiters, extrait) <> 0 Then
  temoin = True
 End If
 vretour = vretour + extrait
Next
MyStrConv = vretour
End Function
' Sauvegarde la liste des sélections cycliques
Public Sub SaveCyclic(laliste As ListBox)
Dim i As Integer
Dim vnb As Integer
ReDim LesCyclic(laliste.ListCount)
VnbCyclic = laliste.ListCount
vnb = laliste.ListCount
For i = 1 To vnb
    LesCyclic(i) = laliste.List(i - 1)
Next
If vnb = 0 Then
    UseCylcic = False
End If
End Sub
' Charge la liste des sélections cycliques
Public Sub LoadCyclic(laliste As ListBox)
 Dim i As Integer
 laliste.Clear
 For i = 1 To VnbCyclic
  laliste.AddItem LesCyclic(i)
 Next
End Sub
' supprime la liste de toutes les abbréviations
Public Sub RemoveAbbrev()
    Do While CollAbrev.Count > 0
        CollAbrev.Remove 1
    Loop
    OkUseAbbrev = False
End Sub

' Procédure chargée de lire un ficiher d'abréviation par défaut
Public Sub OpenAbbrev(fichier As String)
    Dim SIni As New cInifile
    Dim sValue As String
    Dim i As Integer
    Dim vnb As Integer
    Dim Signature As String
    Signature = "THE Rename's abbreviation file by Hervé Thouzard (hthouzard@bigfoot.com) - version 1.00"
    
    If Trim$(fichier) = "" Then
        Exit Sub
    End If
    
    If Not FileExists(fichier) Then
        MsgBox "Sorry you asked to open a default abbreviations file => " + fichier + " but this file does not exist...", vbOKOnly, "Warning"
        Exit Sub
    End If

    OkUseAbbrev = True
    ' Suppression du contenu initial de la collection
    Do While CollAbrev.Count > 0
        CollAbrev.Remove 1
    Loop
    ' Vérification de la signature
    With SIni
        .Path = fichier
        .Section = "General"
        .Key = "Signature"
        sValue = .Value
    End With
    If sValue <> Signature Then
        MsgBox "Sorry but the file => " + fichier + " is not an abbreviation's file coming from THE Rename !"
        Exit Sub
    End If
    ' Lecture du nombre d'abréviations
    With SIni
        .Path = fichier
        .Section = "General"
        .Key = "NumberOfAbbreviations"
        sValue = .Value
    End With
    If Val(sValue) = 0 Then
        MsgBox "Sorry, there's no abbreviations in the file " + fichier
        Exit Sub
    End If
    vnb = Val(sValue)

    ' Chargement dans la collection
    For i = 1 To vnb
        With SIni
            .Path = fichier
            .Section = "Abbreviations"
            .Key = "Abbrev" + Trim$(Str$(i))
            sValue = .Value
        End With
        CollAbrev.Add sValue, Str$(i)
    Next
End Sub
' Procédure chargée de lire un fichier de sélections cycliques par défaut
Public Sub OpenCyclic(fichier As String)
    Dim vligne As String, vnblig As Integer
    Dim ff As Integer

    If Trim$(fichier) = "" Then
        Exit Sub
    End If
    If Not FileExists(fichier) Then
        MsgBox "Sorry you asked to open a default cyclic file => " + fichier + " but this file does not exist...", vbOKOnly, "Warning"
        Exit Sub
    End If
    
    ff = FreeFile
    Open fichier For Input As #ff
    Line Input #ff, vligne
    While Not EOF(ff)
        vnblig = vnblig + 1
        Line Input #ff, vligne
    Wend
    If vligne <> "" Then
        vnblig = vnblig + 1
    End If
    Close #ff
    
    ReDim LesCyclic(vnblig)
    VnbCyclic = vnblig
    
    ff = FreeFile
    Open fichier For Input As #ff
    Line Input #ff, vligne
    vnblig = 0
    While Not EOF(ff)
        vnblig = vnblig + 1
        LesCyclic(vnblig) = vligne
        Line Input #ff, vligne
    Wend
    If vligne <> "" Then
        vnblig = vnblig + 1
        LesCyclic(vnblig) = vligne
    End If
    Close #ff
    UseCylcic = True
End Sub
' Renvoie un nombre en base 26 (notation excel)
Public Function FBase26(nombre As Long) As String
    Dim X As Long
    Dim Y As Long
    Dim z As Long
    Dim cH As String
    Dim b26(26) As String
    Dim i As Integer
    cH = ""
    For i = 1 To 26
        b26(i) = Chr$(64 + i)
    Next
    X = nombre
    If X = 0 Then
        FBase26 = ""
    Else
        While True
            z = X Mod 26
            Y = Int(X / 26)
            If z = 0 Then
                cH = "Z" + cH
                Y = Y - 1
            Else
                cH = b26(z) + cH
            End If
            If Y = 0 Then
                GoTo suite
            End If
            If Y < 26 Then
                cH = b26(Y) + cH
                GoTo suite
            End If
            X = Y
        Wend
    End If
suite:
    FBase26 = cH
End Function
' Détermine si un fichier existe déjà ou pas.
Public Function FileExists(sSource As String) As Boolean
   Dim wfd As WIN32_FIND_DATA
   Dim hFile As Long
   hFile = FindFirstFile(sSource, wfd)
   FileExists = hFile <> INVALID_HANDLE_VALUE
   Call FindClose(hFile)
End Function
' supprime d'une chaine les caractères interdits
Public Function RemIllegals(txt As String, Optional repertoire As Boolean = False) As String
 Dim interdits As String
 Dim i As Integer
 Dim Lng As Integer
 Dim zretour As String
 zretour = ""
 Lng = Len(txt)
 If repertoire = False Then
    interdits = "\/:*?<>|" + Chr$(34)
 Else
    interdits = "/*?<>|" + Chr$(34)
 End If
 For i = 1 To Lng
  If InStr(interdits, Mid$(txt, i, 1)) = 0 Then
   zretour = zretour + Mid$(txt, i, 1)
  End If
 Next
RemIllegals = zretour
End Function
Public Sub ErreurGrave(fonction As String)
 Dim vversion As String
 Dim cH As String
 vversion = App.Major & "." & App.Minor & "." & App.Revision
 cH = "THE Rename, version  " + vversion + ", ERROR Procedure " + fonction & " Error, description : " & Err.Description & ", Error Number : " & Err.Number & ", Error Source : " & Err.Source & ", ErrorDll=" & Err.LastDllError & ". ALL These information have been placed in the clipboard, please paste them in a text file and send me this text file to bug@herve-thouzard.com" & vbCrLf & "Tip, if you have some problems when you launch the program, launch it with the parameter /clean"
 Clipboard.Clear
 Clipboard.SetText cH
 MsgBox cH, vbCritical, "Error in THE Rename..."
End Sub
Public Function CoWbOyS(lachaine As String) As String
 Dim i As Integer
 Dim longueur As Long
 Dim extrait As String
 Dim cpt As Integer
 Dim vretour As String
 cpt = Int(Rnd() * 2)
 vretour = ""
 longueur = Len(lachaine)
 For i = 1 To longueur
    cpt = cpt + 1
    extrait = Mid$(lachaine, i, 1)
    If cpt = 1 Then ' Majuscules
        extrait = UCase$(extrait)
    Else ' minuscules
        extrait = LCase$(extrait)
        cpt = 0
    End If
    vretour = vretour + extrait
 Next
 CoWbOyS = vretour
End Function
' reformate une chaine de facon a ce qu'il n'y ait qu'un seul blanc entre les mots
Public Function RemoveMultipleSpacing(chaine As String) As String
 Dim i As Integer
 Dim longueur As Integer
 Dim cH As String
 Dim temoin As Boolean
 Dim extrait As String
 cH = ""
 longueur = Len(chaine)
 temoin = False
 For i = 1 To longueur
    extrait = Mid$(chaine, i, 1)
    If extrait = " " Then
        If temoin = True Then ' Ce n'est pas le premier blanc, on ne le laisse pas passer
        Else ' C'est le premier blanc, on le laisse passer
            temoin = True
            cH = cH + extrait
        End If
    Else
        If temoin = True Then ' on n'est plus sur un blanc, il n'y a donc plus de raison d'arrêter l'incrémentation
            temoin = False
        End If
        cH = cH + extrait
    End If
 Next
 RemoveMultipleSpacing = cH
End Function
' Sépare les mots d'une chaine en se basant sur les majuscules contenus dans la chaine
Public Function ExtractWords(ligne As String) As String
    Dim i As Integer
    Dim Lng As Integer
    Dim vout As String
    Lng = Len(ligne)
    vout = ""
    vout = UCase$(Left$(ligne, 1))
    For i = 2 To Lng
        If Mid$(ligne, i, 1) = UCase$(Mid$(ligne, i, 1)) And IsCharAlpha(Asc(Mid$(ligne, i, 1))) Then  ' on est sur une majuscule
            If i + 1 <= Lng Then
                If Mid$(ligne, i + 1, 1) <> UCase$(Mid$(ligne, i + 1, 1)) And IsCharAlpha(Asc(Mid$(ligne, i + 1, 1))) Then
                    vout = vout + " " + Mid$(ligne, i, 1)
                Else
                    vout = vout + Mid$(ligne, i, 1)
                End If
            Else
                vout = vout + " " + Mid$(ligne, i, 1)
            End If
        Else
            vout = vout + Mid$(ligne, i, 1)
        End If
    Next
    ExtractWords = vout
End Function
Private Function IsDigit(chaine As String) As Boolean
    IsDigit = IsNumeric(chaine)
End Function
Private Function IsSpace(chaine As String) As Boolean
 Dim Spaces As String
 Spaces = Chr$(9) + Chr$(10) + Chr$(11) + Chr$(12) + Chr$(13) + Chr$(32) + "0"
 If InStr(Spaces, chaine) = 0 Then
    IsSpace = False
 Else
    IsSpace = True
 End If
End Function
Private Function compare_right(a As String, B As String) As Integer
 Dim bias As Integer
 Dim pointerA As Integer
 Dim pointerB As Integer
 bias = 0
 pointerA = 1
 pointerB = 1
 While (True)
    If (Not IsDigit(Mid$(a, pointerA, 1)) And Not IsDigit(Mid$(B, pointerB, 1))) Then
        compare_right = bias
        Exit Function
    Else
        If (Not IsDigit(Mid$(a, pointerA, 1))) Then
            compare_right = -1
            Exit Function
        Else
            If (Not IsDigit(Mid$(B, pointerB, 1))) Then
                compare_right = 1
                Exit Function
            Else
                If (Mid$(a, pointerA, 1) < Mid$(B, pointerB, 1)) Then
                    If (Not bias) Then bias = -1
                Else
                    If Mid$(a, pointerA, 1) > Mid$(B, pointerB, 1) Then
                        If (Not bias) Then bias = 1
                    Else
                        If pointerA = Len(a) And pointerB = Len(B) Then
                            compare_right = bias
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If
    pointerA = pointerA + 1
    pointerB = pointerB + 1
    If pointerA > Len(a) Then
        GoTo suite
    End If
    If pointerB > Len(B) Then
        GoTo suite
    End If
 Wend
suite:
 compare_right = 0
End Function

Private Function compare_left(a As String, B As String) As Integer
 Dim pointerA As Integer
 Dim pointerB As Integer
 pointerA = 1
 pointerB = 1
 While (True)
    If (Not IsDigit(Mid$(a, pointerA, 1)) And Not IsDigit(Mid$(B, pointerB, 1))) Then
        compare_left = 0
        Exit Function
    Else
        If (Not IsDigit(Mid$(a, pointerA, 1))) Then
            compare_left = -1
            Exit Function
        Else
            If (Not IsDigit(Mid$(B, pointerB, 1))) Then
                compare_left = 1
                Exit Function
            Else
                If (Mid$(a, pointerA, 1) < Mid$(B, pointerB, 1)) Then
                    compare_left = -1
                    Exit Function
                Else
                    If (Mid$(a, pointerA, 1) > Mid$(B, pointerB, 1)) Then
                        compare_left = 1
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    pointerA = pointerA + 1
    pointerB = pointerB + 1
    If pointerA > Len(a) Then GoTo suite
    If pointerB > Len(B) Then GoTo suite
 Wend
suite:
 compare_left = 0
End Function

Private Function strnatcmp0(a As String, B As String, fold_case As Integer) As Integer
    Dim ai As Integer
    Dim bi As Integer
    Dim ca As String
    Dim cb As String
    Dim long1 As Integer
    Dim long2 As Integer
    a = LTrim$(a)
    B = LTrim$(B)
    long1 = Len(a)
    long2 = Len(B)
    
    Dim fractional As Integer
    Dim result As Integer
    ai = 1
    bi = 1
  
    If IsNumeric(a) And IsNumeric(B) Then
        If Val(a) < Val(B) Then
            strnatcmp0 = -1
            Exit Function
        Else
            If Val(a) = Val(B) Then
                strnatcmp0 = 0
                Exit Function
            Else
                If Val(a) > Val(B) Then
                    strnatcmp0 = 1
                    Exit Function
                End If
            End If
        End If
    End If
    
    While (True)
        ca = Mid$(a, ai, 1)
        cb = Mid$(B, bi, 1)
'        If ai > long1 Then
'            GoTo Suite0
'        End If
'        If bi > long2 Then
'            GoTo Suite0
'        End If
        
'        While IsSpace(ca)
'            ai = ai + 1
'            ca = Mid$(A, ai, 1)
'            If ai > long1 Then
'                GoTo Suite0
'            End If
'        Wend
'
'Suite0:
'
'        While (IsSpace(cb))
'            bi = bi + 1
'            cb = Mid$(B, bi, 1)
'            If bi > long2 Then
'                GoTo Suite1
'            End If
'        Wend
'
'Suite1:

        ' process run of digits
        If (IsDigit(ca) And IsDigit(cb)) Then
            fractional = (ca = "0" Or cb = "0")
            If fractional Then
                result = compare_left(Mid$(a, ai), Mid$(B, bi))
                If result <> 0 Then
                    strnatcmp0 = result
                    Exit Function
                End If
            Else
                result = compare_right(Mid$(a, ai), Mid$(B, bi))
                If result <> 0 Then
                    strnatcmp0 = result
                    Exit Function
                End If
            End If
        End If
        
        If (ai = Len(a) Or bi = Len(B)) Then
            strnatcmp0 = 0
            Exit Function
        End If
        
        If (fold_case) Then
            ca = UCase$(ca)
            cb = UCase$(cb)
        End If
        If (ca < cb) Then
            strnatcmp0 = -1
            Exit Function
        Else
            If (ca > cb) Then
                strnatcmp0 = 1
                Exit Function
            End If
        End If
        ai = ai + 1
        bi = bi + 1
        If ai > Len(a) Then GoTo suite
        If bi > Len(B) Then GoTo suite
    Wend
suite:
    
End Function
Private Function strnatcmp(a As String, B As String) As Integer
    strnatcmp = strnatcmp0(a, B, 0)
End Function
' Compare, recognizing numeric string and ignoring case.
Public Function strnatcasecmp(a As String, B As String) As Integer
    strnatcasecmp = strnatcmp0(a, B, 1)
End Function

Public Function LesDates(txtFileName As String, QDate As Integer, leformat As String) As String
   Dim Date1(3) As String
   Dim dDate1(3) As Date
   Dim dDate2(3) As Date
   Dim FicTmp As New clsFindFile
   FicTmp.Find txtFileName
    
    FicTmp.GetCreationDate dDate1(1), dDate2(1)
    dDate1(1) = dDate1(1) + dDate2(1)
    FicTmp.GetLastAccessDate dDate1(2), dDate2(2)
    dDate1(2) = dDate1(2) + dDate2(2)
    FicTmp.GetLastWriteDate dDate1(3), dDate2(3)
    dDate1(3) = dDate1(3) + dDate2(3)
    Set FicTmp = Nothing
        
    Date1(1) = Format$(dDate1(1), leformat)
    Date1(2) = Format$(dDate1(2), leformat)
    Date1(3) = Format$(dDate1(3), leformat)
    LesDates = Date1(QDate)
End Function

'Public Function GetFileDateString(CT As FILETIME, leformat As String) As String
'  Dim st As SYSTEMTIME
'  Dim ds As Single
'  If FileTimeToSystemTime(CT, st) Then
'        ds = DateSerial(st.wYear, st.wMonth, st.wDay)
'        GetFileDateString = Format$(ds, leformat)
'  Else
'    GetFileDateString = ""
'  End If
'End Function

' fonction de reformatage de nombres dans une chaine
' Par exemple on a en entrée File1 et on veut File001
Public Function ReformatNumbers(lachaine As String, TailleTot As Integer, CPad As String, Optional NumberToAdd As Integer = 0) As String
    Dim i As Integer
    Dim PosDeb As Integer
    Dim PosFin As Integer
    Dim longueur As Integer
    Dim vtmp As String
    Dim part1 As String
    Dim part2 As String
    Dim part3 As String
    longueur = Len(lachaine)
    If longueur = 0 Then
        ReformatNumbers = ""
        Exit Function
    End If
    PosDeb = 0
    ' On commence par chercher la position de début
    For i = 1 To longueur
        If IsNumeric(Mid$(lachaine, i, 1)) Then
            PosDeb = i
            i = longueur
        End If
    Next
    If PosDeb = 0 Then ' on n'a rien trouvé de numérique dans la chaine, on s'en va
        ReformatNumbers = lachaine
        Exit Function
    End If
    PosFin = PosDeb
    For i = PosDeb To longueur
        If Not IsNumeric(Mid$(lachaine, i, 1)) Then
            PosFin = i - 1
            i = longueur
        Else
            PosFin = i
        End If
    Next
    ' ensuite il ne reste plus qu'à compléter
    longueur = (PosFin - PosDeb) + 1
    If TailleTot - longueur >= 0 Then
'        If PosDeb <> 1 Then
            part1 = Left$(lachaine, PosDeb - 1)
            part3 = Trim$(Str$(Val(Mid$(lachaine, PosDeb)) + NumberToAdd))
            part2 = String$(TailleTot - Len(part3), CPad)
            vtmp = part1 + part2 + part3 + Mid$(lachaine, PosFin + 1)
'        Else
'            part3 = Trim$(Str$(Val(Mid$(lachaine, PosDeb)) + NumberToAdd))
'            part2 = String$(TailleTot - Len(part3), CPad)   ' Complément avec les zéros
'            vtmp = part2 + Mid$(lachaine, PosDeb)
'        End If
    Else
        vtmp = lachaine
    End If
    ReformatNumbers = vtmp
End Function

Public Function SeparateThousands(lachaine As String, CPad As String) As String
    Dim i As Integer
    Dim PosDeb As Integer
    Dim PosFin As Integer
    Dim longueur As Integer
    Dim vtmp As String
    Dim vtmp2 As String
    Dim vtmp3 As String
    Dim j As Integer
    longueur = Len(lachaine)
    If longueur = 0 Then
        SeparateThousands = ""
        Exit Function
    End If
    PosDeb = 0
    ' On commence par chercher la position de début
    For i = 1 To longueur
        If IsNumeric(Mid$(lachaine, i, 1)) Then
            PosDeb = i
            i = longueur
        End If
    Next
    If PosDeb = 0 Then ' on n'a rien trouvé de numérique dans la chaine, on s'en va
        SeparateThousands = lachaine
        Exit Function
    End If
    PosFin = PosDeb
    For i = PosDeb To longueur
        If Not IsNumeric(Mid$(lachaine, i, 1)) Then
            PosFin = i - 1
            i = longueur
        Else
            PosFin = i
        End If
    Next
    ' ensuite il ne reste plus qu'à compléter
    vtmp2 = Mid$(lachaine, PosDeb, (PosFin - PosDeb) + 1)
    vtmp2 = StrReverse(vtmp2)
    vtmp3 = ""
    longueur = Len(vtmp2)
    For i = 1 To longueur
        j = j + 1
        If j = 4 Then
            'If i <> longueur Then  ' On évite de placer un blanc inutilement à la fin d'une chaine, par exemple "001 "
                vtmp3 = vtmp3 + CPad
            'End If
            j = 0
        End If
        vtmp3 = vtmp3 + Mid$(vtmp2, i, 1)
    Next
    vtmp3 = StrReverse(vtmp3)
    vtmp = Left$(lachaine, PosDeb - 1) + vtmp3 + Mid$(lachaine, PosFin + 1)
    SeparateThousands = vtmp
End Function
Public Sub SelAll(zTextbox As Control)
On Error Resume Next
zTextbox.SelStart = 0
zTextbox.SelLength = 300
End Sub
Public Sub EtatHautBas(liste As ListView, cmdUp As CommandButton, cmdDown As CommandButton)
  Dim i As Integer
  i = liste.SelectedItem.Index
  cmdUp.Enabled = (i > 1)
  cmdDown.Enabled = ((i > -1) And (i < (liste.ListItems.Count)))    ' -1
End Sub
Public Sub SetListButtons(liste As ListBox, cmdUp As CommandButton, cmdDown As CommandButton)
  Dim i As Integer
  i = liste.ListIndex
  cmdUp.Enabled = (i > 0)
  cmdDown.Enabled = ((i > -1) And (i < (liste.ListCount - 1)))
End Sub
Public Sub ButtonUp(liste As ListBox)
  On Error Resume Next
  Dim nItem As Integer
  With liste
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = 0 Then Exit Sub  'can't move 1st item up
    .AddItem .Text, nItem - 1
    .RemoveItem nItem + 1
    .Selected(nItem - 1) = True
  End With
End Sub
Public Sub ButtonDown(liste As ListBox)
  On Error Resume Next
  Dim nItem As Integer
  With liste
    If .ListIndex < 0 Then Exit Sub
    nItem = .ListIndex
    If nItem = .ListCount - 1 Then Exit Sub 'can't move last item down
    'move item down
    .AddItem .Text, nItem + 2
    'remove old item
    .RemoveItem nItem
    'select the item that was just moved
    .Selected(nItem + 1) = True
  End With
End Sub
' Change la casse des tags MP3 et VQF selon les paramètres utilisateurs.
Public Function ChangeMP3Case(lachaine As String) As String
    Select Case LesOptions.Mp3VqfOpt3
        Case 0, -1 ' Keep cases
            ChangeMP3Case = lachaine
        Case 1  ' Upper cases
            ChangeMP3Case = UCase$(lachaine)
        Case 2  ' Lower cases
            ChangeMP3Case = LCase$(lachaine)
        Case 3  ' Capitalize first word only
            ChangeMP3Case = UCase$(Left$(lachaine, 1)) + LCase$(Mid$(lachaine, 2))
        Case 4  ' Capitalize all words
            ChangeMP3Case = MyStrConv(lachaine)
    End Select
End Function
' Ajoute un backslash à la fin d'une chaine, si c'est nécessaire
Public Function AddBackSlash(ByVal chaine As String) As String
    If Right$(chaine, 1) <> "\" Then
        chaine = chaine + "\"
    End If
    AddBackSlash = chaine
End Function

'Returns Roman numeral
' Note: this uses the 'maximum compression' method, so
' 1999 becomes MIM rather than MCMXCIX or MDCCCCLXXXXVIIII
' 0 returns a blank string, and modulus is taken of negative numbers
'
Public Function Roman(argNumber As Long) As String
    Dim WorkingNumber As Long
    WorkingNumber = argNumber
    
    If WorkingNumber < 0 Then WorkingNumber = WorkingNumber * -1
    If WorkingNumber = 0 Then Exit Function
    
    Const NumNumerals As Integer = 6
    Dim Numeral(NumNumerals) As String
    Dim Value(NumNumerals) As Long
    
    Numeral(0) = "I": Value(0) = "1"
    Numeral(1) = "V": Value(1) = "5"
    Numeral(2) = "X": Value(2) = "10"
    Numeral(3) = "L": Value(3) = "50"
    Numeral(4) = "C": Value(4) = "100"
    Numeral(5) = "D": Value(5) = "500"
    Numeral(6) = "M": Value(6) = "1000"
    
    Dim i As Integer
    Dim j As Integer
    Dim OutputString As String
    OutputString = ""
    Dim CombinedValue As Long
    Dim CombinedNumeral As String
    Dim FinishedLoop As Boolean
    
    Do
        For j = NumNumerals To 0 Step -1
            Do
                FinishedLoop = True
                If WorkingNumber >= Value(j) Then
                    WorkingNumber = WorkingNumber - Value(j)
                    OutputString = OutputString & Numeral(j)
                    FinishedLoop = False
                End If
            Loop Until FinishedLoop
            
            If (j > 0) Then
                For i = 0 To (Int((j - 1) / 2) * 2) Step 1
                    CombinedNumeral = Numeral(i) & Numeral(j)
                    CombinedValue = Value(j) - Value(i)
                    If WorkingNumber >= CombinedValue Then
                        If CombinedValue <> Value(i) Then
                            WorkingNumber = WorkingNumber - CombinedValue
                            OutputString = OutputString & CombinedNumeral
                        End If
                    End If
                Next
            End If
        Next
    Loop Until WorkingNumber = 0
    Roman = OutputString
End Function
Public Sub InsertTextInTextBox(TxtBox As Control, LstBox As ListBox)
 Dim letexte1 As String, letexte2 As String
 Dim vnewdeb As Integer
 If Len(Trim$(TxtBox.Text)) > 0 Then ' S'il y a déjà du texte
  letexte1 = Left$(TxtBox.Text, TxtBox.SelStart)
  letexte2 = Mid$(TxtBox.Text, TxtBox.SelStart + 1)
  If TxtBox.SelLength = Len(Trim$(TxtBox.Text)) Then ' si tout est sélectionné, tout effacer !
   TxtBox.Text = LstBox.List(LstBox.ListIndex)
   vnewdeb = Len(LstBox.List(LstBox.ListIndex))
  Else ' Tout n'est pas sélectionné, on insère
   If TxtBox.SelLength > 1 Then
    TxtBox.SelText = LstBox.List(LstBox.ListIndex)
   Else
    TxtBox.Text = letexte1 + LstBox.List(LstBox.ListIndex) + letexte2
    vnewdeb = Len(letexte1) + Len(LstBox.List(LstBox.ListIndex)) ' Postionnement du caret à la fin de ce qui vient d'être inséré
   End If
  End If
 Else ' Il n'y a pas de texte
  TxtBox.Text = LstBox.List(LstBox.ListIndex)
  vnewdeb = Len(LstBox.List(LstBox.ListIndex))
 End If
 TxtBox.SelStart = vnewdeb
 TxtBox.SetFocus
End Sub
Public Sub InsertTextInTextBoxFromText(TxtBox As Control, txt As String)
 Dim letexte1 As String, letexte2 As String
 Dim vnewdeb As Integer
 If Len(Trim$(TxtBox.Text)) > 0 Then ' S'il y a déjà du texte
  letexte1 = Left$(TxtBox.Text, TxtBox.SelStart)
  letexte2 = Mid$(TxtBox.Text, TxtBox.SelStart + 1)
  If TxtBox.SelLength = Len(Trim$(TxtBox.Text)) Then ' si tout est sélectionné, tout effacer !
   TxtBox.Text = txt
   vnewdeb = Len(txt)
  Else ' Tout n'est pas sélectionné, on insère
   If TxtBox.SelLength > 1 Then
    TxtBox.SelText = txt
   Else
    TxtBox.Text = letexte1 + txt + letexte2
    vnewdeb = Len(letexte1) + Len(txt) ' Postionnement du caret à la fin de ce qui vient d'être inséré
   End If
  End If
 Else ' Il n'y a pas de texte
  TxtBox.Text = txt
  vnewdeb = Len(txt)
 End If
 TxtBox.SelStart = vnewdeb
 TxtBox.SetFocus
End Sub
Public Sub InsertTextInTextBoxFromMenu(TxtBox As Control, MenuCaption As String)
 Dim letexte1 As String, letexte2 As String
 Dim vnewdeb As Integer
 If Len(Trim$(TxtBox.Text)) > 0 Then ' S'il y a déjà du texte
  letexte1 = Left$(TxtBox.Text, TxtBox.SelStart)
  letexte2 = Mid$(TxtBox.Text, TxtBox.SelStart + 1)
  If TxtBox.SelLength = Len(Trim$(TxtBox.Text)) Then ' si tout est sélectionné, tout effacer !
   TxtBox.Text = MenuCaption
   vnewdeb = Len(MenuCaption)
  Else ' Tout n'est pas sélectionné, on inssère
   TxtBox.Text = letexte1 + MenuCaption + letexte2
   vnewdeb = Len(letexte1) + Len(MenuCaption) ' Postionnement du caret à la fin de ce qui vient d'être insséré
  End If
 Else ' Il n'y a pas de texte
  TxtBox.Text = MenuCaption
  vnewdeb = Len(MenuCaption)
 End If
 TxtBox.SelStart = vnewdeb
End Sub
Sub Main()
    Splash.Show 1
    Unload Splash 'Fermeture de la splash screen
    RENAME.Show 'debut de l'appli
End Sub
Public Sub ChangeTab(KeyCode As Integer, Shift As Integer, TheTab As SSTab)
    Dim vnum As Integer
    Dim vrai As Boolean

    If KeyCode = 120 And Shift = 0 Then ' F9
        vrai = False
        vnum = TheTab.Tab + 1
        While Not vrai
            If vnum >= TheTab.Tabs Then
                vnum = 0
            End If
            If TheTab.TabVisible(vnum) Then
                TheTab.Tab = vnum
                vrai = True
            Else
                vnum = vnum + 1
            End If
        Wend
    End If
    
    If KeyCode = 120 And Shift = 1 Then ' Shift F9
        vrai = False
        vnum = TheTab.Tab - 1
        If vnum < 0 Then
            vnum = TheTab.Tabs - 1
        End If
    
        While Not vrai
            If vnum < 0 Then
                vnum = TheTab.Tabs - 1
            End If
            If TheTab.TabVisible(vnum) Then
                TheTab.Tab = vnum
                vrai = True
            Else
                vnum = vnum - 1
            End If
        Wend
    End If
End Sub
' Retaille la liste des tags automatiquement
Public Sub ResizeLvMp3()
    Dim colonne As ColumnHeader
    Set colonne = RENAME.LvMP3.ColumnHeaders.Item(1)
    AutoSizeColumnHeader RENAME.LvMP3, colonne, True
    Set colonne = RENAME.LvMP3.ColumnHeaders.Item(2)
    AutoSizeColumnHeader RENAME.LvMP3, colonne, True
End Sub
' ********************************************************************************************
' Déplace les fichers sélectionnés vers le haut
' ********************************************************************************************
Public Sub MoveFilesUp(lv As ListView, feuille As Form)
Dim vnb As Long, i As Long
Dim CollFiles As New Collection
Dim CollCheck As New Collection
Dim vret As Boolean
On Error Resume Next
lv.Sorted = False
i = 0
i = LVGetItemSelected(lv, -1)
While i <> -1
    vnb = vnb + 1
    CollFiles.Add i + 1, Str$(i + 1)
    LVSetItemNotSelected lv, i
    If lv.ListItems(i + 1).Checked Then
        CollCheck.Add "1", Str$(i + 1)
    Else
        CollCheck.Add "0", Str$(i + 1)
    End If
    i = LVGetItemSelected(lv, i)
Wend
If vnb = 0 Then
    Exit Sub
End If
feuille.MousePointer = vbHourglass
lv.Visible = False
For i = 1 To vnb
    vret = MoveRow(lv, CollFiles.Item(i), CollFiles.Item(i) - 1)
Next
For i = 1 To vnb
    LVSetItemSelected lv, CollFiles.Item(i) - 2
    If CollCheck.Item(i) = "1" Then
        lv.ListItems(CollFiles.Item(i) - 1).Checked = True
    End If
Next
lv.Visible = True
feuille.MousePointer = vbDefault
lv.SetFocus
End Sub

' ********************************************************************************************
' Déplace les fichers sélectionnés vers le bas
' ********************************************************************************************
Public Sub MoveFilesDown(lv As ListView, feuille As Form)
Dim vnb As Long, i As Long
Dim CollFiles As New Collection
Dim CollCheck As New Collection
Dim vret As Boolean
On Error Resume Next
lv.Sorted = False
i = 0
i = LVGetItemSelected(lv, -1)
While i <> -1
    vnb = vnb + 1
    CollFiles.Add i + 1, Str$(i + 1)
    LVSetItemNotSelected lv, i
    If lv.ListItems(i + 1).Checked Then
        CollCheck.Add "1", Str$(i + 1)
    Else
        CollCheck.Add "0", Str$(i + 1)
    End If
    i = LVGetItemSelected(lv, i)
Wend
If vnb = 0 Then
    Exit Sub
End If
feuille.MousePointer = vbHourglass
lv.Visible = False
For i = vnb To 1 Step -1
    vret = MoveRow(lv, CollFiles.Item(i), CollFiles.Item(i) + 1)
Next
For i = 1 To vnb
    LVSetItemSelected lv, CollFiles.Item(i)
    If CollCheck.Item(i) = "1" Then
        lv.ListItems(CollFiles.Item(i) + 1).Checked = True
    End If
Next
lv.Visible = True
feuille.MousePointer = vbDefault
lv.SetFocus
End Sub

