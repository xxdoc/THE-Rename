; rename.nsi
!define VER_MAJOR 2
!define VER_MINOR 1.3

Name "THE Rename"

OutFile "reninst.exe"
BrandingText "THE Rename installation"
Icon "Rename.ico"
BGGradient off

InstallDir "$PROGRAMFILES\THE Rename"
LicenseText "You must read the following license before installing"
LicenseData "licence.txt"
ComponentText "This will install THE Rename v${VER_MAJOR}.${VER_MINOR} on your computer:"
Caption "THE Rename v${VER_MAJOR}.${VER_MINOR} setup"
SubCaption 1 "THE Rename v${VER_MAJOR}.${VER_MINOR}"
CRCCheck on
DirShow show
DirText "This will install THE Rename v${VER_MAJOR}.${VER_MINOR} on your computer. Choose a directory"
AllowRootDirInstall false
InstType /NOCUSTOM
SetCompress auto
SetDatablockOptimize on


Section "MainSection"
  SetOverwrite on
  SetOutPath $INSTDIR
  File "rename.exe"
  File "THERENAME.HLP"
  File "ccrpftv6.ocx"
  RegDLL "ccrpftv6.ocx"
  File "ccrpDtp6.ocx"
  RegDLL "ccrpDtp6.ocx"  
  File "ccrpUCW6.dll"
  RegDll "ccrpUCW6.dll"  
  File "ccrpftv6.tlb"
  File "commands.ini"
  File "DBGWPROC.DLL"
  RegDLL "DBGWPROC.DLL"
  File "FILE_ID.DIZ"
  File "ISHF_Ex.tlb"
  File "Music.ini"
  File "Rules.ini"
  
  SetOverwrite ifnewer
  SetOutPath $SYSDIR
  File "comcat.dll"
  RegDLL "comcat.dll"
  File "MSCOMCTL.OCX"
  RegDLL "MSCOMCTL.OCX"
  SetOverwrite try
  File "msvbvm60.dll"
  File "oleaut32.dll"
  RegDLL "oleaut32.dll"  
  SetOverwrite ifnewer
  File "TABCTL32.OCX"
  RegDLL "TABCTL32.OCX"  
  File "Ssubtmr6.Dll"
  RegDll "Ssubtmr6.Dll"

  SetOverwrite on
  File "ExifView.dll"
  RegDll "ExifView.dll"
  File "therename.dll"
  File "renogg.dll"
  File "renMM.dll"
  SetOverwrite ifnewer

  ; Clés permettant d'ajouter THE Rename aux menus contextuels de l'explorateur (pour les répertoires et disques)
  WriteRegStr HKCR "Directory\Shell\THE Rename\command" "" "$INSTDIR\rename.exe %1"
  WriteRegStr HKCR "Drive\Shell\THE Rename\command" "" "$INSTDIR\rename.exe %1"
  
  ; Ma clé perso pour savoir (pour les mises à jour) où le programme a été installé
  WriteRegStr HKLM "Software\Herve Thouzard\THE Rename" "InstDir" "$INSTDIR"
  
  ; Clés permettant de faire, dans le menu "Démarrer/Executer" de windows "rename"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\App Paths\rename.exe" "" "$INSTDIR\rename.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\App Paths\rename.exe" "Path" "$INSTDIR"
  
  ; Clés permettant d'afficher le désinstallateur dans la fenêtre standard de Windows
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\THE Rename" "DisplayName" "THE Rename"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\THE Rename" "UninstallString" '"$INSTDIR\uninst.exe"'

  ; Création des raccourcis
  CreateDirectory "$SMPROGRAMS\THE Rename"
  CreateShortCut "$SMPROGRAMS\THE Rename\The Rename.lnk" "$INSTDIR\rename.exe"
  CreateShortCut "$SMPROGRAMS\THE Rename\THE Rename Help.lnk" "$INSTDIR\THERENAME.HLP"
  CreateShortCut "$SMPROGRAMS\THE Rename\Uninstall THE Rename.lnk" "$INSTDIR\uninst.exe"
  WriteUninstaller "uninst.exe"
SectionEnd

Section "Uninstall"
  Delete $INSTDIR\uninst.exe
  Delete $INSTDIR\*.exe
  Delete $INSTDIR\*.ocx
  Delete $INSTDIR\*.oca
  Delete $INSTDIR\*.dll
  Delete $INSTDIR\*.tlb
  Delete $INSTDIR\*.ini
  Delete $INSTDIR\*.diz
  Delete $INSTDIR\*.hlp
  Delete $INSTDIR\*.gid
  Delete "$SMPROGRAMS\THE Rename\The Rename.lnk"
  Delete "$SMPROGRAMS\THE Rename\THE Rename Help.lnk"
  Delete "$SMPROGRAMS\THE Rename\Uninstall THE Rename.lnk"
  RMDir $INSTDIR
  RMDir "$SMPROGRAMS\THE Rename"
  Delete $SYSDIR\therename.dll
  Delete $SYSDIR\renogg.dll
  DeleteRegKey HKEY_CLASSES_ROOT "Directory\Shell\THE Rename"
  DeleteRegKey HKEY_CLASSES_ROOT "Drive\Shell\THE Rename"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\App Paths\rename.exe"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\THE Rename"  
SectionEnd