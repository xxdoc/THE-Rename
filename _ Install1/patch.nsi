; patch.nsi
!define VER_MAJOR 2
!define VER_MINOR 1.3

Name "THE Rename's patch"

OutFile "patch.exe"
BrandingText "THE Rename's patch installation"
Icon "Rename.ico"
BGGradient off

InstallDir "$PROGRAMFILES\THE Rename"
LicenseText "You must read the following license before installing"
LicenseData "licence.txt"
ComponentText "This will install THE Rename's patch v${VER_MAJOR}.${VER_MINOR} on your computer:"
Caption "THE Rename's patch v${VER_MAJOR}.${VER_MINOR} setup"
SubCaption 1 "THE Rename's patch v${VER_MAJOR}.${VER_MINOR}"
CRCCheck on
DirShow show
DirText "This will install THE Rename's patch v${VER_MAJOR}.${VER_MINOR} on your computer. Choose the directory where you have installed THE Rename"
AllowRootDirInstall false
InstType /NOCUSTOM
SetCompress auto
SetDatablockOptimize on


Section "MainSection"
  SetOverwrite on
  SetOutPath $INSTDIR
  File "rename.exe"
  File "THERENAME.HLP"
  File "ccrpUCW6.dll"
  RegDll "ccrpUCW6.dll"  
SectionEnd
