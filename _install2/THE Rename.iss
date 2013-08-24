[Setup]
PrivilegesRequired=poweruser
AppName=THE Rename
AppVerName=THE Rename 2.1.4
AppPublisher=Hervé Thouzard
AppPublisherURL=http://www.herve-thouzard.com/therename.phtml
AppSupportURL=http://www.herve-thouzard.com/therename.phtml
AppUpdatesURL=http://www.herve-thouzard.com/therename.phtml
DefaultDirName={pf}\THE Rename
DefaultGroupName=THE Rename
AllowNoIcons=yes
LicenseFile=licence.txt
DisableStartupPrompt=yes
WizardImageFile=logorv.bmp
Compression=bzip/9
OutputBaseFilename="reninst"

[Messages]
WelcomeLabel1=Welcome to [name] Setup Wizard
[Tasks]
Name: "desktopicon"; Description: "Create a &desktop icon"; GroupDescription: "Additional icons:"
Name: "quicklaunchicon"; Description: "Create a &Quick Launch icon"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
Source: "rename.exe"; DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: "THERENAME.HLP"; DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: "ccrpftv6.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace sharedfile regserver
Source: "ccrpDtp6.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace sharedfile regserver
Source: "ccrpftv6.tlb"; DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: "ISHF_Ex.tlb"; DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: "commands.ini"; DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: "Music.ini"; DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: "Rules.ini"; DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: "DBGWPROC.DLL"; DestDir: "{app}"; CopyMode: alwaysoverwrite; Flags: regserver
Source: "FILE_ID.DIZ"; DestDir: "{app}"; CopyMode: alwaysoverwrite
Source: "msvbvm60.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "oleaut32.dll"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "comcat.dll";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "ccrpUCW6.dll";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "MSCOMCTL.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace sharedfile regserver
Source: "TABCTL32.ocx"; DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace sharedfile regserver
Source: "Ssubtmr6.dll";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "ExifView.dll";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "therename.dll";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder;
Source: "renogg.dll";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder;
Source: "renMM.dll";   DestDir: "{sys}"; CopyMode: alwaysskipifsameorolder;

[Icons]
Name: "{group}\THE Rename"; Filename: "{app}\rename.exe"
Name: "{group}\Uninstall THE Rename"; Filename: "{uninstallexe}"
Name: "{group}\THE Rename Help"; Filename: "{app}\THERENAME.HLP"
Name: "{userdesktop}\THE Rename"; Filename: "{app}\rename.exe"; Tasks: desktopicon
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\THE Rename"; Filename: "{app}\rename.exe"; Tasks: quicklaunchicon

[Run]
Filename: "{app}\rename.exe"; Description: "Launch THE Rename"; Flags: nowait postinstall skipifsilent

[Registry]
Root: HKCR; Subkey: "Directory\Shell\THE Rename";  Flags: uninsdeletekey
Root: HKCR; Subkey: "Directory\Shell\THE Rename\command"; ValueName: ""; ValueType:string; ValueData: """{app}\rename.exe"" ""%1"""; Flags: uninsdeletekey
Root: HKCR; Subkey: "Drive\Shell\THE Rename"; Flags: uninsdeletekey
Root: HKCR; Subkey: "Drive\Shell\THE Rename\command"; ValueName: ""; ValueType:string; ValueData: """{app}\rename.exe"" ""%1"""; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Herve Thouzard\THE Rename"; ValueName: ""; ValueType:string; ValueData: "{app}";
Root: HKLM; Subkey: "Software\Microsoft\Windows\CurrentVersion\App Paths\rename.exe"; ValueName: ""; ValueType:string; ValueData: "{app}\rename.exe"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Windows\CurrentVersion\App Paths\rename.exe"; ValueName: "Path"; ValueType:string; ValueData: "{app}"; Flags: uninsdeletekey
