; Script Inno Setup — Rfill v1.0.0

[Setup]
SourceDir=..
AppName=Rfill
AppVersion=1.2.9
AppPublisher=Jules FAGUET
AppPublisherURL=https://github.com/BigW0lf/Rfill
AppSupportURL=https://github.com/BigW0lf/Rfill/issues
AppUpdatesURL=https://github.com/BigW0lf/Rfill

; Pas besoin de droits administrateur — installe dans %LocalAppData%
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

DefaultDirName={autopf}\Rfill
DefaultGroupName=Rfill
AllowNoIcons=yes

OutputDir=installer
OutputBaseFilename=Rfill_Setup_1.2.9
SetupIconFile=Rfill.ico
UninstallDisplayIcon={app}\Rfill.exe

Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
WizardSizePercent=120

[Languages]
Name: "french"; MessagesFile: "compiler:Languages\French.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\Rfill.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Rfill";                        Filename: "{app}\Rfill.exe"
Name: "{group}\{cm:UninstallProgram,Rfill}";  Filename: "{uninstallexe}"
Name: "{autodesktop}\Rfill";                  Filename: "{app}\Rfill.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Rfill.exe"; Description: "{cm:LaunchProgram,Rfill}"; Flags: nowait postinstall skipifsilent
