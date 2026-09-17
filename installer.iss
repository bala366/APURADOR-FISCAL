#define MyAppName "Apurador Fiscal"
#define MyAppVersion "5.0.0"
#define MyAppPublisher "Apurador Fiscal"
#define MyAppExeName "Apurador Fiscal.exe"

[Setup]
AppId={{D0A8D2B4-7A9C-4A23-9421-B7F82A74A310}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Apurador Fiscal
DefaultGroupName=Apurador Fiscal
OutputDir=installer_output
OutputBaseFilename=Apurador_Fiscal_Setup_V5
Compression=lzma
SolidCompression=yes
WizardStyle=modern
SetupIconFile=assets\app_icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
PrivilegesRequired=admin

[Files]
Source: "dist\Apurador Fiscal\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Apurador Fiscal"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Apurador Fiscal"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Criar atalho Apurador Fiscal na Area de Trabalho"; GroupDescription: "Atalhos:"; Flags: checkedonce

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir Apurador Fiscal"; Flags: nowait postinstall skipifsilent
