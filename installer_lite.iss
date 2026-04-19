; Inno Setup 6 script
#define MyAppName "OrderMatcher-Lite"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "TomReb3l"
#define MyAppExeName "OrderMatcher-Lite.exe"
#define MySourceDir "dist\\OrderMatcher-Lite"

[Setup]
AppId={{3B17812A-4B89-43AF-8CB2-2A3107D6A101}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=OrderMatcher-Lite-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "greek"; MessagesFile: "compiler:Languages\Greek.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Δημιουργία συντόμευσης στην Επιφάνεια Εργασίας"; GroupDescription: "Πρόσθετες επιλογές:"

[Files]
Source: "{#MySourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Εκτέλεση του {#MyAppName}"; Flags: nowait postinstall skipifsilent
