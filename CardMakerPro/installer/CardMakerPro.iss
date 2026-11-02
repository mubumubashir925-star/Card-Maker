; Card Maker Pro preview installer.
; IMPORTANT: This installs the preview assemblies and an optional VBA bootstrap,
; but DOES NOT yet register a visible CorelDRAW docker. Do not label as production.
#define MyAppName "Card Maker Pro (Developer Preview)"
#define MyAppVersion "0.1.0-preview"
[Setup]
AppId={{3C7DCD0A-7C1E-4AA5-9C15-DAA933A9E29A}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\Card Maker Pro
DefaultGroupName=Card Maker Pro
OutputDir=output
OutputBaseFilename=CardMakerPro-Preview-Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
ArchitecturesAllowed=x86 x64
UninstallDisplayIcon={app}\CardMakerPro.Host.dll
WizardStyle=modern
DisableProgramGroupPage=yes
[Files]
Source: "staging\CardMakerPro.Host.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "staging\CardMakerPro.UI.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "CorelDockerBootstrap.bas"; DestDir: "{app}\Developer"; Flags: ignoreversion
Source: "TEST_DOCKER.md"; DestDir: "{app}\Developer"; Flags: ignoreversion
Source: "INSTALLATION.md"; DestDir: "{app}\Developer"; Flags: ignoreversion
[Icons]
Name: "{group}\Uninstall Card Maker Pro Preview"; Filename: "{uninstallexe}"
[Code]
function IsCorelDrawInstalled: Boolean;
var
  Names: TArrayOfString;
  I: Integer;
begin
  Result := False;
  { Only warn; CorelDRAW can be installed later or in nonstandard directories. }
  if RegGetSubkeyNames(HKLM, 'SOFTWARE\Corel', Names) then
    for I := 0 to GetArrayLength(Names)-1 do
      if Pos('CorelDRAW', Names[I]) > 0 then Result := True;
  if RegGetSubkeyNames(HKLM32, 'SOFTWARE\Corel', Names) then
    for I := 0 to GetArrayLength(Names)-1 do
      if Pos('CorelDRAW', Names[I]) > 0 then Result := True;
end;

function InitializeSetup: Boolean;
begin
  Result := True;
  if not IsCorelDrawInstalled then
    MsgBox('CorelDRAW installation was not detected automatically. This preview only copies development DLLs and does not install a working docker.', mbInformation, MB_OK);
end;
