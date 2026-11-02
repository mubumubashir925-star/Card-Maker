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
var
  SelectedAddonDir: string;

function FindAddonDirIn(const Root: string): string;
var
  FindRec: TFindRec;
  Base, Candidate: string;
begin
  Result := '';
  Base := AddBackslash(Root) + 'Corel';
  if not DirExists(Base) then Exit;
  if FindFirst(AddBackslash(Base) + '*', FindRec) then begin
    try
      repeat
        if (FindRec.Attributes and FILE_ATTRIBUTE_DIRECTORY <> 0) and
           (FindRec.Name <> '.') and (FindRec.Name <> '..') and
           (Pos('CorelDRAW Graphics Suite', FindRec.Name) = 1) then begin
          Candidate := AddBackslash(Base) + FindRec.Name + '\\Programs64';
          if FileExists(AddBackslash(Candidate) + 'CorelDRW.exe') then begin
            Result := AddBackslash(Candidate) + 'Addons\\CardMakerPro';
            Exit;
          end;
          Candidate := AddBackslash(Base) + FindRec.Name + '\\Programs';
          if FileExists(AddBackslash(Candidate) + 'CorelDRW.exe') then begin
            Result := AddBackslash(Candidate) + 'Addons\\CardMakerPro';
            Exit;
          end;
        end;
      until not FindNext(FindRec);
    finally
      FindClose(FindRec);
    end;
  end;
end;

function GetAddonDir(Param: string): string;
begin
  Result := SelectedAddonDir;
end;

function HasCorelAddonDir: Boolean;
begin
  Result := SelectedAddonDir <> '';
end;

function InitializeSetup: Boolean;
begin
  SelectedAddonDir := FindAddonDirIn(ExpandConstant('{pf}'));
  if SelectedAddonDir = '' then
    SelectedAddonDir := FindAddonDirIn(ExpandConstant('{pf32}'));
  Result := SelectedAddonDir <> '';
  if not Result then
    MsgBox('CorelDRAW was not found in standard installation folders. Setup stopped without changing CorelDRAW. This preview installer supports only a detected standard installation.', mbCriticalError, MB_OK);
end;
