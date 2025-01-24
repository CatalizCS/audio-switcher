; filepath: /c:/Users/tamaisme/Desktop/audio-switcher/installer.iss
#define MyAppName "Audio Switcher"
#define MyAppVersion "1.0"
#define MyAppPublisher "CatalizCSCatalizCS"
#define MyAppExeName "AudioSwitcher.exe"

[Setup]
AppId={{23E98B40-5B82-4671-B35F-BF241E289432}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
AllowNoIcons=yes
LicenseFile=LICENSE
OutputDir=installer
OutputBaseFilename=AudioSwitcher_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin
SetupIconFile=icon.ico

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startup"; Description: "Start with Windows"; GroupDescription: "Windows Startup"; Flags: unchecked

[Files]
Source: "dist\AudioSwitcher\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon
Name: "{userstartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"; Tasks: startup

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent runascurrentuser

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
begin
    if CurStep = ssPostInstall then
    begin
        // Set RUNASADMIN flag in registry
        RegWriteStringValue(HKEY_CURRENT_USER, 
            'Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers',
            ExpandConstant('{app}\{#MyAppExeName}'),
            'RUNASADMIN');
    end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
    if CurUninstallStep = usPostUninstall then
    begin
        // Remove RUNASADMIN flag from registry
        RegDeleteValue(HKEY_CURRENT_USER,
            'Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers',
            ExpandConstant('{app}\{#MyAppExeName}'));
    end;
end;