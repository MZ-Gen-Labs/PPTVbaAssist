; コマンドラインからバージョンが渡されなかった場合のデフォルト値
#ifndef MyAppVersion
#define MyAppVersion "1.0.0"
#endif

[Setup]
AppName=PPTVbaAssist
AppVersion={#MyAppVersion}
AppPublisher=Your Name or Organization
CreateAppDir=no
; 出力先ディレクトリを2つ上の階層の「Output」フォルダに指定
OutputDir=..\..\Output
OutputBaseFilename=PPTVbaAssist_Setup_{#MyAppVersion}
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
CloseApplications=yes

[Files]
; 2つ上の階層（プロジェクトルート）にあるファイルを指定
Source: "..\..\PPTVbaAssist.ppam"; DestDir: "{userappdata}\Microsoft\AddIns"; Flags: ignoreversion
Source: "..\..\src\UpdateMenu.ps1"; DestDir: "{userappdata}\Microsoft\AddIns"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\PowerPoint\AddIns\PPTVbaAssist"; Flags: uninsdeletekey
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\PowerPoint\AddIns\PPTVbaAssist"; ValueType: string; ValueName: "Path"; ValueData: "{userappdata}\Microsoft\AddIns\PPTVbaAssist.ppam"
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\PowerPoint\AddIns\PPTVbaAssist"; ValueType: dword; ValueName: "AutoLoad"; ValueData: "$ffffffff"

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"