; コマンドラインからバージョンが渡されなかった場合のデフォルト値
#ifndef MyAppVersion
#define MyAppVersion "1.0.0"
#endif

[Setup]
AppId=PPTVbaAssist
AppName=PPTVbaAssist
AppVersion={#MyAppVersion}
AppPublisher=Your Name or Organization

UninstallDisplayName=PPTVbaAssist

; ★修正ポイント: アンインストーラーの保存先をユーザーのLocalAppDataに設定し、エラーを回避
DefaultDirName={localappdata}\PPTVbaAssist
; フォルダ選択画面を非表示にする（サイレントインストール風にするため）
DisableDirPage=yes

; 出力先ディレクトリを2つ上の階層の「Output」フォルダに指定
OutputDir=..\..\Output
OutputBaseFilename=PPTVbaAssist_Setup_{#MyAppVersion}
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
CloseApplications=yes

[Files]
; 2つ上の階層（プロジェクトルート）にあるファイルを指定 (改行せずに1行で記述)
Source: "..\..\PPTVbaAssist.ppam"; DestDir: "{userappdata}\Microsoft\AddIns"; Flags: ignoreversion overwritereadonly
Source: "..\..\src\UpdateMenu.ps1"; DestDir: "{userappdata}\Microsoft\AddIns"; Flags: ignoreversion overwritereadonly

[Registry]
; (改行せずに1行で記述)
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\PowerPoint\AddIns\PPTVbaAssist"; Flags: uninsdeletekey
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\PowerPoint\AddIns\PPTVbaAssist"; ValueType: string; ValueName: "Path"; ValueData: "{userappdata}\Microsoft\AddIns\PPTVbaAssist.ppam"
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\PowerPoint\AddIns\PPTVbaAssist"; ValueType: dword; ValueName: "AutoLoad"; ValueData: "$ffffffff"

[Languages]
Name: "japanese"; MessagesFile: "compiler:Languages\Japanese.isl"

[Code]
// インストール開始時に実行される関数
function InitializeSetup(): Boolean;
var
  WMIService: Variant;
  WbemLocator: Variant;
  WbemObjectSet: Variant;
begin
  Result := True;
  try
    // WMIを使用して現在起動中のプロセスをチェック
    WbemLocator := CreateOleObject('WbemScripting.SWbemLocator');
    WMIService := WbemLocator.ConnectServer('localhost', 'root\CIMV2');
    WbemObjectSet := WMIService.ExecQuery('SELECT * FROM Win32_Process WHERE Name="POWERPNT.EXE"');
    
    // PowerPointが起動している場合
    if WbemObjectSet.Count > 0 then
    begin
      MsgBox('PowerPointが起動しています。' + #13#10 + 'インストールを続行するには、PowerPointを完全に終了してから再度インストーラーを実行してください。', mbError, MB_OK);
      // インストールを中断する
      Result := False;
    end;
  except
    // エラー時は何もしない（そのまま続行）
  end;
end;