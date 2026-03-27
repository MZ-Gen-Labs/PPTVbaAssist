param (
    [string]$pptFilePath,
    [string]$menuXmlPath
)

# ファイルのロックが解除されるまで少し待機
Start-Sleep -Seconds 2

# パスの準備
$pptDir = Split-Path $pptFilePath
$tempZipPath = Join-Path $pptDir "temp_update.zip"

# コピーしてZIPとして扱う
Copy-Item $pptFilePath $tempZipPath -Force

# ZIPファイルの更新処理 (.NETの機能を使用)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zipArchive = [System.IO.Compression.ZipFile]::Open($tempZipPath, [System.IO.Compression.ZipArchiveMode]::Update)

# 既存の customUI/customUI14.xml を探して削除
$entry = $zipArchive.GetEntry("customUI/customUI14.xml")
if ($entry -ne $null) {
    $entry.Delete()
}

# src/menu.xml を customUI/customUI14.xml としてZIPに追加
[System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zipArchive, $menuXmlPath, "customUI/customUI14.xml")
$zipArchive.Dispose()

# 元のPPTMファイルを上書き
Copy-Item $tempZipPath $pptFilePath -Force
Remove-Item $tempZipPath -Force

# 再度PowerPointでファイルを開く
Invoke-Item $pptFilePath