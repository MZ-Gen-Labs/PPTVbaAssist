param (
    [string]$pptFilePath,
    [string]$menuXmlPath
)

Write-Host "Starting menu update process..."

# パスの準備
$pptDir = Split-Path $pptFilePath
$tempZipPath = Join-Path $pptDir "temp_update.zip"

# ファイルのロック解除を待機しながらコピー（最大10秒）
$maxRetries = 10
$retryCount = 0
$fileCopied = $false

while (-not $fileCopied -and $retryCount -lt $maxRetries) {
    try {
        Start-Sleep -Seconds 1
        Copy-Item $pptFilePath $tempZipPath -Force -ErrorAction Stop
        $fileCopied = $true
    } catch {
        $retryCount++
        Write-Host "File is locked. Waiting... ($retryCount / $maxRetries)"
    }
}

if (-not $fileCopied) {
    Write-Host "Error: Could not unlock the PowerPoint file." -ForegroundColor Red
    Write-Host "`nPress Enter to exit..."
    Read-Host
    return
}

try {
    # ZIPファイルの更新処理 (.NETの機能を使用)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Add-Type -AssemblyName System.IO.Compression

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
    Start-Sleep -Seconds 1
    Copy-Item $tempZipPath $pptFilePath -Force
    Remove-Item $tempZipPath -Force

    Write-Host "Update completed successfully! Restarting PowerPoint." -ForegroundColor Green
    Start-Sleep -Seconds 1
    
    # 再度PowerPointでファイルを開く
    Invoke-Item $pptFilePath
    
    # ★正常終了時は10秒カウントダウンしてから自動で閉じる
    Write-Host "`nThis window will close automatically in 10 seconds..."
    for ($i = 10; $i -gt 0; $i--) {
        Write-Host "$i " -NoNewline
        Start-Sleep -Seconds 1
    }
    
} catch {
    Write-Host "An unexpected error occurred: $_" -ForegroundColor Red
    # ★エラー時はエンターキーが押されるまで画面をキープする
    Write-Host "`nPress Enter to exit..."
    Read-Host
}