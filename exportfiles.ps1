Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 設定ファイルのパス
$configFilePath = "$env:USERPROFILE\file_processing_config.json"

# 設定を読み込む関数
function Load-Config {
    if (Test-Path $configFilePath) {
        return Get-Content -Path $configFilePath | ConvertFrom-Json
    }
    return @{ selectedFiles = @(); outputFilePath = "" }
}

# 設定を保存する関数
function Save-Config($config) {
    $config | ConvertTo-Json | Set-Content -Path $configFilePath
}

# 設定の読み込み
$config = Load-Config
$global:selectedFiles = $config.selectedFiles
$global:outputFilePath = $config.outputFilePath

# メインフォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "ファイル処理アプリケーション"
$form.Size = New-Object System.Drawing.Size(500, 300)

# ファイル選択ラベル
$fileLabel = New-Object System.Windows.Forms.Label
$fileLabel.Location = New-Object System.Drawing.Point(120, 20)
$fileLabel.Size = New-Object System.Drawing.Size(350, 60)
$fileLabel.Text = "選択されたファイル: " + ($global:selectedFiles -join ", ")
$form.Controls.Add($fileLabel)

# 出力先ラベル
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(120, 100)
$outputLabel.Size = New-Object System.Drawing.Size(350, 30)
$outputLabel.Text = "出力先: " + $global:outputFilePath
$form.Controls.Add($outputLabel)

# ファイル選択ボタン
$fileButton = New-Object System.Windows.Forms.Button
$fileButton.Location = New-Object System.Drawing.Point(10, 20)
$fileButton.Size = New-Object System.Drawing.Size(100, 30)
$fileButton.Text = "ファイル選択"
$fileButton.Add_Click({
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Multiselect = $true
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:selectedFiles = $fileDialog.FileNames
        # ファイルの存在確認
        $existingFiles = $global:selectedFiles | Where-Object { Test-Path $_ }
        if ($existingFiles.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("有効なファイルが選択されていません。")
            $fileLabel.Text = "選択されたファイル: なし"
        } else {
            $global:selectedFiles = $existingFiles
            $fileLabel.Text = "選択されたファイル: " + ($global:selectedFiles -join ", ")
        }
        # 設定を保存
        Save-Config @{ selectedFiles = $global:selectedFiles; outputFilePath = $global:outputFilePath }
    }
})
$form.Controls.Add($fileButton)

# 出力先選択ボタン
$outputButton = New-Object System.Windows.Forms.Button
$outputButton.Location = New-Object System.Drawing.Point(10, 60)
$outputButton.Size = New-Object System.Drawing.Size(100, 30)
$outputButton.Text = "出力先選択"
$outputButton.Add_Click({
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "All Files (*.*)|*.*"
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $global:outputFilePath = $saveFileDialog.FileName
        # 出力先フォルダの存在確認
        $outputFolder = [System.IO.Path]::GetDirectoryName($global:outputFilePath)
        if (-not (Test-Path $outputFolder)) {
            [System.Windows.Forms.MessageBox]::Show("出力先フォルダが存在しません。")
            $global:outputFilePath = ""
            $outputLabel.Text = "出力先: なし"
        } else {
            $outputLabel.Text = "出力先: $global:outputFilePath"
        }
        # 設定を保存
        Save-Config @{ selectedFiles = $global:selectedFiles; outputFilePath = $global:outputFilePath }
    }
})
$form.Controls.Add($outputButton)

# 処理実行ボタン
$processButton = New-Object System.Windows.Forms.Button
$processButton.Location = New-Object System.Drawing.Point(10, 100)
$processButton.Size = New-Object System.Drawing.Size(100, 30)
$processButton.Text = "処理実行"
$processButton.Add_Click({
    if ($global:selectedFiles.Count -gt 0 -and $global:outputFilePath) {
        try {
            foreach ($file in $global:selectedFiles) {
                # 区切り線とファイル名を追加
                Add-Content -Path $global:outputFilePath -Value "------------------------------------------------------------------------------------------------"
                Add-Content -Path $global:outputFilePath -Value "$(Split-Path -Leaf $file)"
                Add-Content -Path $global:outputFilePath -Value "------------------------------------------------------------------------------------------------"
               
                # ファイルの内容を追加
                Get-Content $file | Add-Content -Path $global:outputFilePath
            }
            [System.Windows.Forms.MessageBox]::Show("処理を実行しました。")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("エラーが発生しました: " + $_.Exception.Message)
        }
    } else {
        [System.Windows.Forms.MessageBox]::Show("ファイルと出力先を選択してください。")
    }
})
$form.Controls.Add($processButton)

# フォームの表示
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
