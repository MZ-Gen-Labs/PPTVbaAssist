Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 設定ファイルのパス
$configFilePath = "$env:USERPROFILE\file_processing_config.json"

# 設定を読み込む関数
function Load-Config {
    if (Test-Path $configFilePath) {
        $config = Get-Content -Path $configFilePath | ConvertFrom-Json
        # outputEncodingがオブジェクトの場合はBodyNameプロパティから文字列を取得
        if ($config.outputEncoding -is [PSCustomObject]) {
            if ($config.outputEncoding.BodyName) {
                $config.outputEncoding = $config.outputEncoding.BodyName.ToLower()
            } else {
                $config.outputEncoding = "shift_jis"
            }
        }
        return $config
    }
    return @{
        selectedFiles = @()
        outputFilePath = ""
        outputEncoding = "shift_jis"
    }
}

# 設定を保存する関数
function Save-Config($config) {
    $config | ConvertTo-Json | Set-Content -Path $configFilePath
}

# エンコード判定関数
function Get-FileEncoding($filePath) {
    $bomUtf8 = [byte[]]@(0xEF, 0xBB, 0xBF)
    $bomUtf16LE = [byte[]]@(0xFF, 0xFE)
    $bomUtf16BE = [byte[]]@(0xFE, 0xFF)
    $bomUtf32LE = [byte[]]@(0xFF, 0xFE, 0x00, 0x00)
    $bomUtf32BE = [byte[]]@(0x00, 0x00, 0xFE, 0xFF)

    $fs = [System.IO.File]::OpenRead($filePath)
    $bytes = New-Object byte[] 4
    $fs.Read($bytes, 0, 4) | Out-Null
    $fs.Close()

    if ($bytes.Length -ge 3 -and $bytes[0] -eq $bomUtf8[0] -and $bytes[1] -eq $bomUtf8[1] -and $bytes[2] -eq $bomUtf8[2]) {
        return "utf8"
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq $bomUtf16LE[0] -and $bytes[1] -eq $bomUtf16LE[1]) {
        return "unicode" # UTF-16 LE
    }
    elseif ($bytes.Length -ge 2 -and $bytes[0] -eq $bomUtf16BE[0] -and $bytes[1] -eq $bomUtf16BE[1]) {
        return "unicodebig"
    }
    elseif ($bytes.Length -ge 4 -and $bytes[0] -eq $bomUtf32LE[0] -and $bytes[1] -eq $bomUtf32LE[1] -and $bytes[2] -eq $bomUtf32LE[2] -and $bytes[3] -eq $bomUtf32LE[3]) {
        return "utf32"
    }
    elseif ($bytes.Length -ge 4 -and $bytes[0] -eq $bomUtf32BE[0] -and $bytes[1] -eq $bomUtf32BE[1] -and $bytes[2] -eq $bomUtf32BE[2] -and $bytes[3] -eq $bomUtf32BE[3]) {
        return "utf32BE"
    }
    else {
        return "shift_jis" # デフォルトエンコード
    }
}

# 設定の読み込み
$config = Load-Config
$global:selectedFiles = $config.selectedFiles
$global:outputFilePath = $config.outputFilePath
$global:outputEncodingName = $config.outputEncoding  # 変数名を変更

# メインフォームの作成
$form = New-Object System.Windows.Forms.Form
$form.Text = "ファイル処理アプリケーション"
$form.Size = New-Object System.Drawing.Size(600, 350)
$form.StartPosition = "CenterScreen"

# ファイル選択ラベル
$fileLabel = New-Object System.Windows.Forms.Label
$fileLabel.Location = New-Object System.Drawing.Point(150, 20)
$fileLabel.Size = New-Object System.Drawing.Size(420, 60)
$fileLabel.Text = "選択されたファイル: " + $(if ($global:selectedFiles.Count -gt 0) { ($global:selectedFiles -join ", ") } else { "なし" })
$form.Controls.Add($fileLabel)

# 出力先ラベル
$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Location = New-Object System.Drawing.Point(150, 100)
$outputLabel.Size = New-Object System.Drawing.Size(420, 30)
$outputLabel.Text = "出力先: " + $(if ($global:outputFilePath) { $global:outputFilePath } else { "なし" })
$form.Controls.Add($outputLabel)

# エンコード選択ラベル
$encodingLabel = New-Object System.Windows.Forms.Label
$encodingLabel.Location = New-Object System.Drawing.Point(150, 140)
$encodingLabel.Size = New-Object System.Drawing.Size(120, 30)
$encodingLabel.Text = "出力エンコード:"
$form.Controls.Add($encodingLabel)

# エンコード選択ドロップダウン
$encodingComboBox = New-Object System.Windows.Forms.ComboBox
$encodingComboBox.Location = New-Object System.Drawing.Point(280, 140)
$encodingComboBox.Size = New-Object System.Drawing.Size(150, 30)
$encodingComboBox.DropDownStyle = "DropDownList"
$encodingComboBox.Items.AddRange(@(
    "utf8",
    "utf8BOM",
    "unicode",      # UTF-16 LE
    "unicodebig",   # UTF-16 BE
    "utf32",
    "utf32BE",
    "shift_jis"
))
# デフォルト選択
if ($config.outputEncoding) {
    $encodingComboBox.SelectedItem = $config.outputEncoding
} else {
    $encodingComboBox.SelectedItem = "shift_jis"
}
$form.Controls.Add($encodingComboBox)

# ファイル選択ボタン
$fileButton = New-Object System.Windows.Forms.Button
$fileButton.Location = New-Object System.Drawing.Point(10, 20)
$fileButton.Size = New-Object System.Drawing.Size(130, 30)
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
        Save-Config @{
            selectedFiles = $global:selectedFiles
            outputFilePath = $global:outputFilePath
            outputEncoding = $global:outputEncodingName  # 変数名を変更
        }
    }
})
$form.Controls.Add($fileButton)

# 出力先選択ボタン
$outputButton = New-Object System.Windows.Forms.Button
$outputButton.Location = New-Object System.Drawing.Point(10, 60)
$outputButton.Size = New-Object System.Drawing.Size(130, 30)
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
        Save-Config @{
            selectedFiles = $global:selectedFiles
            outputFilePath = $global:outputFilePath
            outputEncoding = $global:outputEncodingName  # 変数名を変更
        }
    }
})
$form.Controls.Add($outputButton)

# 処理実行ボタン
$processButton = New-Object System.Windows.Forms.Button
$processButton.Location = New-Object System.Drawing.Point(10, 100)
$processButton.Size = New-Object System.Drawing.Size(130, 30)
$processButton.Text = "処理実行"
$processButton.Add_Click({
    if ($global:selectedFiles.Count -gt 0 -and $global:outputFilePath) {
        try {
            # 出力エンコードの取得
            $selectedEncoding = $encodingComboBox.SelectedItem
            switch ($selectedEncoding) {
                "utf8"      { $encoding = [System.Text.Encoding]::UTF8 }
                "utf8BOM"   { $encoding = New-Object System.Text.UTF8Encoding($true) }
                "unicode"   { $encoding = [System.Text.Encoding]::Unicode }       # UTF-16 LE
                "unicodebig"{ $encoding = [System.Text.Encoding]::BigEndianUnicode }
                "utf32"     { $encoding = [System.Text.Encoding]::UTF32 }
                "utf32BE"   { 
                    try {
                        $encoding = [System.Text.Encoding]::GetEncoding("utf-32BE")
                    } catch {
                        [System.Windows.Forms.MessageBox]::Show("指定されたエンコード 'utf32BE' はサポートされていません。デフォルトの 'shift_jis' を使用します。")
                        $encoding = [System.Text.Encoding]::GetEncoding("shift_jis")
                    }
                }
                "shift_jis" { $encoding = [System.Text.Encoding]::GetEncoding("shift_jis") }
                default     { $encoding = [System.Text.Encoding]::GetEncoding("shift_jis") }
            }

            $outputContent = ""

            foreach ($file in $global:selectedFiles) {
                # 区切り線とファイル名を追加
                $outputContent += "------------------------------------------------------------------------------------------------`r`n"
                $outputContent += "$(Split-Path -Leaf $file)`r`n"
                $outputContent += "------------------------------------------------------------------------------------------------`r`n"
               
                # エンコード判定
                $fileEncoding = Get-FileEncoding $file

                # ファイルの内容を読み込む
                $content = Get-Content -Path $file -Encoding $fileEncoding -Raw
                $outputContent += $content + "`r`n"
            }

            # 出力ファイルに書き込む
            [System.IO.File]::WriteAllText($global:outputFilePath, $outputContent, $encoding)

            [System.Windows.Forms.MessageBox]::Show("処理を実行しました。")
            
            # 設定を保存
            Save-Config @{
                selectedFiles = $global:selectedFiles
                outputFilePath = $global:outputFilePath
                outputEncoding = $selectedEncoding
            }

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
