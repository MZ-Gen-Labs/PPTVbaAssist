Attribute VB_Name = "FileIO"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public ribbon As IRibbonUI

' ★追加: 文字コード管理用のEnumと変数
Public Enum FileEncoding
    encShiftJIS = 1
    encUTF8 = 2
End Enum

Public g_CurrentEncoding As FileEncoding

Sub InitializeVariables()
    ' 初期値はShift-JISとする
    g_CurrentEncoding = encShiftJIS
End Sub

Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
    InitializeVariables
End Sub

' 指定したフォルダにファイルが存在するかどうかをチェックする関数
Function CheckFilesExist(folderPath As String) As Boolean
    Dim fileName As String

    If Dir(folderPath, vbDirectory) = "" Then
        CheckFilesExist = False
        Exit Function
    End If

    fileName = Dir(folderPath & "\*.*")
    Do While fileName <> ""
        Select Case LCase(Right(fileName, 4))
            Case ".bas", ".cls", ".frm"
                CheckFilesExist = True
                Exit Function
        End Select
        fileName = Dir
    Loop

    CheckFilesExist = False
End Function

' ==========================================
'  文字コード切替UI用コールバック
' ==========================================
Sub GetUtf8State(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = (g_CurrentEncoding = encUTF8)
End Sub

Sub ToggleUtf8State(control As IRibbonControl, pressed As Boolean)
    If pressed Then
        g_CurrentEncoding = encUTF8
        MsgBox "入出力を UTF-8 (BOMなし) に切り替えました。", vbInformation
    Else
        g_CurrentEncoding = encShiftJIS
        MsgBox "入出力を Shift-JIS に切り替えました。", vbInformation
    End If
    
    ' ★設定変更時にVS Codeの設定ファイルを自動更新
    UpdateVSCodeSettings g_CurrentEncoding
End Sub

' ==========================================
'  エクスポート処理
' ==========================================
Sub ExportCodeToFile(Optional control As IRibbonControl)
    ExportCodeToFile_
End Sub

Sub ExportCodeToFile_(Optional deleteUnusedFiles As Boolean = True, Optional outputFolder As String = "src")
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim folderPath As String
    Dim currentPresentationPath As String
    Dim exportedFiles As Object ' Scripting.Dictionary

    currentPresentationPath = ActivePresentation.FullName
    If currentPresentationPath = "" Then
        MsgBox "プレゼンテーションが保存されていません。先に保存してください。", vbExclamation
        Exit Sub
    End If

    folderPath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\" & outputFolder & "\"

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    Set exportedFiles = CreateObject("Scripting.Dictionary")

    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule:   exportPath = folderPath & vbComp.Name & ".bas"
            Case vbext_ct_ClassModule: exportPath = folderPath & vbComp.Name & ".cls"
            Case vbext_ct_MSForm:      exportPath = folderPath & vbComp.Name & ".frm"
            Case Else:                 exportPath = ""
        End Select

        If exportPath <> "" Then
            On Error Resume Next
            
            ' ★追加: 選択された文字コードに応じた出力
            If g_CurrentEncoding = encShiftJIS Then
                vbComp.Export exportPath
            ElseIf g_CurrentEncoding = encUTF8 Then
                Dim tempPath As String
                tempPath = exportPath & ".tmp"
                vbComp.Export tempPath
                If Err.Number = 0 Then
                    ConvertSJisToUtf8 tempPath, exportPath
                    Kill tempPath
                End If
            End If
            
            If Err.Number = 0 Then
                Debug.Print vbComp.Type & " がエクスポートされました: " & exportPath
                exportedFiles(exportPath) = True
            Else
                Debug.Print vbComp.Type & " のエクスポートに失敗しました: " & exportPath
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next vbComp

    ' 不要ファイル削除
    If deleteUnusedFiles Then
        Dim fileName As String
        fileName = Dir(folderPath & "*.*")
        Do While fileName <> ""
            Dim filePath As String
            filePath = folderPath & fileName
            If Not exportedFiles.Exists(filePath) Then
                Select Case LCase(Right(fileName, 4))
                    Case ".bas", ".cls", ".frm"
                        On Error Resume Next
                        Kill filePath
                        If Err.Number = 0 Then
                            Debug.Print "ファイルが削除されました: " & filePath
                        Else
                            Debug.Print "ファイルの削除に失敗しました: " & filePath
                            Err.Clear
                        End If
                        On Error GoTo 0
                End Select
            End If
            fileName = Dir
        Loop
    End If

    MsgBox "すべてのモジュールとフォームがエクスポートされました。" & vbCrLf & folderPath, vbInformation
    Set exportedFiles = Nothing
End Sub

' ==========================================
'  インポート処理
' ==========================================
Sub ImportCodeFromFile(Optional control As IRibbonControl)
    Dim folderPath As String
    Dim fileName As String
    Dim vbComp As VBIDE.VBComponent
    Dim moduleName As String
    Dim fileExtension As String
    Dim currentPresentationPath As String

    currentPresentationPath = ActivePresentation.FullName
    If currentPresentationPath = "" Then
        MsgBox "プレゼンテーションが保存されていません。先に保存してください。", vbExclamation
        Exit Sub
    End If

    folderPath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\src\"

    If Not CheckFilesExist(folderPath) Then
        MsgBox "インポートするファイルが見つかりません。", vbExclamation
        Exit Sub
    End If

    fileName = Dir(folderPath & "*.*")
    Do While fileName <> ""
        fileExtension = LCase(Right(fileName, 4))
        moduleName = Left(fileName, Len(fileName) - 4)

        If fileExtension = ".bas" Or fileExtension = ".cls" Or fileExtension = ".frm" Then
            Set vbComp = Nothing
            
            On Error Resume Next
            Set vbComp = Application.VBE.ActiveVBProject.VBComponents(moduleName)
            On Error GoTo 0

            ' 既存モジュールがある場合の処理（名前退避）
            If Not vbComp Is Nothing Then
                On Error Resume Next
                Application.VBE.ActiveVBProject.VBComponents.Remove Application.VBE.ActiveVBProject.VBComponents(moduleName & "_Old")
                vbComp.Name = moduleName & "_Old"
                Application.VBE.ActiveVBProject.VBComponents.Remove vbComp
                If Err.Number <> 0 Then
                    Debug.Print "モジュールの削除に失敗しました: " & moduleName
                    Err.Clear
                End If
                On Error GoTo 0
            End If

            ' ★追加: 選択された文字コードに応じた読込準備
            Dim targetFileToImport As String
            If g_CurrentEncoding = encShiftJIS Then
                targetFileToImport = folderPath & fileName
            ElseIf g_CurrentEncoding = encUTF8 Then
                Dim tempPathImport As String
                tempPathImport = folderPath & fileName & ".tmp"
                ConvertUtf8ToSJis folderPath & fileName, tempPathImport
                targetFileToImport = tempPathImport
            End If

            ' 新しいファイルのインポート
            On Error Resume Next
            Application.VBE.ActiveVBProject.VBComponents.Import targetFileToImport
            
            ' インポート後に一時ファイルを削除 (UTF-8モードの場合)
            If g_CurrentEncoding = encUTF8 And Dir(targetFileToImport) <> "" Then
                Kill targetFileToImport
            End If
            
            If Err.Number = 0 Then
              Debug.Print "モジュール/フォームがインポートされました: " & fileName
            Else
              Debug.Print "モジュール/フォームのインポートに失敗しました: " & fileName
              Err.Clear
            End If
            On Error GoTo 0
        End If

        fileName = Dir
    Loop

    MsgBox "すべてのモジュールとフォームがインポートされました。", vbInformation
End Sub


' ==========================================
'  アドイン化・メニュー更新関連（既存処理）
' ==========================================
Sub SaveAsAddin(Optional control As IRibbonControl)
    ' 既存の処理をそのまま保持
    Dim currentPres As Presentation
    Dim currentPath As String
    Dim addinPath As String
    Dim baseName As String
    Dim extPos As Integer
    Dim slashPos As Integer
    Dim localPath As String
    Dim parentFolder As String
    Dim fileNameOnly As String
    Dim distFolder As String
    Dim targetAddIn As AddIn
    Dim addInFileName As String
    Dim wasLoaded As Boolean
    
    Set currentPres = ActivePresentation
    If currentPres.Path = "" Then
        MsgBox "プレゼンテーションが保存されていません。先に保存してください。", vbExclamation
        Exit Sub
    End If
    
    currentPath = currentPres.FullName
    localPath = OneDriveUrlToLocalPath(currentPath)
    
    slashPos = InStrRev(localPath, "\")
    If slashPos > 0 Then
        parentFolder = Left(localPath, slashPos)
        fileNameOnly = Mid(localPath, slashPos + 1)
    Else
        parentFolder = localPath & "\"
        fileNameOnly = currentPres.Name
    End If
    
    extPos = InStrRev(fileNameOnly, ".")
    If extPos > 0 Then
        baseName = Left(fileNameOnly, extPos - 1)
    Else
        baseName = fileNameOnly
    End If
    
    distFolder = parentFolder & "dist"
    If Dir(distFolder, vbDirectory) = "" Then MkDir distFolder
    
    addinPath = distFolder & "\" & baseName & ".ppam"
    addInFileName = baseName & ".ppam"
    
    wasLoaded = False
    On Error Resume Next
    Set targetAddIn = Application.AddIns(addInFileName)
    If Not targetAddIn Is Nothing Then
        If targetAddIn.Loaded Then
            wasLoaded = True
            targetAddIn.Loaded = msoFalse
        End If
    End If
    On Error GoTo 0
    
    On Error Resume Next
    currentPres.SaveAs addinPath, ppSaveAsOpenXMLAddin
    If Err.Number = 0 Then
        MsgBox "アドインを dist フォルダに保存しました。" & vbCrLf & addinPath, vbInformation
        If wasLoaded And Not targetAddIn Is Nothing Then targetAddIn.Loaded = msoTrue
    Else
        MsgBox "保存に失敗しました。" & vbCrLf & "エラー: " & Err.Description, vbCritical
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Sub ExportMenuToFile(Optional control As IRibbonControl)
    ' 既存の処理をそのまま保持
    Dim currentPresentationPath As String
    Dim folderPath As String, tempZipPath As String
    Dim fso As Object, shellApp As Object, zipFolder As Object
    Dim customUIFolder As Object, uiFile As Object
    Dim vZipPath As Variant, vFolderPath As Variant
    
    currentPresentationPath = ActivePresentation.FullName
    If currentPresentationPath = "" Then
        MsgBox "プレゼンテーションが保存されていません。", vbExclamation
        Exit Sub
    End If
    
    folderPath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\src\"
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempZipPath = folderPath & "temp_for_export.zip"
    
    fso.CopyFile currentPresentationPath, tempZipPath, True
    
    vZipPath = tempZipPath
    vFolderPath = folderPath
    
    Set shellApp = CreateObject("Shell.Application")
    Set zipFolder = shellApp.Namespace(vZipPath)
    
    If zipFolder Is Nothing Then
        MsgBox "一時ZIPファイルの読み込みに失敗しました。", vbCritical
        If fso.FileExists(tempZipPath) Then fso.DeleteFile tempZipPath
        Exit Sub
    End If
    
    Set customUIFolder = zipFolder.ParseName("customUI")
    If Not customUIFolder Is Nothing Then
        Set uiFile = customUIFolder.GetFolder.ParseName("customUI14.xml")
        
        If Not uiFile Is Nothing Then
            shellApp.Namespace(vFolderPath).CopyHere uiFile, 4
            Sleep 1000
            
            If fso.FileExists(folderPath & "menu.xml") Then fso.DeleteFile folderPath & "menu.xml"
            
            If fso.FileExists(folderPath & "customUI14.xml") Then
                Name folderPath & "customUI14.xml" As folderPath & "menu.xml"
                MsgBox "メニューを src\menu.xml に抽出しました。", vbInformation
            Else
                MsgBox "ファイルの抽出に失敗しました。", vbExclamation
            End If
        Else
            MsgBox "ZIP内に customUI14.xml が見つかりません。", vbExclamation
        End If
    Else
        MsgBox "ZIP内に customUI フォルダが見つかりません。", vbExclamation
    End If
    
    If fso.FileExists(tempZipPath) Then fso.DeleteFile tempZipPath
End Sub

Sub ImportMenuFromFile(Optional control As IRibbonControl)
    ' 既存の処理をそのまま保持
    Dim currentPresentationPath As String
    Dim folderPath As String
    Dim menuXmlPath As String
    Dim psScriptPath As String
    Dim psCommand As String
    Dim shellApp As Object
    
    currentPresentationPath = ActivePresentation.FullName
    If currentPresentationPath = "" Then
        MsgBox "プレゼンテーションが保存されていません。", vbExclamation
        Exit Sub
    End If
    
    folderPath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\src\"
    menuXmlPath = folderPath & "menu.xml"
    psScriptPath = folderPath & "UpdateMenu.ps1"
    
    If Dir(psScriptPath) = "" Then
        psScriptPath = Environ("APPDATA") & "\Microsoft\AddIns\UpdateMenu.ps1"
    End If
    
    If Dir(menuXmlPath) = "" Then
        MsgBox "インポートする src\menu.xml が見つかりません。", vbCritical
        Exit Sub
    End If
    If Dir(psScriptPath) = "" Then
        MsgBox "実行用スクリプト UpdateMenu.ps1 が見つかりません。", vbCritical
        Exit Sub
    End If
    
    If MsgBox("メニューを更新するためにファイルを一度閉じます。よろしいですか？" & vbCrLf & _
              "※保存していない変更は失われます。事前に保存してください。", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    ActivePresentation.Save
    
    psCommand = "powershell.exe -ExecutionPolicy Bypass -File """ & psScriptPath & """ -pptFilePath """ & currentPresentationPath & """ -menuXmlPath """ & menuXmlPath & """"
    Set shellApp = CreateObject("WScript.Shell")
    shellApp.Run psCommand, 1, False
    ActivePresentation.Close
End Sub

Private Function GetTargetAddinName() As String
    ' 既存の処理をそのまま保持
    Dim baseName As String
    Dim extPos As Integer
    Dim currentName As String
    
    If Application.Presentations.Count = 0 Then
        GetTargetAddinName = ""
        Exit Function
    End If
    
    currentName = ActivePresentation.Name
    extPos = InStrRev(currentName, ".")
    If extPos > 0 Then
        baseName = Left(currentName, extPos - 1)
    Else
        baseName = currentName
    End If
    GetTargetAddinName = baseName & ".ppam"
End Function

Sub GetAddinState(control As IRibbonControl, ByRef returnedVal As Variant)
    ' 既存の処理をそのまま保持
    Dim targetAddIn As AddIn
    Dim addInName As String
    returnedVal = False
    addInName = GetTargetAddinName()
    If addInName = "" Then Exit Sub
    
    On Error Resume Next
    Set targetAddIn = Application.AddIns(addInName)
    If Not targetAddIn Is Nothing Then returnedVal = targetAddIn.Loaded
    On Error GoTo 0
End Sub

Sub ToggleAddinState(control As IRibbonControl, pressed As Boolean)
    ' 既存の処理をそのまま保持
    Dim targetAddIn As AddIn
    Dim addInName As String
    addInName = GetTargetAddinName()
    
    On Error Resume Next
    Set targetAddIn = Application.AddIns(addInName)
    If Not targetAddIn Is Nothing Then
        targetAddIn.Loaded = pressed
        If pressed Then
            MsgBox addInName & " を有効化しました。", vbInformation
        Else
            MsgBox addInName & " を無効化しました（開発モード）。", vbInformation
        End If
    Else
        MsgBox addInName & " がアドインとしてまだ登録されていません。", vbExclamation
        If Not ribbon Is Nothing Then ribbon.InvalidateControl control.Id
    End If
    On Error GoTo 0
End Sub

' ==========================================
'  文字コード変換 および 設定更新用の補助関数
' ==========================================

' Shift-JISのファイルを読み込み、BOMなしUTF-8で保存する
Private Sub ConvertSJisToUtf8(ByVal sourcePath As String, ByVal destPath As String)
    Dim streamSJIS As Object, streamUTF8 As Object, streamNoBOM As Object
    
    Set streamSJIS = CreateObject("ADODB.Stream")
    streamSJIS.Charset = "Shift_JIS"
    streamSJIS.Open
    streamSJIS.LoadFromFile sourcePath
    
    Set streamUTF8 = CreateObject("ADODB.Stream")
    streamUTF8.Charset = "UTF-8"
    streamUTF8.Open
    streamSJIS.CopyTo streamUTF8
    streamSJIS.Close
    
    streamUTF8.Position = 0
    streamUTF8.Type = 1 ' adTypeBinary
    streamUTF8.Position = 3 ' BOMスキップ
    
    Set streamNoBOM = CreateObject("ADODB.Stream")
    streamNoBOM.Type = 1 ' adTypeBinary
    streamNoBOM.Open
    streamUTF8.CopyTo streamNoBOM
    streamUTF8.Close
    
    streamNoBOM.SaveToFile destPath, 2 ' adSaveCreateOverWrite
    streamNoBOM.Close
    
    Set streamSJIS = Nothing
    Set streamUTF8 = Nothing
    Set streamNoBOM = Nothing
End Sub

' UTF-8のファイルを読み込み、Shift-JISで保存する
Private Sub ConvertUtf8ToSJis(ByVal sourcePath As String, ByVal destPath As String)
    Dim streamUTF8 As Object, streamSJIS As Object
    
    Set streamUTF8 = CreateObject("ADODB.Stream")
    streamUTF8.Charset = "UTF-8"
    streamUTF8.Open
    streamUTF8.LoadFromFile sourcePath
    
    Set streamSJIS = CreateObject("ADODB.Stream")
    streamSJIS.Charset = "Shift_JIS"
    streamSJIS.Open
    
    streamUTF8.CopyTo streamSJIS
    streamSJIS.SaveToFile destPath, 2 ' adSaveCreateOverWrite
    
    streamUTF8.Close
    streamSJIS.Close
    
    Set streamUTF8 = Nothing
    Set streamSJIS = Nothing
End Sub

' VS Codeの設定ファイル(settings.json)を自動更新する処理
Sub UpdateVSCodeSettings(ByVal encoding As FileEncoding)
    Dim rootPath As String, jsonPath As String
    Dim stream As Object, streamNoBOM As Object
    Dim jsonText As String, targetEncodingStr As String
    Dim regEx As Object
    
    ' プロジェクトのルートパスを取得
    rootPath = OneDriveUrlToLocalPath(ActivePresentation.Path)
    jsonPath = rootPath & "\.vscode\settings.json"
    
    ' settings.jsonが存在しない場合は何もしない
    If Dir(jsonPath) = "" Then Exit Sub
    
    If encoding = encShiftJIS Then
        targetEncodingStr = "shiftjis"
    ElseIf encoding = encUTF8 Then
        targetEncodingStr = "utf8"
    End If
    
    Set stream = CreateObject("ADODB.Stream")
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile jsonPath
    jsonText = stream.ReadText
    stream.Close
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = """files\.encoding""\s*:\s*""(shiftjis|utf8)"""
    
    jsonText = regEx.Replace(jsonText, """files.encoding"": """ & targetEncodingStr & """")
    
    stream.Open
    stream.WriteText jsonText
    
    stream.Position = 0
    stream.Type = 1 ' adTypeBinary
    stream.Position = 3 ' BOMスキップ
    
    Set streamNoBOM = CreateObject("ADODB.Stream")
    streamNoBOM.Type = 1 ' adTypeBinary
    streamNoBOM.Open
    stream.CopyTo streamNoBOM
    stream.Close
    
    streamNoBOM.SaveToFile jsonPath, 2
    streamNoBOM.Close
    
    Set stream = Nothing
    Set streamNoBOM = Nothing
    Set regEx = Nothing
    
    Debug.Print ".vscode/settings.json の文字コード設定を " & targetEncodingStr & " に更新しました。"
End Sub