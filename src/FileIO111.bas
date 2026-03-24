Attribute VB_Name = "FileIO111"
Option Explicit

Public ribbon As IRibbonUI

Sub InitializeVariables()
End Sub

Sub RibbonOnLoad(ribbonUI As IRibbonUI)
    Set ribbon = ribbonUI
    InitializeVariables
End Sub

' 指定したフォルダにファイルが存在するかどうかをチェックする関数
'  引数：folderPath - チェックするフォルダのパス
'  戻り値：ファイルが存在する場合はTrue、存在しない場合はFalse
Function CheckFilesExist(folderPath As String) As Boolean
    Dim fileName As String

    ' フォルダが存在しない場合はFalseを返す
    If Dir(folderPath, vbDirectory) = "" Then
        CheckFilesExist = False
        Exit Function
    End If

    ' フォルダ内のファイルをチェック（ワイルドカードを使用）
    fileName = Dir(folderPath & "\*.*")
    Do While fileName <> ""
        Select Case LCase(Right(fileName, 4))
            Case ".bas", ".cls", ".frm"
                CheckFilesExist = True
                Exit Function
        End Select
        fileName = Dir ' 次のファイルを取得
    Loop

    ' ファイルが存在しない場合はFalseを返す
    CheckFilesExist = False
End Function

Sub ExportCodeToFile()
    ExportCodeToFile_
End Sub

Sub ExportCodeToFile_(Optional deleteUnusedFiles As Boolean = True, Optional outputFolder As String = "src")
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim folderPath As String
    Dim currentPresentationPath As String
    Dim exportedFiles As Object ' Scripting.Dictionary

    ' 現在のアクティブなプレゼンテーションのパスを取得
    currentPresentationPath = ActivePresentation.FullName

    ' プレゼンテーションが保存されていない場合は、警告を表示して処理を終了
    If currentPresentationPath = "" Then
        MsgBox "プレゼンテーションが保存されていません。先に保存してください。", vbExclamation
        Exit Sub
    End If

    ' エクスポート先のフォルダパスを指定
    folderPath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\" & outputFolder & "\"

    ' 出力先フォルダが存在しない場合は作成
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If

    ' エクスポートされたファイルのリストを保持するためのDictionaryを作成
    Set exportedFiles = CreateObject("Scripting.Dictionary")

    ' 各標準モジュール、クラスモジュール、フォームをエクスポート
    For Each vbComp In Application.VBE.ActiveVBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule
                exportPath = folderPath & vbComp.Name & ".bas"
            Case vbext_ct_ClassModule
                exportPath = folderPath & vbComp.Name & ".cls"
            Case vbext_ct_MSForm
                exportPath = folderPath & vbComp.Name & ".frm"
            Case Else
                ' 他のタイプは何もしない
                exportPath = ""
        End Select

        If exportPath <> "" Then
            On Error Resume Next
            vbComp.Export exportPath
            If Err.Number = 0 Then
                Debug.Print vbComp.Type & " がエクスポートされました: " & exportPath
                ' エクスポートされたファイルのパスをDictionaryに追加
                exportedFiles(exportPath) = True
            Else
                Debug.Print vbComp.Type & " のエクスポートに失敗しました: " & exportPath
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next vbComp

    ' エクスポート先フォルダ内のファイルをチェックし、
    ' エクスポートされたファイルリストにないファイルを削除（オプション）
    If deleteUnusedFiles Then
        Dim fileName As String
        fileName = Dir(folderPath & "*.*")
        Do While fileName <> ""
            Dim filePath As String
            filePath = folderPath & fileName
            If Not exportedFiles.Exists(filePath) Then
                Select Case LCase(Right(fileName, 4))
                    Case ".bas", ".cls", ".frm"
                        ' ファイルを削除
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

    ' オブジェクトを解放
    Set exportedFiles = Nothing
End Sub

Sub ImportCodeFromFile()
    Dim folderPath As String
    Dim fileName As String
    Dim vbComp As VBIDE.VBComponent
    Dim moduleName As String
    Dim fileExtension As String
    Dim currentPresentationPath As String

    ' 現在のアクティブなプレゼンテーションのパスを取得
    currentPresentationPath = ActivePresentation.FullName

    ' プレゼンテーションが保存されていない場合は、警告を表示して処理を終了
    If currentPresentationPath = "" Then
        MsgBox "プレゼンテーションが保存されていません。先に保存してください。", vbExclamation
        Exit Sub
    End If

    ' インポート元のフォルダパスを指定
    folderPath = OneDriveUrlToLocalPath(ActivePresentation.Path) & "\src\"

    ' インポートするファイルが存在するかどうかをチェック
    If Not CheckFilesExist(folderPath) Then
        MsgBox "インポートするファイルが見つかりません。", vbExclamation
        Exit Sub
    End If

    ' 標準モジュール、クラスモジュール、フォームのインポート
    fileName = Dir(folderPath & "*.*")
    Do While fileName <> ""
        fileExtension = LCase(Right(fileName, 4))
        moduleName = Left(fileName, Len(fileName) - 4)

        If fileExtension = ".bas" Or fileExtension = ".cls" Or fileExtension = ".frm" Then
            On Error Resume Next
            Set vbComp = Application.VBE.ActiveVBProject.VBComponents(moduleName)
            If Not vbComp Is Nothing Then
                Application.VBE.ActiveVBProject.VBComponents.Remove vbComp
            End If

            If Err.Number <> 0 Then
                Debug.Print "モジュールの削除に失敗しました: " & moduleName
                Err.Clear
            End If

            Application.VBE.ActiveVBProject.VBComponents.Import folderPath & fileName
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




' Onedriveフォルダ取得関数
' https://kuroihako.com/vba/onedriveurltolocalpath/
' パワーポイント用に以下のみ修正
'        PathSeparator = "/"
'        ' パワーポイントでは以下の処理がないためハードコード
'        ' PathSeparator = Application.PathSeparator


' [VBA]OneDriveで同期しているファイルまたはフォルダのURLをローカルパスに変換する関数
' Copyright (c) 2020-2023  黒箱
' This software is released under the GPLv3.
' このソフトウェアはGNU GPLv3の下でリリースされています。

'* @fn Public Function OneDriveUrlToLocalPath(ByRef Url As String) As String
'* @brief OneDriveのファイルURL又はフォルダURLをローカルパスに変換します。
'* @param[in] Url OneDrive内に保存されたのファイル又はフォルダのURL
'* @return Variant ローカルパスを返します。引数Urlにローカルパスに"https://"以外から始まる文字列を指定した場合、引数Urlを返します。
'* @details OneDriveのファイルURL又はフォルダURLをローカルパスに変換します。本関数は、ExcelブックがOneDrive内に格納されている場合に、Workbook.Path又はWorkbook.FullNameがURLを返す問題を解決するためのものです。
'*
Public Function OneDriveUrlToLocalPath(ByRef url As String) As String
Const OneDriveCommercialUrlPattern As String = "*my.sharepoint.com*" '法人向けOneDriveのURLか否かを判定するためのLike右辺値

    '引数がURLでない場合、引数はローカルパスと判断してそのまま返す。
    If Not (url Like "https://*") Then
        OneDriveUrlToLocalPath = url
        Exit Function
    End If
    
    'OneDriveのパスを取得しておく(パフォーマンス優先)。
    Static PathSeparator As String
    Static OneDriveCommercialPath As String
    Static OneDriveConsumerPath As String
    
    If (PathSeparator = "") Then
        PathSeparator = "/"
        ' パワーポイントでは以下の処理がないためハードコード
        ' PathSeparator = Application.PathSeparator
        
        '法人向けOneDrive(OneDrive for Business)のパス
        OneDriveCommercialPath = Environ("OneDriveCommercial")
        If (OneDriveCommercialPath = "") Then OneDriveCommercialPath = Environ("OneDrive")
        
        '個人向けOneDriveのパス
        OneDriveConsumerPath = Environ("OneDriveConsumer")
        If (OneDriveConsumerPath = "") Then OneDriveConsumerPath = Environ("OneDrive")

    End If
    
    '法人向けOneDrive：URL＝"https://会社名-my.sharepoint.com/personal/ユーザー名_domain_com/Documentsファイルパス")
    Dim FilePathPos As Long
    If (url Like OneDriveCommercialUrlPattern) Then
        FilePathPos = InStr(1, url, "/Documents") + 10 '10 = Len("/Documents")
        OneDriveUrlToLocalPath = OneDriveCommercialPath & Replace(Mid(url, FilePathPos), "/", PathSeparator)
        
    '個人向けOneDrive：URL＝"https://d.docs.live.net/CID番号/ファイルパス"
    Else
        FilePathPos = InStr(9, url, "/") '9 == Len("https://") + 1
        FilePathPos = InStr(FilePathPos + 1, url, "/")

        If (FilePathPos = 0) Then
            OneDriveUrlToLocalPath = OneDriveConsumerPath
        Else
            OneDriveUrlToLocalPath = OneDriveConsumerPath & Replace(Mid(url, FilePathPos), "/", PathSeparator)
        End If
    End If

End Function

