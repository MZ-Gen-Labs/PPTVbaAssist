Attribute VB_Name = "UtilOneDrivePath"
Option Explicit

' ==============================================================================
' OneDriveのURLをローカルのファイルパスに変換する関数
' (フルスクラッチによる独自実装のため、ライセンスの制約なく自由に利用可能)
' ==============================================================================
Public Function OneDriveUrlToLocalPath(ByVal FilePath As String) As String
    ' "https://" から始まらない場合は、既にローカルパスとみなしてそのまま返す
    If InStr(1, FilePath, "https://", vbTextCompare) <> 1 Then
        OneDriveUrlToLocalPath = FilePath
        Exit Function
    End If

    Dim basePath As String
    Dim relativePath As String
    Dim searchPos As Long
    
    ' 1. 法人向け OneDrive (SharePoint) の場合
    If InStr(1, FilePath, "my.sharepoint.com", vbTextCompare) > 0 Then
        ' 環境変数からベースパスを取得（フォールバック付き）
        basePath = Environ("OneDriveCommercial")
        If basePath = "" Then basePath = Environ("OneDrive")
        
        ' "/Documents/" 以降の文字列を相対パスとして抽出
        searchPos = InStr(1, FilePath, "/Documents/", vbTextCompare)
        If searchPos > 0 Then
            relativePath = Mid(FilePath, searchPos + 11) ' "/Documents/" の文字数(11)を加算
        End If
        
    ' 2. 個人向け OneDrive の場合
    ElseIf InStr(1, FilePath, "d.docs.live.net", vbTextCompare) > 0 Then
        basePath = Environ("OneDriveConsumer")
        If basePath = "" Then basePath = Environ("OneDrive")
        
        ' "d.docs.live.net/" の後に続くユーザーIDの次の "/" を探す
        searchPos = InStr(1, FilePath, "d.docs.live.net/", vbTextCompare)
        If searchPos > 0 Then
            searchPos = InStr(searchPos + 16, FilePath, "/") ' "d.docs.live.net/" の文字数(16)を加算
            If searchPos > 0 Then
                relativePath = Mid(FilePath, searchPos + 1)
            End If
        End If
        
    ' 3. それ以外のURL形式の場合はそのまま返す
    Else
        OneDriveUrlToLocalPath = FilePath
        Exit Function
    End If
    
    ' ローカルパスとして結合し、区切り文字の "/" を "\" に置換して返す
    If relativePath <> "" Then
        OneDriveUrlToLocalPath = basePath & "\" & Replace(relativePath, "/", "\")
    Else
        OneDriveUrlToLocalPath = basePath
    End If
End Function