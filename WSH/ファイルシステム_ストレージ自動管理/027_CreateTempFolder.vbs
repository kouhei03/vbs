' 概要: WindowsのTEMP配下にランダム名の一意な作業フォルダーを作成します。自動削除はしません。
' 実行方法: cscript //nologo 027_CreateTempFolder.vbs
' 処理結果: TEMP配下の一意な作業フォルダー作成。処理結果や状態はコンソールに表示されます。
' コード解説: FileSystemObjectで対象パスを確認し、ファイルやフォルダーを処理します。
Option Explicit
Dim fso, tempRoot, path, attempts
' 主処理: TEMP配下の一意な作業フォルダー作成。
Set fso = CreateObject("Scripting.FileSystemObject")
tempRoot = fso.GetSpecialFolder(2).Path
For attempts = 1 To 10
    path = fso.BuildPath(tempRoot, "wsh_work_" & fso.GetTempName)
    If Not fso.FolderExists(path) And Not fso.FileExists(path) Then
        fso.CreateFolder path
        WScript.Echo path
        WScript.Quit 0
    End If
Next
WScript.Echo "Could not create a unique temporary folder."
WScript.Quit 1
