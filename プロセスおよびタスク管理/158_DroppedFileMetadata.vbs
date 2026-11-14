' ファイル概要: ドロップファイルのメタデータ取得。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 158_DroppedFileMetadata.vbs <ファイルパス> [ファイルパス ...]
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: ドロップファイルのメタデータ取得処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim fso, i, path, file
Set fso=CreateObject("Scripting.FileSystemObject")
If WScript.Arguments.Count=0 Then WScript.Echo "Drop one or more files onto the script or pass file paths.": WScript.Quit 2
' [処理解説] 各引数をファイルパスとして確認し、存在するファイルのメタデータを表示します。
For i=0 To WScript.Arguments.Count-1
    path=WScript.Arguments(i)
    If fso.FileExists(path) Then
        Set file=fso.GetFile(path)
        WScript.Echo file.Path & " bytes=" & file.Size & " modified=" & file.DateLastModified
    Else
        WScript.Echo "Not a file: " & path
    End If
Next
