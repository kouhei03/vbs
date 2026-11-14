' ファイル概要: 共有名とローカルフォルダーを検証する（共有作成なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 192_CreateSharePreview.vbs <share-name> <existing-folder>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 共有名とローカルフォルダーを検証する（共有作成なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim shareName, folderPath, fso, expression
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 192_CreateSharePreview.vbs <share-name> <existing-folder>"
    WScript.Quit 2
End If
shareName = WScript.Arguments(0)
folderPath = WScript.Arguments(1)
If Left(folderPath, 2) = "\\" Then
    WScript.Echo "UNC paths are not allowed; select a local folder."
    WScript.Quit 2
End If
Set expression = New RegExp
expression.Pattern = "^[A-Za-z0-9][A-Za-z0-9._$ -]{0,79}$"
If Not expression.Test(shareName) Then
    WScript.Echo "Share name contains unsupported characters."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FolderExists(folderPath) Then
    WScript.Echo "Folder does not exist."
    WScript.Quit 1
End If
WScript.Echo "READ-ONLY share preview; no network share was created."
WScript.Echo "Proposed share name: " & shareName
WScript.Echo "Existing folder: " & folderPath
WScript.Echo "Access restrictions must be reviewed separately."
