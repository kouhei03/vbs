' ファイル概要: アクセス権を削除する対象を確認する（変更なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 188_AclRemovalPreview.vbs <file-or-folder> <principal>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: アクセス権を削除する対象を確認する（変更なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim targetPath, principal, fso
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 188_AclRemovalPreview.vbs <file-or-folder> <principal>"
    WScript.Quit 2
End If
targetPath = WScript.Arguments(0)
principal = WScript.Arguments(1)
If Left(targetPath, 2) = "\\" Then
    WScript.Echo "UNC paths are not allowed in this local-only preview."
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(targetPath) And Not fso.FolderExists(targetPath) Then
    WScript.Echo "Target path does not exist."
    WScript.Quit 1
End If
If Len(principal) = 0 Or Len(principal) > 256 Then
    WScript.Echo "Principal name length is invalid."
    WScript.Quit 2
End If
WScript.Echo "READ-ONLY ACL preview; no permission was removed or reset."
WScript.Echo "Target: " & targetPath
WScript.Echo "Principal to review: " & principal
