' ファイル概要: アクセス権を付与する対象と権限を確認する（変更なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 187_AclGrantPreview.vbs <file-or-folder> <principal> <Read|Write>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: アクセス権を付与する対象と権限を確認する（変更なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim targetPath, principal, rights, fso
If WScript.Arguments.Count <> 3 Then
    WScript.Echo "Usage: cscript //nologo 187_AclGrantPreview.vbs <file-or-folder> <principal> <Read|Write>"
    WScript.Quit 2
End If
targetPath = WScript.Arguments(0)
principal = WScript.Arguments(1)
rights = WScript.Arguments(2)
If Left(targetPath, 2) = "\\" Then
    WScript.Echo "UNC paths are not allowed in this local-only preview."
    WScript.Quit 2
End If
If StrComp(rights, "Read", vbTextCompare) <> 0 And _
   StrComp(rights, "Write", vbTextCompare) <> 0 Then
    WScript.Echo "Rights must be Read or Write."
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
WScript.Echo "READ-ONLY ACL preview; no permission was granted."
WScript.Echo "Target: " & targetPath
WScript.Echo "Principal: " & principal
WScript.Echo "Requested access: " & rights
