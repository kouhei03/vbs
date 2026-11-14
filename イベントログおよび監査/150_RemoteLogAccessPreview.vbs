' ファイル概要: 接続先名を検証する（ネットワーク接続なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 150_RemoteLogAccessPreview.vbs <computer-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 接続先名を検証する（ネットワーク接続なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim computerName, pattern
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 150_RemoteLogAccessPreview.vbs <computer-name>"
    WScript.Quit 2
End If
computerName = WScript.Arguments(0)
Set pattern = New RegExp
pattern.Pattern = "^[A-Za-z0-9][A-Za-z0-9.-]{0,252}$"
pattern.IgnoreCase = True
If Not pattern.Test(computerName) Then
    WScript.Echo "Computer name contains unsupported characters."
    WScript.Quit 2
End If
WScript.Echo "Target name accepted for planning: " & computerName
WScript.Echo "No network connection or remote event-log query was made."
