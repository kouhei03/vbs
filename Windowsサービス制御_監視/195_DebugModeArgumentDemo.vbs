' ファイル概要: 詳細表示引数の安全な解析例。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 195_DebugModeArgumentDemo.vbs [/verbose]
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 詳細表示引数の安全な解析例処理の結果やエラーをコンソールに表示します。

' /verbose引数の解釈例です。サービス構成や状態は変更しません。
Option Explicit
Dim debugMode, argument
 debugMode = False
For Each argument In WScript.Arguments
    If LCase(argument) = "/verbose" Then debugMode = True
Next
If debugMode Then
    WScript.Echo "Verbose mode enabled for this local read-only demo."
Else
    WScript.Echo "Normal mode. Pass /verbose for an extra diagnostic line."
End If
WScript.Echo "No service operation was performed."
