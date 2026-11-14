' ファイル概要: コマンドライン引数表示デモ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 157_CliArgumentsDemo.vbs [引数 ...]
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: コマンドライン引数表示デモ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim i
WScript.Echo "Argument count=" & WScript.Arguments.Count
For i=0 To WScript.Arguments.Count-1
    WScript.Echo "Argument " & i & "=" & WScript.Arguments(i)
Next
