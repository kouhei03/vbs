' ファイル概要: verboseモード切替デモ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 162_VerboseModeDemo.vbs [/verbose]
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: verboseモード切替デモ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim verbose, i
verbose=False
For i=0 To WScript.Arguments.Count-1
    If LCase(WScript.Arguments(i))="/verbose" Then verbose=True
Next
WScript.Echo "Process diagnostics started."
If verbose Then
    WScript.Echo "Host=" & CreateObject("WScript.Network").ComputerName
    WScript.Echo "Timestamp=" & Now
    WScript.Echo "Arguments=" & WScript.Arguments.Count
End If
