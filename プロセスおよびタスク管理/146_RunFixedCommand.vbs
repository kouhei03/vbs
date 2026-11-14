' ファイル概要: 固定コマンドの同期/非同期呼び出し例。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 146_RunFixedCommand.vbs <sync|async>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 固定コマンドの同期/非同期呼び出し例処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, waitForExit, result, mode
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count<>1 Then WScript.Echo "Usage: cscript //nologo 146_RunFixedCommand.vbs <sync|async>": WScript.Quit 2
mode=LCase(WScript.Arguments(0))
If mode="sync" Then
    waitForExit=True
ElseIf mode="async" Then
    waitForExit=False
Else
    WScript.Echo "Mode must be sync or async."
    WScript.Quit 2
End If
Set shell=CreateObject("WScript.Shell")
result=shell.Run("cmd.exe /c ver",1,waitForExit)
If waitForExit Then WScript.Echo "Exit code=" & result Else WScript.Echo "Command launched asynchronously."
