' ファイル概要: スクリプトの開始・終了時刻をコンソールに記録する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 156_RunLifecycleConsoleLog.vbs
' 実行結果: スクリプトの開始・終了時刻をコンソールに記録する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim startedAt, endedAt
startedAt = Now
WScript.Echo "Script started: " & CStr(startedAt)

' Place the script's normal work between these two lifecycle messages.

endedAt = Now
WScript.Echo "Script ended: " & CStr(endedAt)
WScript.Echo "This example writes to the console only; it creates no event source or event."
