' ファイル概要: スプーラーサービス状態照会。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 130_PrintSpoolerStatus.vbs
' 実行結果: スプーラーサービス状態照会処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, items, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set items = service.ExecQuery("SELECT Name,State,ProcessId,StartMode FROM Win32_Service WHERE Name='Spooler'")
If items.Count = 0 Then WScript.Echo "Spooler service not found.": WScript.Quit 1
For Each item In items
    WScript.Echo "State=" & item.State & " PID=" & item.ProcessId & " StartMode=" & item.StartMode
Next
