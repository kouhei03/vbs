' ファイル概要: メモリ不足イベントの取得（ダンプなし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 098_MemoryExhaustionEvents.vbs
' 実行結果: メモリ不足イベントの取得（ダンプなし）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, events, eventRecord, count
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery("SELECT EventCode,TimeGenerated,RecordNumber FROM Win32_NTLogEvent WHERE Logfile='System' AND EventCode='2004'")
count = 0
For Each eventRecord In events
    WScript.Echo "Event 2004 Time=" & eventRecord.TimeGenerated & " Record=" & eventRecord.RecordNumber
    count = count + 1
    If count >= 20 Then Exit For
Next
If count = 0 Then WScript.Echo "No resource exhaustion event 2004 returned."
