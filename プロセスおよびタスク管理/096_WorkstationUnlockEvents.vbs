' ファイル概要: 画面ロック解除イベント照会。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 096_WorkstationUnlockEvents.vbs
' 実行結果: 画面ロック解除イベント照会処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, events, eventRecord, query, count
Set service = GetObject("winmgmts:\\.\root\cimv2")
query = "SELECT EventCode,TimeGenerated,RecordNumber,SourceName FROM Win32_NTLogEvent WHERE Logfile='Security' AND EventCode='4801'"
Set events = service.ExecQuery(query)
count = 0
For Each eventRecord In events
    WScript.Echo "Event=" & eventRecord.EventCode & " Time=" & eventRecord.TimeGenerated & " Record=" & eventRecord.RecordNumber
    count = count + 1
    If count >= 20 Then Exit For
Next
If count = 0 Then WScript.Echo "No event returned; log permissions/audit policy may limit results."
