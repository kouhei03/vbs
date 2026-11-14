' ファイル概要: シャットダウン履歴照会のみ。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 164_ShutdownEventHistory.vbs
' 実行結果: シャットダウン履歴照会のみ処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, events, eventRecord, count
Set service=GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode,TimeGenerated,RecordNumber を取得します。
Set events=service.ExecQuery("SELECT EventCode,TimeGenerated,RecordNumber FROM Win32_NTLogEvent WHERE Logfile='System' AND (EventCode='1074' OR EventCode='6006' OR EventCode='6008')")
count=0
For Each eventRecord In events
    WScript.Echo "Event=" & eventRecord.EventCode & " Time=" & eventRecord.TimeGenerated & " Record=" & eventRecord.RecordNumber
    count=count+1
    If count>=20 Then Exit For
Next
If count=0 Then WScript.Echo "No shutdown-related event returned."
