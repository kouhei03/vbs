' ファイル概要: 過去24時間のApplicationイベント1000と1001を表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 143_ApplicationErrorDetector.vbs
' 実行結果: 過去24時間のApplicationイベント1000と1001を表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, events, eventItem, filter, count
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("h", -24, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Application' AND (EventCode=1000 OR EventCode=1001) " & _
    "AND TimeGenerated >= '" & filter.Value & "'")

count = 0
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & vbTab & _
        eventItem.SourceName & vbTab & eventItem.RecordNumber
    count = count + 1
Next
WScript.Echo "Application error/crash events in the last 24 hours: " & count
WScript.Echo "Event message text was not read."
