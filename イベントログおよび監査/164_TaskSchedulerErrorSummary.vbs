' ファイル概要: タスクスケジューラのエラーイベントを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 164_TaskSchedulerErrorSummary.vbs
' 実行結果: タスクスケジューラのエラーイベントを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim filter, service, events, eventItem, count
Set filter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
filter.SetVarDate DateAdd("d", -7, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode, TimeGenerated を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Microsoft-Windows-TaskScheduler/Operational' " & _
    "AND (EventCode=101 OR EventCode=103 OR EventCode=203) " & _
    "AND TimeGenerated >= '" & filter.Value & "'")
count = 0
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode
    count = count + 1
Next
WScript.Echo "Task Scheduler error events in the last seven days: " & count
