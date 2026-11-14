' ファイル概要: ネットワークアダプターの状態変化に関するイベントを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetNetworkChangeEvents.vbs
' 実行結果: ネットワークアダプターの状態変化に関するイベントを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim timeFilter, service, events, eventItem
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
timeFilter.SetVarDate DateAdd("h", -24, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode, SourceName, TimeGenerated, RecordNumber を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM " & _
    "Win32_NTLogEvent WHERE Logfile='System' " & _
    "AND (EventCode=27 OR EventCode=4201) " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

WScript.Echo "Potential network adapter state changes (last 24 hours)"
WScript.Echo "Time" & vbTab & "Event ID" & vbTab & "Source" & vbTab & "Record"
For Each eventItem In events
    WScript.Echo EventTime(eventItem.TimeGenerated) & vbTab & _
        eventItem.EventCode & vbTab & eventItem.SourceName & vbTab & _
        eventItem.RecordNumber
Next

Function EventTime(cimTime)
    Dim dateTime
    Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
    dateTime.Value = cimTime
    EventTime = CStr(dateTime.GetVarDate(True))
End Function
