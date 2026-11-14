' ファイル概要: ログオン・ログオフイベントの件数を集計する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetLogonEventCounts.vbs
' 実行結果: ログオン・ログオフイベントの件数を集計する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim timeFilter, service, events, eventItem, logonCount, logoffCount
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("h", -24, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent WHERE Logfile='Security' " & _
    "AND (EventCode=4624 OR EventCode=4634) " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

logonCount = 0
logoffCount = 0
For Each eventItem In events
    If eventItem.EventCode = 4624 Then
        logonCount = logonCount + 1
    ElseIf eventItem.EventCode = 4634 Then
        logoffCount = logoffCount + 1
    End If
Next
WScript.Echo "Successful logon events (4624), last 24 hours: " & logonCount
WScript.Echo "Logoff events (4634), last 24 hours: " & logoffCount
WScript.Echo "Counts only; account, session, and network details are omitted."
