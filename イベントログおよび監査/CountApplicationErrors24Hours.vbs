' ファイル概要: 過去24時間のアプリケーションエラーを集計する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo CountApplicationErrors24Hours.vbs
' 実行結果: 過去24時間のアプリケーションエラーを集計する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim timeFilter, service, events, eventItem, count
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
timeFilter.SetVarDate DateAdd("h", -24, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Application' AND (EventCode=1000 OR EventCode=1001) " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

count = 0
For Each eventItem In events
    count = count + 1
Next
WScript.Echo "Application crash/error events in the last 24 hours: " & count
