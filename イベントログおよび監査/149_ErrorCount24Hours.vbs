' ファイル概要: 過去24時間のシステムとアプリケーションのエラーを集計する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 149_ErrorCount24Hours.vbs
' 実行結果: 過去24時間のシステムとアプリケーションのエラーを集計する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, events, eventItem, filter, count
Set filter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
filter.SetVarDate DateAdd("h", -24, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode FROM Win32_NTLogEvent " & _
    "WHERE (Logfile='System' OR Logfile='Application') " & _
    "AND Type='Error' AND TimeGenerated >= '" & filter.Value & "'")
count = 0
For Each eventItem In events
    count = count + 1
Next
WScript.Echo "System/Application error count in the last 24 hours: " & count
