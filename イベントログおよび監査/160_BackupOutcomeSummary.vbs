' ファイル概要: バックアップに関連するイベントを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 160_BackupOutcomeSummary.vbs
' 実行結果: バックアップに関連するイベントを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim filter, service, events, eventItem, count
Set filter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
filter.SetVarDate DateAdd("d", -7, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode, SourceName, TimeGenerated を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE (Logfile='Application' OR Logfile='System' OR " & _
    "Logfile='Microsoft-Windows-Backup' OR " & _
    "Logfile='Microsoft-Windows-Backup/Operational') " & _
    "AND (SourceName='Microsoft-Windows-Backup' OR SourceName='Backup') " & _
    "AND TimeGenerated >= '" & filter.Value & "'")
count = 0
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & _
        vbTab & eventItem.SourceName
    count = count + 1
Next
WScript.Echo "Backup-related events found: " & count
WScript.Echo "Event messages were not read; verify outcome in the backup application."
