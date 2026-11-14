' ファイル概要: Windows Installerのエラーイベントを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 159_MsiErrorSummary.vbs
' 実行結果: Windows Installerのエラーイベントを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim dateFilter, service, events, eventItem, count
Set dateFilter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
dateFilter.SetVarDate DateAdd("d", -7, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode, TimeGenerated, RecordNumber を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode, TimeGenerated, RecordNumber FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Application' AND SourceName='MsiInstaller' " & _
    "AND Type='Error' AND TimeGenerated >= '" & dateFilter.Value & "'")
count = 0
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & _
        vbTab & eventItem.RecordNumber
    count = count + 1
Next
WScript.Echo "Windows Installer error count in the last seven days: " & count
