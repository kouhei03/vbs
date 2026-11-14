' ファイル概要: Windows Installerのエラーイベントを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 159_MsiErrorSummary.vbs
' 実行結果: Windows Installerのエラーイベントを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim filter, service, events, eventItem, count
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("d", -7, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, TimeGenerated, RecordNumber FROM Win32_NTLogEvent " & _
    "WHERE Logfile='Application' AND SourceName='MsiInstaller' " & _
    "AND Type='Error' AND TimeGenerated >= '" & filter.Value & "'")
count = 0
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & _
        vbTab & eventItem.RecordNumber
    count = count + 1
Next
WScript.Echo "Windows Installer error count in the last seven days: " & count
