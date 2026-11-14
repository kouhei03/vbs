' ファイル概要: 直近7日間のWindows Installerエラーを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo FindMsiInstallerErrors.vbs
' 実行結果: 直近7日間のWindows Installerエラーを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim timeFilter, service, events, eventItem
Set timeFilter = CreateObject("WbemScripting.SWbemDateTime")
timeFilter.SetVarDate DateAdd("h", -168, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated, RecordNumber FROM " & _
    "Win32_NTLogEvent WHERE Logfile='Application' " & _
    "AND SourceName='MsiInstaller' AND Type='Error' " & _
    "AND TimeGenerated >= '" & timeFilter.Value & "'")

WScript.Echo "MsiInstaller errors (last 7 days)"
WScript.Echo "Time" & vbTab & "Event ID" & vbTab & "Record"
For Each eventItem In events
    WScript.Echo EventTime(eventItem.TimeGenerated) & vbTab & _
        eventItem.EventCode & vbTab & eventItem.RecordNumber
Next

Function EventTime(cimTime)
    Dim dateTime
    Set dateTime = CreateObject("WbemScripting.SWbemDateTime")
    dateTime.Value = cimTime
    EventTime = CStr(dateTime.GetVarDate(True))
End Function
