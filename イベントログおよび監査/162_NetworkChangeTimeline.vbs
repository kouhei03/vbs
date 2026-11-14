' ファイル概要: ネットワーク関連のイベントメタデータを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 162_NetworkChangeTimeline.vbs
' 実行結果: ネットワーク関連のイベントメタデータを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim filter, service, events, eventItem
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("d", -7, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='System' AND (SourceName='Tcpip' OR SourceName='Dhcp' " & _
    "OR SourceName='Microsoft-Windows-NetworkProfile') " & _
    "AND TimeGenerated >= '" & filter.Value & "'")
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & _
        vbTab & eventItem.SourceName
Next
WScript.Echo "Network event messages and addresses were omitted."
