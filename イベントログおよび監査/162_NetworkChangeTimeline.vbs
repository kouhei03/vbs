' ファイル概要: ネットワーク関連のイベントメタデータを表示する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 162_NetworkChangeTimeline.vbs
' 実行結果: ネットワーク関連のイベントメタデータを表示する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim dateFilter, service, events, eventItem
Set dateFilter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
dateFilter.SetVarDate DateAdd("d", -7, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode, SourceName, TimeGenerated を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='System' AND (SourceName='Tcpip' OR SourceName='Dhcp' " & _
    "OR SourceName='Microsoft-Windows-NetworkProfile') " & _
    "AND TimeGenerated >= '" & dateFilter.Value & "'")
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & _
        vbTab & eventItem.SourceName
Next
WScript.Echo "Network event messages and addresses were omitted."
