' ファイル概要: Plug and Playのイベントを表示する（デバイス識別子は非表示）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 163_UsbEventTimeline.vbs
' 実行結果: Plug and Playのイベントを表示する（デバイス識別子は非表示）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim dateFilter, service, events, eventItem
Set dateFilter = CreateObject("WbemScripting.SWbemDateTime")
' [処理解説] 指定した期間の開始時刻をWMIの日時形式に変換して照会条件に使用します。
dateFilter.SetVarDate DateAdd("d", -7, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_NTLogEvent を照会し、EventCode, SourceName, TimeGenerated を取得します。
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='System' AND (SourceName='Kernel-PnP' OR SourceName='UserPnp') " & _
    "AND TimeGenerated >= '" & dateFilter.Value & "'")
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & _
        vbTab & eventItem.SourceName
Next
WScript.Echo "Device identifiers and event message text were not collected."
