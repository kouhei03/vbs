' ファイル概要: Plug and Playのイベントを表示する（デバイス識別子は非表示）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 163_UsbEventTimeline.vbs
' 実行結果: Plug and Playのイベントを表示する（デバイス識別子は非表示）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim filter, service, events, eventItem
Set filter = CreateObject("WbemScripting.SWbemDateTime")
filter.SetVarDate DateAdd("d", -7, Now), True
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set events = service.ExecQuery( _
    "SELECT EventCode, SourceName, TimeGenerated FROM Win32_NTLogEvent " & _
    "WHERE Logfile='System' AND (SourceName='Kernel-PnP' OR SourceName='UserPnp') " & _
    "AND TimeGenerated >= '" & filter.Value & "'")
For Each eventItem In events
    WScript.Echo eventItem.TimeGenerated & vbTab & eventItem.EventCode & _
        vbTab & eventItem.SourceName
Next
WScript.Echo "Device identifiers and event message text were not collected."
