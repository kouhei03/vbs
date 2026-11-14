' ファイル概要: タイムゾーンとUTCオフセットを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetTimeZoneInfo.vbs
' 実行結果: タイムゾーンとUTCオフセットを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, zones, zone, services, timeService
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_TimeZone を照会し、Description, Bias を取得します。
Set zones = service.ExecQuery("SELECT Description, Bias FROM Win32_TimeZone")

For Each zone In zones
    WScript.Echo "Time zone: " & zone.Description
    WScript.Echo "UTC bias (minutes): " & zone.Bias
Next

' [処理解説] WMIの Win32_Service を照会し、State, StartMode を取得します。
Set services = service.ExecQuery( _
    "SELECT State, StartMode FROM Win32_Service WHERE Name='W32Time'")
For Each timeService In services
    WScript.Echo "Windows Time service state: " & timeService.State
    WScript.Echo "Windows Time service start mode: " & timeService.StartMode
Next
