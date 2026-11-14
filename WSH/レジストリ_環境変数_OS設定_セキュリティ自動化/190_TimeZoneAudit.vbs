' ローカルのタイムゾーン名とUTCオフセットを表示します。時刻・タイムゾーン変更はしません。
Option Explicit
Dim service, zones, zone
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set zones = service.ExecQuery("SELECT Caption, Bias, StandardName, DaylightName FROM Win32_TimeZone")
For Each zone In zones
    WScript.Echo "Caption=" & zone.Caption & " Standard=" & zone.StandardName
    WScript.Echo "Daylight=" & zone.DaylightName & " BiasMinutes=" & zone.Bias
Next
