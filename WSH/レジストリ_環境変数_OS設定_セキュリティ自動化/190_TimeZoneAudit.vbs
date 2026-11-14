' 概要: ローカルのタイムゾーン名とUTCオフセットを表示します。時刻・タイムゾーン変更はしません。
' 実行方法: cscript //nologo 190_TimeZoneAudit.vbs
' 処理結果: ローカルタイムゾーン情報確認。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, zones, zone
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set zones = service.ExecQuery("SELECT Caption, Bias, StandardName, DaylightName FROM Win32_TimeZone")
For Each zone In zones
    WScript.Echo "Caption=" & zone.Caption & " Standard=" & zone.StandardName
    WScript.Echo "Daylight=" & zone.DaylightName & " BiasMinutes=" & zone.Bias
Next
