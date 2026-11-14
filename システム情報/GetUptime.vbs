' ファイル概要: 最終起動時刻と稼働時間を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetUptime.vbs
' 実行結果: 最終起動時刻と稼働時間を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, systems, system, dateTime, bootTime, elapsedMinutes
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")

For Each system In systems
    dateTime.Value = system.LastBootUpTime
    bootTime = dateTime.GetVarDate(True)
    elapsedMinutes = CLng((Now - bootTime) * 1440)
    WScript.Echo "Last boot: " & CStr(bootTime)
    WScript.Echo "Uptime: " & (elapsedMinutes \ 1440) & " days, " & _
        ((elapsedMinutes Mod 1440) \ 60) & " hours, " & _
        (elapsedMinutes Mod 60) & " minutes"
Next
