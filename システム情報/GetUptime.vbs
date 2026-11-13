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
