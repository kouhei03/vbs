Option Explicit

Dim service, zones, zone, services, timeService
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set zones = service.ExecQuery("SELECT Description, Bias FROM Win32_TimeZone")

For Each zone In zones
    WScript.Echo "Time zone: " & zone.Description
    WScript.Echo "UTC bias (minutes): " & zone.Bias
Next

Set services = service.ExecQuery( _
    "SELECT State, StartMode FROM Win32_Service WHERE Name='W32Time'")
For Each timeService In services
    WScript.Echo "Windows Time service state: " & timeService.State
    WScript.Echo "Windows Time service start mode: " & timeService.StartMode
Next
