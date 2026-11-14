Option Explicit
Dim service, items, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set items = service.ExecQuery("SELECT Name,State,ProcessId,StartMode FROM Win32_Service WHERE Name='Spooler'")
If items.Count = 0 Then WScript.Echo "Spooler service not found.": WScript.Quit 1
For Each item In items
    WScript.Echo "State=" & item.State & " PID=" & item.ProcessId & " StartMode=" & item.StartMode
Next
