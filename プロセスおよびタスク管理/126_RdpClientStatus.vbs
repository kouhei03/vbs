Option Explicit
Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery("SELECT Name,ProcessId,CreationDate FROM Win32_Process WHERE Name='mstsc.exe'")
If processes.Count = 0 Then WScript.Echo "mstsc.exe is not running."
For Each process In processes
    WScript.Echo process.Name & " PID=" & process.ProcessId & " Created=" & process.CreationDate
Next
