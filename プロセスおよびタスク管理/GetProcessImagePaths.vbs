Option Explicit

Dim service, processes, process, imagePath
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, ExecutablePath FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Executable path"
For Each process In processes
    imagePath = process.ExecutablePath
    If IsNull(imagePath) Then imagePath = "Unavailable"
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & imagePath
Next
