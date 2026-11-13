Option Explicit

Dim service, processes, process, bytes
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, PrivatePageCount FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "Private memory (bytes)"
For Each process In processes
    bytes = process.PrivatePageCount
    If IsNull(bytes) Then bytes = "Unavailable"
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & bytes
Next
