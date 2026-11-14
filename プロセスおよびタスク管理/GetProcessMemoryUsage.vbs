Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, WorkingSetSize, VirtualSize FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "Working set (bytes)" & vbTab & "Virtual size (bytes)"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        ByteCount(process.WorkingSetSize) & vbTab & _
        ByteCount(process.VirtualSize)
Next

Function ByteCount(value)
    If IsNull(value) Or IsEmpty(value) Then
        ByteCount = "N/A"
    Else
        ByteCount = FormatNumber(CDbl(value), 0)
    End If
End Function
