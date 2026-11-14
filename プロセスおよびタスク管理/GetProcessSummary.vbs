Option Explicit

Dim service, processes, process, count, totalWorkingSet
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT ProcessId, WorkingSetSize FROM Win32_Process")

count = 0
totalWorkingSet = 0
For Each process In processes
    count = count + 1
    If Not IsNull(process.WorkingSetSize) Then
        totalWorkingSet = totalWorkingSet + CDbl(process.WorkingSetSize)
    End If
Next

WScript.Echo "Process count: " & count
WScript.Echo "Combined working set (bytes): " & FormatNumber(totalWorkingSet, 0)
