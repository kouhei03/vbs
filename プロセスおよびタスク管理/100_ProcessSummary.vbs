Option Explicit
Dim service, processes, process, count, bytes
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery("SELECT ProcessId,WorkingSetSize FROM Win32_Process")
count = 0
bytes = CDbl(0)
For Each process In processes
    count = count + 1
    If Not IsNull(process.WorkingSetSize) Then bytes = bytes + CDbl(process.WorkingSetSize)
Next
WScript.Echo "Process count=" & count
WScript.Echo "Total working set bytes=" & FormatNumber(bytes,0)
