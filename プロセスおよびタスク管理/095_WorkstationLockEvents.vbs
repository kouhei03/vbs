Option Explicit
Dim service, events, eventRecord, query, count
Set service = GetObject("winmgmts:\\.\root\cimv2")
query = "SELECT EventCode,TimeGenerated,RecordNumber,SourceName FROM Win32_NTLogEvent WHERE Logfile='Security' AND EventCode='4800'"
Set events = service.ExecQuery(query)
count = 0
For Each eventRecord In events
    WScript.Echo "Event=" & eventRecord.EventCode & " Time=" & eventRecord.TimeGenerated & " Record=" & eventRecord.RecordNumber
    count = count + 1
    If count >= 20 Then Exit For
Next
If count = 0 Then WScript.Echo "No event returned; log permissions/audit policy may limit results."
