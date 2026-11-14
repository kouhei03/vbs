Option Explicit
Dim service, events, eventRecord, count
Set service=GetObject("winmgmts:\\.\root\cimv2")
Set events=service.ExecQuery("SELECT EventCode,TimeGenerated,RecordNumber FROM Win32_NTLogEvent WHERE Logfile='System' AND (EventCode='1074' OR EventCode='6006' OR EventCode='6008')")
count=0
For Each eventRecord In events
    WScript.Echo "Event=" & eventRecord.EventCode & " Time=" & eventRecord.TimeGenerated & " Record=" & eventRecord.RecordNumber
    count=count+1
    If count>=20 Then Exit For
Next
If count=0 Then WScript.Echo "No shutdown-related event returned."
