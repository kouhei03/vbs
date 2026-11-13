Option Explicit

Dim service, processes, process, userName, domainName, result
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Owner"
For Each process In processes
    userName = ""
    domainName = ""
    result = process.GetOwner(userName, domainName)
    If result = 0 Then
        WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
            domainName & "\" & userName
    Else
        WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
            "Unavailable (code " & result & ")"
    End If
Next
