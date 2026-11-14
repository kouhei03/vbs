Option Explicit
Dim service, processName, validator, records, process, fieldName, value, userName, domainName, result
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 077_ProcessOwner.vbs <process.exe>"
    WScript.Quit 2
End If
processName = WScript.Arguments(0)
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Then
    WScript.Echo "Use a simple executable name such as notepad.exe."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
If records.Count = 0 Then
    WScript.Echo "Process not found: " & processName
    WScript.Quit 1
End If
For Each process In records
    WScript.Echo "Name=" & process.Name & " PID=" & process.ProcessId
    userName = ""
    domainName = ""
    result = process.GetOwner(userName, domainName)
    If result = 0 Then
        WScript.Echo "Owner=" & domainName & "\" & userName
    Else
        WScript.Echo "Owner unavailable; WMI return code=" & result
    End If
Next
