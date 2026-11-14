Option Explicit

Dim service, services, item, serviceName, timeoutText, timeoutSeconds
Dim found, running, startTime, index, character
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo WaitForServiceRunning.vbs <service-name> <timeout-seconds 1-300>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
timeoutText = Trim(WScript.Arguments(1))
If Len(serviceName) = 0 Or Len(timeoutText) = 0 Or Len(timeoutText) > 3 Then
    WScript.Echo "A service name and timeout from 1 to 300 are required."
    WScript.Quit 2
End If
For index = 1 To Len(timeoutText)
    character = Mid(timeoutText, index, 1)
    If InStr(1, "0123456789", character, vbBinaryCompare) = 0 Then
        WScript.Echo "Timeout must be a whole number."
        WScript.Quit 2
    End If
Next
timeoutSeconds = CLng(timeoutText)
If timeoutSeconds < 1 Or timeoutSeconds > 300 Then
    WScript.Echo "Timeout must be between 1 and 300 seconds."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
startTime = Now
Do
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    found = False
    running = False
    For Each item In services
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            found = True
            running = (StrComp(item.State, "Running", vbTextCompare) = 0)
            Exit For
        End If
    Next
    If Not found Then
        WScript.Echo "Service not found: " & serviceName
        WScript.Quit 3
    End If
    If running Then
        WScript.Echo "Service is running: " & serviceName
        WScript.Quit 0
    End If
    If DateDiff("s", startTime, Now) >= timeoutSeconds Then
        WScript.Echo "Timed out waiting for service: " & serviceName
        WScript.Quit 1
    End If
    WScript.Sleep 1000
Loop
