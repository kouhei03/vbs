' 指定サービスの状態を、最大5回の段階的間隔で観察します。再起動はしません。
Option Explicit
Dim serviceName, validator, service, services, item, attempt, delaySeconds, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 187_ServiceBackoffObservation.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Invalid service name."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
For attempt = 1 To 5
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    found = False
    For Each item In services
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            WScript.Echo "Sample " & attempt & vbTab & item.State
            found = True
            Exit For
        End If
    Next
    If Not found Then
        WScript.Echo "Service not found."
        WScript.Quit 3
    End If
    delaySeconds = attempt * 2
    If attempt < 5 Then WScript.Sleep delaySeconds * 1000
Next
WScript.Echo "Observation complete; no recovery action was taken."
