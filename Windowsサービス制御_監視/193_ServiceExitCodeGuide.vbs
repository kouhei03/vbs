' サービスの終了コードを取得し、よくあるWin32コードの説明を表示します。
Option Explicit
Dim serviceName, validator, service, services, item, found, codeValue, description
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 193_ServiceExitCodeGuide.vbs <service-name>"
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
Set services = service.ExecQuery("SELECT Name, State, ExitCode, ServiceSpecificExitCode FROM Win32_Service")
found = False
For Each item In services
    If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
        codeValue = CLng(item.ExitCode)
        description = "Unmapped Win32 service exit code"
        Select Case codeValue
            Case 0: description = "The operation completed successfully."
            Case 2: description = "The system cannot find the file specified."
            Case 5: description = "Access is denied."
            Case 1053: description = "The service did not respond to the start or control request in time."
            Case 1067: description = "The process terminated unexpectedly."
            Case 1068: description = "A dependency service or group failed to start."
        End Select
        WScript.Echo "State=" & item.State & " ExitCode=" & codeValue & " Meaning=" & description
        WScript.Echo "ServiceSpecificExitCode=" & item.ServiceSpecificExitCode
        found = True
        Exit For
    End If
Next
If Not found Then
    WScript.Echo "Service not found."
    WScript.Quit 3
End If
