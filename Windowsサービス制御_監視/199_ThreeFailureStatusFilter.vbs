' 連続3回の非実行状態をコンソールで検知します。自動復旧や通知送信は行いません。
Option Explicit
Dim serviceName, validator, service, services, item, failures, sample, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 199_ThreeFailureStatusFilter.vbs <service-name>"
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
failures = 0
For sample = 1 To 3
    Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
    found = False
    For Each item In services
        If StrComp(item.Name, serviceName, vbTextCompare) = 0 Then
            found = True
            WScript.Echo "Sample " & sample & "=" & item.State
            If StrComp(item.State, "Running", vbTextCompare) = 0 Then
                failures = 0
            Else
                failures = failures + 1
            End If
            Exit For
        End If
    Next
    If Not found Then
        WScript.Echo "Service not found."
        WScript.Quit 3
    End If
    If sample < 3 Then WScript.Sleep 5000
Next
If failures = 3 Then
    WScript.Echo "ALERT: not Running in all 3 samples; no action was taken."
Else
    WScript.Echo "No three-consecutive-sample failure."
End If
