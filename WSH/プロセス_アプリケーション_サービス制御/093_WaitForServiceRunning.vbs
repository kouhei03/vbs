' サービスがRunningになるまで最大60秒ポーリングします（サービス制御なし）。
Option Explicit
Dim name, re, service, deadline, records, item
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 093_WaitForServiceRunning.vbs <service-name>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
deadline = DateAdd("s", 60, Now)
Do While Now < deadline
    Set records = service.ExecQuery("SELECT State FROM Win32_Service WHERE Name='" & name & "'")
    If records.Count = 0 Then WScript.Echo "Service not found: " & name: WScript.Quit 1
    For Each item In records
        If item.State = "Running" Then
            WScript.Echo name & " is Running."
            WScript.Quit 0
        End If
    Next
    WScript.Sleep 1000
Loop
WScript.Echo "Timed out waiting for Running state; service was not started."
WScript.Quit 1
