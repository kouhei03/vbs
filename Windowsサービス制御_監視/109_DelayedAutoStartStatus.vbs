' 指定サービスのDelayedAutoStartレジストリ値を読み取り専用で表示します。
Option Explicit
Dim serviceName, validator, shell, registryPath, delayedValue
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 109_DelayedAutoStartStatus.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Invalid service name."
    WScript.Quit 2
End If
Set shell = CreateObject("WScript.Shell")
registryPath = "HKLM\SYSTEM\CurrentControlSet\Services" & serviceName & "\DelayedAutoStart"
On Error Resume Next
delayedValue = shell.RegRead(registryPath)
If Err.Number <> 0 Then
    WScript.Echo "DelayedAutoStart value is absent or unreadable; this alone does not establish startup mode."
    Err.Clear
Else
    WScript.Echo "DelayedAutoStart=" & delayedValue & " (1 usually indicates delayed automatic start)"
End If
On Error GoTo 0
