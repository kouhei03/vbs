Option Explicit

Dim shell, process, line, ruleCount, netshPath
Set shell = CreateObject("WScript.Shell")
netshPath = shell.ExpandEnvironmentStrings("%SystemRoot%\System32\netsh.exe")
Set process = shell.Exec(Chr(34) & netshPath & Chr(34) & _
    " advfirewall consec show rule name=all")
ruleCount = 0
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    If InStr(1, line, "Rule Name", vbTextCompare) > 0 Then ruleCount = ruleCount + 1
Next
Do While Not process.StdErr.AtEndOfStream
    WScript.Echo "ERROR: " & process.StdErr.ReadLine
Loop
If process.ExitCode <> 0 Then
    WScript.Echo "IPsec/connections rule query failed with exit code " & process.ExitCode
    WScript.Quit process.ExitCode
End If
WScript.Echo "Connection security rule labels found: " & ruleCount
WScript.Echo "Read-only query; IPsec policy was not changed."
