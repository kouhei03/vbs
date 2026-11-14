Option Explicit
Dim shell, process, line, found
Set shell = CreateObject("WScript.Shell")
Set process = shell.Exec("tasklist.exe /v /fo csv")
found = False
Do While Not process.StdOut.AtEndOfStream
    line = process.StdOut.ReadLine
    If InStr(1, line, "Not Responding", vbTextCompare) > 0 Or InStr(line, "応答なし") > 0 Then
        WScript.Echo line
        found = True
    End If
Loop
If Not found Then WScript.Echo "No not-responding process row found."
If process.ExitCode <> 0 Then WScript.Quit process.ExitCode
