' メンテナンス候補プロセスとPIDを表示します。強制終了やバックアップは行いません。
Option Explicit
Dim name, re, service, records, process
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 073_MaintenanceProcessPreview.vbs <process.exe>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, ProcessId, CreationDate FROM Win32_Process WHERE Name='" & name & "'")
For Each process In records
    WScript.Echo process.Name & " PID=" & process.ProcessId & " Started=" & process.CreationDate
Next
WScript.Echo "Preview only; no process was terminated."
