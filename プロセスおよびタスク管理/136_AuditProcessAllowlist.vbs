Option Explicit
Dim fso, allowFile, allowed, line, service, processes, process, found
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 136_AuditProcessAllowlist.vbs <allowlist-file>": WScript.Quit 2
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(WScript.Arguments(0)) Then WScript.Echo "Allowlist file not found.": WScript.Quit 1
Set allowed = CreateObject("Scripting.Dictionary")
allowed.CompareMode = vbTextCompare
Set allowFile = fso.OpenTextFile(WScript.Arguments(0),1,False)
Do Until allowFile.AtEndOfStream
    line = Trim(allowFile.ReadLine)
    If Len(line)>0 Then allowed(line)=True
Loop
allowFile.Close
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery("SELECT Name,ProcessId FROM Win32_Process")
For Each process In processes
    If Not allowed.Exists(process.Name) Then WScript.Echo "Not in allowlist: " & process.Name & " PID=" & process.ProcessId
Next
WScript.Echo "Audit only; no process was blocked."
