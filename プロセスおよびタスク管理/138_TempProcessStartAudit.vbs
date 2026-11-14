Option Explicit
Dim service, processes, process, shell, tempPath, pathValue, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery("SELECT Name,ProcessId,ExecutablePath,CreationDate FROM Win32_Process")
Set shell = CreateObject("WScript.Shell")
tempPath = LCase(shell.ExpandEnvironmentStrings("%TEMP%"))
found = False
For Each process In processes
    If Not IsNull(process.ExecutablePath) Then
        pathValue = LCase(process.ExecutablePath)
        If Left(pathValue, Len(tempPath)) = tempPath Then
            WScript.Echo process.Name & " PID=" & process.ProcessId & " Created=" & process.CreationDate
            found = True
        End If
    End If
Next
If Not found Then WScript.Echo "No process image under the current TEMP path was found."
