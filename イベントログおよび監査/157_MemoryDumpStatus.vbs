Option Explicit

Dim shell, fso, systemRoot, dumpPath, dumpFile
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
systemRoot = shell.ExpandEnvironmentStrings("%SystemRoot%")
dumpPath = fso.BuildPath(systemRoot, "MEMORY.DMP")

If fso.FileExists(dumpPath) Then
    Set dumpFile = fso.GetFile(dumpPath)
    WScript.Echo "MEMORY.DMP exists; bytes=" & dumpFile.Size & _
        "; modified=" & dumpFile.DateLastModified
Else
    WScript.Echo "MEMORY.DMP was not found at the standard path."
End If
WScript.Echo "The dump was not opened, copied, or uploaded."
