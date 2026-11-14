' PATH/TEMP/USERPROFILEのみを過去スナップショットと比較します。値そのものは表示しません。
Option Explicit
Dim fso, path, file, previous, shell, env, names, name, currentValue, changed
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 186_SafeEnvironmentSnapshotCompare.vbs <snapshot.txt>"
    WScript.Quit 2
End If
Set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName(WScript.Arguments(0))
If Not fso.FileExists(path) Then WScript.Quit 2
Set previous = CreateObject("Scripting.Dictionary")
previous.CompareMode = 1
Set file = fso.OpenTextFile(path, 1, False, -2)
Do Until file.AtEndOfStream
    ParseSnapshotLine file.ReadLine, previous
Loop
file.Close
Set shell = CreateObject("WScript.Shell")
Set env = shell.Environment("PROCESS")
names = Array("PATH", "TEMP", "USERPROFILE")
For Each name In names
    currentValue = env(name)
    changed = (Not previous.Exists(name))
    If Not changed Then changed = (StrComp(previous(name), currentValue, 0) <> 0)
    If changed Then WScript.Echo name & " changed or missing from snapshot."
Next
Sub ParseSnapshotLine(line, dictionary)
    Dim separator, key
    separator = InStr(line, "=")
    If separator > 1 Then
        key = UCase(Left(line, separator - 1))
        If key = "PATH" Or key = "TEMP" Or key = "USERPROFILE" Then
            dictionary(key) = Mid(line, separator + 1)
        End If
    End If
End Sub
