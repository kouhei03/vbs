Option Explicit
Dim shell, environment, variableName, entry, separator, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo InspectEnvironmentRemovalTarget.vbs <variable-name>"
    WScript.Quit 2
End If
variableName = WScript.Arguments(0)
Set shell = CreateObject("WScript.Shell")
Set environment = shell.Environment("USER")
found = False
For Each entry In environment
    separator = InStr(entry, "=")
    If separator > 0 Then
        If StrComp(Left(entry, separator - 1), variableName, vbTextCompare) = 0 Then found = True
    End If
Next
If found Then
    WScript.Echo "Variable exists in the current user environment: " & variableName
Else
    WScript.Echo "Variable is not present in the current user environment: " & variableName
End If
WScript.Echo "No environment variable was removed."
