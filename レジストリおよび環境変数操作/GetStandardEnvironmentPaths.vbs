Option Explicit

Dim shell, environment, variableName
Set shell = CreateObject("WScript.Shell")
Set environment = shell.Environment("PROCESS")

For Each variableName In Array("SystemRoot", "TEMP", "TMP")
    WScript.Echo variableName & "=" & environment(variableName)
Next
