Option Explicit
Dim verbose, i
verbose=False
For i=0 To WScript.Arguments.Count-1
    If LCase(WScript.Arguments(i))="/verbose" Then verbose=True
Next
WScript.Echo "Process diagnostics started."
If verbose Then
    WScript.Echo "Host=" & CreateObject("WScript.Network").ComputerName
    WScript.Echo "Timestamp=" & Now
    WScript.Echo "Arguments=" & WScript.Arguments.Count
End If
