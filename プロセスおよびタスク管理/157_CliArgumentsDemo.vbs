Option Explicit
Dim i
WScript.Echo "Argument count=" & WScript.Arguments.Count
For i=0 To WScript.Arguments.Count-1
    WScript.Echo "Argument " & i & "=" & WScript.Arguments(i)
Next
