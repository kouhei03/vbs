Option Explicit
Dim shell, process, outputText, errorText
Set shell=CreateObject("WScript.Shell")
Set process=shell.Exec("cmd.exe /c ver")
outputText=""
errorText=""
Do While Not process.StdOut.AtEndOfStream
    outputText=outputText & process.StdOut.ReadLine & vbCrLf
Loop
Do While Not process.StdErr.AtEndOfStream
    errorText=errorText & process.StdErr.ReadLine & vbCrLf
Loop
WScript.Echo "STDOUT:" & vbCrLf & outputText
WScript.Echo "STDERR:" & vbCrLf & errorText
WScript.Echo "ExitCode=" & process.ExitCode
