Option Explicit
Dim shell, waitForExit, result, mode
If WScript.Arguments.Count<>1 Then WScript.Echo "Usage: cscript //nologo 146_RunFixedCommand.vbs <sync|async>": WScript.Quit 2
mode=LCase(WScript.Arguments(0))
If mode="sync" Then
    waitForExit=True
ElseIf mode="async" Then
    waitForExit=False
Else
    WScript.Echo "Mode must be sync or async."
    WScript.Quit 2
End If
Set shell=CreateObject("WScript.Shell")
result=shell.Run("cmd.exe /c ver",1,waitForExit)
If waitForExit Then WScript.Echo "Exit code=" & result Else WScript.Echo "Command launched asynchronously."
