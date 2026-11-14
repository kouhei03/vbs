Option Explicit
Dim shell, pid
Set shell = CreateObject("WScript.Shell")
pid = shell.Run("notepad.exe", 7, False)
WScript.Echo "Notepad launch requested in minimized mode; return value=" & pid
