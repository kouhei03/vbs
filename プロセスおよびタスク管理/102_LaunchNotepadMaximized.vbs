Option Explicit
Dim shell, pid
Set shell = CreateObject("WScript.Shell")
pid = shell.Run("notepad.exe", 3, False)
WScript.Echo "Notepad launch requested in maximized mode; return value=" & pid
