' Applicationログへ固定の情報イベントを1件記録します。Security監査ログには書き込みません。
Option Explicit
Dim shell
Set shell = CreateObject("WScript.Shell")
If Not shell.LogEvent(4, "WSH safe example: informational configuration audit marker.") Then
    WScript.Echo "Could not write an Application information event."
    WScript.Quit 1
End If
WScript.Echo "One informational event was written to the Application log."
