' ローカルのログオンセッション種別と開始時刻を列挙します（切断・ロック操作なし）。
Option Explicit
Dim service, sessions, session
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set sessions = service.ExecQuery("SELECT LogonId, LogonType, StartTime FROM Win32_LogonSession")
For Each session In sessions
    WScript.Echo "LogonId=" & session.LogonId & " Type=" & session.LogonType & " Started=" & session.StartTime
Next
