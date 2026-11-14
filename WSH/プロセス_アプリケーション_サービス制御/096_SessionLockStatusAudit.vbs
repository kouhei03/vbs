' 対話ユーザーとログオンセッション情報を表示します。画面ロックは行いません。
Option Explicit
Dim service, systems, item, sessions, session
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT Name, UserName FROM Win32_ComputerSystem")
For Each item In systems
    If IsNull(item.UserName) Or item.UserName = "" Then
        WScript.Echo "Computer=" & item.Name & " InteractiveUser=(none)"
    Else
        WScript.Echo "Computer=" & item.Name & " InteractiveUser=" & item.UserName
    End If
Next
Set sessions = service.ExecQuery("SELECT LogonId, LogonType, StartTime FROM Win32_LogonSession WHERE LogonType=2 OR LogonType=10")
For Each session In sessions
    WScript.Echo "InteractiveLogonId=" & session.LogonId & " Type=" & session.LogonType & " Started=" & session.StartTime
Next
