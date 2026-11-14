' 概要: 対話ユーザーとログオンセッション情報を表示します。画面ロックは行いません。
' 実行方法: cscript //nologo 096_SessionLockStatusAudit.vbs
' 処理結果: 対話セッション情報の表示（ロックなし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, systems, item, sessions, session
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
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
