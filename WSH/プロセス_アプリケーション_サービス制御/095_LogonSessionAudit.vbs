' 概要: ローカルのログオンセッション種別と開始時刻を列挙します（切断・ロック操作なし）。
' 実行方法: cscript //nologo 095_LogonSessionAudit.vbs
' 処理結果: ログオンセッション種別の列挙。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, sessions, session
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set sessions = service.ExecQuery("SELECT LogonId, LogonType, StartTime FROM Win32_LogonSession")
For Each session In sessions
    WScript.Echo "LogonId=" & session.LogonId & " Type=" & session.LogonType & " Started=" & session.StartTime
Next
