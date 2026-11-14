' 概要: 指定プロセスのSessionIdを照会します。別セッションでの起動は行いません。
' 実行方法: Usage: cscript //nologo 061_ProcessSessionAudit.vbs <process.exe>
' 処理結果: プロセスのSessionId照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim name, re, service, records, process
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 061_ProcessSessionAudit.vbs <process.exe>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT ProcessId, SessionId FROM Win32_Process WHERE Name='" & name & "'")
For Each process In records
    WScript.Echo name & " PID=" & process.ProcessId & " SessionId=" & process.SessionId
Next
