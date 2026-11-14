' 概要: 指定プロセスの親子関係を表示します。プロセス終了は行いません。
' 実行方法: Usage: cscript //nologo 067_ProcessTreeAudit.vbs <process.exe>
' 処理結果: 指定プロセスの親PID監査（終了なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim name, re, service, records, process, names, parents, key, pid
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 067_ProcessTreeAudit.vbs <process.exe>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, ProcessId, ParentProcessId FROM Win32_Process WHERE Name='" & name & "'")
For Each process In records
    WScript.Echo "Name=" & process.Name & " PID=" & process.ProcessId & " ParentPID=" & process.ParentProcessId
Next
WScript.Echo "Audit only; process tree was not modified."
