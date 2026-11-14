' 概要: 指定プロセスのメモリ量を確認し、閾値超過を報告します。再起動は行いません。
' 実行方法: Usage: cscript //nologo 094_ProcessMemoryThresholdAudit.vbs <process.exe> <threshold-MB>
' 処理結果: メモリ閾値監査（再起動なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim name, re, threshold, service, records, process, found
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 2 Then
    WScript.Echo "Usage: cscript //nologo 094_ProcessMemoryThresholdAudit.vbs <process.exe> <threshold-MB>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
If Not IsNumeric(WScript.Arguments(1)) Then WScript.Quit 2
If CDbl(WScript.Arguments(1)) <= 0 Or CDbl(WScript.Arguments(1)) > 1048576 Then WScript.Quit 2
threshold = CDbl(WScript.Arguments(1)) * 1048576
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT ProcessId, WorkingSetSize FROM Win32_Process WHERE Name='" & name & "'")
found = False
For Each process In records
    found = True
    WScript.Echo name & " PID=" & process.ProcessId & " WorkingSetBytes=" & process.WorkingSetSize
    If CDbl(process.WorkingSetSize) > threshold Then WScript.Echo "WARNING: memory threshold exceeded."
Next
If Not found Then WScript.Echo "Process not found."
