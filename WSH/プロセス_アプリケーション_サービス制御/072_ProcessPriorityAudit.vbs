' 概要: プロセス優先度を照会します。優先度・CPU affinityの変更は行いません。
' 実行方法: cscript //nologo 072_ProcessPriorityAudit.vbs
' 処理結果: プロセス優先度の照会（変更なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, records, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, ProcessId, Priority FROM Win32_Process")
For Each process In records
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & "Priority=" & process.Priority
Next
WScript.Echo "Read-only audit; no scheduling settings changed."
