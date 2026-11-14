' 概要: メンテナンス候補プロセスとPIDを表示します。強制終了やバックアップは行いません。
' 実行方法: Usage: cscript //nologo 073_MaintenanceProcessPreview.vbs <process.exe>
' 処理結果: メンテナンス候補の表示（終了なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim name, re, service, records, process
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 073_MaintenanceProcessPreview.vbs <process.exe>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, ProcessId, CreationDate FROM Win32_Process WHERE Name='" & name & "'")
For Each process In records
    WScript.Echo process.Name & " PID=" & process.ProcessId & " Started=" & process.CreationDate
Next
WScript.Echo "Preview only; no process was terminated."
