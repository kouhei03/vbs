' 概要: 指定Windowsサービスの状態を照会します。開始・停止・再起動は行いません。
' 実行方法: Usage: cscript //nologo 054_ServiceStatusAudit.vbs <service-name>
' 処理結果: サービス状態照会（制御なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim name, re, service, records, item
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 054_ServiceStatusAudit.vbs <service-name>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, State, StartMode, ProcessId FROM Win32_Service WHERE Name='" & name & "'")
If records.Count = 0 Then
    WScript.Echo "Service not found: " & name
    WScript.Quit 1
End If
For Each item In records
    WScript.Echo "Name=" & item.Name & " State=" & item.State & " StartMode=" & item.StartMode & " PID=" & item.ProcessId
Next
