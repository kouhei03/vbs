' 概要: サービスの現在のスタートアップ種類を照会します。変更は行いません。
' 実行方法: Usage: cscript //nologo 055_ServiceStartupTypeAudit.vbs <service-name>
' 処理結果: サービス起動種類の照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim name, re, service, records, item
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 055_ServiceStartupTypeAudit.vbs <service-name>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, StartMode, State FROM Win32_Service WHERE Name='" & name & "'")
If records.Count = 0 Then WScript.Echo "Service not found: " & name: WScript.Quit 1
For Each item In records
    WScript.Echo item.Name & " StartMode=" & item.StartMode & " State=" & item.State
Next
