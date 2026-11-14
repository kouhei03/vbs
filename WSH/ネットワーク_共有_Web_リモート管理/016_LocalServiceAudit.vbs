' 概要: ローカルの指定サービス状態を読み取ります。リモート接続・制御はしません。
' 実行方法: Usage: cscript //nologo 016_LocalServiceAudit.vbs <service-name>
' 処理結果: 単一ローカルサービス状態照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim name, re, service, records, item
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 016_LocalServiceAudit.vbs <service-name>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, State, StartMode FROM Win32_Service WHERE Name='" & name & "'")
If records.Count = 0 Then WScript.Echo "Local service not found.": WScript.Quit 1
For Each item In records
    WScript.Echo item.Name & vbTab & item.State & vbTab & item.StartMode
Next
