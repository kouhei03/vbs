' 概要: Windows Timeサービスの状態を照会します。時刻同期コマンドは実行しません。
' 実行方法: cscript //nologo 025_TimeServiceAudit.vbs
' 処理結果: Windows Timeサービス状態照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, State, StartMode FROM Win32_Service WHERE Name='W32Time'")
If records.Count = 0 Then WScript.Echo "W32Time service not found.": WScript.Quit 1
For Each item In records
    WScript.Echo item.Name & vbTab & item.State & vbTab & item.StartMode
Next
