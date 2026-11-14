' 概要: 印刷スプーラーサービスの状態を照会します。サービスや印刷キューは変更しません。
' 実行方法: cscript //nologo 080_PrintSpoolerAudit.vbs
' 処理結果: スプーラー状態照会（再起動なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Name, State, StartMode, ProcessId FROM Win32_Service WHERE Name='Spooler'")
If records.Count = 0 Then WScript.Echo "Spooler service not found.": WScript.Quit 1
For Each item In records
    WScript.Echo "State=" & item.State & " StartMode=" & item.StartMode & " PID=" & item.ProcessId
Next
