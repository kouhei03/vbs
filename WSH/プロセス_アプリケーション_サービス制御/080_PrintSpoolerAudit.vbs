' 印刷スプーラーサービスの状態を照会します。サービスや印刷キューは変更しません。
Option Explicit
Dim service, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, State, StartMode, ProcessId FROM Win32_Service WHERE Name='Spooler'")
If records.Count = 0 Then WScript.Echo "Spooler service not found.": WScript.Quit 1
For Each item In records
    WScript.Echo "State=" & item.State & " StartMode=" & item.StartMode & " PID=" & item.ProcessId
Next
