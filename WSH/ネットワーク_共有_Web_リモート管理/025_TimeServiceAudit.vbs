' Windows Timeサービスの状態を照会します。時刻同期コマンドは実行しません。
Option Explicit
Dim service, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name, State, StartMode FROM Win32_Service WHERE Name='W32Time'")
If records.Count = 0 Then WScript.Echo "W32Time service not found.": WScript.Quit 1
For Each item In records
    WScript.Echo item.Name & vbTab & item.State & vbTab & item.StartMode
Next
