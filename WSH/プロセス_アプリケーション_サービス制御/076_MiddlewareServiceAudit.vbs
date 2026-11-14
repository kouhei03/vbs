' 代表的なミドルウェアサービスの状態を照会します。管理コマンドは実行しません。
Option Explicit
Dim service, name, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
For Each name In Array("W3SVC", "MSSQLSERVER", "Spooler")
    Set records = service.ExecQuery("SELECT Name, State, StartMode FROM Win32_Service WHERE Name='" & name & "'")
    If records.Count = 0 Then
        WScript.Echo name & vbTab & "Not installed"
    Else
        For Each item In records
            WScript.Echo item.Name & vbTab & item.State & vbTab & item.StartMode
        Next
    End If
Next
