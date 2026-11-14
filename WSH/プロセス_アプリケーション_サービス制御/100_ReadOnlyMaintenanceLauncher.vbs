' 概要: ローカルのプロセス件数と代表サービス状態をまとめて表示します。制御操作はしません。
' 実行方法: cscript //nologo 100_ReadOnlyMaintenanceLauncher.vbs
' 処理結果: プロセス数・代表サービス状態の読み取り専用サマリー。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, processes, process, processCount, name, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set processes = service.ExecQuery("SELECT ProcessId FROM Win32_Process")
processCount = 0
For Each process In processes
    processCount = processCount + 1
Next
WScript.Echo "Local process count=" & processCount
For Each name In Array("Spooler", "W3SVC", "MSSQLSERVER")
    Set records = service.ExecQuery("SELECT State, StartMode FROM Win32_Service WHERE Name='" & name & "'")
    If records.Count = 0 Then
        WScript.Echo name & "=Not installed"
    Else
        For Each item In records
            WScript.Echo name & "=" & item.State & " StartMode=" & item.StartMode
        Next
    End If
Next
WScript.Echo "Read-only maintenance summary; no process, service, security, or startup changes."
