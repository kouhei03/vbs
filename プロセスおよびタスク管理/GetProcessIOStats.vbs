' ファイル概要: プロセスごとのI/O統計を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessIOStats.vbs
' 実行結果: プロセスごとのI/O統計を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、Name, ProcessId, ReadOperationCount, WriteOperationCount, ReadTransferCount, WriteTransferCount を取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, ReadOperationCount, WriteOperationCount, " & _
    "ReadTransferCount, WriteTransferCount FROM Win32_Process")

WScript.Echo "Process" & vbTab & "PID" & vbTab & "Read ops" & vbTab & _
    "Write ops" & vbTab & "Read bytes" & vbTab & "Written bytes"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.ReadOperationCount & vbTab & process.WriteOperationCount & _
        vbTab & process.ReadTransferCount & vbTab & process.WriteTransferCount
Next
