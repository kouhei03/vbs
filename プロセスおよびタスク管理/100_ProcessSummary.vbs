' ファイル概要: プロセス集計サマリー。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 100_ProcessSummary.vbs
' 実行結果: プロセス集計サマリー処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, processes, process, count, bytes
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、ProcessId,WorkingSetSize を取得します。
Set processes = service.ExecQuery("SELECT ProcessId,WorkingSetSize FROM Win32_Process")
count = 0
bytes = CDbl(0)
For Each process In processes
    count = count + 1
    If Not IsNull(process.WorkingSetSize) Then bytes = bytes + CDbl(process.WorkingSetSize)
Next
WScript.Echo "Process count=" & count
WScript.Echo "Total working set bytes=" & FormatNumber(bytes,0)
