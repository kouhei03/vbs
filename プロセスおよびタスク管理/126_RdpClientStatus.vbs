' ファイル概要: ローカルRDPクライアント状態。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 126_RdpClientStatus.vbs
' 実行結果: ローカルRDPクライアント状態処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery("SELECT Name,ProcessId,CreationDate FROM Win32_Process WHERE Name='mstsc.exe'")
If processes.Count = 0 Then WScript.Echo "mstsc.exe is not running."
For Each process In processes
    WScript.Echo process.Name & " PID=" & process.ProcessId & " Created=" & process.CreationDate
Next
