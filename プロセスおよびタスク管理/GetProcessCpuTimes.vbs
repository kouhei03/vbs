' ファイル概要: プロセスごとのユーザーモード・カーネルモードCPU時間を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessCpuTimes.vbs
' 実行結果: プロセスごとのユーザーモード・カーネルモードCPU時間を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、Name, ProcessId, UserModeTime, KernelModeTime を取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, UserModeTime, KernelModeTime FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "User CPU time (100ns)" & vbTab & "Kernel CPU time (100ns)"
For Each process In processes
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
        process.UserModeTime & vbTab & process.KernelModeTime
Next
