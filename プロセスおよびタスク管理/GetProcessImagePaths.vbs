' ファイル概要: プロセスごとの実行ファイルパスを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessImagePaths.vbs
' 実行結果: プロセスごとの実行ファイルパスを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process, imagePath
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、Name, ProcessId, ExecutablePath を取得します。
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, ExecutablePath FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Executable path"
For Each process In processes
    imagePath = process.ExecutablePath
    If IsNull(imagePath) Then imagePath = "Unavailable"
    WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & imagePath
Next
