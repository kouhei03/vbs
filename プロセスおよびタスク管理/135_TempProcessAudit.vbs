' ファイル概要: TEMP配下の実行中プロセス監査。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 135_TempProcessAudit.vbs
' 実行結果: TEMP配下の実行中プロセス監査処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, processes, process, shell, tempPath, pathValue, found
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Process を照会し、Name,ProcessId,ExecutablePath,CreationDate を取得します。
Set processes = service.ExecQuery("SELECT Name,ProcessId,ExecutablePath,CreationDate FROM Win32_Process")
Set shell = CreateObject("WScript.Shell")
tempPath = LCase(shell.ExpandEnvironmentStrings("%TEMP%"))
found = False
For Each process In processes
    If Not IsNull(process.ExecutablePath) Then
        pathValue = LCase(process.ExecutablePath)
        If Left(pathValue, Len(tempPath)) = tempPath Then
            WScript.Echo process.Name & " PID=" & process.ProcessId & " Path=" & process.ExecutablePath
            found = True
        End If
    End If
Next
If Not found Then WScript.Echo "No process image under the current TEMP path was found."
