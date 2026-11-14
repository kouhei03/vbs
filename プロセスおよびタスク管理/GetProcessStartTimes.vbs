' ファイル概要: プロセスごとの起動時刻を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessStartTimes.vbs
' 実行結果: プロセスごとの起動時刻を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process, dateTime, startTime
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId, CreationDate FROM Win32_Process")
Set dateTime = CreateObject("WbemScripting.SWbemDateTime")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Start time"
For Each process In processes
    If Not IsNull(process.CreationDate) Then
        dateTime.Value = process.CreationDate
        startTime = dateTime.GetVarDate(True)
        WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
            CStr(startTime)
    End If
Next
