' ファイル概要: プロセスごとの所有者を取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetProcessOwners.vbs
' 実行結果: プロセスごとの所有者を取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, processes, process, userName, domainName, result
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set processes = service.ExecQuery( _
    "SELECT Name, ProcessId FROM Win32_Process")

WScript.Echo "Process name" & vbTab & "PID" & vbTab & "Owner"
For Each process In processes
    userName = ""
    domainName = ""
    result = process.GetOwner(userName, domainName)
    If result = 0 Then
        WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
            domainName & "\" & userName
    Else
        WScript.Echo process.Name & vbTab & process.ProcessId & vbTab & _
            "Unavailable (code " & result & ")"
    End If
Next
