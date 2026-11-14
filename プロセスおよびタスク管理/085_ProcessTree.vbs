' ファイル概要: 親子プロセス関係のツリー表示。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 085_ProcessTree.vbs
' 実行結果: 親子プロセス関係のツリー表示処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, process, names, parents, pid, parentPid, chain, depth
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Name,ProcessId,ParentProcessId FROM Win32_Process")
Set names = CreateObject("Scripting.Dictionary")
Set parents = CreateObject("Scripting.Dictionary")
For Each process In records
    names(CStr(process.ProcessId)) = process.Name
    parents(CStr(process.ProcessId)) = CStr(process.ParentProcessId)
Next
For Each process In records
    pid = CStr(process.ProcessId)
    parentPid = parents(pid)
    chain = process.Name & " (PID " & pid & ")"
    depth = 0
    Do While names.Exists(parentPid) And depth < 16
        chain = names(parentPid) & " (PID " & parentPid & ") > " & chain
        parentPid = parents(parentPid)
        depth = depth + 1
    Loop
    WScript.Echo chain
Next
