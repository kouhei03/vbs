' ファイル概要: プロセスメモリの有界サンプリング。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 083_MemoryTrend.vbs <process.exe>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: プロセスメモリの有界サンプリング処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, process, processName, validator, i, found
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 083_MemoryTrend.vbs <process.exe>": WScript.Quit 2
processName = WScript.Arguments(0)
Set validator = New RegExp
' [処理解説] 引数に許可する文字を限定して名前を検証します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Then WScript.Echo "Invalid executable name.": WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
For i = 1 To 5
    ' [処理解説] WMIの Win32_Process を照会し、ProcessId,WorkingSetSize を取得します。
    Set records = service.ExecQuery("SELECT ProcessId,WorkingSetSize FROM Win32_Process WHERE Name = '" & processName & "'")
    found = False
    For Each process In records
        WScript.Echo "Sample=" & i & " PID=" & process.ProcessId & " WorkingSetBytes=" & process.WorkingSetSize
        found = True
    Next
    If Not found Then WScript.Echo "Process not running at sample " & i
    ' [処理解説] 次の処理まで指定されたミリ秒だけ待機します。
    If i < 5 Then WScript.Sleep 2000
Next
WScript.Echo "Samples are diagnostic only and do not assert a memory leak."
