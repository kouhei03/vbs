' ファイル概要: CPU高負荷の有界サンプリング。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 082_HighCpuMonitor.vbs <process.exe>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: CPU高負荷の有界サンプリング処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, targets, target, samples, sample, counter, processName, validator, i, found
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 082_HighCpuMonitor.vbs <process.exe>": WScript.Quit 2
processName = WScript.Arguments(0)
Set validator = New RegExp
' [処理解説] 引数に許可する文字を限定して名前を検証します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Then WScript.Echo "Invalid executable name.": WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
For i = 1 To 5
    ' [処理解説] WMIの Win32_Process を照会し、ProcessId を取得します。
    Set targets = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
    found = False
    For Each target In targets
        ' [処理解説] WMIの Win32_PerfFormattedData_PerfProc_Process を照会し、Name,IDProcess,PercentProcessorTime を取得します。
        Set samples = service.ExecQuery("SELECT Name,IDProcess,PercentProcessorTime FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess = " & target.ProcessId)
        For Each counter In samples
            WScript.Echo "PID=" & target.ProcessId & " CPU counter=" & counter.PercentProcessorTime & "% (sample " & i & ")"
            found = True
        Next
    Next
    If Not found Then WScript.Echo "No CPU sample returned for " & processName
    ' [処理解説] 次の処理まで指定されたミリ秒だけ待機します。
    If i < 5 Then WScript.Sleep 1000
Next
WScript.Echo "Counters are provider-reported and may exceed 100% across cores; no action was taken."
