' ファイル概要: CPU高負荷の有界サンプリング。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 082_HighCpuMonitor.vbs <process.exe>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: CPU高負荷の有界サンプリング処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, targets, target, samples, sample, counter, processName, validator, i, found
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 082_HighCpuMonitor.vbs <process.exe>": WScript.Quit 2
processName = WScript.Arguments(0)
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(processName) Then WScript.Echo "Invalid executable name.": WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
For i = 1 To 5
    Set targets = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
    found = False
    For Each target In targets
        Set samples = service.ExecQuery("SELECT Name,IDProcess,PercentProcessorTime FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess = " & target.ProcessId)
        For Each counter In samples
            WScript.Echo "PID=" & target.ProcessId & " CPU counter=" & counter.PercentProcessorTime & "% (sample " & i & ")"
            found = True
        Next
    Next
    If Not found Then WScript.Echo "No CPU sample returned for " & processName
    If i < 5 Then WScript.Sleep 1000
Next
WScript.Echo "Counters are provider-reported and may exceed 100% across cores; no action was taken."
