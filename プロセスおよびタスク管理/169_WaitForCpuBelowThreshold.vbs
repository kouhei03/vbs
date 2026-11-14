' ファイル概要: CPU負荷が下がるまでの期限付き待機。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 169_WaitForCpuBelowThreshold.vbs <threshold-0-100> <timeout-1-300>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: CPU負荷が下がるまでの期限付き待機処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, counters, counter, threshold, timeoutSeconds, elapsed, cpu, found
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count<>2 Then WScript.Echo "Usage: cscript //nologo 169_WaitForCpuBelowThreshold.vbs <threshold-0-100> <timeout-1-300>": WScript.Quit 2
If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Echo "Both arguments must be numeric.": WScript.Quit 2
threshold=CDbl(WScript.Arguments(0))
timeoutSeconds=CLng(WScript.Arguments(1))
If threshold<0 Or threshold>100 Or timeoutSeconds<1 Or timeoutSeconds>300 Then WScript.Echo "Threshold/timeout out of range.": WScript.Quit 2
Set service=GetObject("winmgmts:\\.\root\cimv2")
For elapsed=0 To timeoutSeconds Step 2
    ' [処理解説] WMIの Win32_PerfFormattedData_PerfOS_Processor を照会し、Name,PercentProcessorTime を取得します。
    Set counters=service.ExecQuery("SELECT Name,PercentProcessorTime FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name='_Total'")
    found=False
    For Each counter In counters
        cpu=CDbl(counter.PercentProcessorTime)
        found=True
        WScript.Echo "CPU=" & cpu & "% elapsed=" & elapsed & "s"
    Next
    If Not found Then WScript.Echo "CPU counter unavailable.": WScript.Quit 1
    If cpu<threshold Then WScript.Echo "Below threshold; safe to continue.": WScript.Quit 0
    ' [処理解説] 次の処理まで指定されたミリ秒だけ待機します。
    If elapsed<timeoutSeconds Then WScript.Sleep 2000
Next
WScript.Echo "Timed out while CPU remained above threshold."
WScript.Quit 1
