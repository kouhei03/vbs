' 解説: CPU総使用率がしきい値未満になるまで最大300秒監視し、負荷が高ければ処理を待機します。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, counters, counter, threshold, timeoutSeconds, elapsed, cpu, found
Dim service, counters, counter, threshold, timeoutSeconds, elapsed, cpu, found
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count<>2
If WScript.Arguments.Count<>2 Then WScript.Echo "Usage: cscript //nologo 169_WaitForCpuBelowThreshold.vbs <threshold-0-100> <timeout-1-300>": WScript.Quit 2
' 解説: 条件を判定し、必要な分岐を選択します: If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1))
If Not IsNumeric(WScript.Arguments(0)) Or Not IsNumeric(WScript.Arguments(1)) Then WScript.Echo "Both arguments must be numeric.": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
threshold=CDbl(WScript.Arguments(0))
' 解説: 取得値または計算結果を変数に格納します。
timeoutSeconds=CLng(WScript.Arguments(1))
' 解説: 条件を判定し、必要な分岐を選択します: If threshold<0 Or threshold>100 Or timeoutSeconds<1 Or timeoutSeconds>300
If threshold<0 Or threshold>100 Or timeoutSeconds<1 Or timeoutSeconds>300 Then WScript.Echo "Threshold/timeout out of range.": WScript.Quit 2
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service=GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定範囲の要素を順番に処理します: For elapsed=0 To timeoutSeconds Step 2
For elapsed=0 To timeoutSeconds Step 2
    ' 解説: 指定条件に合う情報を読み取り専用で照会します。
    Set counters=service.ExecQuery("SELECT Name,PercentProcessorTime FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name='_Total'")
    ' 解説: 取得値または計算結果を変数に格納します。
    found=False
    ' 解説: コレクション内の各項目を順番に処理します: For Each counter In counters
    For Each counter In counters
        ' 解説: 取得値または計算結果を変数に格納します。
        cpu=CDbl(counter.PercentProcessorTime)
        ' 解説: 取得値または計算結果を変数に格納します。
        found=True
        ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
        WScript.Echo "CPU=" & cpu & "% elapsed=" & elapsed & "s"
    ' 解説: 現在の反復を終え、次の要素へ進みます。
    Next
    ' 解説: 条件を判定し、必要な分岐を選択します: If Not found
    If Not found Then WScript.Echo "CPU counter unavailable.": WScript.Quit 1
    ' 解説: 条件を判定し、必要な分岐を選択します: If cpu<threshold
    If cpu<threshold Then WScript.Echo "Below threshold; safe to continue.": WScript.Quit 0
    ' 解説: 条件を判定し、必要な分岐を選択します: If elapsed<timeoutSeconds
    If elapsed<timeoutSeconds Then WScript.Sleep 2000
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Timed out while CPU remained above threshold."
' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
WScript.Quit 1
