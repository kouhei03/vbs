' 解説: 指定プロセスのCPUカウンターを最大5回サンプリングし、高負荷状態を報告します。終了・再起動は行いません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, targets, target, samples, sample, counter, processName, validator, i, found
Dim service, targets, target, samples, sample, counter, processName, validator, i, found
' 解説: 条件を判定し、必要な分岐を選択します: If WScript.Arguments.Count <> 1
If WScript.Arguments.Count <> 1 Then WScript.Echo "Usage: cscript //nologo 082_HighCpuMonitor.vbs <process.exe>": WScript.Quit 2
' 解説: 取得値または計算結果を変数に格納します。
processName = WScript.Arguments(0)
' 解説: 後続処理で利用するオブジェクト参照を設定します。
Set validator = New RegExp
' 解説: 取得値または計算結果を変数に格納します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
' 解説: 条件を判定し、必要な分岐を選択します: If Not validator.Test(processName)
If Not validator.Test(processName) Then WScript.Echo "Invalid executable name.": WScript.Quit 2
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定範囲の要素を順番に処理します: For i = 1 To 5
For i = 1 To 5
    ' 解説: 指定条件に合う情報を読み取り専用で照会します。
    Set targets = service.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = '" & processName & "'")
    ' 解説: 取得値または計算結果を変数に格納します。
    found = False
    ' 解説: コレクション内の各項目を順番に処理します: For Each target In targets
    For Each target In targets
        ' 解説: 指定条件に合う情報を読み取り専用で照会します。
        Set samples = service.ExecQuery("SELECT Name,IDProcess,PercentProcessorTime FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess = " & target.ProcessId)
        ' 解説: コレクション内の各項目を順番に処理します: For Each counter In samples
        For Each counter In samples
            ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
            WScript.Echo "PID=" & target.ProcessId & " CPU counter=" & counter.PercentProcessorTime & "% (sample " & i & ")"
            ' 解説: 取得値または計算結果を変数に格納します。
            found = True
        ' 解説: 現在の反復を終え、次の要素へ進みます。
        Next
    ' 解説: 現在の反復を終え、次の要素へ進みます。
    Next
    ' 解説: 条件を判定し、必要な分岐を選択します: If Not found
    If Not found Then WScript.Echo "No CPU sample returned for " & processName
    ' 解説: 条件を判定し、必要な分岐を選択します: If i < 5
    If i < 5 Then WScript.Sleep 1000
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
WScript.Echo "Counters are provider-reported and may exceed 100% across cores; no action was taken."
