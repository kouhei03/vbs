' 解説: メモリ使用率と指定しきい値を読み取り、警告判定用データを表示します。通知や設定変更はしません。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: service, systems, memories, system, memory, totalBytes, availableBytes, threshold, usedPercent, hasTotal, hasAvailable
Dim service, systems, memories, system, memory, totalBytes, availableBytes, threshold, usedPercent, hasTotal, hasAvailable
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: WScript.Arguments.Count > 1
If WScript.Arguments.Count > 1 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Usage: cscript //nologo 064_MemoryThresholdData.vbs [threshold-percent]"
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 取得値または計算結果を変数へ格納します。
threshold = 85
' 解説: 取得値または計算結果を変数へ格納します。
hasTotal = False
' 解説: 取得値または計算結果を変数へ格納します。
hasAvailable = False
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: WScript.Arguments.Count = 1
If WScript.Arguments.Count = 1 Then
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: Not IsNumeric(WScript.Arguments(0))
    If Not IsNumeric(WScript.Arguments(0)) Then
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo "Threshold must be a number from 0 to 100."
        ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
        WScript.Quit 2
    ' 解説: 条件分岐ブロックを終了します。
    End If
    ' 解説: 取得値または計算結果を変数へ格納します。
    threshold = CDbl(WScript.Arguments(0))
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: threshold < 0 Or threshold > 100
    If threshold < 0 Or threshold > 100 Then
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo "Threshold must be between 0 and 100."
        ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
        WScript.Quit 2
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set systems = service.ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
' 解説: 指定した情報を読み取り専用で照会します。
Set memories = service.ExecQuery("SELECT AvailableBytes FROM Win32_PerfFormattedData_PerfOS_Memory")
' 解説: コレクション内の各項目を順に処理します: For Each system In systems
For Each system In systems
    ' 解説: 取得値または計算結果を変数へ格納します。
    totalBytes = CDbl(system.TotalPhysicalMemory)
    ' 解説: 取得値または計算結果を変数へ格納します。
    hasTotal = True
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: コレクション内の各項目を順に処理します: For Each memory In memories
For Each memory In memories
    ' 解説: 取得値または計算結果を変数へ格納します。
    availableBytes = CDbl(memory.AvailableBytes)
    ' 解説: 取得値または計算結果を変数へ格納します。
    hasAvailable = True
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: Not hasTotal Or Not hasAvailable
If Not hasTotal Or Not hasAvailable Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Total or available memory data is unavailable."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: totalBytes <= 0
If totalBytes <= 0 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Total memory is invalid."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 取得値または計算結果を変数へ格納します。
usedPercent = ((totalBytes - availableBytes) / totalBytes) * 100
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "Total bytes: " & totalBytes
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "Available bytes: " & availableBytes
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "Used percent: " & FormatNumber(usedPercent, 2)
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "Warning threshold percent: " & threshold
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "Threshold exceeded: " & (usedPercent >= threshold)
