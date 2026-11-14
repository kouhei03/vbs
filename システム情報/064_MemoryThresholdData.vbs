' ファイル概要: しきい値超過判定用データ（通知なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 064_MemoryThresholdData.vbs [threshold-percent]
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: しきい値超過判定用データ（通知なし）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, systems, memories, system, memory, totalBytes, availableBytes, threshold, usedPercent, hasTotal, hasAvailable
If WScript.Arguments.Count > 1 Then
    WScript.Echo "Usage: cscript //nologo 064_MemoryThresholdData.vbs [threshold-percent]"
    WScript.Quit 2
End If
threshold = 85
hasTotal = False
hasAvailable = False
If WScript.Arguments.Count = 1 Then
    If Not IsNumeric(WScript.Arguments(0)) Then
        WScript.Echo "Threshold must be a number from 0 to 100."
        WScript.Quit 2
    End If
    threshold = CDbl(WScript.Arguments(0))
    If threshold < 0 Or threshold > 100 Then
        WScript.Echo "Threshold must be between 0 and 100."
        WScript.Quit 2
    End If
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set systems = service.ExecQuery("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem")
Set memories = service.ExecQuery("SELECT AvailableBytes FROM Win32_PerfFormattedData_PerfOS_Memory")
For Each system In systems
    totalBytes = CDbl(system.TotalPhysicalMemory)
    hasTotal = True
Next
For Each memory In memories
    availableBytes = CDbl(memory.AvailableBytes)
    hasAvailable = True
Next
If Not hasTotal Or Not hasAvailable Then
    WScript.Echo "Total or available memory data is unavailable."
    WScript.Quit 1
End If
If totalBytes <= 0 Then
    WScript.Echo "Total memory is invalid."
    WScript.Quit 1
End If
usedPercent = ((totalBytes - availableBytes) / totalBytes) * 100
WScript.Echo "Total bytes: " & totalBytes
WScript.Echo "Available bytes: " & availableBytes
WScript.Echo "Used percent: " & FormatNumber(usedPercent, 2)
WScript.Echo "Warning threshold percent: " & threshold
WScript.Echo "Threshold exceeded: " & (usedPercent >= threshold)
