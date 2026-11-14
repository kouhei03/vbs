' ファイル概要: バッテリー劣化度算出。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 109_BatteryHealth.vbs
' 実行結果: バッテリー劣化度算出処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, designRows, fullRows, row, designCapacity, fullCapacity, foundDesign, foundFull
Set service = GetObject("winmgmts:\\.\root\wmi")
' [処理解説] WMIの BatteryStaticData を照会し、DesignedCapacity を取得します。
Set designRows = service.ExecQuery("SELECT DesignedCapacity FROM BatteryStaticData")
' [処理解説] WMIの BatteryFullChargedCapacity を照会し、FullChargedCapacity を取得します。
Set fullRows = service.ExecQuery("SELECT FullChargedCapacity FROM BatteryFullChargedCapacity")
foundDesign = False
foundFull = False
For Each row In designRows
    designCapacity = CDbl(row.DesignedCapacity)
    foundDesign = True
    Exit For
Next
For Each row In fullRows
    fullCapacity = CDbl(row.FullChargedCapacity)
    foundFull = True
    Exit For
Next
If Not foundDesign Or Not foundFull Or designCapacity <= 0 Then
    WScript.Echo "Battery capacity data is unavailable or invalid."
    WScript.Quit 1
End If
WScript.Echo "Health percent=" & FormatNumber((fullCapacity / designCapacity) * 100, 2)
WScript.Echo "Degradation percent=" & FormatNumber((1 - fullCapacity / designCapacity) * 100, 2)
