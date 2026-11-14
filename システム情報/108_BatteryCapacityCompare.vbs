' ファイル概要: 設計容量と満充電容量。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 108_BatteryCapacityCompare.vbs
' 実行結果: 設計容量と満充電容量処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, designRows, fullRows, row, foundDesign, foundFull
Set service = GetObject("winmgmts:\\.\root\wmi")
' [処理解説] WMIの BatteryStaticData を照会し、DesignedCapacity を取得します。
Set designRows = service.ExecQuery("SELECT DesignedCapacity FROM BatteryStaticData")
' [処理解説] WMIの BatteryFullChargedCapacity を照会し、FullChargedCapacity を取得します。
Set fullRows = service.ExecQuery("SELECT FullChargedCapacity FROM BatteryFullChargedCapacity")
foundDesign = False
foundFull = False
For Each row In designRows
    WScript.Echo "DesignedCapacity=" & row.DesignedCapacity
    foundDesign = True
Next
For Each row In fullRows
    WScript.Echo "FullChargedCapacity=" & row.FullChargedCapacity
    foundFull = True
Next
If Not foundDesign Or Not foundFull Then
    WScript.Echo "Battery capacity data is not available from this device/provider."
    WScript.Quit 1
End If
