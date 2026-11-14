' ファイル概要: メモリ規格と速度。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 058_MemoryStandardSpeed.vbs
' 実行結果: メモリ規格と速度処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, modules, module, memoryTypeName
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_PhysicalMemory を照会し、DeviceLocator, SMBIOSMemoryType, MemoryType, Speed, ConfiguredClockSpeed を取得します。
Set modules = service.ExecQuery("SELECT DeviceLocator, SMBIOSMemoryType, MemoryType, Speed, ConfiguredClockSpeed FROM Win32_PhysicalMemory")
For Each module In modules
    Select Case module.SMBIOSMemoryType
        Case 24: memoryTypeName = "DDR3"
        Case 26: memoryTypeName = "DDR4"
        Case 30: memoryTypeName = "LPDDR4"
        Case 34: memoryTypeName = "DDR5"
        Case 35: memoryTypeName = "LPDDR5"
        Case 21: memoryTypeName = "DDR2"
        Case 20: memoryTypeName = "DDR"
        Case Else: memoryTypeName = "Other/unknown SMBIOS code " & module.SMBIOSMemoryType
    End Select
    WScript.Echo module.DeviceLocator & ": " & memoryTypeName
    WScript.Echo "Speed=" & module.Speed & " MHz; ConfiguredClockSpeed=" & module.ConfiguredClockSpeed & " MHz"
Next
