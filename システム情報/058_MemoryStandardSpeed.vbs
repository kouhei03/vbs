' 解説: DIMMのSMBIOSメモリ種別コードと速度を読み取り、主要な規格コードを名称に変換します。
Option Explicit
' 解説: WMI接続、メモリモジュール、および種別名称を保持します。
Dim service, modules, module, memoryTypeName
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set modules = service.ExecQuery("SELECT DeviceLocator, SMBIOSMemoryType, MemoryType, Speed, ConfiguredClockSpeed FROM Win32_PhysicalMemory")
' 解説: コレクション内の各項目を順に処理します: For Each module In modules
For Each module In modules
    ' 解説: 値に応じて対応するケースを選択します: module.SMBIOSMemoryType
    Select Case module.SMBIOSMemoryType
        ' 解説: この値または条件に該当する場合の分岐です: Case 24: memoryTypeName = "DDR3"
        Case 24: memoryTypeName = "DDR3"
        ' 解説: この値または条件に該当する場合の分岐です: Case 26: memoryTypeName = "DDR4"
        Case 26: memoryTypeName = "DDR4"
        ' 解説: この値または条件に該当する場合の分岐です: Case 30: memoryTypeName = "LPDDR4"
        Case 30: memoryTypeName = "LPDDR4"
        ' 解説: この値または条件に該当する場合の分岐です: Case 34: memoryTypeName = "DDR5"
        Case 34: memoryTypeName = "DDR5"
        ' 解説: この値または条件に該当する場合の分岐です: Case 35: memoryTypeName = "LPDDR5"
        Case 35: memoryTypeName = "LPDDR5"
        ' 解説: この値または条件に該当する場合の分岐です: Case 21: memoryTypeName = "DDR2"
        Case 21: memoryTypeName = "DDR2"
        ' 解説: この値または条件に該当する場合の分岐です: Case 20: memoryTypeName = "DDR"
        Case 20: memoryTypeName = "DDR"
        ' 解説: この値または条件に該当する場合の分岐です: Case Else: memoryTypeName = "Other/unknown SMBIOS code " & module.SMBIOSMemoryType
        Case Else: memoryTypeName = "Other/unknown SMBIOS code " & module.SMBIOSMemoryType
    ' 解説: Select Caseの選択分岐を終了します。
    End Select
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo module.DeviceLocator & ": " & memoryTypeName
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Speed=" & module.Speed & " MHz; ConfiguredClockSpeed=" & module.ConfiguredClockSpeed & " MHz"
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
