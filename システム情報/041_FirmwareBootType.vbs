' 解説: UEFI/レガシーブート方式をWMIから読み取り、既知のFirmwareType値を名称に変換します。
Option Explicit
' 解説: WMI接続、システム情報、および表示名を保持します。
Dim service, systems, system, bootType
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set systems = service.ExecQuery("SELECT FirmwareType FROM Win32_ComputerSystem")
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: systems.Count = 0
If systems.Count = 0 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "FirmwareType is unavailable."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: コレクション内の各項目を順に処理します: For Each system In systems
For Each system In systems
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: IsNull(system.FirmwareType)
    If IsNull(system.FirmwareType) Then
        ' 解説: 取得値または計算結果を変数へ格納します。
        bootType = "Unknown; firmware type was not reported"
    ' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
    Else
        ' 解説: 値に応じて対応するケースを選択します: system.FirmwareType
        Select Case system.FirmwareType
        ' 解説: この値または条件に該当する場合の分岐です: Case 2: bootType = "Legacy BIOS"
        Case 2: bootType = "Legacy BIOS"
        ' 解説: この値または条件に該当する場合の分岐です: Case 3: bootType = "UEFI"
        Case 3: bootType = "UEFI"
        ' 解説: この値または条件に該当する場合の分岐です: Case Else: bootType = "Unknown or unsupported: " & system.FirmwareType
        Case Else: bootType = "Unknown or unsupported: " & system.FirmwareType
        ' 解説: Select Caseの選択分岐を終了します。
        End Select
    ' 解説: 条件分岐ブロックを終了します。
    End If
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Firmware boot type: " & bootType
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
