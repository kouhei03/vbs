' 解説: メモリフォームファクターコードを読み取り、DIMM/SODIMM等の名称に変換します。
Option Explicit
' 解説: WMI接続、メモリモジュール、およびフォームファクター名を保持します。
Dim service, modules, module, formName
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set modules = service.ExecQuery("SELECT DeviceLocator, FormFactor FROM Win32_PhysicalMemory")
' 解説: コレクション内の各項目を順に処理します: For Each module In modules
For Each module In modules
    ' 解説: 値に応じて対応するケースを選択します: module.FormFactor
    Select Case module.FormFactor
        ' 解説: この値または条件に該当する場合の分岐です: Case 8: formName = "DIMM"
        Case 8: formName = "DIMM"
        ' 解説: この値または条件に該当する場合の分岐です: Case 12: formName = "SODIMM"
        Case 12: formName = "SODIMM"
        ' 解説: この値または条件に該当する場合の分岐です: Case 9: formName = "RIMM"
        Case 9: formName = "RIMM"
        ' 解説: この値または条件に該当する場合の分岐です: Case 13: formName = "SRIMM"
        Case 13: formName = "SRIMM"
        ' 解説: この値または条件に該当する場合の分岐です: Case Else: formName = "Other/unknown code " & module.FormFactor
        Case Else: formName = "Other/unknown code " & module.FormFactor
    ' 解説: Select Caseの選択分岐を終了します。
    End Select
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo module.DeviceLocator & ": " & formName
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
