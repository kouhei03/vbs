' 解説: CPUアーキテクチャコードを読み取り、代表的な値を名称に変換します。
Option Explicit
' 解説: WMI接続、CPU情報、およびアーキテクチャ表示名を保持します。
Dim service, processors, processor, architectureName
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set processors = service.ExecQuery("SELECT Architecture, AddressWidth, DataWidth FROM Win32_Processor")
' 解説: コレクション内の各項目を順に処理します: For Each processor In processors
For Each processor In processors
    ' 解説: 値に応じて対応するケースを選択します: processor.Architecture
    Select Case processor.Architecture
        ' 解説: この値または条件に該当する場合の分岐です: Case 0: architectureName = "x86"
        Case 0: architectureName = "x86"
        ' 解説: この値または条件に該当する場合の分岐です: Case 5: architectureName = "ARM"
        Case 5: architectureName = "ARM"
        ' 解説: この値または条件に該当する場合の分岐です: Case 9: architectureName = "x64"
        Case 9: architectureName = "x64"
        ' 解説: この値または条件に該当する場合の分岐です: Case 12: architectureName = "ARM64"
        Case 12: architectureName = "ARM64"
        ' 解説: この値または条件に該当する場合の分岐です: Case 6: architectureName = "Itanium"
        Case 6: architectureName = "Itanium"
        ' 解説: この値または条件に該当する場合の分岐です: Case Else: architectureName = "Unknown (" & processor.Architecture & ")"
        Case Else: architectureName = "Unknown (" & processor.Architecture & ")"
    ' 解説: Select Caseの選択分岐を終了します。
    End Select
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Architecture=" & architectureName
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "AddressWidth=" & processor.AddressWidth & " DataWidth=" & processor.DataWidth
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
