' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、batteries、battery、found、batteryState）。主な用途: WMIサービス接続、バッテリー一覧、バッテリー情報、対象が見つかったかを示すフラグ。
Dim service, batteries, battery, found, batteryState
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのバッテリー（Win32_Battery）を読み取り専用で照会し、EstimatedChargeRemaining, BatteryStatus, Statusを取得します。
Set batteries = service.ExecQuery( _
    "SELECT EstimatedChargeRemaining, BatteryStatus, Status FROM Win32_Battery")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
found = False
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each battery In batteries
For Each battery In batteries
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
    found = True
    ' 解説: 値ごとに処理を分ける選択分岐を開始します: battery.BatteryStatus
    Select Case battery.BatteryStatus
        ' 解説: Select Caseの値が「1: batteryState = "Other"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 1: batteryState = "Other"
        ' 解説: Select Caseの値が「2: batteryState = "Unknown"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 2: batteryState = "Unknown"
        ' 解説: Select Caseの値が「3: batteryState = "Fully charged"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 3: batteryState = "Fully charged"
        ' 解説: Select Caseの値が「4: batteryState = "Low"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 4: batteryState = "Low"
        ' 解説: Select Caseの値が「5: batteryState = "Critical"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 5: batteryState = "Critical"
        ' 解説: Select Caseの値が「6: batteryState = "Charging"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 6: batteryState = "Charging"
        ' 解説: Select Caseの値が「7: batteryState = "Charging and high"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 7: batteryState = "Charging and high"
        ' 解説: Select Caseの値が「8: batteryState = "Charging and low"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 8: batteryState = "Charging and low"
        ' 解説: Select Caseの値が「9: batteryState = "Charging and critical"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 9: batteryState = "Charging and critical"
        ' 解説: Select Caseの値が「10: batteryState = "Undefined"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 10: batteryState = "Undefined"
        ' 解説: Select Caseの値が「11: batteryState = "Partially charged"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
        Case 11: batteryState = "Partially charged"
        ' 解説: Select CaseのどのCase値にも一致しない場合に実行する既定分岐です。
        Case Else: batteryState = "Not reported"
    ' 解説: Select Caseの選択分岐を終了します。
    End Select
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Charge remaining (%): " & battery.EstimatedChargeRemaining
    WScript.Echo "Charge remaining (%): " & battery.EstimatedChargeRemaining
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Battery state: " & batteryState
    WScript.Echo "Battery state: " & batteryState
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Health/status: " & battery.Status
    WScript.Echo "Health/status: " & battery.Status
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件「Not found Then WScript.Echo "No battery reported (or battery data is unavailable)."」を評価し、成立した場合に続く分岐処理を実行します。
If Not found Then WScript.Echo "No battery reported (or battery data is unavailable)."
