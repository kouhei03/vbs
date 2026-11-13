' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、devices、device、found）。主な用途: WMIサービス接続、デバイス一覧、デバイス情報、対象が見つかったかを示すフラグ。
Dim service, devices, device, found
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのPnPデバイス（Win32_PnPEntity）を読み取り専用で照会し、Name, Manufacturer, Statusを取得します。 固定絞り込み条件: PNPClass='USB'。
Set devices = service.ExecQuery( _
    "SELECT Name, Manufacturer, Status FROM Win32_PnPEntity " & _
    "WHERE PNPClass='USB'")

' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
found = False
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each device In devices
For Each device In devices
    ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
    found = True
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Device: " & device.Name
    WScript.Echo "Device: " & device.Name
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Manufacturer: " & device.Manufacturer
    WScript.Echo "Manufacturer: " & device.Manufacturer
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Status: " & device.Status
    WScript.Echo "Status: " & device.Status
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: ""
    WScript.Echo ""
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件「Not found Then WScript.Echo "No USB devices returned by WMI."」を評価し、成立した場合に続く分岐処理を実行します。
If Not found Then WScript.Echo "No USB devices returned by WMI."
