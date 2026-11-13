' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、devices、device）。主な用途: WMIサービス接続、デバイス一覧、デバイス情報。
Dim service, devices, device
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのサウンドデバイス（Win32_SoundDevice）を読み取り専用で照会し、Name, Manufacturer, Statusを取得します。
Set devices = service.ExecQuery( _
    "SELECT Name, Manufacturer, Status FROM Win32_SoundDevice")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each device In devices
For Each device In devices
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
