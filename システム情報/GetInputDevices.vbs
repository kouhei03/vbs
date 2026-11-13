' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、devices、device）。主な用途: WMIサービス接続、デバイス一覧、デバイス情報。
Dim service, devices, device
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Keyboards:"
WScript.Echo "Keyboards:"
' 解説: WMIのキーボード（Win32_Keyboard）を読み取り専用で照会し、Descriptionを取得します。
Set devices = service.ExecQuery("SELECT Description FROM Win32_Keyboard")
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each device In devices
For Each device In devices
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: " " & device.Description
    WScript.Echo "  " & device.Description
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Pointing devices:"
WScript.Echo "Pointing devices:"
' 解説: WMIのポインティングデバイス（Win32_PointingDevice）を読み取り専用で照会し、Descriptionを取得します。
Set devices = service.ExecQuery("SELECT Description FROM Win32_PointingDevice")
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each device In devices
For Each device In devices
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: " " & device.Description
    WScript.Echo "  " & device.Description
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
