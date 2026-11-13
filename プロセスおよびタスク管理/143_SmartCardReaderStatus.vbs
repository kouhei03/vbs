' 解説: ローカルPnP情報からスマートカードリーダー候補と状態を照会します。認証やカード操作はしません。
Option Explicit
' 解説: このスクリプトで使う変数を宣言します: service, devices, device
Dim service, devices, device
' 解説: ローカルWMI/CIMまたはWindowsオブジェクトに接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定条件に合う情報を読み取り専用で照会します。
Set devices = service.ExecQuery("SELECT Name,PNPClass,Status,PNPDeviceID FROM Win32_PnPEntity WHERE PNPClass='SmartCardReader' OR Name LIKE '%smart card%'")
' 解説: 条件を判定し、必要な分岐を選択します: If devices.Count=0
If devices.Count=0 Then WScript.Echo "No smart-card reader device was reported."
' 解説: コレクション内の各項目を順番に処理します: For Each device In devices
For Each device In devices
    ' 解説: 状態、診断結果、またはエラーをコンソールへ表示します。
    WScript.Echo device.Name & " Status=" & device.Status & " PNPClass=" & device.PNPClass
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
