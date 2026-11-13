' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、bioses、bios）。主な用途: WMIサービス接続、BIOS情報。
Dim service, bioses, bios
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのBIOS（Win32_BIOS）を読み取り専用で照会し、Manufacturer, SMBIOSBIOSVersion, ReleaseDateを取得します。
Set bioses = service.ExecQuery( _
    "SELECT Manufacturer, SMBIOSBIOSVersion, ReleaseDate FROM Win32_BIOS")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each bios In bioses
For Each bios In bioses
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Manufacturer: " & bios.Manufacturer
    WScript.Echo "Manufacturer: " & bios.Manufacturer
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Version: " & bios.SMBIOSBIOSVersion
    WScript.Echo "Version: " & bios.SMBIOSBIOSVersion
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Release date (WMI format): " & bios.ReleaseDate
    WScript.Echo "Release date (WMI format): " & bios.ReleaseDate
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
