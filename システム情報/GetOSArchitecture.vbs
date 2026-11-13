' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、systems、system）。主な用途: WMIサービス接続、OS情報一覧、OS情報。
Dim service, systems, system
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindows OS（Win32_OperatingSystem）を読み取り専用で照会し、OSArchitectureを取得します。
Set systems = service.ExecQuery("SELECT OSArchitecture FROM Win32_OperatingSystem")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each system In systems
For Each system In systems
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "OS architecture: " & system.OSArchitecture
    WScript.Echo "OS architecture: " & system.OSArchitecture
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
