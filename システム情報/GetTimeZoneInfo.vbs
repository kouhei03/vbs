' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、zones、zone、services、timeService）。主な用途: WMIサービス接続、タイムゾーン一覧、タイムゾーン情報、サービス一覧、Windows Timeサービス情報。
Dim service, zones, zone, services, timeService
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのタイムゾーン（Win32_TimeZone）を読み取り専用で照会し、Description, Biasを取得します。
Set zones = service.ExecQuery("SELECT Description, Bias FROM Win32_TimeZone")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each zone In zones
For Each zone In zones
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Time zone: " & zone.Description
    WScript.Echo "Time zone: " & zone.Description
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "UTC bias (minutes): " & zone.Bias
    WScript.Echo "UTC bias (minutes): " & zone.Bias
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: WMIのWindowsサービス（Win32_Service）を読み取り専用で照会し、State, StartModeを取得します。 固定絞り込み条件: Name='W32Time'。
Set services = service.ExecQuery( _
    "SELECT State, StartMode FROM Win32_Service WHERE Name='W32Time'")
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each timeService In services
For Each timeService In services
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Windows Time service state: " & timeService.State
    WScript.Echo "Windows Time service state: " & timeService.State
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Windows Time service start mode: " & timeService.StartMode
    WScript.Echo "Windows Time service start mode: " & timeService.StartMode
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
