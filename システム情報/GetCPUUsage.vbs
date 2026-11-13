' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、counters、counter）。主な用途: WMIサービス接続、パフォーマンスカウンター一覧、パフォーマンスカウンター。
Dim service, counters, counter
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのCPU使用率カウンター（Win32_PerfFormattedData_PerfOS_Processor）を読み取り専用で照会し、Name, PercentProcessorTimeを取得します。 固定絞り込み条件: Name='_Total'。
Set counters = service.ExecQuery( _
    "SELECT Name, PercentProcessorTime FROM " & _
    "Win32_PerfFormattedData_PerfOS_Processor WHERE Name='_Total'")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each counter In counters
For Each counter In counters
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Current total CPU usage (%): " & counter.PercentProcessorTime
    WScript.Echo "Current total CPU usage (%): " & counter.PercentProcessorTime
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
