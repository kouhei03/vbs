' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、counters、counter）。主な用途: WMIサービス接続、パフォーマンスカウンター一覧、パフォーマンスカウンター。
Dim service, counters, counter
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのプロセス別CPUカウンター（Win32_PerfFormattedData_PerfProc_Process）を読み取り専用で照会し、Name, IDProcess, PercentProcessorTimeを取得します。 固定絞り込み条件: IDProcess > 0。
Set counters = service.ExecQuery( _
    "SELECT Name, IDProcess, PercentProcessorTime " & _
    "FROM Win32_PerfFormattedData_PerfProc_Process WHERE IDProcess > 0")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Process name" & vbTab & "PID" & vbTab & "CPU counter (%)"
WScript.Echo "Process name" & vbTab & "PID" & vbTab & _
    "CPU counter (%)"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each counter In counters
For Each counter In counters
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: counter.Name & vbTab & counter.IDProcess & vbTab & counter.PercentProcessorTime
    WScript.Echo counter.Name & vbTab & counter.IDProcess & vbTab & _
        counter.PercentProcessorTime
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
