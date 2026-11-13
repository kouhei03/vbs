' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、processors、processor）。主な用途: WMIサービス接続、CPU一覧、CPU情報。
Dim service, processors, processor
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのCPU（Win32_Processor）を読み取り専用で照会し、Name, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessorsを取得します。
Set processors = service.ExecQuery( _
    "SELECT Name, MaxClockSpeed, NumberOfCores, NumberOfLogicalProcessors " & _
    "FROM Win32_Processor")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each processor In processors
For Each processor In processors
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Model: " & processor.Name
    WScript.Echo "Model: " & processor.Name
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Maximum clock speed (MHz): " & processor.MaxClockSpeed
    WScript.Echo "Maximum clock speed (MHz): " & processor.MaxClockSpeed
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Cores: " & processor.NumberOfCores
    WScript.Echo "Cores: " & processor.NumberOfCores
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Logical processors: " & processor.NumberOfLogicalProcessors
    WScript.Echo "Logical processors: " & processor.NumberOfLogicalProcessors
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: ""
    WScript.Echo ""
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
