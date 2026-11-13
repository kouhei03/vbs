' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、printers、printer）。主な用途: WMIサービス接続、プリンター一覧、プリンター情報。
Dim service, printers, printer
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのプリンター（Win32_Printer）を読み取り専用で照会し、Name, Default, Network, Localを取得します。
Set printers = service.ExecQuery( _
    "SELECT Name, Default, Network, Local FROM Win32_Printer")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each printer In printers
For Each printer In printers
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Printer: " & printer.Name
    WScript.Echo "Printer: " & printer.Name
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Default: " & printer.Default
    WScript.Echo "Default: " & printer.Default
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Network printer: " & printer.Network
    WScript.Echo "Network printer: " & printer.Network
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Local printer: " & printer.Local
    WScript.Echo "Local printer: " & printer.Local
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: ""
    WScript.Echo ""
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
