' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、boards、board）。主な用途: WMIサービス接続、マザーボード一覧、マザーボード情報。
Dim service, boards, board
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのマザーボード（Win32_BaseBoard）を読み取り専用で照会し、Manufacturer, Productを取得します。
Set boards = service.ExecQuery( _
    "SELECT Manufacturer, Product FROM Win32_BaseBoard")

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each board In boards
For Each board In boards
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Manufacturer: " & board.Manufacturer
    WScript.Echo "Manufacturer: " & board.Manufacturer
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Model: " & board.Product
    WScript.Echo "Model: " & board.Product
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
