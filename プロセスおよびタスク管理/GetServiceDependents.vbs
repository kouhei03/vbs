' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、relations、relation）。主な用途: WMIサービス接続。
Dim service, relations, relation
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのサービス依存関係（Win32_DependentService）を読み取り専用で照会し、Antecedent, Dependentを取得します。
Set relations = service.ExecQuery("SELECT Antecedent, Dependent FROM Win32_DependentService")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Service reference" & vbTab & "Depends-on service reference"
WScript.Echo "Service reference" & vbTab & "Depends-on service reference"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each relation In relations
For Each relation In relations
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: relation.Dependent & vbTab & relation.Antecedent
    WScript.Echo relation.Dependent & vbTab & relation.Antecedent
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
