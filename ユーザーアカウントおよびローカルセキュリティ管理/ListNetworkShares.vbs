' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、shares、share）。主な用途: WMIサービス接続、共有一覧、共有情報。
Dim service, shares, share
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの共有フォルダー（Win32_Share）を読み取り専用で照会し、Name, Path, Type, Descriptionを取得します。
Set shares = service.ExecQuery( _
    "SELECT Name, Path, Type, Description FROM Win32_Share")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Share name" & vbTab & "Path" & vbTab & "Type" & vbTab & "Description"
WScript.Echo "Share name" & vbTab & "Path" & vbTab & "Type" & vbTab & "Description"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each share In shares
For Each share In shares
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: share.Name & vbTab & share.Path & vbTab & share.Type & vbTab & share.Description
    WScript.Echo share.Name & vbTab & share.Path & vbTab & _
        share.Type & vbTab & share.Description
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
