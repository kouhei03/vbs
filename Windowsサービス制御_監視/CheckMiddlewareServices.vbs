' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、services、item、nameUpper、found）。主な用途: WMIサービス接続、サービス一覧、現在処理中のWMI項目、対象が見つかったかを示すフラグ。
Dim service, services, item, nameUpper, found
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのWindowsサービス（Win32_Service）を読み取り専用で照会し、Name, DisplayName, Stateを取得します。
Set services = service.ExecQuery( _
    "SELECT Name, DisplayName, State FROM Win32_Service")
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
found = False

' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each item In services
For Each item In services
    ' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: nameUpper = UCase(item.Name)
    nameUpper = UCase(item.Name)
    ' 解説: 条件「nameUpper = "W3SVC" Or nameUpper = "WAS" Or nameUpper = "MSSQLSERVER" Or nameUpper = "SQLSERVERAGENT" Or Left(nameUpper, 6) = "MSSQL$" Or Left(nameUpper, 9) = "SQLAGENT$" Then」を評価し、成立した場合に続く分岐処理を実行します。
    If nameUpper = "W3SVC" Or nameUpper = "WAS" Or _
        nameUpper = "MSSQLSERVER" Or nameUpper = "SQLSERVERAGENT" Or _
        Left(nameUpper, 6) = "MSSQL$" Or Left(nameUpper, 9) = "SQLAGENT$" Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
        found = True
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: item.Name & vbTab & item.DisplayName & vbTab & item.State
        WScript.Echo item.Name & vbTab & item.DisplayName & vbTab & item.State
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件「Not found Then WScript.Echo "No matching SQL Server or IIS services found."」を評価し、成立した場合に続く分岐処理を実行します。
If Not found Then WScript.Echo "No matching SQL Server or IIS services found."
