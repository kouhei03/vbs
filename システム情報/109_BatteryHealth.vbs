' 解説: 設計容量と満充電容量からバッテリー健康度の算出用比率を表示します。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: service, designRows, fullRows, row, designCapacity, fullCapacity, foundDesign, foundFull
Dim service, designRows, fullRows, row, designCapacity, fullCapacity, foundDesign, foundFull
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\wmi")
' 解説: 指定した情報を読み取り専用で照会します。
Set designRows = service.ExecQuery("SELECT DesignedCapacity FROM BatteryStaticData")
' 解説: 指定した情報を読み取り専用で照会します。
Set fullRows = service.ExecQuery("SELECT FullChargedCapacity FROM BatteryFullChargedCapacity")
' 解説: 取得値または計算結果を変数へ格納します。
foundDesign = False
' 解説: 取得値または計算結果を変数へ格納します。
foundFull = False
' 解説: コレクション内の各項目を順に処理します: For Each row In designRows
For Each row In designRows
    ' 解説: 取得値または計算結果を変数へ格納します。
    designCapacity = CDbl(row.DesignedCapacity)
    ' 解説: 取得値または計算結果を変数へ格納します。
    foundDesign = True
    ' 解説: 現在のループまたは処理を終了します。
    Exit For
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: コレクション内の各項目を順に処理します: For Each row In fullRows
For Each row In fullRows
    ' 解説: 取得値または計算結果を変数へ格納します。
    fullCapacity = CDbl(row.FullChargedCapacity)
    ' 解説: 取得値または計算結果を変数へ格納します。
    foundFull = True
    ' 解説: 現在のループまたは処理を終了します。
    Exit For
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: Not foundDesign Or Not foundFull Or designCapacity <= 0
If Not foundDesign Or Not foundFull Or designCapacity <= 0 Then
    ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
    WScript.Echo "Battery capacity data is unavailable or invalid."
    ' 解説: スクリプトを終了し、呼び出し元へ終了コードを返します。
    WScript.Quit 1
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "Health percent=" & FormatNumber((fullCapacity / designCapacity) * 100, 2)
' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
WScript.Echo "Degradation percent=" & FormatNumber((1 - fullCapacity / designCapacity) * 100, 2)
