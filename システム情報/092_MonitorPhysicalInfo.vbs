' 解説: WmiMonitorIDの数値配列を文字列に変換して、接続モニター情報を表示します。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: service, monitors, monitor, fieldName, values, code, decoded, i
Dim service, monitors, monitor, fieldName, values, code, decoded, i
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\wmi")
' 解説: 指定した情報を読み取り専用で照会します。
Set monitors = service.ExecQuery("SELECT ManufacturerName, ProductCodeID, SerialNumberID, UserFriendlyName FROM WmiMonitorID")
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: monitors.Count = 0
If monitors.Count = 0 Then WScript.Echo "No monitor identification data returned."
' 解説: コレクション内の各項目を順に処理します: For Each monitor In monitors
For Each monitor In monitors
    ' 解説: コレクション内の各項目を順に処理します: For Each fieldName In Array("ManufacturerName", "ProductCodeID", "SerialNumberID", "UserFriendlyName")
    For Each fieldName In Array("ManufacturerName", "ProductCodeID", "SerialNumberID", "UserFriendlyName")
        ' 解説: 取得値または計算結果を変数へ格納します。
        values = monitor.Properties_.Item(fieldName).Value
        ' 解説: 取得値または計算結果を変数へ格納します。
        decoded = ""
        ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: IsArray(values)
        If IsArray(values) Then
            ' 解説: コレクション内の各項目を順に処理します: For Each code In values
            For Each code In values
                ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: CLng(code) <> 0
                If CLng(code) <> 0 Then decoded = decoded & ChrW(CLng(code))
            ' 解説: 現在の反復を終え、次の要素へ進みます。
            Next
        ' 解説: 条件分岐ブロックを終了します。
        End If
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo fieldName & ": " & decoded
    ' 解説: 現在の反復を終え、次の要素へ進みます。
    Next
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
