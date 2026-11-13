' 解説: IPEnabledな各アダプターからIPv4アドレスと対応するサブネット情報を読み取ります。
Option Explicit
' 解説: WMI接続、アダプター情報、配列要素、およびサブネット配列を保持します。
Dim service, adapters, adapter, addresses, subnets, index, found
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set adapters = service.ExecQuery("SELECT Description, IPAddress, IPSubnet FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
' 解説: 取得値または計算結果を変数へ格納します。
found = False
' 解説: コレクション内の各項目を順に処理します: For Each adapter In adapters
For Each adapter In adapters
    ' 解説: 取得値または計算結果を変数へ格納します。
    addresses = adapter.IPAddress
    ' 解説: 取得値または計算結果を変数へ格納します。
    subnets = adapter.IPSubnet
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: IsArray(addresses)
    If IsArray(addresses) Then
        ' 解説: 指定範囲の各要素を順番に処理します: For index = LBound(addresses) To UBound(addresses)
        For index = LBound(addresses) To UBound(addresses)
            ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: InStr(addresses(index), ":") = 0
            If InStr(addresses(index), ":") = 0 Then
                ' 解説: 取得値または計算結果を変数へ格納します。
                found = True
                ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
                WScript.Echo adapter.Description & ": " & addresses(index)
                ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: IsArray(subnets)
                If IsArray(subnets) Then
                    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: index <= UBound(subnets)
                    If index <= UBound(subnets) Then WScript.Echo "Subnet: " & subnets(index)
                ' 解説: 条件分岐ブロックを終了します。
                End If
            ' 解説: 条件分岐ブロックを終了します。
            End If
        ' 解説: 現在の反復を終え、次の要素へ進みます。
        Next
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: Not found
If Not found Then WScript.Echo "No IPv4 address was returned."
