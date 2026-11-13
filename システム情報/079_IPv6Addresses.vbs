' 解説: IPEnabledな各アダプターのアドレスからIPv6形式だけを選んで表示します。
Option Explicit
' 解説: WMI接続、アダプター情報、およびアドレス配列を保持します。
Dim service, adapters, adapter, address, found
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: 指定した情報を読み取り専用で照会します。
Set adapters = service.ExecQuery("SELECT Description, IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
' 解説: 取得値または計算結果を変数へ格納します。
found = False
' 解説: コレクション内の各項目を順に処理します: For Each adapter In adapters
For Each adapter In adapters
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: IsArray(adapter.IPAddress)
    If IsArray(adapter.IPAddress) Then
        ' 解説: コレクション内の各項目を順に処理します: For Each address In adapter.IPAddress
        For Each address In adapter.IPAddress
            ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: InStr(address, ":") > 0
            If InStr(address, ":") > 0 Then
                ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
                WScript.Echo adapter.Description & ": " & address
                ' 解説: 取得値または計算結果を変数へ格納します。
                found = True
            ' 解説: 条件分岐ブロックを終了します。
            End If
        ' 解説: 現在の反復を終え、次の要素へ進みます。
        Next
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件を評価し、該当する場合の分岐処理を選択します: Not found
If Not found Then WScript.Echo "No IPv6 addresses were returned."
