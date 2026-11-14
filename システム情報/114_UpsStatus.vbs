' ファイル概要: UPSデバイス状態。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 114_UpsStatus.vbs
' 実行結果: UPSデバイス状態処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, record, fieldName, fieldValue, arrayItem, joinedValue
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_PnPEntity を照会し、Name, Status, PNPDeviceID を取得します。
Set records = service.ExecQuery("SELECT Name, Status, PNPDeviceID FROM Win32_PnPEntity WHERE Name LIKE '%UPS%' OR Name LIKE '%Uninterruptible%'")
If records.Count = 0 Then
    WScript.Echo "No data returned for: UPSデバイス状態"
    WScript.Quit 0
End If
For Each record In records
    WScript.Echo "---"
    For Each fieldName In Array("Name", "Status", "PNPDeviceID")
        fieldValue = record.Properties_.Item(fieldName).Value
        ' [処理解説] 取得値がNULLまたは配列の場合を分け、表示できる文字列に整えます。
        If IsNull(fieldValue) Then
            WScript.Echo fieldName & ": (not reported)"
        ElseIf IsArray(fieldValue) Then
            joinedValue = ""
            For Each arrayItem In fieldValue
                If Len(joinedValue) > 0 Then joinedValue = joinedValue & ", "
                joinedValue = joinedValue & CStr(arrayItem)
            Next
            WScript.Echo fieldName & ": " & joinedValue
        Else
            WScript.Echo fieldName & ": " & CStr(fieldValue)
        End If
    Next
Next
