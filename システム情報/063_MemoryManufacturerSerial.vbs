' ファイル概要: メモリメーカーとシリアル番号。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 063_MemoryManufacturerSerial.vbs
' 実行結果: メモリメーカーとシリアル番号処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, record, fieldName, fieldValue, arrayItem, joinedValue
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_PhysicalMemory を照会し、DeviceLocator, Manufacturer, SerialNumber, PartNumber を取得します。
Set records = service.ExecQuery("SELECT DeviceLocator, Manufacturer, SerialNumber, PartNumber FROM Win32_PhysicalMemory")
If records.Count = 0 Then
    WScript.Echo "No data returned for: メモリメーカーとシリアル番号"
    WScript.Quit 0
End If
For Each record In records
    WScript.Echo "---"
    For Each fieldName In Array("DeviceLocator", "Manufacturer", "SerialNumber", "PartNumber")
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
