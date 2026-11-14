' ファイル概要: GPUドライバ情報。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 088_GpuDriverInfo.vbs
' 実行結果: GPUドライバ情報処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, record, fieldName, fieldValue, arrayItem, joinedValue
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_PnPSignedDriver を照会し、DeviceName, DriverVersion, DriverDate を取得します。
Set records = service.ExecQuery("SELECT DeviceName, DriverVersion, DriverDate FROM Win32_PnPSignedDriver WHERE DeviceClass = 'DISPLAY'")
If records.Count = 0 Then
    WScript.Echo "No data returned for: GPUドライバ情報"
    WScript.Quit 0
End If
For Each record In records
    WScript.Echo "---"
    For Each fieldName In Array("DeviceName", "DriverVersion", "DriverDate")
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
