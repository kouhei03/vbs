' ファイル概要: ページング等の性能カウンター（帯域幅そのものではありません）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 065_MemoryPerformanceCounters.vbs
' 実行結果: ページング等の性能カウンター（帯域幅そのものではありません）処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, record, fieldName, fieldValue, arrayItem, joinedValue
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT AvailableBytes, CacheBytes, PagesPersec, PoolNonpagedBytes, PoolPagedBytes FROM Win32_PerfFormattedData_PerfOS_Memory")
If records.Count = 0 Then
    WScript.Echo "No data returned for: メモリ性能カウンター"
    WScript.Quit 0
End If
For Each record In records
    WScript.Echo "---"
    For Each fieldName In Array("AvailableBytes", "CacheBytes", "PagesPersec", "PoolNonpagedBytes", "PoolPagedBytes")
        fieldValue = record.Properties_.Item(fieldName).Value
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
