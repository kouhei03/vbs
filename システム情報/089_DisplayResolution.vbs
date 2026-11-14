' ファイル概要: ディスプレイ解像度。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 089_DisplayResolution.vbs
' 実行結果: ディスプレイ解像度処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, record, fieldName, fieldValue, arrayItem, joinedValue
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_VideoController を照会し、Name, CurrentHorizontalResolution, CurrentVerticalResolution を取得します。
Set records = service.ExecQuery("SELECT Name, CurrentHorizontalResolution, CurrentVerticalResolution FROM Win32_VideoController")
If records.Count = 0 Then
    WScript.Echo "No data returned for: ディスプレイ解像度"
    WScript.Quit 0
End If
For Each record In records
    WScript.Echo "---"
    For Each fieldName In Array("Name", "CurrentHorizontalResolution", "CurrentVerticalResolution")
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
