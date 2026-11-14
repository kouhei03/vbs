' ファイル概要: WMIプロバイダーが公開するS.M.A.R.T.予測状態。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 073_SmartStatus.vbs
' 実行結果: WMIプロバイダーが公開するS.M.A.R.T.予測状態処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, record, fieldName, fieldValue, arrayItem, joinedValue
Set service = GetObject("winmgmts:\\.\root\wmi")
' [処理解説] WMIの MSStorageDriver_FailurePredictStatus を照会し、InstanceName, PredictFailure, Reason を取得します。
Set records = service.ExecQuery("SELECT InstanceName, PredictFailure, Reason FROM MSStorageDriver_FailurePredictStatus")
If records.Count = 0 Then
    WScript.Echo "No data returned for: S.M.A.R.T.予測状態"
    WScript.Quit 0
End If
For Each record In records
    WScript.Echo "---"
    For Each fieldName In Array("InstanceName", "PredictFailure", "Reason")
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
