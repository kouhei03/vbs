' ファイル概要: Windowsインストール先フォルダー。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 119_WindowsInstallFolders.vbs
' 実行結果: Windowsインストール先フォルダー処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, records, record, fieldName, fieldValue, arrayItem, joinedValue
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_OperatingSystem を照会し、WindowsDirectory, SystemDirectory, SystemDrive を取得します。
Set records = service.ExecQuery("SELECT WindowsDirectory, SystemDirectory, SystemDrive FROM Win32_OperatingSystem")
If records.Count = 0 Then
    WScript.Echo "No data returned for: Windowsインストール先フォルダー"
    WScript.Quit 0
End If
For Each record In records
    WScript.Echo "---"
    For Each fieldName In Array("WindowsDirectory", "SystemDirectory", "SystemDrive")
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
