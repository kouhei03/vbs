' ファイル概要: モニター物理情報。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 092_MonitorPhysicalInfo.vbs
' 実行結果: モニター物理情報処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, monitors, monitor, fieldName, values, code, decoded, i
Set service = GetObject("winmgmts:\\.\root\wmi")
Set monitors = service.ExecQuery("SELECT ManufacturerName, ProductCodeID, SerialNumberID, UserFriendlyName FROM WmiMonitorID")
If monitors.Count = 0 Then WScript.Echo "No monitor identification data returned."
For Each monitor In monitors
    For Each fieldName In Array("ManufacturerName", "ProductCodeID", "SerialNumberID", "UserFriendlyName")
        values = monitor.Properties_.Item(fieldName).Value
        decoded = ""
        If IsArray(values) Then
            For Each code In values
                If CLng(code) <> 0 Then decoded = decoded & ChrW(CLng(code))
            Next
        End If
        WScript.Echo fieldName & ": " & decoded
    Next
Next
