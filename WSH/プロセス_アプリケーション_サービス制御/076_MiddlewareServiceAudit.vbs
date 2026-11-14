' 概要: 代表的なミドルウェアサービスの状態を照会します。管理コマンドは実行しません。
' 実行方法: cscript //nologo 076_MiddlewareServiceAudit.vbs
' 処理結果: 代表的ミドルウェアサービスの状態照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim service, name, records, item
Set service = GetObject("winmgmts:\\.\root\cimv2")
For Each name In Array("W3SVC", "MSSQLSERVER", "Spooler")
    ' 主処理: WMIに問い合わせて対象情報を取得します。
    Set records = service.ExecQuery("SELECT Name, State, StartMode FROM Win32_Service WHERE Name='" & name & "'")
    If records.Count = 0 Then
        WScript.Echo name & vbTab & "Not installed"
    Else
        For Each item In records
            WScript.Echo item.Name & vbTab & item.State & vbTab & item.StartMode
        Next
    End If
Next
