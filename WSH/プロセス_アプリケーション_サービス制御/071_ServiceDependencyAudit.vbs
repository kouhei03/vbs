' 概要: 指定サービスの依存先／依存元を表示します。サービス状態は変更しません。
' 実行方法: Usage: cscript //nologo 071_ServiceDependencyAudit.vbs <service-name>
' 処理結果: サービス依存関係の照会（制御なし）。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim name, re, service, records, relation
' 引数の数を検証し、誤った呼び出しではUsageを表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 071_ServiceDependencyAudit.vbs <service-name>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 主処理: WMIの照会結果を処理します。
Set records = service.ExecQuery("SELECT Antecedent, Dependent FROM Win32_DependentService")
For Each relation In records
    If InStr(1, relation.Antecedent, "Name=""" & name & """", 1) > 0 Or _
        InStr(1, relation.Dependent, "Name=""" & name & """", 1) > 0 Then
        WScript.Echo relation.Antecedent & vbTab & "->" & vbTab & relation.Dependent
    End If
Next
