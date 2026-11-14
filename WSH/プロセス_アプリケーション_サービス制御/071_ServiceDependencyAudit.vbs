' 指定サービスの依存先／依存元を表示します。サービス状態は変更しません。
Option Explicit
Dim name, re, service, records, relation
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 071_ServiceDependencyAudit.vbs <service-name>"
    WScript.Quit 2
End If
name = WScript.Arguments(0)
Set re = New RegExp
re.Pattern = "^[A-Za-z0-9_.-]+$"
If Not re.Test(name) Then WScript.Quit 2
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set records = service.ExecQuery("SELECT Antecedent, Dependent FROM Win32_DependentService")
For Each relation In records
    If InStr(1, relation.Antecedent, "Name=""" & name & """", 1) > 0 Or _
        InStr(1, relation.Dependent, "Name=""" & name & """", 1) > 0 Then
        WScript.Echo relation.Antecedent & vbTab & "->" & vbTab & relation.Dependent
    End If
Next
