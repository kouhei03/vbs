' 起動順の参考として、対象サービスが依存するサービスを表示します。起動は行いません。
Option Explicit
Dim serviceName, validator, service, relations, relation, marker, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 107_DependencyStartOrderPreview.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Invalid service name."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set relations = service.ExecQuery("SELECT Antecedent, Dependent FROM Win32_DependentService")
marker = "Name=""" & serviceName & """"
found = False
For Each relation In relations
    If InStr(1, relation.Dependent, marker, vbTextCompare) > 0 Then
        WScript.Echo "Required service reference=" & relation.Antecedent
        found = True
    End If
Next
If Not found Then WScript.Echo "No dependency references found (or service is not visible)."
