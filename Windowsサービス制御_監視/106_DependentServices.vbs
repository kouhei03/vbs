' 一つのサービスに依存しているサービスをローカルWMIから列挙します。変更操作はしません。
Option Explicit
Dim serviceName, validator, service, relations, relation, marker, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 106_DependentServices.vbs <service-name>"
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
    If InStr(1, relation.Antecedent, marker, vbTextCompare) > 0 Then
        WScript.Echo "Dependent reference=" & relation.Dependent
        found = True
    End If
Next
If Not found Then WScript.Echo "No dependent services found (or service is not visible)."
