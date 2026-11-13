Option Explicit

Dim service, relations, relation
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set relations = service.ExecQuery("SELECT Antecedent, Dependent FROM Win32_DependentService")

WScript.Echo "Service reference" & vbTab & "Depends-on service reference"
For Each relation In relations
    WScript.Echo relation.Dependent & vbTab & relation.Antecedent
Next
