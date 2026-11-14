Option Explicit

Dim service, relations, relation
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set relations = service.ExecQuery( _
    "SELECT Antecedent, Dependent FROM Win32_DependentService")

WScript.Echo "Required service reference" & vbTab & _
    "Dependent service reference"
For Each relation In relations
    WScript.Echo relation.Antecedent & vbTab & relation.Dependent
Next
