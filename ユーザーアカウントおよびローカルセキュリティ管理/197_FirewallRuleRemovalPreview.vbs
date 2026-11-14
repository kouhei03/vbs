Option Explicit

Dim ruleName, policy, rules, rule, found
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 197_FirewallRuleRemovalPreview.vbs <rule-name>"
    WScript.Quit 2
End If
ruleName = WScript.Arguments(0)
Set policy = CreateObject("HNetCfg.FwPolicy2")
Set rules = policy.Rules
found = False
For Each rule In rules
    If StrComp(rule.Name, ruleName, vbTextCompare) = 0 Then
        found = True
        WScript.Echo "Rule: " & rule.Name
        WScript.Echo "Enabled: " & rule.Enabled
        WScript.Echo "Direction: " & rule.Direction
        WScript.Echo "Action: " & rule.Action
    End If
Next
If Not found Then WScript.Echo "Firewall rule was not found."
WScript.Echo "Preview only; no rule was removed or disabled."
