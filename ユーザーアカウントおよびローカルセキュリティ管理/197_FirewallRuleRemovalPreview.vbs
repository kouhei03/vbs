' ファイル概要: 指定ファイアウォールルールの状態を確認する（変更なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 197_FirewallRuleRemovalPreview.vbs <rule-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定ファイアウォールルールの状態を確認する（変更なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim ruleName, policy, rules, rule, found
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
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
