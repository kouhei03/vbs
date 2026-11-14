' ファイル概要: 依存しているサービスの列挙。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 106_DependentServices.vbs <service-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 依存しているサービスの列挙処理の結果やエラーをコンソールに表示します。

' 一つのサービスに依存しているサービスをローカルWMIから列挙します。変更操作はしません。
Option Explicit
Dim serviceName, validator, service, relations, relation, marker, found
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 106_DependentServices.vbs <service-name>"
    WScript.Quit 2
End If
serviceName = Trim(WScript.Arguments(0))
Set validator = New RegExp
' [処理解説] 引数に許可する文字を限定して名前を検証します。
validator.Pattern = "^[A-Za-z0-9_.-]+$"
If Not validator.Test(serviceName) Then
    WScript.Echo "Invalid service name."
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_DependentService を照会し、Antecedent, Dependent を取得します。
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
