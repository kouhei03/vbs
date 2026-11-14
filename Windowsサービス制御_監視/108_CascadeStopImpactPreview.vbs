' ファイル概要: 停止時に影響を受けるサービスの状態プレビュー（停止なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 108_CascadeStopImpactPreview.vbs <service-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 停止時に影響を受けるサービスの状態プレビュー（停止なし）処理の結果やエラーをコンソールに表示します。

' 停止した場合に影響を受け得る依存サービスを表示します。停止は行いません。
Option Explicit
Dim serviceName, validator, service, relations, relation, services, item, marker, found, stateText
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 108_CascadeStopImpactPreview.vbs <service-name>"
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
Set services = service.ExecQuery("SELECT Name, State FROM Win32_Service")
marker = "Name=""" & serviceName & """"
found = False
For Each relation In relations
    If InStr(1, relation.Antecedent, marker, vbTextCompare) > 0 Then
        stateText = "Unknown"
        For Each item In services
            If InStr(1, relation.Dependent, "Name=""" & item.Name & """", vbTextCompare) > 0 Then stateText = item.State
        Next
        WScript.Echo relation.Dependent & vbTab & stateText
        found = True
    End If
Next
If Not found Then WScript.Echo "No dependent services found."
