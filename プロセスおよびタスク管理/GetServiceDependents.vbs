' ファイル概要: サービスに依存するサービスを列挙する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetServiceDependents.vbs
' 実行結果: サービスに依存するサービスを列挙する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, relations, relation
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_DependentService を照会し、Antecedent, Dependent を取得します。
Set relations = service.ExecQuery("SELECT Antecedent, Dependent FROM Win32_DependentService")

WScript.Echo "Service reference" & vbTab & "Depends-on service reference"
For Each relation In relations
    WScript.Echo relation.Dependent & vbTab & relation.Antecedent
Next
