' ファイル概要: WMIクラス利用可否の互換性確認。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 204_WmiServiceCompatibility.vbs
' 実行結果: WMIクラス利用可否の互換性確認処理の結果やエラーをコンソールに表示します。

' 必要なWMIクラスへのアクセス可否を照会し、互換性確認の簡単な結果を出します。
Option Explicit
Dim service, services, relations
On Error Resume Next
Set service = GetObject("winmgmts:\\.\root\cimv2")
If Err.Number <> 0 Then
    WScript.Echo "WMI namespace unavailable: " & Err.Description
    WScript.Quit 1
End If
Err.Clear
' [処理解説] WMIの Win32_Service を照会し、Name を取得します。
Set services = service.ExecQuery("SELECT Name FROM Win32_Service")
If Err.Number <> 0 Then
    WScript.Echo "Win32_Service query unavailable: " & Err.Description
    WScript.Quit 1
End If
WScript.Echo "Win32_Service available; count=" & services.Count
Err.Clear
' [処理解説] WMIの Win32_DependentService を照会し、Antecedent, Dependent を取得します。
Set relations = service.ExecQuery("SELECT Antecedent, Dependent FROM Win32_DependentService")
If Err.Number <> 0 Then
    WScript.Echo "Win32_DependentService unavailable: " & Err.Description
    Err.Clear
Else
    WScript.Echo "Win32_DependentService available; count=" & relations.Count
End If
On Error GoTo 0
