' ファイル概要: 新規HTMLサービスダッシュボード。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 201_ServiceDashboard.vbs <new-html-path>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: サービス状態のHTMLダッシュボードを新規作成し、処理結果を表示します。

' ローカルサービス状態のHTMLダッシュボードを新規作成します。既存ファイルは上書きしません。
Option Explicit
Dim fso, path, file, service, services, item
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 201_ServiceDashboard.vbs <new-html-path>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(path) Then
    WScript.Echo "Refusing to overwrite: " & path
    WScript.Quit 2
End If
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set services = service.ExecQuery("SELECT Name, DisplayName, State, StartMode FROM Win32_Service")
Set file = fso.CreateTextFile(path, False, True)
file.WriteLine "<!doctype html><html><head><meta charset=""utf-8""><title>Service report</title></head><body><h1>Service dashboard</h1><table><tr><th>Name</th><th>Display name</th><th>State</th><th>Start mode</th></tr>"
For Each item In services
    file.WriteLine "<tr><td>" & Html(item.Name) & "</td><td>" & Html(item.DisplayName) & "</td><td>" & Html(item.State) & "</td><td>" & Html(item.StartMode) & "</td></tr>"
Next
file.WriteLine "</table></body></html>"
file.Close
WScript.Echo "Report written: " & path
Function Html(value)
    Dim text
    text = CStr(value)
    text = Replace(text, "&", "&amp;")
    text = Replace(text, "<", "&lt;")
    text = Replace(text, ">", "&gt;")
    text = Replace(text, Chr(34), "&quot;")
    Html = text
End Function
