' ファイル概要: 許可リストとサービス名の読み取り専用比較。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 169_ServiceAllowlistAudit.vbs <allowlist-file>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 許可リストとサービス名の読み取り専用比較処理の結果やエラーをコンソールに表示します。

' 外部許可リストとローカルサービス名を比較し、不一致を表示します。削除や無効化はしません。
Option Explicit
Dim fso, inputFile, allowed, service, services, item, path, lineText, found, allowedName
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo 169_ServiceAllowlistAudit.vbs <allowlist-file>"
    WScript.Quit 2
End If
path = WScript.Arguments(0)
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(path) Then
    WScript.Echo "Allowlist file not found: " & path
    WScript.Quit 2
End If
Set allowed = CreateObject("Scripting.Dictionary")
allowed.CompareMode = 1
Set inputFile = fso.OpenTextFile(path, 1, False)
Do Until inputFile.AtEndOfStream
    lineText = Trim(inputFile.ReadLine)
    If Len(lineText) > 0 Then allowed(lineText) = True
Loop
inputFile.Close
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_Service を照会し、Name, DisplayName, State を取得します。
Set services = service.ExecQuery("SELECT Name, DisplayName, State FROM Win32_Service")
For Each item In services
    If Not allowed.Exists(item.Name) Then WScript.Echo "Not in allowlist: " & item.Name & vbTab & item.DisplayName & vbTab & item.State
Next
WScript.Echo "Audit completed; no service settings were changed."
