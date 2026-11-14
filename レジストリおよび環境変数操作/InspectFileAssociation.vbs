' ファイル概要: 拡張子の既定およびユーザー選択のProgIDを確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo InspectFileAssociation.vbs <.extension>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 拡張子の既定およびユーザー選択のProgIDを確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim registry, extension, progId, result, hive
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo InspectFileAssociation.vbs <.extension>"
    WScript.Quit 2
End If
extension = WScript.Arguments(0)
If Left(extension, 1) <> "." Or InStr(extension, "\") > 0 Then
    WScript.Echo "Provide a single extension beginning with a period."
    WScript.Quit 2
End If
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
result = registry.GetStringValue(&H80000000, extension, "", progId)
If result = 0 Then WScript.Echo "HKCR default ProgID=" & progId
result = registry.GetStringValue(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & extension & "\UserChoice", "ProgId", progId)
If result = 0 Then
    WScript.Echo "UserChoice ProgID=" & progId
Else
    WScript.Echo "UserChoice ProgID is not available; WMI return code: " & result
End If
