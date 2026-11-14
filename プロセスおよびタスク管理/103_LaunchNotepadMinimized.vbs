' ファイル概要: Notepad最小化起動の固定例。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 103_LaunchNotepadMinimized.vbs
' 実行結果: Notepad最小化起動の固定例処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, pid
Set shell = CreateObject("WScript.Shell")
pid = shell.Run("notepad.exe", 7, False)
WScript.Echo "Notepad launch requested in minimized mode; return value=" & pid
