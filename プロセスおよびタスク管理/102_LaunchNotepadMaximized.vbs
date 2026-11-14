' ファイル概要: Notepad最大化起動の固定例。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 102_LaunchNotepadMaximized.vbs
' 実行結果: Notepad最大化起動の固定例処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim shell, pid
Set shell = CreateObject("WScript.Shell")
' [処理解説] ウィンドウ表示形式3で最大化を指定し、終了を待たずに制御を戻します。
pid = shell.Run("notepad.exe", 3, False)
WScript.Echo "Notepad launch requested in maximized mode; return value=" & pid
