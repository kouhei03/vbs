' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: WScript.Shellのシステム環境変数コレクションと項目を保持します。
Dim shell, environment, variable
Set shell = CreateObject("WScript.Shell")
Set environment = shell.Environment("SYSTEM")
' 解説: システム環境変数を読み取り、値を表示します。変更操作は行いません。
For Each variable In environment
    WScript.Echo variable
Next
