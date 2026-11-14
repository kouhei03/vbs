' 概要: IFEOデバッガー設定を固定のnotepad.exeキーで監査します。新規設定は行いません。
' 実行方法: cscript //nologo 169_IfeoAudit.vbs
' 処理結果: notepad.exeのIFEO Debugger値確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\notepad.exe\Debugger"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "No debugger value reported for notepad.exe."
Else
    WScript.Echo "Debugger value exists for notepad.exe: " & value
End If
On Error GoTo 0
WScript.Echo "Read-only audit; no IFEO policy was created."
