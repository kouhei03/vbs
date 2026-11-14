' 概要: 固定のデモアプリキーに初回起動フラグがあるか照会します。値は削除しません。
' 実行方法: cscript //nologo 166_FirstRunFlagAudit.vbs
' 処理結果: デモ用初回起動フラグの照会。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, value
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\WSH\SafeExamples\FirstRunComplete"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
If Err.Number <> 0 Then
    WScript.Echo "Demo first-run flag is absent."
Else
    WScript.Echo "Demo first-run flag exists: " & value
End If
On Error GoTo 0
WScript.Echo "No application state was reset."
