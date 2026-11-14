' 概要: 固定のデモ値を読み、存在しない場合とアクセスエラーを明示して処理します。
' 実行方法: cscript //nologo 193_RegistryErrorHandlerDemo.vbs
' 処理結果: レジストリ読取エラー処理デモ。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, path, value, errorNumber, errorText
Set shell = CreateObject("WScript.Shell")
path = "HKEY_CURRENT_USER\Software\WSH\SafeExamples\ErrorDemo"
On Error Resume Next
' 主処理: 対象のレジストリ値を読み取ります。
value = shell.RegRead(path)
errorNumber = Err.Number
errorText = Err.Description
Err.Clear
On Error GoTo 0
If errorNumber = 0 Then
    WScript.Echo "Value=" & value
ElseIf errorNumber = 1 Or errorNumber = -2147024894 Then
    WScript.Echo "Demo value is absent."
Else
    WScript.Echo "Registry read failed. Error=" & errorNumber & " " & errorText
    WScript.Quit 1
End If
