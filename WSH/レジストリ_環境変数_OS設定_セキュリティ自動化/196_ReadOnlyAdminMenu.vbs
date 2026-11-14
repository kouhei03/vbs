' 概要: 対話メニューから読み取り専用の環境・UAC監査を選択します。
' 実行方法: cscript //nologo 196_ReadOnlyAdminMenu.vbs
' 処理結果: 対話式の読み取り専用監査メニュー。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim choice, shell, env, value, keyPath
WScript.Echo "1. Current user PATH (value omitted)"
WScript.Echo "2. UAC EnableLUA status"
WScript.Echo "0. Exit"
WScript.StdOut.Write "Select: "
choice = Trim(WScript.StdIn.ReadLine)
Set shell = CreateObject("WScript.Shell")
Select Case choice
    Case "1"
        Set env = shell.Environment("USER")
        If env("PATH") = "" Then
            WScript.Echo "User PATH is empty or not set."
        Else
            WScript.Echo "User PATH is configured; value omitted."
        End If
    Case "2"
        keyPath = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\EnableLUA"
        On Error Resume Next
        ' 主処理: 対象のレジストリ値を読み取ります。
        value = shell.RegRead(keyPath)
        If Err.Number = 0 Then
            WScript.Echo "EnableLUA=" & value
        Else
            WScript.Echo "EnableLUA is unavailable: " & Err.Description
            WScript.Quit 1
        End If
        On Error GoTo 0
    Case "0"
        WScript.Echo "No changes made."
    Case Else
        WScript.Echo "Invalid selection."
        WScript.Quit 2
End Select
