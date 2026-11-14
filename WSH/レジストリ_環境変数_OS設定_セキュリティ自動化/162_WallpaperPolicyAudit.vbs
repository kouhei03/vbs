' 概要: ユーザー壁紙と壁紙ポリシーの設定値を照会します。壁紙は変更しません。
' 実行方法: cscript //nologo 162_WallpaperPolicyAudit.vbs
' 処理結果: 壁紙設定値確認。処理結果や状態はコンソールに表示されます。
' コード解説: レジストリ値を読み取り、値または読み取りエラーを表示します。
Option Explicit
Dim shell, paths, path, value
Set shell = CreateObject("WScript.Shell")
paths = Array( _
    "HKEY_CURRENT_USER\Control Panel\Desktop\WallPaper", _
    "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\Wallpaper")
For Each path In paths
    On Error Resume Next
    ' 主処理: 対象のレジストリ値を読み取ります。
    value = shell.RegRead(path)
    If Err.Number = 0 Then
        WScript.Echo path & "=" & value
    Else
        WScript.Echo path & "=(not set)"
    End If
    Err.Clear
    On Error GoTo 0
Next
