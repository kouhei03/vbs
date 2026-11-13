' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、users、user、userName、found）。主な用途: WMIサービス接続、ユーザー一覧、ユーザー情報、対象ユーザー名、対象が見つかったかを示すフラグ。
Dim service, users, user, userName, found
' 解説: 条件「WScript.Arguments.Count <> 1 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo GetLocalUserStatus.vbs <user-name>"
    WScript.Echo "Usage: cscript //nologo GetLocalUserStatus.vbs <user-name>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: userName = Trim(WScript.Arguments(0))
userName = Trim(WScript.Arguments(0))
' 解説: 条件「Len(userName) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Len(userName) = 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "User name must not be empty."
    WScript.Echo "User name must not be empty."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのユーザーアカウント（Win32_UserAccount）を読み取り専用で照会し、Name, Disabled, Lockout, PasswordExpiresを取得します。 固定絞り込み条件: LocalAccount=True。
Set users = service.ExecQuery( _
    "SELECT Name, Disabled, Lockout, PasswordExpires FROM Win32_UserAccount " & _
    "WHERE LocalAccount=True")
' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = False
found = False
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each user In users
For Each user In users
    ' 解説: 条件「StrComp(user.Name, userName, vbTextCompare) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
    If StrComp(user.Name, userName, vbTextCompare) = 0 Then
        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: found = True
        found = True
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "User: " & user.Name
        WScript.Echo "User: " & user.Name
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Disabled: " & user.Disabled
        WScript.Echo "Disabled: " & user.Disabled
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Locked out: " & user.Lockout
        WScript.Echo "Locked out: " & user.Lockout
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Password expires: " & user.PasswordExpires
        WScript.Echo "Password expires: " & user.PasswordExpires
        ' 解説: 現在のForループを直ちに抜け、後続処理へ進みます。
        Exit For
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
' 解説: 条件「Not found Then」を評価し、成立した場合に続く分岐処理を実行します。
If Not found Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Local user not found: " & userName
    WScript.Echo "Local user not found: " & userName
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 3（対象未検出）を返します。
    WScript.Quit 3
' 解説: 条件分岐ブロックを終了します。
End If
