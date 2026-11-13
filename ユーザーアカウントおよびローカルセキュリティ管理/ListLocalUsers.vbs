' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、users、user）。主な用途: WMIサービス接続、ユーザー一覧、ユーザー情報。
Dim service, users, user
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのユーザーアカウント（Win32_UserAccount）を読み取り専用で照会し、Name, Disabled, Lockout, PasswordExpiresを取得します。 固定絞り込み条件: LocalAccount=True。
Set users = service.ExecQuery( _
    "SELECT Name, Disabled, Lockout, PasswordExpires FROM Win32_UserAccount " & _
    "WHERE LocalAccount=True")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Local user" & vbTab & "Disabled" & vbTab & "Locked out" & vbTab & "Password expires"
WScript.Echo "Local user" & vbTab & "Disabled" & vbTab & _
    "Locked out" & vbTab & "Password expires"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each user In users
For Each user In users
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: user.Name & vbTab & user.Disabled & vbTab & user.Lockout & vbTab & user.PasswordExpires
    WScript.Echo user.Name & vbTab & user.Disabled & vbTab & _
        user.Lockout & vbTab & user.PasswordExpires
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
