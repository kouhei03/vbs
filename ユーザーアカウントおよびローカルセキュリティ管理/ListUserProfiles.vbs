' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、profiles、profile）。主な用途: WMIサービス接続、プロファイル一覧、ユーザープロファイル。
Dim service, profiles, profile
' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIのユーザープロファイル（Win32_UserProfile）を読み取り専用で照会し、SID, LocalPath, Loaded, Specialを取得します。 固定絞り込み条件: Special=False。
Set profiles = service.ExecQuery( _
    "SELECT SID, LocalPath, Loaded, Special FROM Win32_UserProfile " & _
    "WHERE Special=False")

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "SID" & vbTab & "Profile path" & vbTab & "Loaded"
WScript.Echo "SID" & vbTab & "Profile path" & vbTab & "Loaded"
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each profile In profiles
For Each profile In profiles
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: profile.SID & vbTab & profile.LocalPath & vbTab & profile.Loaded
    WScript.Echo profile.SID & vbTab & profile.LocalPath & vbTab & profile.Loaded
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next
