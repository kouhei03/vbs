' 解説: UAC通知レベルに関連するDWORD値を読み取ります。設定は変更しません。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: registry, result, value, settingName
Dim registry, result, value, settingName
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: コレクション内の各項目を順に処理します: For Each settingName In Array("EnableLUA", "ConsentPromptBehaviorAdmin", "PromptOnSecureDesktop", "FilterAdministratorToken")
For Each settingName In Array("EnableLUA", "ConsentPromptBehaviorAdmin", "PromptOnSecureDesktop", "FilterAdministratorToken")
    ' 解説: 取得値または計算結果を変数へ格納します。
    result = registry.GetDWORDValue(&H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", settingName, value)
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: result = 0
    If result = 0 Then
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo settingName & "=" & value
    ' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
    Else
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo settingName & " unavailable; WMI code=" & result
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
