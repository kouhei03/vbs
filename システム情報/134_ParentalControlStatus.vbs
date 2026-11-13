' 解説: ペアレンタルコントロール関連キーの存在だけを確認し、ポリシーは変更しません。
Option Explicit
' 解説: このスクリプト内で使用する変数を宣言します: registry, result, names, keyName
Dim registry, result, names, keyName
' 解説: 必要なローカルWMIまたはWindowsオブジェクトへ接続します。
Set registry = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
' 解説: コレクション内の各項目を順に処理します: For Each keyName In Array("SOFTWARE\Microsoft\Parental Controls", "SOFTWARE\Policies\Microsoft\Windows\ParentalControls")
For Each keyName In Array("SOFTWARE\Microsoft\Parental Controls", "SOFTWARE\Policies\Microsoft\Windows\ParentalControls")
    ' 解説: 取得値または計算結果を変数へ格納します。
    result = registry.EnumKey(&H80000002, keyName, names)
    ' 解説: 条件を評価し、該当する場合の分岐処理を選択します: result = 0
    If result = 0 Then
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo keyName & ": key present"
    ' 解説: 先行条件が不成立の場合に追加条件を評価します: result = 2
    ElseIf result = 2 Then
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo keyName & ": not present"
    ' 解説: 先行条件がいずれも不成立の場合の代替処理へ進みます。
    Else
        ' 解説: 取得結果、状態、またはエラー情報をコンソールに表示します。
        WScript.Echo keyName & ": WMI code=" & result
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: 現在の反復を終え、次の要素へ進みます。
Next
