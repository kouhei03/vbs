' 解説: 変数の宣言漏れを実行前に検出するため、明示的な変数宣言を必須にします。
Option Explicit

' 解説: このスクリプトで使う変数を宣言します（service、settings、setting、shareName、descriptor、result）。主な用途: WMIサービス接続、設定一覧、設定情報、対象共有名、アクセス制御記述子、処理結果コード。
Dim service, settings, setting, shareName, descriptor, result
' 解説: このスクリプトで使う変数を宣言します（entry、trustee、trusteeName、accessType）。主な用途: アクセス制御エントリ、アクセス許可の対象主体。
Dim entry, trustee, trusteeName, accessType
' 解説: 条件「WScript.Arguments.Count <> 1 Then」を評価し、成立した場合に続く分岐処理を実行します。
If WScript.Arguments.Count <> 1 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Usage: cscript //nologo GetShareAcl.vbs <share-name>"
    WScript.Echo "Usage: cscript //nologo GetShareAcl.vbs <share-name>"
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If
' 解説: 関数の戻り値を受け取り、後続の判定や出力に利用します: shareName = Trim(WScript.Arguments(0))
shareName = Trim(WScript.Arguments(0))
' 解説: 条件「Len(shareName) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
If Len(shareName) = 0 Then
    ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Share name must not be empty."
    WScript.Echo "Share name must not be empty."
    ' 解説: スクリプトを終了し、呼び出し元に終了コード 2（引数不正）を返します。
    WScript.Quit 2
' 解説: 条件分岐ブロックを終了します。
End If

' 解説: ローカルのWMI/CIM名前空間へ接続し、以降の情報照会に使うサービスオブジェクトを取得します。
Set service = GetObject("winmgmts:\\.\root\cimv2")
' 解説: WMIの共有ACL設定（Win32_LogicalShareSecuritySetting）を読み取り専用で照会し、Nameを取得します。
Set settings = service.ExecQuery( _
    "SELECT Name FROM Win32_LogicalShareSecuritySetting")
' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each setting In settings
For Each setting In settings
    ' 解説: 条件「StrComp(setting.Name, shareName, vbTextCompare) = 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
    If StrComp(setting.Name, shareName, vbTextCompare) = 0 Then
        ' 解説: 共有のセキュリティ記述子を読み取り、DACL/ACEのアクセス情報を確認します。アクセス権は変更しません。
        result = setting.GetSecurityDescriptor(descriptor)
        ' 解説: 条件「result <> 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
        If result <> 0 Then
            ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Could not read share ACL; return code: " & result
            WScript.Echo "Could not read share ACL; return code: " & result
            ' 解説: スクリプトを終了し、呼び出し元に終了コード 1（処理失敗）を返します。
            WScript.Quit 1
        ' 解説: 条件分岐ブロックを終了します。
        End If
        ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Share: " & setting.Name
        WScript.Echo "Share: " & setting.Name
        ' 解説: 条件「IsArray(descriptor.DACL) Then」を評価し、成立した場合に続く分岐処理を実行します。
        If IsArray(descriptor.DACL) Then
            ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Trustee" & vbTab & "ACE type" & vbTab & "Access mask"
            WScript.Echo "Trustee" & vbTab & "ACE type" & vbTab & "Access mask"
            ' 解説: コレクション内の各要素を順番に取り出し、個々の項目に対する処理を行います: For Each entry In descriptor.DACL
            For Each entry In descriptor.DACL
                ' 解説: アクセス制御エントリから許可対象のユーザー/グループ情報を取り出します。
                Set trustee = entry.Trustee
                ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: trusteeName = trustee.Name
                trusteeName = trustee.Name
                ' 解説: 条件「Not IsNull(trustee.Domain) Then」を評価し、成立した場合に続く分岐処理を実行します。
                If Not IsNull(trustee.Domain) Then
                    ' 解説: 条件「Len(trustee.Domain) > 0 Then」を評価し、成立した場合に続く分岐処理を実行します。
                    If Len(trustee.Domain) > 0 Then
                        ' 解説: 計算結果または取得値を変数へ保存し、後続の処理で利用します: trusteeName = trustee.Domain & "\" & trusteeName
                        trusteeName = trustee.Domain & "\" & trusteeName
                    ' 解説: 条件分岐ブロックを終了します。
                    End If
                ' 解説: 条件分岐ブロックを終了します。
                End If
                ' 解説: 値ごとに処理を分ける選択分岐を開始します: entry.AceType
                Select Case entry.AceType
                    ' 解説: Select Caseの値が「0: accessType = "Allow"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
                    Case 0: accessType = "Allow"
                    ' 解説: Select Caseの値が「1: accessType = "Deny"」に一致した場合の分岐です。右辺に状態名がある場合は、その状態を表示用の名称へ対応付けます。
                    Case 1: accessType = "Deny"
                    ' 解説: Select CaseのどのCase値にも一致しない場合に実行する既定分岐です。
                    Case Else: accessType = "Other (" & entry.AceType & ")"
                ' 解説: Select Caseの選択分岐を終了します。
                End Select
                ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: trusteeName & vbTab & accessType & vbTab & entry.AccessMask
                WScript.Echo trusteeName & vbTab & accessType & vbTab & _
                    entry.AccessMask
            ' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
            Next
        ' 解説: 直前までの条件がどれも成立しない場合の代替処理に入ります。
        Else
            ' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "No DACL entries returned."
            WScript.Echo "No DACL entries returned."
        ' 解説: 条件分岐ブロックを終了します。
        End If
        ' 解説: スクリプトを終了し、呼び出し元に終了コード 0（正常終了）を返します。
        WScript.Quit 0
    ' 解説: 条件分岐ブロックを終了します。
    End If
' 解説: For Eachの現在の反復を終え、次の要素へ進みます。
Next

' 解説: 結果、案内、またはエラーをコンソールへ表示します。表示式: "Share ACL not found: " & shareName
WScript.Echo "Share ACL not found: " & shareName
' 解説: スクリプトを終了し、呼び出し元に終了コード 3（対象未検出）を返します。
WScript.Quit 3
