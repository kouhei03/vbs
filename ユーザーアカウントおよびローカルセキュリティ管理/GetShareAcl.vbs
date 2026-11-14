' ファイル概要: 指定した共有フォルダーのACLを取得する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo GetShareAcl.vbs <share-name>
'   <>は必須引数、[]は省略可能な引数を示します。値を置き換えて入力してください。
' 実行結果: 指定した共有フォルダーのACLを取得する処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, settings, setting, shareName, descriptor, result
Dim entry, trustee, trusteeName, accessType
' [処理解説] 必要な引数の数が合わない場合、使い方を表示して終了します。
If WScript.Arguments.Count <> 1 Then
    WScript.Echo "Usage: cscript //nologo GetShareAcl.vbs <share-name>"
    WScript.Quit 2
End If
shareName = Trim(WScript.Arguments(0))
If Len(shareName) = 0 Then
    WScript.Echo "Share name must not be empty."
    WScript.Quit 2
End If

Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_LogicalShareSecuritySetting を照会し、Name を取得します。
Set settings = service.ExecQuery( _
    "SELECT Name FROM Win32_LogicalShareSecuritySetting")
For Each setting In settings
    If StrComp(setting.Name, shareName, vbTextCompare) = 0 Then
        result = setting.GetSecurityDescriptor(descriptor)
        If result <> 0 Then
            WScript.Echo "Could not read share ACL; return code: " & result
            WScript.Quit 1
        End If
        WScript.Echo "Share: " & setting.Name
        If IsArray(descriptor.DACL) Then
            WScript.Echo "Trustee" & vbTab & "ACE type" & vbTab & "Access mask"
            For Each entry In descriptor.DACL
                Set trustee = entry.Trustee
                trusteeName = trustee.Name
                If Not IsNull(trustee.Domain) Then
                    If Len(trustee.Domain) > 0 Then
                        trusteeName = trustee.Domain & "\" & trusteeName
                    End If
                End If
                Select Case entry.AceType
                    Case 0: accessType = "Allow"
                    Case 1: accessType = "Deny"
                    Case Else: accessType = "Other (" & entry.AceType & ")"
                End Select
                WScript.Echo trusteeName & vbTab & accessType & vbTab & _
                    entry.AccessMask
            Next
        Else
            WScript.Echo "No DACL entries returned."
        End If
        WScript.Quit 0
    End If
Next

WScript.Echo "Share ACL not found: " & shareName
WScript.Quit 3
