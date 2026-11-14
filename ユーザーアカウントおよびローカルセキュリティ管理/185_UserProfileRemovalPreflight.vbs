' ファイル概要: 90日以上未使用の非特殊・未ロードプロファイルを確認する（削除なし）。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo 185_UserProfileRemovalPreflight.vbs
' 実行結果: 90日以上未使用の非特殊・未ロードプロファイルを確認する（削除なし）処理の結果やエラーをコンソールに表示します。

Option Explicit

Dim service, profiles, profile, cutoff, lastUse, oldCount
Set service = GetObject("winmgmts:\\.\root\cimv2")
' [処理解説] WMIの Win32_UserProfile を照会し、SID, LocalPath, Loaded, Special, LastUseTime を取得します。
Set profiles = service.ExecQuery( _
    "SELECT SID, LocalPath, Loaded, Special, LastUseTime FROM Win32_UserProfile")
cutoff = DateAdd("d", -90, Now)
oldCount = 0
For Each profile In profiles
    If Not profile.Special Then
        If Not profile.Loaded Then
            If Not IsNull(profile.LastUseTime) Then
                If Len(CStr(profile.LastUseTime)) > 0 Then
                    Set lastUse = CreateObject("WbemScripting.SWbemDateTime")
                    lastUse.Value = profile.LastUseTime
                    If lastUse.GetVarDate(True) < cutoff Then
                        oldCount = oldCount + 1
                        WScript.Echo "Review candidate: " & profile.LocalPath & _
                            "; SID=" & profile.SID & "; last used=" & lastUse.GetVarDate(True)
                    End If
                End If
            End If
        End If
    End If
Next
WScript.Echo "Profiles older than 90 days: " & oldCount
WScript.Echo "Preflight only; no profile or user data was deleted."
