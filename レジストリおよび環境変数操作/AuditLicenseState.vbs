' ファイル概要: プロダクトキーを取得せずライセンス状態を確認する。
' 実行方法: Windowsのコマンドプロンプトで、このファイルのフォルダーから以下を実行。
'   cscript //nologo AuditLicenseState.vbs
' 実行結果: プロダクトキーを取得せずライセンス状態を確認する処理の結果やエラーをコンソールに表示します。

Option Explicit
Dim service, products, product, licensedCount, statusName
Set service = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set products = service.ExecQuery("SELECT Name, LicenseStatus FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL")
licensedCount = 0
For Each product In products
    Select Case product.LicenseStatus
        Case 0: statusName = "Unlicensed"
        Case 1: statusName = "Licensed"
        Case 2: statusName = "Initial grace"
        Case 3: statusName = "Additional grace"
        Case 4: statusName = "Non-genuine grace"
        Case 5: statusName = "Notification"
        Case 6: statusName = "Extended grace"
        Case Else: statusName = "Unknown"
    End Select
    WScript.Echo product.Name & ": " & statusName
    If product.LicenseStatus = 1 Then licensedCount = licensedCount + 1
Next
If licensedCount = 0 Then WScript.Echo "No licensed product was reported; verify with your organization's licensing tools."
