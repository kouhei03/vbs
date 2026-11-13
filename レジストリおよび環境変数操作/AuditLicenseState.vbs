' 解説: 明示的な変数宣言を必須にします。
Option Explicit
' 解説: Windowsライセンス情報の照会と結果の判定に使う変数を宣言します。
Dim service, products, product, licensedCount, statusName
Set service = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
' 解説: プロダクトキーを取得せず、キー情報の存在するライセンス製品の状態だけを照会します。
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
