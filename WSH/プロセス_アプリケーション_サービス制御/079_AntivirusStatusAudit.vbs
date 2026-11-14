' 概要: SecurityCenter2が公開するウイルス対策製品の状態を照会します。保護機能は変更しません。
' 実行方法: cscript //nologo 079_AntivirusStatusAudit.vbs
' 処理結果: SecurityCenter2の製品状態照会。処理結果や状態はコンソールに表示されます。
' コード解説: WMIの照会結果を列挙し、条件に一致する情報を処理します。
Option Explicit
Dim securityCenter, products, product
Set securityCenter = GetObject("winmgmts:\\.\root\SecurityCenter2")
' 主処理: WMIの照会結果を処理します。
Set products = securityCenter.ExecQuery("SELECT displayName, productState, timestamp FROM AntiVirusProduct")
For Each product In products
    WScript.Echo product.displayName & vbTab & "State=" & product.productState & vbTab & "Updated=" & product.timestamp
Next
If products.Count = 0 Then WScript.Echo "No products reported by SecurityCenter2."
