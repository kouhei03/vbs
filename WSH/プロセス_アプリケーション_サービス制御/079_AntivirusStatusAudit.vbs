' SecurityCenter2が公開するウイルス対策製品の状態を照会します。保護機能は変更しません。
Option Explicit
Dim securityCenter, products, product
Set securityCenter = GetObject("winmgmts:\\.\root\SecurityCenter2")
Set products = securityCenter.ExecQuery("SELECT displayName, productState, timestamp FROM AntiVirusProduct")
For Each product In products
    WScript.Echo product.displayName & vbTab & "State=" & product.productState & vbTab & "Updated=" & product.timestamp
Next
If products.Count = 0 Then WScript.Echo "No products reported by SecurityCenter2."
