Option Explicit

Dim network, mappings, index
Set network = CreateObject("WScript.Network")
Set mappings = network.EnumNetworkDrives

If mappings.Count = 0 Then
    WScript.Echo "No network drives are mapped."
Else
    WScript.Echo "Drive" & vbTab & "Network path"
    For index = 0 To mappings.Count - 1 Step 2
        WScript.Echo mappings.Item(index) & vbTab & mappings.Item(index + 1)
    Next
End If
