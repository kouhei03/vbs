Option Explicit
Dim service, adapters, adapter, identity, classification
Set service = GetObject("winmgmts:\\.\root\cimv2")
Set adapters = service.ExecQuery("SELECT Name, AdapterCompatibility, VideoProcessor, PNPDeviceID FROM Win32_VideoController")
For Each adapter In adapters
    identity = LCase(SafeText(adapter.Name) & " " & SafeText(adapter.AdapterCompatibility) & " " & SafeText(adapter.VideoProcessor))
    If InStr(identity, "intel") > 0 Or InStr(identity, "uhd graphics") > 0 Or InStr(identity, "iris") > 0 Then
        classification = "Likely integrated (heuristic)"
    ElseIf InStr(identity, "nvidia") > 0 Or InStr(identity, "geforce") > 0 Or InStr(identity, "radeon") > 0 Then
        classification = "Likely discrete (heuristic; some Radeon GPUs are integrated)"
    Else
        classification = "Unknown; inspect device documentation"
    End If
    WScript.Echo SafeText(adapter.Name) & " / " & SafeText(adapter.AdapterCompatibility)
    WScript.Echo "Classification: " & classification
    WScript.Echo "PNPDeviceID=" & SafeText(adapter.PNPDeviceID)
Next

Function SafeText(value)
    If IsNull(value) Then
        SafeText = ""
    Else
        SafeText = CStr(value)
    End If
End Function
