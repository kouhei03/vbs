Option Explicit
Dim service, records, record, fieldName, fieldValue, arrayItem, joinedValue
Set service = GetObject("winmgmts:\\.\root\cimv2\Security\MicrosoftTpm")
Set records = service.ExecQuery("SELECT SpecVersion, IsEnabled_InitialValue, IsActivated_InitialValue, ManufacturerIdTxt FROM Win32_Tpm")
If records.Count = 0 Then
    WScript.Echo "No data returned for: TPM搭載/有効状態"
    WScript.Quit 0
End If
For Each record In records
    WScript.Echo "---"
    For Each fieldName In Array("SpecVersion", "IsEnabled_InitialValue", "IsActivated_InitialValue", "ManufacturerIdTxt")
        fieldValue = record.Properties_.Item(fieldName).Value
        If IsNull(fieldValue) Then
            WScript.Echo fieldName & ": (not reported)"
        ElseIf IsArray(fieldValue) Then
            joinedValue = ""
            For Each arrayItem In fieldValue
                If Len(joinedValue) > 0 Then joinedValue = joinedValue & ", "
                joinedValue = joinedValue & CStr(arrayItem)
            Next
            WScript.Echo fieldName & ": " & joinedValue
        Else
            WScript.Echo fieldName & ": " & CStr(fieldValue)
        End If
    Next
Next
