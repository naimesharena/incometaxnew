Attribute VB_Name = "checkBoxModule"
Sub SelectCheckBox()
Attribute SelectCheckBox.VB_ProcData.VB_Invoke_Func = "q\n14"

If ActiveSheet.name = "Taxes Paid and Verification" Then
    For Each chk In ActiveSheet.CheckBoxes
        If UCase(Mid(chk.name, 1, 9)) = "CHECK BOX" Then
            If Not Intersect(Range(ActiveCell.Address), chk.TopLeftCell) Is Nothing Then
                chk.Value = 1
            End If
        End If
    Next chk
End If
End Sub
Sub DeselectCheckBox()
Attribute DeselectCheckBox.VB_ProcData.VB_Invoke_Func = "r\n14"

If ActiveSheet.name = "Taxes Paid and Verification" Then
    For Each chk In ActiveSheet.CheckBoxes
        If UCase(Mid(chk.name, 1, 9)) = "CHECK BOX" Then
            If Not Intersect(Range(ActiveCell.Address), chk.TopLeftCell) Is Nothing Then
                chk.Value = 0
            End If
        End If
    Next chk
End If
End Sub

