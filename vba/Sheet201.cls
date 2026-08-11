Attribute VB_Name = "Sheet201"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Sub UNLOCKRANGE_General_1398A(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = rangenamestring
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet201.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = False
                    chcell.MergeArea.Interior.Color = "&HCCFFCC"
                    chcell.Value = ""
                Else
                    chcell.Locked = False
                    chcell.Interior.Color = "&HCCFFCC"
                    chcell.Value = ""
                End If
            Next
        Next
End Sub

Sub temp()
Sheet202.Protect Password:=getmsgstate
'ThisWorkbook.Protect Password:=getmsgstate
Application.EnableEvents = True

End Sub
Sub LOCKRANGE_General_1398A(rangenamestring As Variant)
Dim i As Long
Dim chcell As Range
    rangenamestring = Split(rangenamestring, "||")
        For i = 0 To UBound(rangenamestring)
            For Each chcell In Sheet201.Range(rangenamestring(i))
                If chcell.MergeCells Then
                    chcell.MergeArea.Locked = True
                    chcell.MergeArea.Interior.Color = "&HD8D8D8"
                    chcell.Value = ""
                Else
                    chcell.Locked = True
                    chcell.Interior.Color = "&HD8D8D8"
                    chcell.Value = ""
                End If
            Next
        Next
End Sub
Sub LockUnlockBATI()
    Application.EnableEvents = False
    Sheet202.Unprotect Password:=getmsgstate
    Sheet201.Unprotect Password:=getmsgstate
    
    If Sheet201.Range("U_PreviouslyFiledForThisAY").Value = "No" Then
        Sheet202.Range("U_LatestTotInc").MergeArea.ClearContents
        Sheet202.Range("U_LatestTotInc").MergeArea.Locked = True
        Sheet202.Range("U_LatestTotInc").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_LatestTotInc").Value = ""
        
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.ClearContents
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.Locked = True
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Value = ""
        
        Sheet201.Range("U_ITRForm").MergeArea.ClearContents
        Sheet201.Range("U_ITRForm").MergeArea.Locked = True
        Sheet201.Range("U_ITRForm").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet201.Range("U_ITRForm").Value = ""
        
        Sheet201.Range("U_AcknowledgementNo").MergeArea.ClearContents
        Sheet201.Range("U_AcknowledgementNo").MergeArea.Locked = True
        Sheet201.Range("U_AcknowledgementNo").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet201.Range("U_AcknowledgementNo").Value = ""
        
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.ClearContents
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.Locked = True
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet201.Range("U_OrigRetFiledDate").Value = ""
        
        'Ankita_UR
        LOCKRANGEIT2
    Else
        Sheet202.Range("U_LatestTotInc").MergeArea.Locked = False
        Sheet202.Range("U_LatestTotInc").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.Locked = False
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Value = "(Select)"
        
        Sheet201.Range("U_ITRForm").MergeArea.Locked = False
        Sheet201.Range("U_ITRForm").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet201.Range("U_ITRForm").Value = "(Select)"
        
        Sheet201.Range("U_AcknowledgementNo").MergeArea.Locked = False
        Sheet201.Range("U_AcknowledgementNo").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.Locked = False
        Sheet201.Range("U_OrigRetFiledDate").MergeArea.Interior.Color = "&HCCFFCC"
           'Ankita_UR
        UNLOCKRANGEIT2
    End If
    
    If Sheet201.Range("U_PreviouslyFiledForThisAY").Value = "(Select)" Then
        Sheet202.Range("U_LatestTotInc").MergeArea.ClearContents
        Sheet202.Range("U_LatestTotInc").MergeArea.Locked = True
        Sheet202.Range("U_LatestTotInc").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_LatestTotInc").Value = ""
    End If
    Application.EnableEvents = True
    Sheet202.Protect Password:=getmsgstate
    Sheet201.Protect Password:=getmsgstate
End Sub
Sub LockUnlockGenYear()
    Application.EnableEvents = False
    Sheet201.Unprotect Password:=getmsgstate
    
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value <> "1" Then
'
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = False
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet201.Range("U_RevisedReturnFile").Value = "(Select)"
'
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.ClearContents
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_UpdatedReturnFile").Value = ""
'
'    ElseIf Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value = "1" Then
'
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Locked = False
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet201.Range("U_UpdatedReturnFile").Value = "(Select)"
'
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.ClearContents
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_RevisedReturnFile").Value = ""
'    Else
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.ClearContents
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_RevisedReturnFile").Value = ""
'
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.ClearContents
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_UpdatedReturnFile").Value = ""
'    End If
    
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value >= "1" Then
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Locked = False
'        Sheet201.Range("U_UpdatedReturnFile").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet201.Range("U_UpdatedReturnFile").Value = "(Select)"
'    End If
'
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_2").Value >= "1" Then
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = False
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet201.Range("U_RevisedReturnFile").Value = "(Select)"
'    End If
    
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value = "0" Then
'        Sheet201.Range("U_UpdatedReturnFile").ClearContents
'        Sheet201.Range("U_UpdatedReturnFile").Locked = True
'        Sheet201.Range("U_UpdatedReturnFile").Interior.Color = "&HD8D8D8"
'    End If
'
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_2").Value = "0" Then
'        Sheet201.Range("U_RevisedReturnFile").ClearContents
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Locked = True
'        Sheet201.Range("U_RevisedReturnFile").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet201.Range("U_RevisedReturnFile").Value = ""
'    End If
    
    Application.EnableEvents = True
    Sheet201.Protect Password:=getmsgstate
End Sub
Sub LockUnlockGenReason()
    Application.EnableEvents = False
    Sheet201.Unprotect Password:=getmsgstate
    
    If Sheet201.Range("Reason1").Value = 1 Or Sheet201.Range("Reason2").Value = 1 Or Sheet201.Range("Reason3").Value = 1 Then
        
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Locked = False
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Interior.Color = "&HCCFFCC"
        'Sheet201.Range("U_UnabsorbedDepreciation").Value = "Yes"
        'Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Locked = True
    Else
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.ClearContents
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Locked = True
        Sheet201.Range("U_UnabsorbedDepreciation").MergeArea.Interior.Color = "&HD8D8D8"
                   
        Sheet201.Range("U_RevisedReturnFile").ClearContents
        Sheet201.Range("U_RevisedReturnFile").Locked = True
        Sheet201.Range("U_RevisedReturnFile").Interior.Color = "&HD8D8D8"

        
        Sheet201.Range("U_UpdatedReturnFile").ClearContents
        Sheet201.Range("U_UpdatedReturnFile").Locked = True
        Sheet201.Range("U_UpdatedReturnFile").Interior.Color = "&HD8D8D8"
        
        Sheet201.Range("U_UnabsorbedDepreciationYear").ClearContents
        Sheet201.Range("U_UnabsorbedDepreciationYear").Locked = True
        Sheet201.Range("U_UnabsorbedDepreciationYear").Interior.Color = "&HD8D8D8"

        
    End If
    
            
    
    Application.EnableEvents = True
    Sheet201.Protect Password:=getmsgstate
End Sub
Function chkNumeric_1398A(field As Variant) As Boolean
    Dim k As Long
    Dim chkchar As String

    chkNumeric_1398A = True
    For k = 1 To Len(field)
    chkchar = Mid(field, k, 1)
        If Not IsNumeric(chkchar) Then
            chkNumeric_1398A = False
            Exit Function
        End If
    Next
End Function

'Sub ReturnDropdownA12(RowIndex As Long)
'On Error GoTo endline
'Application.EnableEvents = False
'Sheet201.Unprotect Password:=getmsgstate
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = "2024-25" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'End If
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = "2025-26" Then
'    If ChkMinInclusiveDate(Dformat("31/12/2024", ""), Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
'    End If
'End If
'endline:
'Sheet201.Protect Password:=getmsgstate
'    Application.EnableEvents = True
'End Sub

Sub ReturnDropdownA12(RowIndex As Long)
On Error GoTo endline
Application.EnableEvents = False
Sheet201.Unprotect Password:=getmsgstate



'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = "2024-25" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).Value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2024", "") <= Dformat(Sheet3.Range("Ver.Date").Value, "")) And (Dformat("31/03/2025", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2023", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If

'Ankita_UR

'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = "2025-26" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).Value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("31/12/2025", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        ElseIf (Dformat("01/01/2026", "") < Dformat(Sheet3.Range("Ver.Date").Value, "")) Or (Dformat("31/03/2026", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
'        End If
'    End If
'End If
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = " 2026-27" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).Value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        'Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("31/12/2026", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        ElseIf (Dformat("01/01/2027", "") < Dformat(Sheet3.Range("Ver.Date").Value, "")) Or (Dformat("31/03/2027", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
'        End If
'    End If
'End If



'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = "2025-26" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).Value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2025", "") <= Dformat(Sheet3.Range("Ver.Date").Value, "")) And (Dformat("31/03/202", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2024", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If
endline:
Sheet201.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
'Ankita_16/04/2025
Sheet1.Range("taxcheck").Value = "Y"


Application.EnableEvents = False



'If Not Application.Intersect(Sheet201.Range("U_AadhaarCardNo"), Sheet201.Range(Target.Address)) Is Nothing Then
'    If Not chkNumeric(Target.Value) Then
'        MsgBox "Please enter valid 12 digit aadhar number"
'        Target.Value = ""
'    End If
'End If
'Change28.11.2022.102.06B
Table12_139_23_24Dropdown (Target.row)
'ReturnDropdownA12 (Target.row)
'---end change


'Ankita_UR

If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
 
Dim dd1 As String

Dim dd2 As String
 
'dd1 = "2025-03-31"
'dd2 = "2026-03-31"
 
'Ankita_03/11/2025===========================
 
dd1 = "2026-03-31"
dd2 = "2027-03-31"

If Target.Value = "2026-27" Then
If Dformat(Sheet3.Range("Ver.Date"), "") <= dd1 Then
MsgBox "AY ""2026-27"" can be selected after 31/03/2026", vbOKOnly, "Alert"
Target.Value = "(Select)"
End If
Else
If Target.Value = "2027-28" Then
If Dformat(Sheet3.Range("Ver.Date"), "") <= dd2 Then
MsgBox "AY ""2027-28"" can be selected after 31/03/2027", vbOKOnly, "Alert"
Target.Value = "(Select)"
End If
End If
End If
End If
'============================================
 
 
'If Target.Value = "2025-26" Then
'If Dformat(Sheet3.Range("Ver.Date"), "") <= dd1 Then
'MsgBox "AY ""2025-26"" can be selected on or after 01/04/2025", vbOKOnly, "Alert"
'Target.Value = "(Select)"
'End If
'Else
'If Target.Value = "2026-27" Then
'If Dformat(Sheet3.Range("Ver.Date"), "") <= dd2 Then
'MsgBox "AY ""2026-27"" can be selected on or after 01/04/2026", vbOKOnly, "Alert"
'Target.Value = "(Select)"
'End If
'End If
'End If
'End If

If Not Application.Intersect(Sheet201.Range("U_PreviouslyFiledForThisAY"), Sheet201.Range(Target.Address)) Is Nothing Then
LockUnlockBATI
End If

If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
'LockUnlockGenYear
    Application.EnableEvents = False
    Sheet201.Unprotect Password:=getmsgstate
    'Change28.11.2022.102.06A
'    If Target.Value = "2023-24" Or Target.Value = "2024-25" Then
'Ankita_UR

'    If Target.Value = "2024-25" Or Target.Value = "2025-26" Then
     If Target.Value = "2026-27" Or Target.Value = "2027-28" Then   'Ankita_03/11/2025
'     If Target.Value = "2026-27" Or Target.Value = "2025-26" Then
'   ---end change


'   Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
  
'   Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.ClearContents
    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Locked = False
    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"

'    ElseIf Target.Value = "2024" Then
'
''    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.ClearContents
'    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Locked = False
'    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'
'    '    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
    
    Else
    
    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
    Sheet201.Cells(Target.row, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"

    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.ClearContents
    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Locked = True
    Sheet201.Cells(Target.row, Sheet201.Range("U_RevisedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
    End If
    Application.EnableEvents = True
    Sheet201.Protect Password:=getmsgstate
    Application.EnableEvents = False
End If



'Enhancement

'If Not Application.Intersect(Sheet201.Range("U_UpdatedReturnDuringPeriod"), Sheet201.Range(Target.Address)) Is Nothing Then
'
'Dim dd3 As String
'Dim dd4 As String
'
'dd3 = "2026-04-01"
'dd4 = "2026-03-31"
'
''If Target.Value = "Up to 12 months from the end of Relevant Assessment Year" Then
''If Dformat(Sheet3.Range("Ver.Date"), "") >= dd3 Then
''MsgBox """Please select appropriate option in 'Are you filing the updated return during the period.'""", vbOKOnly, "Alert"
''Target.Value = "(Select)"
''End If
''Else
'If Target.Value = "Between 12 to 24 Months from the end of Relevant Assessment  Year" Then
'If Dformat(Sheet3.Range("Ver.Date"), "") <= dd4 Then
'MsgBox """Please select appropriate option in 'Are you filing the updated return during the period.'""", vbOKOnly, "Alert"
'Target.Value = "(Select)"
'End If
'End If
'End If
''End If




''Ankita_UR
'If Not Application.Intersect(Sheet201.Range("U_PreviouslyFiledForThisAY"), Sheet201.Range(Target.Address)) Is Nothing Then
'If UCase(Target.Value) = "NO" Then
'If Sheet5.Range("BacValue").Value = 2 Then
''MsgBox "* ""Please select ""If no, whether you are choosing to opt for taxation under section 115BA/115BAA/115BAB this year?"" as ""No"" in Part A General if Sl.no A5 in Part A General 139(8A) is selected as ""No""", vbOKOnly, "Alert"
'MsgBox "*""Please select  Are you opting for new tax regime u/s 115BAC ? as ""NO"" in Part A General if Sl.no A5 in Part A General 139(8A)  is selected as ""No""", vbOKOnly, "Alert"
'End If
'End If
'End If

'Ankita_UR
If Not Application.Intersect(Sheet201.Range("U_ReasonsForUpdatingIncome"), Sheet201.Range(Target.Address)) Is Nothing Then
   If UCase(Target.Value) <> UCase("(Select)") Then
    If WorksheetFunction.CountIf(Sheet201.Range("U_ReasonsForUpdatingIncome"), Target.Value) > 1 Then
        Application.EnableEvents = False
            MsgBox "You cannot select the same reason for updating your income twice" & Chr(13)
            Target.Value = ""
        Application.EnableEvents = True
    End If
    End If
LockUnlockGenReason
End If


If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
      If UCase(Target.Value) <> UCase("(Select)") Then
        If WorksheetFunction.CountIf(Sheet201.Range("U_UnabsorbedDepreciationYear"), Target.Value) > 1 Then
            Application.EnableEvents = False
                MsgBox "You can not select the same dropdown for twice", , "Error"
                Target.Value = ""
            Application.EnableEvents = True
        End If
      End If
End If

If Not Application.Intersect(Sheet201.Range("U_AcknowledgementNo"), Sheet201.Range(Target.Address)) Is Nothing Then
Application.EnableEvents = False
Sheet1.Range("sheet1.ReceiptNo").Value = Sheet201.Range("U_AcknowledgementNo").Value
Application.EnableEvents = True
End If

If Not Application.Intersect(Sheet201.Range("U_OrigRetFiledDate"), Sheet201.Range(Target.Address)) Is Nothing Then
Sheet1.Range("sheet1.OrigRetFiledDate").Value = Sheet201.Range("U_OrigRetFiledDate").Value
End If

'Ankita_03/11/2025=======================
If Not Application.Intersect(Sheet201.Range("U_OrigRetFiledDate"), Sheet201.Range(Target.Address)) Is Nothing Then
    Dim filingDate As Date
    Dim minDate As Date
   ' minDate = DateSerial(2025, 4, 1)
     minDate = Dformat("01-04-2025", "")
    'filingDate = Sheet201.Range(Target.Address).Value
    filingDate = Dformat(Sheet201.Range(Target.Address).Value, "")
    
    ' Check condition
    If filingDate < minDate Then
        MsgBox "Date of filing cannot be prior to 01/04/2025 in Part A General 139(8A)" & Chr(13)
        Sheet201.Range(Target.Address).Value = ""
    End If
End If
'========================================


'Ankita_03/11/2025===============================
If Not Application.Intersect(Sheet201.Range("U_LaidOutIn_139_8A"), Sheet201.Range(Target.Address)) Is Nothing Then
If Sheet201.Range(Target.Address).Value = "No" Then
'MsgBox "Dear user, you will not be able to proceed to file Updated Return u/s 139(8A) as you are not eligible for filing an updated return as per the conditions laid out in first, second and third provisos to section 139(8A)" & Chr(13)
MsgBox "Dear user, you will not be able to proceed to file Updated Return u/s 139(8A) as you are not eligible for filing an updated return as per the conditions laid out in first, second, third and fourth provisos to section 139(8A)" & Chr(13)
Sheet201.Range(Target.Address).Value = "(Select)"
End If
End If
'================================================

''If (Target.name.name) = "U_UnabsorbedDepreciation" Then
If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciation"), Sheet201.Range(Target.Address)) Is Nothing Then
    Application.EnableEvents = False
        Sheet201.Unprotect Password:=getmsgstate
    If Target.Value = "Yes" Then
    
'    Sheet201.Range("U_UnabsorbedDepreciationYear").ClearContents
    
    Sheet201.Range("U_UnabsorbedDepreciationYear").Locked = False
    Sheet201.Range("U_UnabsorbedDepreciationYear").Interior.Color = "&HCCFFCC"
        
    Else
    
    Sheet201.Range("U_UnabsorbedDepreciationYear").ClearContents
    Sheet201.Range("U_UnabsorbedDepreciationYear").Locked = True
    Sheet201.Range("U_UnabsorbedDepreciationYear").Interior.Color = "&HD8D8D8"
    
    Sheet201.Range("U_RevisedReturnFile").ClearContents
    Sheet201.Range("U_RevisedReturnFile").Locked = True
    Sheet201.Range("U_RevisedReturnFile").Interior.Color = "&HD8D8D8"
    
    Sheet201.Range("U_UpdatedReturnFile").ClearContents
    Sheet201.Range("U_UpdatedReturnFile").Locked = True
    Sheet201.Range("U_UpdatedReturnFile").Interior.Color = "&HD8D8D8"
    End If
    Application.EnableEvents = True
        Sheet201.Protect Password:=getmsgstate
End If
'
'     Dim RangeArray As Variant
'     RangeArray = "U_UnabsorbedDepreciationYear"
'     Sheet201.Unprotect Password:=getmsgstate
'        If Sheet201.Range("Reason1").Value = 1 Or Sheet201.Range("Reason2").Value = 1 Or Sheet201.Range("Reason3").Value = 1 Then
'
'        UNLOCKRANGE_General_1398A (RangeArray)
'           Else
'
'          LOCKRANGE_General_1398A (RangeArray)
'
'         End If
'       Sheet201.Protect Password:=getmsgstate
'End If

'If Not Application.Intersect(Sheet201.Range("U_UnabsorbedDepreciationYear"), Sheet201.Range(Target.Address)) Is Nothing Then
'    If Sheet201.Range("U_UnabsorbedDepreciation").Value <> "Yes" Then
'    If Sheet201.Range("U_UnabsorbedDepreciationYear_1").Value = 1 Or Sheet201.Range("U_UnabsorbedDepreciationYear_2").Value = 1 Then
'    fmsgbox "* Year you can not select if updated return is not to reduce carried forward loss or unabsorbed depreciation or tax credit" & Chr(13)
'    Sheet201.Range(Target.Address).Value = ""
'    End If
'    End If
'End If



'If Not Application.Intersect(Sheet201.Range("U_AadhaarEnrolmentId"), Sheet201.Range(Target.Address)) Is Nothing Then
''If (Target.name.name) = "U_AadhaarEnrolmentId" Then
'    If Len(Sheet201.Range("U_AadhaarEnrolmentId").Value) >= 12 Then
'        If IsNumeric(Sheet201.Range("U_AadhaarEnrolmentId").Value) Then
'
'            If Not checkfieldSuperSpecialcharacterDot_1398A(Sheet201.Range("U_AadhaarEnrolmentId").Value) Then
'                fmsgbox "* Please enter the valid Aadhaar Enrolment ID in Sheet Income Details"
'                Sheet1.Range("Sheet1.AadhaarEnrol").Value = ""
'            End If
'
'            If Sheet201.Range("U_AadhaarEnrolmentId").Value <> "0000000000000000000000000000" Then
'              Sheet201.Range("U_AadhaarEnrolmentId").Value = Sheet201.Range("U_AadhaarEnrolmentId").Value
'            Else
'              fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'              Sheet1.Range("U_AadhaarEnrolmentId").Value = ""
'            End If
'
'
'            If Sheet201.Range("U_AadhaarEnrolmentId").Value < 0 Then
'                fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'                Sheet201.Range("U_AadhaarEnrolmentId").Value = ""
'            End If
'
'            If Sheet201.Range("U_AadhaarEnrolmentId").Value <> "1111111111111111111111111111" Then
'                Sheet201.Range("U_AadhaarEnrolmentId").Value = Sheet201.Range("U_AadhaarEnrolmentId").Value
'            Else
'                fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'                Sheet201.Range("U_AadhaarEnrolmentId").Value = ""
'            End If
'
'        Else
'            fmsgbox "* Please enter the valid Aadhaar Enrolment ID in schedule Part A General 139(8A)"
'            Sheet201.Range("U_AadhaarEnrolmentId").Value = ""
'        End If
'    Else
'        Sheet201.Range("U_AadhaarEnrolmentId").Value = ""
'    End If
'End If
If (Target.name.name) = "U_AcknowledgementNo" Then
    Dim urecieptno As Variant
    Application.EnableEvents = False
    
    urecieptno = Sheet201.Range("U_AcknowledgementNo").Value
    
    If Not chkNumeric_1398A(urecieptno) Then
        fmsgbox "* Invalid Acknowledgement Number. Please retry "
        Sheet201.Range("U_AcknowledgementNo").Value = ""
        Sheet201.Range("U_AcknowledgementNo").Select
    End If
    
    If Sheet201.Range("U_AcknowledgementNo").Value <> "000000000000000" Then
        Sheet201.Range("U_AcknowledgementNo").Value = Sheet201.Range("U_AcknowledgementNo").Value
    Else
        fmsgbox "* Invalid Acknowledgement Number. Please retry "
        Sheet201.Range("U_AcknowledgementNo").Value = ""
        Sheet201.Range("U_AcknowledgementNo").Select
    
    End If
    Application.EnableEvents = True
End If



endline:
Application.EnableEvents = True

End Sub



