Attribute VB_Name = "Sheet202"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Sub LockUnlockATI_1()
    Application.EnableEvents = False
    Sheet202.Unprotect Password:=getmsgstate
    
    If Sheet202.Range("U_LastAmtPayable").Value > 0 Then
        Sheet202.Range("U_Refund").MergeArea.ClearContents
        Sheet202.Range("U_Refund").MergeArea.Locked = True
        Sheet202.Range("U_Refund").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_Refund").Value = ""
        
        Sheet202.Range("U_TotRefund").MergeArea.ClearContents
        Sheet202.Range("U_TotRefund").MergeArea.Locked = True
        Sheet202.Range("U_TotRefund").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_TotRefund").Value = ""
    Else
        Sheet202.Range("U_Refund").MergeArea.Locked = False
        Sheet202.Range("U_Refund").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet202.Range("U_Refund").Value = 0
        Sheet202.Range("U_TotRefund").MergeArea.Locked = False
        Sheet202.Range("U_TotRefund").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet202.Range("U_TotRefund").Value = 0
    End If
    
    Application.EnableEvents = True
    Sheet201.Protect Password:=getmsgstate
End Sub
Sub a()
Application.EnableEvents = True

End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub Worksheet_Calculate()

   
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
'Change28.11.2022.102.01IT
PartBATI_IT1DateValidation Target
PartBATI_IT2DateValidation Target
'---end
Sheet1.Range("taxcheck").Value = ""
    
Sheet202.Unprotect Password:=getmsgstate
Application.EnableEvents = False

'Added by Riyaz on 05/12/2025

If Not Application.Intersect(Sheet202.Range("U_DateDep1"), Sheet202.Range(Target.Address)) Is Nothing Then

Dim dateOfdeposit As Variant
Dim AssYear As Long
Dim yrdop As Long
Dim mthdop As Long

dateOfdeposit = Target.Value
        
        AssYear = 2025
        dateOfdeposit = Dformat(dateOfdeposit, "yyyy-mm-dd")
        
        yrdop = CLng(Mid(dateOfdeposit, 1, 4))
        mthdop = CLng(Mid(dateOfdeposit, 6, 2))
        
            If (yrdop >= AssYear) Then
            
                dateOfdeposit = mthdop - 4 + (yrdop - AssYear) * 12
        
            End If
        
        dateOfdeposit = dateOfdeposit + 1
    Sheet202.Range(Replace(Replace(Target.Address, "$", ""), "G", "T")).Value = dateOfdeposit
End If

'End

If Target.Column = 6 Then
    Target.Value = UCase(Target.Value)
End If


'=========================================
'Ankita_UR
If Not Application.Intersect(Sheet202.Range("U_BSRCode2"), Sheet202.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet202.Range("U_DateDep2"), Sheet202.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet202.Range("U_SrlNoOfChaln2"), Sheet202.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet202.Range("U_Amt2"), Sheet202.Range(Target.Address)) Is Nothing Then
 
    If Sheet201.Range("U_PreviouslyFiledForThisAY").Value = "No" Then
         LOCKRANGEIT2
    End If
End If
'================================================================
If Not Application.Intersect(Sheet202.Range("U_LastAmtPayable"), Sheet202.Range(Target.Address)) Is Nothing Then
LockUnlockATI_1
End If

If Not Application.Intersect(Sheet202.Range("U_Refund"), Sheet202.Range(Target.Address)) Is Nothing Then
    If Sheet202.Range("U_Refund").Value > 0 Or Sheet202.Range("U_TotRefund").Value > 0 Then
        Sheet202.Range("U_LastAmtPayable").MergeArea.ClearContents
        Sheet202.Range("U_LastAmtPayable").MergeArea.Locked = True
        Sheet202.Range("U_LastAmtPayable").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_LastAmtPayable").Value = ""
    Else
        Sheet202.Range("U_LastAmtPayable").MergeArea.Locked = False
        Sheet202.Range("U_LastAmtPayable").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet202.Range("U_LastAmtPayable").Value = 0
    End If
End If

If Not Application.Intersect(Sheet202.Range("U_TotRefund"), Sheet202.Range(Target.Address)) Is Nothing Then
    If Sheet202.Range("U_TotRefund").Value > 0 Or Sheet202.Range("U_Refund").Value > 0 Then
        Sheet202.Range("U_LastAmtPayable").MergeArea.ClearContents
        Sheet202.Range("U_LastAmtPayable").MergeArea.Locked = True
        Sheet202.Range("U_LastAmtPayable").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet202.Range("U_LastAmtPayable").Value = ""

    Else
        Sheet202.Range("U_LastAmtPayable").MergeArea.Locked = False
        Sheet202.Range("U_LastAmtPayable").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet202.Range("U_LastAmtPayable").Value = 0
    End If
End If

'Ankita_UR
If Not Application.Intersect(Sheet202.Range("U_LastAmtPayable"), Sheet202.Range(Target.Address)) Is Nothing Then
     
     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
        If Sheet202.Range("U_LastAmtPayable").Value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then ""Amount payable as per prior ITR can’t be more than 0""", , "Error"
           Sheet202.Range("U_LastAmtPayable").ClearContents
        
        End If
     End If
End If

If Not Application.Intersect(Sheet202.Range("U_Refund"), Sheet202.Range(Target.Address)) Is Nothing Then
     
     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
        If Sheet202.Range("U_Refund").Value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then ""Refund claimed can’t be more than 0""", , "Error"
           Sheet202.Range("U_Refund").ClearContents
        
        End If
     End If
End If

If Not Application.Intersect(Sheet202.Range("U_TotRefund"), Sheet202.Range(Target.Address)) Is Nothing Then
     
     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
        If Sheet202.Range("U_TotRefund").Value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           MsgBox "If A5 in Part A gen 139(8A)  is ""No"", then ""Refund issued can’t be more than 0""", , "Error"
           Sheet202.Range("U_TotRefund").ClearContents
        
        End If
     End If
End If

If Not Application.Intersect(Sheet202.Range("U_RegAssessementTAX"), Sheet202.Range(Target.Address)) Is Nothing Then
     
     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
        If Sheet202.Range("U_RegAssessementTAX").Value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
'           MsgBox "If A5 in Part A gen 139(8A)  is ""No"", then ""Regular Assessment Tax"" can't be claimed""", , "Error"
            MsgBox "If A5 in Part A gen 139(8A)  is ""No"", then ""Regular Assessment Tax"" can't be claimed", , "Error"
           
           Sheet202.Range("U_RegAssessementTAX").ClearContents
        
        End If
     End If
End If

'added by Chetan C M on 13/11/2025 for SIT-103528
'start--
If Not Application.Intersect(Sheet202.Range("U_DateDep1"), Sheet202.Range(Target.Address)) Is Nothing Then
 
    If Not CheckDateddmmyyyy(Sheet202.Range(Target.Address).Value) Then
     MsgBox ("""Please enter valid date in dd/mm/yyyy format in schedule IT-1. Also date of credit to Central government cannot be prior to 01/01/2026 in schedule IT-1.""")

    Sheet202.Range(Target.Address).Value = ""
    Target.Select
    End If
End If

If Not Application.Intersect(Sheet202.Range("U_DateDep2"), Sheet202.Range(Target.Address)) Is Nothing Then
 
    If Not CheckDateddmmyyyy(Sheet202.Range(Target.Address).Value) Then
     MsgBox ("""Please enter valid date in dd/mm/yyyy format in schedule IT-2. Also date of credit to Central government cannot be prior to 01/04/2024 in schedule IT-2.""")

    Sheet202.Range(Target.Address).Value = ""
    Target.Select
    End If
End If
'--end

endline:
Application.EnableEvents = True
Sheet202.Protect Password:=getmsgstate
End Sub






