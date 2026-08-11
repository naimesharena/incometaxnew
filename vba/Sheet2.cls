Attribute VB_Name = "Sheet2"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
''Application.OnKey "{TAB}"
''MsgBox ActiveCell.Address
'
'
'If Not Range(ActiveCell.Address).Locked = True Then
'       If ActiveCell.Validation.Type = 6 Then
'    Application.OnKey "%{DOWN}", ""
'    'Sheet1.Range(Target.Address).Value = ""
'ElseIf ActiveCell.Validation.Type = 1 Then
' Application.OnKey "%{DOWN}", ""
'ElseIf ActiveCell.Validation.Type = 3 Then
'
'    Application.OnKey "%{DOWN}"
'End If
'Else
'Application.OnKey "%{DOWN}", ""
'End If
'End Sub


Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False
ActiveWindow.DisplayHeadings = False

 
Sheet2.Unprotect Password:=EfilingCommon.getmsgstate


fmsgboxoK "Please verify the TAN details and the Amount of credit available in your Form 16 and 26AS" & Chr(10) & Chr(10) & "Please verify the Tax payment details as available in your form 26 AS"

If Mid(Sheet1.Range("sheet1.ReturnFileSec"), 1, 7) = "139(8A)" Then
Dim message As String
'message = "If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax, credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT and also at Sl. No. 14B of Part B-ATI. Please be noted that credit will be allowed only once. Also, REGULAR ASSESSMENT TAX if any paid needs to be disclosed at Sl. No. 8 & Sl. No. 14B of Part B-ATI. If the return is being filed u/s 139(8A) and any payments are made u/s 140B, please disclose those payments at Sl. No. 14A Part B-ATI "
'message = "If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax / credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT and also in separately in Schedule IT-2 as the credit for the same is not to be allowed under section 140B (2). Please be noted that credit will be allowed only once."
'Change 01, Satya, 13.09.2022
message = "If the return is being filed u/s 139(8A), then Details of payments of Advance Tax / Self Assessment Tax , credit for which has not been claimed in the earlier return needs to be disclosed in Schedule IT and also in Schedule IT-2.Please be noted that credit will be allowed only once."
'End change---
message = message & Chr(13) & "Also, ""REGULAR ASSESSMENT TAX"" if any paid needs to be disclosed in sl.no.8 of Part B_ATI and schedule IT-2 at sl.no.14B"
'message = message & Chr(13) & "If the return is being filed u/s 139(8A) and any payments are made u/s 140B, please disclose those payments in Schedule IT-1"
MsgBox (message)
MsgBox "If the return is being filed u/s 139(8A) and any payments are made u/s 140B, please disclose those payments in Schedule IT-1"

End If

   ActiveSheet.Shapes("Button 57").Left = Range("SchTDS.Button1").Left + 6
    ActiveSheet.Shapes("Button 57").Top = Range("SchTDS.Button1").Top + 5

    ActiveSheet.Shapes("Button 58").Left = Range("SchTDS.Button2").Left + 6
    ActiveSheet.Shapes("Button 58").Top = Range("SchTDS.Button2").Top + 1


    ActiveSheet.Shapes("Button 59").Left = Range("SchTDS.Button3").Left + 6
    ActiveSheet.Shapes("Button 59").Top = Range("SchTDS.Button3").Top + 1


     ActiveSheet.Shapes("Button_60").Left = Range("SchTDS.Button4").Left + 6
    ActiveSheet.Shapes("Button_60").Top = Range("SchTDS.Button4").Top + 1

Sheet2.Protect Password:=EfilingCommon.getmsgstate

endline:
Application.EnableEvents = True
Sheet2.Protect Password:=EfilingCommon.getmsgstate

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo exit1
    Dim cell As Range
    Dim cell1 As Range
    Dim cell80G1Amount As Range
    Dim cell80G2Amount As Range
    Dim cell80G3Amount As Range
    Dim incRangeCells, tottaxRangeCells As Range
    Dim i, mIntCells As Long
    Dim porRangeCells, amtClaimRangecells, totDeductRagecells
    Dim k, mIntCtr As Long
    Dim DateRangeCells As Range
    Dim x, dateCtr As Long

Application.ScreenUpdating = False
Application.EnableEvents = False
sPassword = EfilingCommon.getmsgstate
Sheet1.Unprotect Password:=sPassword
Sheet1.Range("taxcheck").Value = ""
Sheet1.Protect Password:=sPassword


'Ankita_28_04_2025
Dim TDS_Section_warning1, TDS_Section_warning2, TDS_Section_warning3, TDS_section_error As Range
Dim TDS_Section_warning_count1, TDS_Section_warning_count2, TDS_Section_warning_count3, TDS_section_error_count, iCount_warning1, iCount_warning2, iCount_warning3, iCount_error As Long

'Newly added by Bindu on 24th April As per DE V6
If Not Intersect(Range("TDsOthr2.SectionTDSDeducted"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("TDsOthr.SectionTDS"), Range(Target.Address)) Is Nothing Then

     If Target.Value <> "" Or UCase(Target.Value) <> UCase("(Select)") Then
       Set TDS_Section_warning1 = Sheet5.Range("TDS_Section_warning1").Cells
       TDS_Section_warning_count1 = Sheet5.Range("TDS_Section_warning1").count
       
       For iCount_warning1 = 1 To TDS_Section_warning_count1
         
         'Ankita_28/04/2025_SIT-92368
         If UCase(Trim(Target.Value)) = UCase(Trim(TDS_Section_warning1.item(iCount_warning1).Value)) Then
'            MsgBox "*Incomes chargeable to tax at special rates are not expected to be reported in ITR  4. Accordingly, TDS relating to the special rates income namely 194B,194BB,194BA,194IA,194IC,194LA,194S are also not expected to be claimed. Thus, please ensure that the return being filed in ITR 4 is not defective.", vbInformation
             MsgBox "*Incomes chargeable to tax at special rates are not expected to be reported in ITR 1. Accordingly, TDS relating to the special rates income namely 194B,194BB,194BA,194IA,194IC,194LA,194S are also not expected to be claimed. Thus, please ensure that the return being filed in ITR 1 is not defective." & Chr(13)
          End If
       Next
       
       Set TDS_Section_warning2 = Sheet5.Range("TDS_Section_warning2").Cells
       TDS_Section_warning_count2 = Sheet5.Range("TDS_Section_warning2").count
       
       For iCount_warning2 = 1 To TDS_Section_warning_count2
       
       'Ankita_28/04/2025_SIT-92368
         If UCase(Trim(Target.Value)) = UCase(Trim(TDS_Section_warning2.item(iCount_warning2).Value)) Then
'           MsgBox "*The Income on which the TDS under this section has been deducted is payable to a non-resident.Returns for income of non-residents are required to be filed in ITR 3/5. Thus, kindly ensure that the return being filed in ITR 4 is not defective.", vbInformation
            MsgBox "*The Income on which the TDS under this section has been deducted is payable to a non-resident.Returns for income of non-residents are required to be filed in ITR 2/3. Thus, kindly ensure that the return being filed in ITR 1 is not defective." & Chr(13)
  
          End If
       Next
       
       'Ankita_28/04/2025
       
       Set TDS_Section_warning3 = Sheet5.Range("TDS_Section_warning3").Cells
       TDS_Section_warning_count3 = Sheet5.Range("TDS_Section_warning3").count
       
       For iCount_warning3 = 1 To TDS_Section_warning_count3
       
          If UCase(Trim(Target.Value)) = UCase(Trim(TDS_Section_warning3.item(iCount_warning3).Value)) Then
            MsgBox "*The income on which TDS under this section has been deducted is not expected to be reported in ITR 1. Kindly ensure that the return being filed in ITR 1 is not defective." & Chr(13)
          End If
       Next
  
       
       Set TDS_section_error = Sheet5.Range("TDS_Section_Error").Cells
       TDS_section_error_count = Sheet5.Range("TDS_Section_Error").count
       
        For iCount_error = 1 To TDS_section_error_count
       
          If UCase(Trim(Target.Value)) = UCase(Trim(TDS_section_error.item(iCount_error).Value)) Then
            fmsgbox "*TDS relating to salaries is required to be claimed in Schedule TDS 1."
            Target.Value = "(Select)"
          End If
       
        Next
     
     
     End If

End If

'Ankita_07/05/2026=====
If Not Application.Intersect(Sheet2.Range("TDSal.IncChrgSalary"), Sheet2.Range(Target.Address)) Is Nothing Then
    If Sheet1.Range("grossSalary_exempt_familyPension").Value < Sheet2.Range("TDS.IncSum").Value Then
        fmsgbox ("""Sum of ""Gross Salary"" , ""Family Pension"" and ""exempt incomes which are part of salary"" declared is less than the amount of Gross Salary declared in Schedule TDS 1.""")
    End If
End If

  'Target.Value = UCase(Target.Value)
If Not Intersect(Range("TDSal.TAN"), Range(Target.Address)) Is Nothing Then
If Not checkfieldSuperSpecialcharacterDot(Range(Target.Address).Value) Then
fmsgbox "* Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet in Schedule TDS-1"
Range(Target.Address).Value = ""
Else
   Range(Target.Address).Value = UCase(Range(Target.Address))
End If
End If

If Not Intersect(Range("TDS26QB.Aadhar_Number"), Range(Target.Address)) Is Nothing Then
If Not checkfieldSuperSpecialcharacterDot(Range(Target.Address).Value) Then
fmsgbox "* Please enter the valid Aadhaar number"
Range(Target.Address).Value = ""
Else
   Range(Target.Address).Value = UCase(Range(Target.Address))
End If
End If

If Not Intersect(Range("TDSoth.TAN"), Range(Target.Address)) Is Nothing Then
   If Not checkfieldSuperSpecialcharacterDot(Range(Target.Address).Value) Then
fmsgbox "* Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet in Schedule TDS-2"
Range(Target.Address).Value = ""
Else
   Range(Target.Address).Value = UCase(Range(Target.Address))
   End If
End If

If Not Intersect(Range("TDS26QB.PAN"), Range(Target.Address)) Is Nothing Then
   Range(Target.Address).Value = UCase(Range(Target.Address))
End If



If Not Intersect(Range("TDSal.EmployerOrDeductorOrCollecterName"), Range(Target.Address)) Is Nothing Then
If Not checkfieldSuperSpecialcharactername(Range(Target.Address).Value) Then
         fmsgbox ("* Please do not enter Special characters < , >" & ",& in Schedule TDS-1" & Chr(13))
        Range(Target.Address).Value = ""
End If
End If

If Not Intersect(Range("TDSoth.EmployerOrDeductorOrCollecterName"), Range(Target.Address)) Is Nothing Then
If Not checkfieldSuperSpecialcharactername(Range(Target.Address).Value) Then
         fmsgbox ("* Please do not enter Special characters < , >" & ",& in Schedule TDS-2" & Chr(13))
        Range(Target.Address).Value = ""
End If
End If

If Not Intersect(Range("TDS26QB.EmployerOrDeductorName"), Range(Target.Address)) Is Nothing Then
If Not checkfieldSuperSpecialcharactername(Range(Target.Address).Value) Then
         fmsgbox ("* Please do not enter Special characters < , >" & ",& in Schedule TDS 3" & Chr(13))
        Range(Target.Address).Value = ""
End If
End If


If Not Intersect(Range("TaxP.DateDep"), Range(Target.Address)) Is Nothing Then
Set DateRangeCells = Sheet2.Range("TaxP.DateDep").Cells
dateCtr = Sheet2.Range("TaxP.DateDep").count ''''

For x = 1 To dateCtr
    If Not DateRangeCells.item(x).Value = "" Then
    
    'TAX DETAILS-C5  Bindu 2024-25
    If Not FormatNCheckDate(DateRangeCells.item(x).Value) Then
        fmsgbox ("*Please enter a valid date in dd/mm/yyyy format.")
        DateRangeCells.item(x).Value = ""
        DateRangeCells.item(x).Select
        Exit Sub
    End If
    
       'Ankita_06/01/2026=========================
       
       If Not CheckDateBefore3(DateRangeCells.item(x).Value) Then 'Bindu
'        Change28.11.2022.102.06B
'        fmsgbox "* Date of deposit cannot be prior to 01-04-2021 in Schedule IT"
            
            'TAX DETAILS-E5
            'fmsgbox "* Date of deposit cannot be prior to 01-04-2022 in Schedule IT"
            
            'TAX DETAILS-C5 2024-25 Bindu
            ' "Date of deposit cannot be prior to 01-04-20XX ".
            
            'Ankita_29/12/2025==========================
            Dim cutoff10 As Date
            cutoff10 = CDate(Sheet5.Range("IT_DOD1").Value)
'            fmsgbox "* ""Date of deposit cannot be prior to " & Dformat2(cutoff10, "dd/mm/yyyy") & " in Schedule IT"""
            fmsgbox "* ""Date of deposit cannot be prior to " & Dformat4(cutoff10, "dd/mm/yyyy") & " in Schedule IT"""
            '=========================
            
  '            fmsgbox "* ""Date of deposit cannot be prior to 01-04-2024 in Schedule IT" 'Year changed from 2023 to 2024 by Ankita on 29/11/2024
            
'            ---End
            DateRangeCells.item(x).Value = ""
            DateRangeCells.item(x).Select
       Else
'            If Dformat(Trim(DateRangeCells.item(x).Value), "yyyy-mm-dd") > Dformat(Trim(Sheet3.Range("Ver.Date").Value), "yyyy-mm-dd") Then
            If Not ChkMinInclusiveDate(Dformat(Trim(Sheet3.Range("Ver.Date").Value), ""), Dformat(Trim(DateRangeCells.item(x).Value), "")) Then
'            If Not ChkMinInclusiveDate("2023-05-05", Dformat(Trim(DateRangeCells.item(x).Value), "")) Then
                fmsgbox "* Date of deposit can not be more than system date in Income Details"
                DateRangeCells.item(x).Value = ""
                DateRangeCells.item(x).Select
            Else
                DateRangeCells.item(x).Value = CreditDate
            End If

        End If
        If Mid(Sheet1.Range("sheet1.ReturnFileSec1").Value, 1, 2) = "21" And (Not DateRangeCells.item(x).Value = "") Then
            
            'PAG_E3
            'If Not ChkMinInclusiveDate("2024-01-01", Dformat(DateRangeCells.item(x).Value, "")) Then
            
            'PAG_C3 2024-25 Bindu
              'Ankita_25-26
'            If Not ChkMinInclusiveDate("2024-12-31", Dformat(DateRangeCells.item(x).Value, "")) Then
              '  fmsgbox "Please enter the details of 140B in Sl No 14A of Part B-ATI and not in this schedule, so as to get the correct computation of your tax liability"
               'Ankita_25-26
'                fmsgbox "Date of deposit can't be after 31-12-2024 in schedule IT. Please enter payments made after 31st December 2024 in schedule IT 1."
'                 If Not ChkMinInclusiveDate("2025-12-31", Dformat(DateRangeCells.item(x).Value, "")) Then
'                 fmsgbox "Date of deposit can't be after 31-12-2025 in schedule IT. Please enter payments made after 31st December 2025 in schedule IT 1."
'Ankita_23/12/2025=======================

Dim cutoff As Date
cutoff = CDate(Sheet5.Range("DOB_1").Value)
'                If Not ChkMinInclusiveDate("2025-12-31", Dformat(DateRangeCells.item(x).Value, "")) Then
                 If Not ChkMinInclusiveDate(Dformat(Trim(Sheet5.Range("DOB_1").Value), "yyyy-mm-dd"), Dformat(DateRangeCells.item(x).Value, "")) Then
                 fmsgbox "Date of deposit can't be after  " & Format(cutoff, "dd/mm/yyyy") & " in schedule IT. Please enter payments made after 31st December 2026 in schedule IT 1."
            End If
        End If
    End If
Next
End If

exit1:
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub
Sub enable()
Application.EnableEvents = True
End Sub
 
