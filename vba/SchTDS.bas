Attribute VB_Name = "SchTDS"
Option Explicit

Public end_I, end_I2, end_I3, end_I4, rngname_I, rngname_I2 As Variant
Public TAN_TDS As Variant
Public Tan2_TDS, DeductName2_TDS, Year2_TDS, TotTaxDedct_TDS, ClaimAmt2_TDS, TaxDeducted_TDS, BSR_TDS, DateCredit_TDS, SerialNum_TDS, TaxPaid3_TDS As Variant
Public PAN_TDS, Tenant_Aadhar_TDS, DeductName2_TDS3, Year2_TDS3, TotTaxDedct_TDS3, inownhands6_TDS3, ClaimAmt2_TDS3, TaxDeducted_TDS3 As Variant
Public Portuguese2_TDS As Variant
Public Tot_chrg As Variant
Public TsEmpName As Variant
Public Amount4_I As Variant
Public UTN_TDS2 As Variant
Public ColCount3b_3, ColCount2b_2, Section_TDS2, Section_TDS3 As Variant 'Malli 23/04/2025
Public Msgbox_TDS, MsgBox_TDS2, MsgBox_TDS3 As String
Public ColCount1, ColCountAadhar, ColCount1_1, ColCount1_2, ColCount1_3, ColCount1_4, ColCount1_5, ColCount1_6, ColCount1_7, ColCount1_8, ColCount1_9, ColCount1_10, ColCount1_11, UTNCount As Variant '03/02/2025   ColCount1_12_removed on 12/02/2025 as er V0.4
Public ColCount2, ColCount2_1, ColCount2_2, ColCount2_3 As Variant
Public ColCount3, ColCount3_1, ColCount3_2, ColCount3_3, ColCount3_4, ColCount3_5, ColCount3_6, ColCount3_7, ColCount3_8, ColCount3_9, ColCount3_10, ColCount3_11 As Variant  '03/02/2025  ColCount3_12_ removed on 12/02/2025 as per V0.4
Public TAN_TDS1 As Variant
Public CreditDate As String






'Sub Procedure to Validate All Required the Sheet "TDS"

Sub ValidateTDS_All()
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("TDS")
     
     If Not ValidateSheetTDS1 Then
        sourceSheet.Activate
        'MsgBox (Msgbox_I), vbOKOnly, vbMessgaeCaption
        fmsgbox (Msgbox_I)
        CloseMsg
    End If
    
   If Not ValidateSheetTDS2 Then
        sourceSheet.Activate
        'MsgBox (Msgbox_TDS), vbOKOnly, vbMessgaeCaption
        fmsgbox (Msgbox_TDS)
        CloseMsg
    End If
    
    
    If Not ValidateSheetTDS3 Then
        sourceSheet.Activate
        'MsgBox (MsgBox_TDS3), vbOKOnly, vbMessgaeCaption
        fmsgbox (MsgBox_TDS3)
        CloseMsg
    End If
    
    
 If Not ValidateAdvanceTax Then
    sourceSheet.Activate
'    MsgBox (MsgBox_TDS2), vbOKOnly, vbMessgaeCaption
    fmsgbox (MsgBox_TDS2)
    CloseMsg
 End If
 
 'Ankita_04/06/2025
'  If Not ValidateMandatoryShTDS1 Then
'    sourceSheet.Activate
''    MsgBox (MsgBox_TDS2), vbOKOnly, vbMessgaeCaption
'    fmsgbox (Msgbox_I)
'    CloseMsg
' End If

  
End Sub


Function ValidateSheetTDS1() As Boolean
    ValidateSheetTDS1 = True
'    Msgbox_I = "TDS 1" & Chr(10)
    subProcCaption = "Validating TDS1"
    
    
    
    If Not ValidateTAN1_TDS Then ValidateSheetTDS1 = False
    If (Len(Range("TDSal.TAN").item(1).Value) > 0) Then
        If Not ValidateEmp1_TDS() Then ValidateSheetTDS1 = False
        If Not ValidateIncomeCharg1_TDS Then ValidateSheetTDS1 = False
        If Not ValidateTaxDedct1_TDS() Then ValidateSheetTDS1 = False
        If Not ValidateIncChargeSal() Then ValidateSheetTDS1 = False
        If Not ValidateTotTaxDeducted() Then ValidateSheetTDS1 = False
        If Not ValidateMandatoryShTDS1() Then ValidateSheetTDS1 = False  'Ankita_02/06/2025
    End If
    
    setTblinfo_I2
    setTblinfo_I3
    setTblinfo_I4
    
    'Malli-----
'    Dim end_Tan As Variant
'    'end80GGC = WorksheetFunction.Max(0, end80GGC, end80GGC1, end80GGC2, end80GGC4, end80GGC5)
'    end_Tan = WorksheetFunction.Max(0, end_I2, end_I3, end_I4)
'    If end_Tan >= 1 Then
'    If end_I = 0 Then
'       Msgbox_I = Msgbox_I + "* ""Please enter the TAN of Employer"" in Schedule TDS1." & Chr(13)
'    End If
'    End If
    '----------
    'Ankita_02/06/2025
   If ((end_I <> end_I2) Or (end_I <> end_I3) Or (end_I <> end_I4)) Then
'                Msgbox_I = Msgbox_I + "* ""Please enter the TAN of Employer"" in Schedule TDS1." & Chr(13)
'                Msgbox_I = Msgbox_I + "* Enter All mandatory Fields in Schedule TDS1." & Chr(13)
             ValidateSheetTDS1 = False
             Exit Function
    End If
End Function

'Ankita_02/06/2025
Function ValidateMandatoryShTDS1()
ValidateMandatoryShTDS1 = True
Dim i As Long
Dim flag As Boolean
flag = True
For i = 1 To Sheet2.Range("TDSal.TAN").Rows.count
    If Sheet2.Range("TDSal.TAN").item(i).Value <> "" Or Sheet2.Range("TDSal.EmployerOrDeductorOrCollecterName").item(i).Value <> "" Or _
        Sheet2.Range("TDSal.IncChrgSalary").item(i).Value <> "" Or Sheet2.Range("TDSal.TotalTDSSalary").item(i).Value <> "" Then
        If Sheet2.Range("TDSal.TAN").item(i).Value = "" Then
            flag = False
                Msgbox_I = Msgbox_I + "* ""Please enter the TAN of Employer at Sr. No  " & i & "  in Schedule TDS1." & Chr(13)
        End If
        If flag = False Then
            ValidateMandatoryShTDS1 = False
            Exit Function
        End If
    End If
Next i

End Function



Function ValidateTAN1_TDS() As Boolean
    ValidateTAN1_TDS = True
    setTblinfo_I
    noOfProcessSub = end_I
    Dim rangecells As Range
    Set rangecells = Range("TDSal.TAN").Cells
    Dim i As Long
    ReDim TAN_TDS(end_I)
    'Malli-------------
    
    setTblinfo_I2
    setTblinfo_I3
    setTblinfo_I4
    
    'Malli-----
    Dim end_Tan As Variant
    'end80GGC = WorksheetFunction.Max(0, end80GGC, end80GGC1, end80GGC2, end80GGC4, end80GGC5)
    end_Tan = WorksheetFunction.Max(0, end_I2, end_I3, end_I4)
    If end_Tan >= 1 Then
    If end_I = 0 Then
       Msgbox_I = Msgbox_I + "* ""Please enter the TAN of Employer"" in Schedule TDS1." & Chr(13)
'       ValidateTAN1_TDS = False
'                Exit Function
    End If
    End If
    
    '------------------
    For i = 1 To end_I
        TAN_TDS(i) = rangecells.item(i).Value
        If Not Len(TAN_TDS(i)) = 0 Then
            If Not (ValidateTantype_text(Mid(TAN_TDS(i), 1, 4)) And IsNumeric(Mid(TAN_TDS(i), 5, 5)) And ValidateTantype_text(Right(TAN_TDS(i), 1))) Then
                Msgbox_I = Msgbox_I + "* Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
                ValidateTAN1_TDS = False
                Exit Function
          End If
         ' If Not IsNumeric(Mid(TAN_TDS(i), 5, 5)) Then
         '       Msgbox_I = Msgbox_I + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
          '      ValidateTAN1_TDS = False
          '      Exit Function
          ''  End If
          '  If Not ValidateTantype_text(Right(TAN_TDS(i), 1)) Then
          '      Msgbox_I = Msgbox_I + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
          '      ValidateTAN1_TDS = False
          '      Exit Function
           ' End If
        
        
     ElseIf Not chkCompulsory(TAN_TDS(i)) Then
         Msgbox_I = Msgbox_I + "* Please enter the TAN of Employer at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateTAN1_TDS = False
         Exit Function
     End If
     
     If Not ValidateTANCodes(UCase(TAN_TDS(i))) Then
     
        Msgbox_I = Msgbox_I + "* Invalid TAN at Sr.No " & i & "  Schedule TDS-1.Please enter valid TAN" & Chr(13)
        ValidateTAN1_TDS = False
        Exit Function
    End If
    UpdateProgressBar
 Next
End Function

Function ValidateIncChargeSal() As Boolean
Dim incRangeCells As Variant, tottaxRangeCells As Variant, mIntCells As Variant, i As Variant
ValidateIncChargeSal = True
Set incRangeCells = Sheet2.Range("TDSal.IncChrgSalary").Cells
 Set tottaxRangeCells = Sheet2.Range("TDSal.TotalTDSSalary").Cells
 mIntCells = Sheet2.Range("TDSal.IncChrgSalary").count
 
 For i = 1 To mIntCells
    If (incRangeCells.item(i).Value <> "") Then
    
    
    
'    If Len(incRangeCells.Item(i).Value) > 14 Then
'    Msgbox_I = Msgbox_I + "Total tax deducted cannot be more than Income chargeable under salaries cannot exceed 14 Digits at Sr.No " & i
'    ValidateIncChargeSal = False
'    End If
    
    
    
    If (tottaxRangeCells.item(i).Value > incRangeCells.item(i).Value) Then
          Msgbox_I = Msgbox_I + "* Total tax deducted cannot be more than Income chargeable under salaries at Sr.No " & i & " in Schedule TDS1."
            'tottaxRangeCells.Item(i).Value = ""
            'tottaxRangeCells.Item(i).Activate
           ValidateIncChargeSal = False
        End If
    End If
Next

'TAX DETAILS-C5 2024-25 Bindu
'Ayush_19/08
'If Sheet1.Range("IncD.IncomeFromSal").Value < Sheet2.Range("TDS.IncSum").Value Then
'If Sheet1.Range("grosssalandexemptincome_hidden").Value < Sheet2.Range("TDS.IncSum").Value Then
   ' Msgbox_I = Msgbox_I + "*Amount of Gross Salary declared is less than the amount of Gross Salary declared in Schedule TDS 1”"
'   Msgbox_I = Msgbox_I + "* ""Amount of ""Gross Salary"" and ""exempt incomes"" which are part of salary declared is less than the amount of Gross Salary declared in Schedule TDS 1"""
'
'    ValidateIncChargeSal = False
    'Changed by Ankita on 29/10/2025_IPIP-74303
'     MsgBox "* ""Amount of ""Gross Salary"" and ""exempt incomes"" which are part of salary declared is less than the amount of Gross Salary declared in Schedule TDS 1""", , "Alert:"
       
       'Ankita_27/01/2026===
       
'If Sheet1.Range("grossSalary_exempt_familyPension").Value < Sheet2.Range("TDS.IncSum").Value Then
'      Msgbox_I = Msgbox_I + "* ""Sum of ""Gross Salary"" ,""Family Pension"" and ""exempt incomes which are part of salary"" declared is less than the amount of Gross Salary declared in Schedule TDS 1"""
'           ValidateIncChargeSal = False
'End If

End Function

Function ValidateTotTaxDeducted() As Boolean
ValidateTotTaxDeducted = True
Dim val As Variant
val = Sheet2.Range("TDSal.Sum").Value

If Len(val) > 14 Then
    Msgbox_I = Msgbox_I + "* Total tax deducted cannot be more than 14 digit in Schedule TDS1."
    ValidateTotTaxDeducted = False
Exit Function

End If
End Function

Function ValidateEmp1_TDS() As Boolean
    ValidateEmp1_TDS = True
    setTblinfo_I
    Dim rangecells As Range
    Set rangecells = Range("TDSal.EmployerOrDeductorOrCollecterName").Cells
    Dim i As Long
    ReDim TsEmpName(end_I)
    For i = 1 To end_I
        TsEmpName(i) = rangecells.item(i).Value
        
     If Not chkCompulsory(TsEmpName(i)) Then
         Msgbox_I = Msgbox_I + "* Please enter the Name of Employer at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateEmp1_TDS = False
         Exit Function
     End If
     
     If Len(TsEmpName(i)) > 125 Then
         Msgbox_I = Msgbox_I + "* Name of Deductor cannot exceed 125 characters at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateEmp1_TDS = False
         Exit Function
     End If

 Next
End Function


Function ValidateIncomeCharg1_TDS() As Boolean
    ValidateIncomeCharg1_TDS = True
    setTblinfo_I
    Dim rangecells As Range
    Set rangecells = Range("TDSal.IncChrgSalary").Cells
    Dim i As Long
    ReDim Tot_chrg(end_I)
    For i = 1 To end_I
        Tot_chrg(i) = rangecells.item(i).Value
        
     If Not chkCompulsory(Tot_chrg(i)) Then
        ' Msgbox_I = Msgbox_I + "* Income Chargeable is mandatory at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
        'Added by Aavula Naresh
          Msgbox_I = Msgbox_I + "* Please enter the Income chargeable under Salaries at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateIncomeCharg1_TDS = False
         Exit Function
     End If
     
     If Not IsNumeric(Tot_chrg(i)) Then
         Msgbox_I = Msgbox_I + "* Income Chargeable should be Numeric Value at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateIncomeCharg1_TDS = False
         Exit Function
     End If
     
     
     If Tot_chrg(i) < 0 Then
         Msgbox_I = Msgbox_I + "* Income Chargeable cannot be negative at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateIncomeCharg1_TDS = False
         Exit Function
     End If
     
     If (Tot_chrg(i) > 99999999999999#) Then
         Msgbox_I = Msgbox_I + "* Income Chargeable cannot exceed 14 digits at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateIncomeCharg1_TDS = False
         Exit Function
     End If

 Next
End Function


Function ValidateTaxDedct1_TDS() As Boolean
    ValidateTaxDedct1_TDS = True
    setTblinfo_I
    Dim rangecells As Range
    Set rangecells = Range("TDSal.TotalTDSSalary").Cells
    Dim i As Long
    ReDim Amount4_I(end_I)
    For i = 1 To end_I
        Amount4_I(i) = rangecells.item(i).Value
        
     If Not chkCompulsory(Amount4_I(i)) Then
         'Msgbox_I = Msgbox_I + "* Total Tax Deducted is mandatory at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         'Added by Aavula Naresh
         Msgbox_I = Msgbox_I + "* Please enter the Total Tax Deducted at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
  
         ValidateTaxDedct1_TDS = False
         Exit Function
     End If
     
     If Amount4_I(i) < 0 Then
         Msgbox_I = Msgbox_I + "* Total Tax Deducted cannot be negative at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateTaxDedct1_TDS = False
         Exit Function
     End If
     
     If Not IsNumeric(Amount4_I(i)) Then
         Msgbox_I = Msgbox_I + "* Total Tax Deducted should be Numeric value at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateTaxDedct1_TDS = False
         Exit Function
     End If
     
     If Amount4_I(i) > 99999999999999# Then
         Msgbox_I = Msgbox_I + "* Total Tax Deducted cannot exceed 14 digits at Sr. No  " & i & " in Schedule TDS-1." & Chr(13)
         ValidateTaxDedct1_TDS = False
         Exit Function
     End If

 Next
End Function


Sub setTblinfo_I2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSal.TotalTDSSalary").count
    Set rangecells = Range("TDSal.TotalTDSSalary").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
 end_I2 = ccount
End Sub

Sub setTblinfo_I3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSal.IncChrgSalary").count
    Set rangecells = Range("TDSal.IncChrgSalary").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end_I3 = ccount
End Sub

Sub setTblinfo_I()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSal.TAN").count
    Set rangecells = Range("TDSal.TAN").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
 end_I = ccount
 End Sub
 
 Sub setTblinfo_I4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSal.EmployerOrDeductorOrCollecterName").count
    Set rangecells = Range("TDSal.EmployerOrDeductorOrCollecterName").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not Range("TDSal.EmployerOrDeductorOrCollecterName").item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
 end_I4 = ccount
 End Sub
 
 
 '<-------------Checking Validations For the Grid 1 in TDS Ends----------------------------->
 

'---------------Checking Validation for the Grid Two in the TDS Begins---------------------


Function ValidateSheetTDS2() As Boolean
subProcCaption = "Validating TDS2"

ValidateSheetTDS2 = True
'Msgbox_TDS = "TDS2" & Chr(10)
    If Not ValidateTAN2_TDS Then ValidateSheetTDS2 = False
If Len((Range("TDSoth.TAN").item(1).Value) > 0) Then
        If Not ValidateDeductor_TDS Then ValidateSheetTDS2 = False
          If Not ValidateSectionTDSDeducted_TDS2i Then ValidateSheetTDS2 = False 'Jyoti2025-26 19/04/2025
        If Not ValidateYear_TDS Then ValidateSheetTDS2 = False
        If Not ValidateTaxDeducted_TDS2 Then ValidateSheetTDS2 = False
        If Not ValidateTotTaxDeducted_TDS Then ValidateSheetTDS2 = False
        If Not Validateinownhands6_TDS Then ValidateSheetTDS2 = False
        If Not ValidateMandatoryShTDS2 Then ValidateSheetTDS2 = False   'Ankita_03/06/2025
'        If Not ValidateSectionTDS_Tds_TDS2 Then ValidateSheetTDS2 = False  '03/02/2025 Ankita Commented on 12/02/2025 as per V0/4
'        If Not ValidateTDS6_TDS Then ValidateSheetTDS2 = False
'       ' If Not ValidatePAN6other_TDS Then ValidateSheetTDS2 = False
'
'        If Not ValidateTotTaxDeducted7_TDS Then ValidateSheetTDS2 = False
'        If Not Validateinownhands8_TDS Then ValidateSheetTDS2 = False
'        If Not ValidateTDS8_TDS Then ValidateSheetTDS2 = False
'        'If Not ValidatePAN8other_TDS Then ValidateSheetTDS2 = False
'        'If Not ValidateClaimOutofTotAmount_TDS Then ValidateSheetTDS2 = False
'        'If Not ValidateUniqueTDSCert_TDS2 Then ValidateSheetTDS2 = False
'
'        'If Not ValidateAmtClaimedBySpouse Then ValidateSheetTDS2 = False
      If Not ValidateTotAmtClaimed Then ValidateSheetTDS2 = False
'            'If (Sheet1.Range("sheet1.PortugeseCC5A").Value = "Yes") Then
'        If Not ValidatePAN6_TDS Then ValidateSheetTDS2 = False
'        If Not ValidatePAN8_TDS Then ValidateSheetTDS2 = False
'            'End If
End If
     setTableInfo_Grid2
     setTableInfo1_Grid2
     setTableInfo2_Grid2
     setTableInfo3_Grid2
     setTableInfo4_Grid2
'    setTableInfo5_Grid2  '03/02/2025  Commented on 12/02/2025 as per V0/4
     setTableInfo6_Grid2
     setTableInfo2b_Grid2  'Malli 23/04/2025
     
     
'     setTableInfo7_Grid2
'     setTableInfo8_Grid2
'     setTableInfo9_Grid2
'     setTableInfo10_Grid2
'     setTableInfo11_Grid2
'
     'setUTNtable
     
     
    ' If (Sheet1.Range("sheet1.PortugeseCC5A").Value = "Yes") Then
          'setTableInfo5_Grid2
          'setTableInfo9_Grid2
    'End If
    
'    If (Sheet1.Range("sheet1.PortugeseCC5A").Value = "Yes") Then
'        If ((ColCount1 <> ColCount1_1) Or (ColCount1 <> ColCount1_2) Or (ColCount1 <> ColCount1_4) Or (ColCount1 <> ColCount1_6) Or (ColCount1 <> ColCount1_8) Or (ColCount1 <> ColCount1_11) Or (ColCount1 <> ColCount1_9)) Then
'            Msgbox_TDS = Msgbox_TDS + "Enter all mandatory fields for Schedule TDS Other than Salary" & Chr(13)
'            ValidateSheetTDS2 = False
'        End If
             
    'If ((ColCount1 <> ColCount1_1) Or (ColCount1 <> ColCount1_2) Or (ColCount1 <> ColCount1_4) Or (ColCount1 <> ColCount1_6) Or ColCount1 <> ColCount1_3) Then
'Ankita_03/06/2025

'     If ((ColCount1 <> ColCount1_1) Or (ColCount1 <> ColCount1_2) Or (ColCount1 <> ColCount1_4) Or (ColCount1 <> ColCount1_6) Or ColCount1 <> ColCount1_3 Or (ColCount1 <> ColCount2b_2)) Then
'        'Msgbox_TDS = Msgbox_TDS + "* Enter all mandatory fields for in Schedule TDS-2." & Chr(13)
'        'Ankita
'         Msgbox_TDS = Msgbox_TDS + "* Please enter the TAN of Deductor." & Chr(13)
'         ValidateSheetTDS2 = False
'     End If
        
        
'   If (ColCount1_3 <> ColCount1_9 Or ColCount1_3 <> ColCount1_7) Then
'    Msgbox_TDS = Msgbox_TDS + "Enter all mandatory fields in  Col(6) for Schedule TDS Other than Salary" & Chr(13)
'    ValidateSheetTDS2 = False
'    End If
'
'    If (ColCount1_10 <> ColCount1_5 Or ColCount1_10 <> ColCount1_11) Then
'    Msgbox_TDS = Msgbox_TDS + "Enter all mandatory fields in col(8) Schedule TDS Other than Salary" & Chr(13)
'    ValidateSheetTDS2 = False
'    End If
End Function
'Function ValidateAmtClaimedBySpouse() As Boolean
'ValidateAmtClaimedBySpouse = True
'Dim porRangeCells As Variant, amtTaxDeducted As Variant, amtTaxDeducted37BA As Variant, amtClaimRangecells As Variant, amtTaxDeducteamtClaimRangecells As Variant, totDeductRagecells As Variant, mIntCtr As Variant, k As Variant
'Set porRangeCells = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Cells
'Set amtClaimRangecells = Sheet2.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells
'Set totDeductRagecells = Sheet2.Range("TDSoth.TotTDSOnAmtPaid").Cells
'Set amtTaxDeducted = Sheet2.Range("TDSoth.AmountDeducted").Cells
'Set amtTaxDeducted37BA = Sheet2.Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid37BA").Cells
'mIntCtr = Sheet2.Range("TDSoth.TotTDSOnAmtPaid").Count
'
'For k = 1 To mIntCtr
'    If (totDeductRagecells.Item(k).Value <> "") Then
'        If ((amtClaimRangecells.Item(k).Value + porRangeCells.Item(k).Value + amtTaxDeducted37BA.Item(k).Value) > totDeductRagecells.Item(k).Value) Then
'           Msgbox_TDS = Msgbox_TDS + "Amount claimed in Col 6 + Col 7 + Col 8 cannot exceed Total tax deducted (Col 5) at Sr.No " & k
'            ValidateAmtClaimedBySpouse = False
'            'amtClaimRangecells.Item(k).Value = ""
'            'porRangeCells.Item(k).Value = ""
'
'        End If
'
'        'If ((amtTaxDeducted.Item(k).Value) < ((2 * totDeductRagecells.Item(k).Value))) Then
'        '    Msgbox_TDS = Msgbox_TDS + "Amount which is subject to tax deduction could not be less than twice the amount of tax deducted (Col 5) at Sr.No " & k
'        '    ValidateAmtClaimedBySpouse = False
'        'End If
'    End If
'Next
'End Function


'Ankita_03/06/2025
'..............................................................................................................
Function ValidateMandatoryShTDS2()
ValidateMandatoryShTDS2 = True
Dim i As Long
Dim flag As Boolean
flag = True
Dim rangecell1 As Range
Dim rangecell2 As Range
Dim rangecell3 As Range
Dim rangecell4 As Range
Dim rangecell5 As Range
Dim rangecell6 As Range
Dim rangecells As Range

Set rangecell1 = Range("TDSoth.TAN").Cells
Set rangecell2 = Sheet2.Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells
Set rangecell3 = Sheet2.Range("TDsOthr.SectionTDS").Cells
Set rangecell4 = Sheet2.Range("TDSoth.AmountDeducted").Cells
Set rangecell5 = Sheet2.Range("TDSoth.DeductedYear").Cells
Set rangecell6 = Sheet2.Range("TDSoth.TotTDSOnAmtPaid").Cells
Set rangecells = Sheet2.Range("TDSoth.6income").Cells



'For i = 1 To Sheet2.Range("TDSal.TAN").Rows.count
For i = 1 To Sheet2.Range("TDSoth.TAN").Rows.count

'Ankita_28/01/2026=========
'    If rangecell1.item(i).Value <> "" Or rangecell2.item(i).Value <> "" Or rangecell3.item(i).Value <> "" Or rangecell4.item(i).Value <> "" Or _
'        rangecell5.item(i).Value <> "(Select)" Or rangecell6.item(i).Value <> "" Or rangecells.item(i).Value <> "" Then

'IPIP-87788
     If rangecell1.item(i).Value <> "" Or rangecell2.item(i).Value <> "" Or rangecell3.item(i).Value <> "" Or rangecell4.item(i).Value <> "" Or _
        (rangecell5.item(i).Value <> "(Select)" And rangecell5.item(i).Value <> "") Or rangecell6.item(i).Value <> "" Or rangecells.item(i).Value <> "" Then

 
        If rangecell1.item(i).Value = "" Then
            flag = False
                Msgbox_TDS = Msgbox_TDS + "* ""Please enter the TAN of Employer at Sr. No  " & i & "  in Schedule TDS2." & Chr(13)
        ValidateMandatoryShTDS2 = False
        End If
        If flag = False Then
            ValidateMandatoryShTDS2 = False
            Exit Function
        End If
    End If
Next i

End Function




Function ValidateTotAmtClaimed() As Boolean
ValidateTotAmtClaimed = True
Dim val As Variant
val = Sheet2.Range("TDSoth.Sum").Value

If Len(val) > 14 Then
    Msgbox_TDS = Msgbox_TDS + "* Total Amount claimed this year cannot be more than 14 digit in Schedule TDS-2."
    ValidateTotAmtClaimed = False
Exit Function
End If
End Function
'Validation for TAN Number in Grid 2

Function ValidateTAN2_TDS() As Boolean
ValidateTAN2_TDS = True
setTableInfo_Grid2
noOfProcessSub = ColCount1

Dim rangecells As Range
Set rangecells = Range("TDSoth.TAN").Cells
Dim i As Long
ReDim Tan2_TDS(ColCount1)
For i = 1 To ColCount1
    Tan2_TDS(i) = rangecells.item(i).Value
    If Not Len(Tan2_TDS(i)) = 0 Then
    
            If Not ValidateTantype_text(Mid(Tan2_TDS(i), 1, 4)) Then
                Msgbox_TDS = Msgbox_TDS + "* Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet at Sr. No  " & i & "" & Chr(13)
                ValidateTAN2_TDS = False
                Exit Function
            End If
            If Not IsNumeric(Mid(Tan2_TDS(i), 5, 5)) Then
                Msgbox_TDS = Msgbox_TDS + "* Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet at Sr. No  " & i & "" & Chr(13)
                ValidateTAN2_TDS = False
                Exit Function
            End If
            If Not ValidateTantype_text(Right(Tan2_TDS(i), 1)) Then
                Msgbox_TDS = Msgbox_TDS + "* Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet at Sr. No  " & i & "" & Chr(13)
                ValidateTAN2_TDS = False
               Exit Function
            End If
            
             If Not ValidateTANCodes(UCase(Tan2_TDS(i))) Then
              Msgbox_TDS = Msgbox_TDS + "* Invalid TAN. TAN format should be First 4 alphabets, then 5 digits, then alphabet at Sr. No  " & i & "" & Chr(13)
'                Msgbox_TDS = Msgbox_TDS + "* Invalid TAN at Sr.No " & i & "  Schedule TDS-2.Please enter valid TAN" & Chr(13)
                ValidateTAN2_TDS = False
                Exit Function
            End If
    ElseIf Not chkCompulsory(Tan2_TDS(i)) Then
        'Msgbox_TDS = Msgbox_TDS + "* TAN of the Deductor is mandatory at Sr. No  " & i & "  in Schedule TDS-2." & Chr(13)
        Msgbox_TDS = Msgbox_TDS + "* Please enter the TAN of Deductor." & Chr(13)

        ValidateTAN2_TDS = False
        Exit Function
    End If
    UpdateProgressBar
    Next
End Function

'03/02/2025
'Commented on 12/02/2025 by Ankita as V0.4
'Function ValidateSectionTDS_Tds_TDS2() As Boolean
'ValidateSectionTDS_Tds_TDS2 = True
'setTableInfo5_Grid2
'noOfProcessSub = ColCount1
'
'Dim rangecells As Range
'Dim i As Long
'Set rangecells = Range("TDsOthr.SectionTDS").Cells
'ReDim SectionTDS_TDS2_1(ColCount1)
'
'  For i = 1 To ColCount1
'
'    SectionTDS_TDS2_1(i) = rangecells.item(i).Value
'    If isdropdownblank(SectionTDS_TDS2_1(i)) Then
'    Msgbox_TDS = Msgbox_TDS + "* ""Please select the section under which TDS is deducted from dropdown"" at Sr.No " & i & " in Schedule TDS2" & Chr(13)
'
'                ValidateSectionTDS_Tds_TDS2 = False
'
'                Exit Function
'
'    End If
'
'
'
'UpdateProgressBar
'
'Next
'
'End Function
 
 
 
 'Ankita_03/02/2025
'Commented on 12/02/2025 by Ankita as V0.4
'Function ValidateSectionTDS_Tds_TDS3() As Boolean
'ValidateSectionTDS_Tds_TDS3 = True
'setTableInfo12_Grid4
'noOfProcessSub = ColCount3
'
'Dim rangecells As Range
'Dim i As Long
'Set rangecells = Range("TDsOthr.SectionTDS_ii").Cells
'ReDim SectionTDS_TDS2_2(ColCount3)
'
'  For i = 1 To ColCount3
'
'    SectionTDS_TDS2_2(i) = rangecells.item(i).Value
'    If isdropdownblank(SectionTDS_TDS2_2(i)) Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "* ""Please select the section under which TDS is deducted from dropdown"" at Sr.No " & i & " in Schedule TDS3" & Chr(13)
'
'                ValidateSectionTDS_Tds_TDS3 = False
'
'                Exit Function
'
'    End If
'
'
'
'UpdateProgressBar
'
'Next
'
'End Function

'Validation for name of the Deductor

Function ValidateDeductor_TDS() As Boolean
ValidateDeductor_TDS = True
setTableInfo_Grid2
Dim rangecells As Range
Set rangecells = Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells
Dim i As Long
ReDim DeductName2_TDS(ColCount1)
For i = 1 To ColCount1
DeductName2_TDS(i) = rangecells.item(i).Value
If Len(DeductName2_TDS(i)) = 0 Then
End If

If Not chkCompulsory(DeductName2_TDS(i)) Then
    'Msgbox_TDS = Msgbox_TDS + "* Name of the Deductor is mandatory at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
   Msgbox_TDS = Msgbox_TDS + "* Please enter the Name of deductor." & Chr(13)

    ValidateDeductor_TDS = False
    Exit Function
End If
Next
End Function

'Function ValidateUniqueTDSCert_TDS2() As Boolean
'ValidateUniqueTDSCert_TDS2 = True
'Dim rangecells As Range
'Dim i, UTNintCount As Long
'Set rangecells = Range("TDSoth.UTN").Cells
'UTNintCount = Range("TDSoth.UTN").Count'
'
'ReDim UTN_TDS2(UTNintCount)
'For i = 1 To UTNintCount
'If Not rangecells.Item(i).Value = "" Then
'UTN_TDS2(i) = rangecells.Item(i).Value
'  If Not checkfieldspecialcharacter(rangecells.Item(i).Value) Then
'        Msgbox_TDS = Msgbox_TDS + "Unique TDS Certificate Numbers Cannot contain Special Characters at Sr.No " & i & "in Sch TDS2" & Chr(10)
'        ValidateUniqueTDSCert_TDS2 = False
'    Exit Function
'    End If
'      If Len(UTN_TDS2(i)) < 6 Or Len(UTN_TDS2(i)) > 8 Then
'        Msgbox_TDS = Msgbox_TDS + "Unique TDS Certificate Numbers Should 8 digits at Sr.No " & i & "in Sch TDS2" & Chr(10)
'        ValidateUniqueTDSCert_TDS2 = False
'    Exit Function
'    End If
'End If
'Next
'End Function


Function ValidateYear_TDS() As Boolean
ValidateYear_TDS = True
setTableInfo_Grid2
Dim rangecells As Range
Set rangecells = Range("TDSoth.DeductedYear").Cells
Dim i As Long
ReDim Year2_TDS(ColCount1)
For i = 1 To ColCount1
Year2_TDS(i) = rangecells.item(i).Value
If Len(Year2_TDS(i)) = 0 Then
End If

If (Year2_TDS(i) = "" Or UCase(Year2_TDS(i)) = "(SELECT)") Then
   ' Msgbox_TDS = Msgbox_TDS + "* Deduction Year is mandatory at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    Msgbox_TDS = Msgbox_TDS + "* Please select deduction year from dropdown." & Chr(13)

    ValidateYear_TDS = False
    Exit Function
End If

If Not checkfieldspecialcharacter_TDS_TCS(Year2_TDS(i)) Then
     Msgbox_TDS = Msgbox_TDS + "* Deduction Year ,characters < > & ' " & Chr(34) & " are not allowed at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
     ValidateYear_TDS = False
     Exit Function
  End If

Next
End Function


Function ValidateTaxDeducted_TDS2() As Boolean
ValidateTaxDeducted_TDS2 = True
    setTableInfo_Grid2
    Dim rangecells As Range
    Set rangecells = Range("TDSoth.AmountDeducted").Cells
    Dim i As Long
    ReDim TaxDeducted_TDS(ColCount1)
    For i = 1 To ColCount1
        TaxDeducted_TDS(i) = rangecells.item(i).Value
        If Len(TaxDeducted_TDS(i)) = 0 Then
            TaxDeducted_TDS(i) = ""
        End If
        
        If TaxDeducted_TDS(i) < 0 Then
         Msgbox_TDS = Msgbox_TDS + "* Gross receipt which is subject to tax deduction cannot be negative at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
            ValidateTaxDeducted_TDS2 = False
            Exit Function
        End If
        
        If Len(TaxDeducted_TDS(i)) > 14 Then
         Msgbox_TDS = Msgbox_TDS + "* Gross receipt which is subject to tax deduction cannot exceed 14 Digits at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
            ValidateTaxDeducted_TDS2 = False
            Exit Function
        End If
        
        
        If Not chkCompulsory(TaxDeducted_TDS(i)) Then
            'Msgbox_TDS = Msgbox_TDS + "* Gross receipt which is subject to tax deduction is mandatory at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
            Msgbox_TDS = Msgbox_TDS + "* Please enter Gross receipt which is subject to tax deduction." & Chr(13)

            ValidateTaxDeducted_TDS2 = False
            Exit Function
        End If
        
        
        'Ankita_02/06/2025
    If Not IsNumeric(TaxDeducted_TDS(i)) Then
            Msgbox_TDS = Msgbox_TDS & "* Gross Amount at Sr. No  " & i & "  in Schedule TDS-2 should be Numeric value." & Chr(13)
            ValidateTaxDeducted_TDS2 = False
            Exit Function
        End If
        '-------------
    Next
End Function






Function ValidateTotTaxDeducted_TDS() As Boolean
ValidateTotTaxDeducted_TDS = True
setTableInfo_Grid2
Dim rangecells As Range
Set rangecells = Range("TDSoth.TotTDSOnAmtPaid").Cells
Dim i As Long
ReDim TotTaxDedct_TDS(ColCount1)
For i = 1 To ColCount1
TotTaxDedct_TDS(i) = rangecells.item(i).Value
If Len(TotTaxDedct_TDS(i)) = 0 Then
End If

If TotTaxDedct_TDS(i) < 0 Then
    Msgbox_TDS = Msgbox_TDS + "* Tax deducted cannot be negative at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    ValidateTotTaxDeducted_TDS = False
    Exit Function
End If


If Len(TotTaxDedct_TDS(i)) > 14 Then
    Msgbox_TDS = Msgbox_TDS + "* Tax deducted cannot exceed 14 Digits at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    ValidateTotTaxDeducted_TDS = False
    Exit Function
End If


If Not chkCompulsory(TotTaxDedct_TDS(i)) Then
'Change.27.02.2023.102.IDS.16
    'Msgbox_TDS = Msgbox_TDS + "* Please enter Tax deducted is mandatory at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    Msgbox_TDS = Msgbox_TDS + "* Please enter Tax deducted at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    
'end ChangeIDS
    ValidateTotTaxDeducted_TDS = False
    Exit Function
End If
Next
End Function

Function Validateinownhands6_TDS() As Boolean
Validateinownhands6_TDS = True
setTableInfo_Grid2
Dim rangecells As Range
Dim rangecells1 As Range

Set rangecells = Range("TDSoth.6income").Cells
Set rangecells1 = Range("TDSoth.TotTDSOnAmtPaid").Cells
Dim i As Long
ReDim inownhands6_TDS(ColCount1)
ReDim Final_TDS(ColCount1)
For i = 1 To ColCount1
inownhands6_TDS(i) = rangecells.item(i).Value
Final_TDS(i) = rangecells1.item(i).Value
If Len(inownhands6_TDS(i)) = 0 Then
End If

If inownhands6_TDS(i) < 0 Then
    Msgbox_TDS = Msgbox_TDS + "* Please enter TDS Credit out of (5) claimed this Year cannot be negative at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    Validateinownhands6_TDS = False
    Exit Function
End If

If Len(inownhands6_TDS(i)) > 14 Then
    Msgbox_TDS = Msgbox_TDS + "* Please enter TDS Credit out of (5) claimed this Year cannot exceed 14 Digits at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    Validateinownhands6_TDS = False
    Exit Function
End If

If Not chkCompulsory(inownhands6_TDS(i)) Then
'Change.27.02.2023.102.IDS.15
'    Msgbox_TDS = Msgbox_TDS + "* Please enter TDS Credit out of (5) claimed this Year is mandatory at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    Msgbox_TDS = Msgbox_TDS + "* Please enter TDS Credit out of (5) claimed this Year at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    
'End ChangeIDS
    Validateinownhands6_TDS = False
    Exit Function
End If


If Final_TDS(i) < inownhands6_TDS(i) Then
    Msgbox_TDS = Msgbox_TDS + "* TDS credit claimed cannot be more than Tax deducted at Sr.NO " & i & " in Schedule TDS-2." & Chr(13)
    Validateinownhands6_TDS = False
    Exit Function
End If
Next
End Function


'Function ValidateTDS6_TDS() As Boolean
'ValidateTDS6_TDS = True
'setTableInfo3_Grid2
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.6TDS").Cells
'Dim i As Long
'ReDim TDS6_TDS(ColCount1_3)
'For i = 1 To ColCount1_3
'TDS6_TDS(i) = rangecells.Item(i).Value
'If Len(TDS6_TDS(i)) = 0 Then
'End If
'
'
'If TDS6_TDS(i) < 0 Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter TDS amount col(6) in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (19 TDS2) cannot be negative" & Chr(13)
'    ValidateTDS6_TDS = False
'    Exit Function
'End If
'
'If Len(TDS6_TDS(i)) > 14 Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter TDS amount col(6) in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (19 TDS2) cannot exceed 14 Digits" & Chr(13)
'    ValidateTDS6_TDS = False
'    Exit Function
'End If
'
'
'
'If Not chkCompulsory(TDS6_TDS(i)) Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter TDS amount col(6) in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'    ValidateTDS6_TDS = False
'    Exit Function
'End If
'Next
'End Function
'
'Function ValidatePAN6_TDS() As Boolean
'ValidatePAN6_TDS = True
'setTableInfo3_Grid2
'noOfProcessSub = ColCount1_3
'
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.6panspouse").Cells
'Dim i As Long
'ReDim PAN6_TDS(ColCount1_3)
'
'For i = 1 To ColCount1_3
'    PAN6_TDS(i) = rangecells.Item(i).Value
'    If Not Len(PAN6_TDS(i)) = 0 Then
'        If Not CheckIFSCTDS(Mid(PAN6_TDS(i), 1, 10)) Then
'        Msgbox_TDS = Msgbox_TDS + "PAN of the spouse or other person at Sr. No  " & i & "  in Sheet TDS2  is invalid. First 5 alphabets, next 4 digits, then alphabet " & Chr(13)
'        ValidatePAN6_TDS = False
'        End If
'
'
'    ElseIf Not chkCompulsory(PAN6_TDS(i)) Then
'        Msgbox_TDS = Msgbox_TDS + " PAN of the spouse or other person at Sr.No " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'        ValidatePAN6_TDS = False
'        Exit Function
'    End If
'    UpdateProgressBar
'    Next
'End Function
'
'
'
'
''Function ValidatePAN6other_TDS() As Boolean
''ValidatePAN6other_TDS = True
''setTableInfo_Grid2
''noOfProcessSub = ColCount1
''
''Dim rangecells As Range
''Set rangecells = Range("TDSoth.6panperson").Cells
''Dim i As Long
''ReDim PAN6other_TDS(ColCount1)
''
''For i = 1 To ColCount1
''    PAN6other_TDS(i) = rangecells.Item(i).Value
''    If Not Len(PAN6other_TDS(i)) = 0 Then
''        If Not mIncmDtls.CheckPAN(Mid(PAN6other_TDS(i), 1, 10)) Then
''        Msgbox_TDS = Msgbox_TDS + "PAN of the other person at Sr. No  " & i & "  in Sheet TDS2  is invalid. First 5 alphabets, next 4 digits, then alphabet (4th alphabet must be ""P"")  " & Chr(13)
''        ValidatePAN6other_TDS = False
''        End If
''
''   End If
''
''    UpdateProgressBar
''Next
''End Function
'
'' new
'
'
'Function ValidateTotTaxDeducted7_TDS() As Boolean
'ValidateTotTaxDeducted7_TDS = True
'setTableInfo_Grid2
'Dim rangecells As Range
'Dim mintr As Variant
'Set rangecells = Range("TDSoth.TotTDSOnAmtPaid7").Cells
'mintr = Range("TDSoth.TotTDSOnAmtPaid7").Cells.Count
'Dim i As Long
'ReDim TotTaxDedct7_TDS(ColCount1)
'For i = 1 To ColCount1
'TotTaxDedct7_TDS(i) = rangecells.Item(i).Value
'If Len(TotTaxDedct7_TDS(i)) = 0 Then
'End If
'
'If Not chkCompulsory(TotTaxDedct7_TDS(i)) Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter TDS credit Amount out of (5) or (6) being claimed this Year in own hands at Sr.NO " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'    ValidateTotTaxDeducted7_TDS = False
'    Exit Function
'End If
'
'If TotTaxDedct7_TDS(i) < 0 Then
'    Msgbox_TDS = Msgbox_TDS + " TDS credit Amount out of (5) or (6) being claimed this Year in own hands at Sr.NO " & i & " cannot be negative in Sheet TDS, (19 TDS2) " & Chr(13)
'    ValidateTotTaxDeducted7_TDS = False
'
'End If
'
'If Len(TotTaxDedct7_TDS(i)) > 14 Then
'    Msgbox_TDS = Msgbox_TDS + " TDS credit Amount out of (5) or (6) being claimed this Year in own hands at Sr.NO " & i & " cannot exceed 14 Digits in Sheet TDS, (19 TDS2) " & Chr(13)
'    ValidateTotTaxDeducted7_TDS = False
'
'End If
'
'If (Range("TDSoth.TotTDSOnAmtPaid7").Cells.Item(i).Value + Range("TDSoth.8TDS").Cells.Item(i).Value > (Range("TDSoth.6TDS").Cells.Item(i).Value + Range("TDSoth.totTDSOnAmtPaid").Cells.Item(i).Value)) Then
'Msgbox_TDS = Msgbox_TDS + "Amount in field Col(7) and TDS at col(8) cannot be more than sum of field Col(5) and TDS at col(6) in sch TDS2 " & Chr(13)
'ValidateTotTaxDeducted7_TDS = False
'Exit Function
'End If
'Next
'End Function
'
'Function Validateinownhands8_TDS() As Boolean
'Validateinownhands8_TDS = True
'setTableInfo10_Grid2
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid37BA").Cells
'Dim i As Long
'ReDim inownhands8_TDS(ColCount1_10)
'For i = 1 To ColCount1_10
'inownhands8_TDS(i) = rangecells.Item(i).Value
'If Len(inownhands8_TDS(i)) = 0 Then
'End If
'
'
'
'
'If inownhands8_TDS(i) < 0 Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter Income in col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (19 TDS2) cannot be negative" & Chr(13)
'    Validateinownhands8_TDS = False
'    Exit Function
'End If
'
'
'If Len(inownhands8_TDS(i)) > 14 Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter Income in col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (19 TDS2) cannot exceed 14 Digits" & Chr(13)
'    Validateinownhands8_TDS = False
'    Exit Function
'End If
'
'If Not chkCompulsory(inownhands8_TDS(i)) Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter Income in col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'    Validateinownhands8_TDS = False
'    Exit Function
'End If
'Next
'End Function
'
'
'Function ValidateTDS8_TDS() As Boolean
'ValidateTDS8_TDS = True
'setTableInfo10_Grid2
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.8TDS").Cells
'Dim i As Long
'ReDim TDS8_TDS(ColCount1_10)
'For i = 1 To ColCount1_10
'TDS8_TDS(i) = rangecells.Item(i).Value
'If Len(TDS8_TDS(i)) = 0 Then
'End If
'
'
'If TDS8_TDS(i) < 0 Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter TDS amount col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (19 TDS2) cannot be negative" & Chr(13)
'    ValidateTDS8_TDS = False
'    Exit Function
'End If
'
'If Len(TDS8_TDS(i)) > 14 Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter TDS amount col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (19 TDS2) cannot exceed 14 Digits" & Chr(13)
'    ValidateTDS8_TDS = False
'    Exit Function
'End If
'
'If Not chkCompulsory(TDS8_TDS(i)) Then
'    Msgbox_TDS = Msgbox_TDS + "Please enter TDS amount col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'    ValidateTDS8_TDS = False
'    Exit Function
'End If
'Next
'End Function
'
'Function ValidatePAN8_TDS() As Boolean
'ValidatePAN8_TDS = True
'setTableInfo10_Grid2
'noOfProcessSub = ColCount1_10
'
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.AmtClaimedBySpouse").Cells
'Dim i As Long
'ReDim PAN8_TDS(ColCount1_10)
'
'For i = 1 To ColCount1_10
'    PAN8_TDS(i) = rangecells.Item(i).Value
'    If Not Len(PAN8_TDS(i)) = 0 Then
'        If Not CheckIFSCTDS(Mid(PAN8_TDS(i), 1, 10)) Then
'        Msgbox_TDS = Msgbox_TDS + "PAN of the spouse or other person at Sr. No  " & i & "  in Sheet TDS2  is invalid. First 5 alphabets, next 4 digits, then alphabet " & Chr(13)
'        ValidatePAN8_TDS = False
'        End If
'
'
'        ElseIf Not chkCompulsory(PAN8_TDS(i)) Then
'        Msgbox_TDS = Msgbox_TDS + " PAN of the spouse or other person  at Sr.No " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'        ValidatePAN8_TDS = False
'        Exit Function
'    End If
'    UpdateProgressBar
'    Next
'End Function
'



'Function ValidatePAN8other_TDS() As Boolean
'ValidatePAN8other_TDS = True
'setTableInfo_Grid2
'noOfProcessSub = ColCount1
'
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.8panperson").Cells
'Dim i As Long
'ReDim PAN8other_TDS(ColCount1)
'
'For i = 1 To ColCount1
'    PAN8other_TDS(i) = rangecells.Item(i).Value
'    If Not Len(PAN8other_TDS(i)) = 0 Then
'        If Not mIncmDtls.CheckPAN(Mid(PAN8other_TDS(i), 1, 10)) Then
'        Msgbox_TDS = Msgbox_TDS + "PAN of the other person at Sr. No  " & i & "  in Sheet TDS2  is invalid. First 5 alphabets, next 4 digits, then alphabet (4th alphabet must be ""P"") " & Chr(13)
'        ValidatePAN8other_TDS = False
'        End If
'    End If
'
'
'    UpdateProgressBar
'Next
'End Function








'
'Function ValidateClaimOutofTotAmount_TDS() As Boolean
'ValidateClaimOutofTotAmount_TDS = True
'setTableInfo_Grid2
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells
'Dim i As Long
'ReDim ClaimAmt2_TDS(ColCount1)
'For i = 1 To ColCount1
'ClaimAmt2_TDS(i) = rangecells.Item(i).Value
'If Len(ClaimAmt2_TDS(i)) = 0 Then
'End If
'
'If Not chkCompulsory(ClaimAmt2_TDS(i)) Then
'    Msgbox_TDS = Msgbox_TDS + "Amount Claimed at Sr.NO " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'    ValidateClaimOutofTotAmount_TDS = False
'    Exit Function
'End If
'Next
'End Function


'Function ValidatePortugeseCC5A_TDS() As Boolean
'ValidatePortugeseCC5A_TDS = True
'setTableInfo_Grid2
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.AmtClaimedBySpouse").Cells
'Dim i As Long
'ReDim Portuguese2_TDS(ColCount1)
'For i = 1 To ColCount1
'Portuguese2_TDS(i) = rangecells.Item(i).Value
'If Len(Portuguese2_TDS(i)) = 0 Then
'End If
'If Not chkCompulsory(Portuguese2_TDS(i)) Then
'    Msgbox_TDS = Msgbox_TDS + "Amount claimed in the hands of Spouse at Sr.NO " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'    ValidatePortugeseCC5A_TDS = False
'    Exit Function
'End If
'Next
'End Function



'Common Table count finder

Sub setTableInfo_Grid2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSoth.TAN").count
    Set rangecells = Range("TDSoth.TAN").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount1 = ccount
End Sub

Sub setTableInfo1_Grid2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSoth.DeductedYear").count
    Set rangecells = Range("TDSoth.DeductedYear").Cells
    For mIntCtr = 1 To mIntCells
        If Not (rangecells.item(mIntCtr).Value = "" Or UCase(rangecells.item(mIntCtr).Value) = "(SELECT)") Then
        ccount = ccount + 1
        End If
    Next
    ColCount1_1 = ccount
End Sub

Sub setTableInfo2_Grid2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSoth.TotTDSOnAmtPaid").count
    Set rangecells = Range("TDSoth.TotTDSOnAmtPaid").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount1_2 = ccount
End Sub

Sub setTableInfo3_Grid2()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSoth.6income").count
    Set rangecells = Range("TDSoth.6income").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount1_3 = ccount
End Sub

Sub setTableInfo4_Grid2()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSoth.EmployerOrDeductorOrCollecterName").count
    Set rangecells = Range("TDSoth.EmployerOrDeductorOrCollecterName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount1_4 = ccount
End Sub

'03/02/2025
'Commented on 12/02/2025 by Ankita as V0.4

'Sub setTableInfo5_Grid2()
'Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDsOthr.SectionTDS").count
'    Set rangecells = Range("TDsOthr.SectionTDS").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    ColCount1_12 = ccount
'End Sub


'Sub setTableInfo5_Grid2()
'Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDSoth.AmtClaimedBySpouse").Count
'    Set rangecells = Range("TDSoth.AmtClaimedBySpouse").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.Item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    ColCount1_5 = ccount
'End Sub

Sub setTableInfo6_Grid2()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSoth.AmountDeducted").count
    Set rangecells = Range("TDSoth.AmountDeducted").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount1_6 = ccount
End Sub


'Sub setTableInfo7_Grid2()
'Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDSoth.6TDS").Count
'    Set rangecells = Range("TDSoth.6TDS").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.Item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    ColCount1_7 = ccount
'End Sub
'
'Sub setTableInfo8_Grid2()
'Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDSoth.TotTDSOnAmtPaid7").Count
'    Set rangecells = Range("TDSoth.TotTDSOnAmtPaid7").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.Item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    ColCount1_8 = ccount
'End Sub
'
'Sub setTableInfo9_Grid2()
'Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDSoth.6panspouse").Count
'    Set rangecells = Range("TDSoth.6panspouse").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.Item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    ColCount1_9 = ccount
'End Sub
'
'
'
'
'
'
'
'
'Sub setTableInfo10_Grid2()
'Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid37BA").Count
'    Set rangecells = Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid37BA").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.Item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    ColCount1_10 = ccount
'End Sub
'
'Sub setTableInfo11_Grid2()
'Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDSoth.8TDS").Count
'    Set rangecells = Range("TDSoth.8TDS").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.Item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    ColCount1_11 = ccount
'End Sub

Sub setUTNtable()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDSoth.UTN").count
    Set rangecells = Range("TDSoth.UTN").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    If ccount <= ColCount1 Then
    UTNCount = ColCount1
    Else:
    UTNCount = ccount
    End If
End Sub

'<--------Completion of Grid 2 Validation-------------->

Function ValidateAdvanceTax() As Boolean
subProcCaption = "Validating IT"
ValidateAdvanceTax = True
'MsgBox_TDS2 = "21 IT" & Chr(10)
If Not ValidateBSR_TDS Then ValidateAdvanceTax = False
    If Len((Range("TaxP.BSRCode").item(1).Value) > 0) Then
        If Not ValidateDateCreditToGovt Then ValidateAdvanceTax = False
        If Not ValidateSerialNum Then ValidateAdvanceTax = False
        If Not ValidateTaxPaid Then ValidateAdvanceTax = False
        If Not ValidateTottaxPaid Then ValidateAdvanceTax = False
        If Not ValidateMandatoryShIT Then ValidateAdvanceTax = False  'Ankita_02/06/2025

    End If
    setTableInfo1_Grid3
    setTableInfo2_Grid3
    setTableInfo3_Grid3
    
'    If ((ColCount2 <> ColCount2_1) Or (ColCount2 <> ColCount2_2) Or (ColCount2 <> ColCount2_3)) Then
''    MsgBox_TDS2 = MsgBox_TDS2 + "* Enter All Mandatory Details for Schedule Advance TAX / Self Assessment TAX."
'    'Ankita
'     MsgBox_TDS2 = MsgBox_TDS2 + "*Please enter the BSR code"
'    ValidateAdvanceTax = False
'    End If
End Function


'Ankita_02/06/2025
Function ValidateMandatoryShIT() As Boolean
ValidateMandatoryShIT = True
Dim i As Long
Dim flag As Boolean
flag = True
For i = 1 To Sheet2.Range("TaxP.BSRCode").Rows.count
    If Sheet2.Range("TaxP.BSRCode").item(i).Value <> "" Or Sheet2.Range("TaxP.DateDep").item(i).Value <> "" Or _
        Sheet2.Range("TaxP.SrlNoOfChaln").item(i).Value <> "" Or Sheet2.Range("TaxP.Amt").item(i).Value <> "" Then
        If Sheet2.Range("TaxP.BSRCode").item(i).Value = "" Then
            flag = False
'            MsgSchIT = MsgSchIT + "* Please enter the BSR code in Schedule IT" & Chr(13)
             'SIT-56341
                MsgBox_TDS2 = MsgBox_TDS2 + "* Please enter the BSR code " & Chr(13)
        End If
        If flag = False Then
            ValidateMandatoryShIT = False
            Exit Function
        End If
    End If
Next i

End Function

Function ValidateBSR_TDS() As Boolean
ValidateBSR_TDS = True
setTableInfo_Grid3
noOfProcessSub = ColCount2
Dim rangecells As Range
Set rangecells = Range("TaxP.BSRCode").Cells
Dim i As Long
ReDim BSR_TDS(ColCount2)
For i = 1 To ColCount2
BSR_TDS(i) = rangecells.item(i).Value



If Not chkCompulsory(BSR_TDS(i)) Then
    MsgBox_TDS2 = MsgBox_TDS2 + "* BSR Code at Sr.NO " & i & " in Schedule IT is mandatory" & Chr(13)
    ValidateBSR_TDS = False
    Exit Function
End If

If Not Len(BSR_TDS(i)) = 0 Then
            If Not (chkNumeric(Mid(BSR_TDS(i), 1, 3)) And checkfieldspecialcharacter(Mid(BSR_TDS(i), 4, 4))) Then
                MsgBox_TDS2 = MsgBox_TDS2 + "* Invalid BSR code in schedule TDS. BSR format should be First 3 numeric and next 4 alphanumeric, at Sr. No  " & i & " in Schedule IT.." & Chr(13)
                ValidateBSR_TDS = False
                Exit Function
                End If
          End If

'---------
'Jyoti2025-26 19/04/2025

'   If Not BSR_TDS(i) = "0000000" Then
    If BSR_TDS(i) = "0000000" Then
   MsgBox_TDS2 = MsgBox_TDS2 + "*Please enter a valid 7 characters BSR Code" & Chr(13)
    ValidateBSR_TDS = False
    Exit Function
    End If
'------

'If Not chkNumeric(BSR_TDS(i)) Then
'    MsgBox_TDS2 = MsgBox_TDS2 + "* Please enter a valid 7 digit BSR Code at Serial No " & i & " & Chr(10)"
'    ValidateBSR_TDS = False
'    Exit Function
'End If
'
'If Not CLng(BSR_TDS(i)) >= 0 Then
'    MsgBox_TDS2 = MsgBox_TDS2 + "* BSR Code at Sr.NO " & i & " in Sheet TDS, (21 IT) Cannot be all 0" & Chr(13)
'    ValidateBSR_TDS = False
'    Exit Function
'End If

'If Not checkfieldspecialcharacter(BSR_TDS(i)) Then
'        MsgBox_TDS2 = MsgBox_TDS2 + "* BSR Code cannot contain Special Characters at Sr.No " & i & "in Sch Schedule IT" & Chr(10)
'        ValidateBSR_TDS = False
'    Exit Function
'End If

'If Not IsNumeric(BSR_TDS(i)) Then
'    MsgBox_TDS2 = MsgBox_TDS2 + "* Please enter a valid 7 digit BSR Code at Serial No " & i & " & Chr(10)"
'    ValidateBSR_TDS = False
'    Exit Function
'End If
UpdateProgressBar
Next
End Function


Function ValidateDateCreditToGovt() As Boolean
ValidateDateCreditToGovt = True
setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Range("TaxP.DateDep").Cells
Dim i As Long
ReDim DateCredit_TDS(ColCount2)
For i = 1 To ColCount2
DateCredit_TDS(i) = rangecells.item(i).Value
If Len(DateCredit_TDS(i)) = 0 Then
End If

If Not chkCompulsory(DateCredit_TDS(i)) Then
    MsgBox_TDS2 = MsgBox_TDS2 + "* Please enter the date of Deposit at Sr.NO " & i & " in Schedule IT is mandatory" & Chr(13)
    ValidateDateCreditToGovt = False
    Exit Function
End If



If Not CheckDateBefore1(DateCredit_TDS(i)) Then
    MsgBox_TDS2 = MsgBox_TDS2 + "* Date of Credit into Govt Account at Sr.NO " & i & " in Schedule IT should be after 31/03/2017 at Sr.No " & i & Chr(13)
    ValidateDateCreditToGovt = False
    Exit Function
End If



Next
End Function


Function ValidateSerialNum() As Boolean
ValidateSerialNum = True
setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Range("TaxP.SrlNoOfChaln").Cells
Dim i As Long
ReDim SerialNum_TDS(ColCount2)
For i = 1 To ColCount2
SerialNum_TDS(i) = rangecells.item(i).Value
If Len(SerialNum_TDS(i)) = 0 Then
End If

If Not chkCompulsory(SerialNum_TDS(i)) Then
    'MsgBox_TDS2 = MsgBox_TDS2 + "* Serial Number of Challan at Sr.NO " & i & " is mandatory in Schedule IT." & Chr(13)
    MsgBox_TDS2 = MsgBox_TDS2 + "* Please enter the serial number of Challan." & Chr(13)

    ValidateSerialNum = False
    Exit Function
End If


If Not chkNumeric(SerialNum_TDS(i)) Then
    MsgBox_TDS2 = MsgBox_TDS2 + "* Serial Number of Challan at Sr.NO " & i & " in Schedule IT can allow only Numeric Values" & Chr(13)
    ValidateSerialNum = False
    Exit Function
End If
Next
End Function

Function ValidateTaxPaid() As Boolean
ValidateTaxPaid = True
setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Range("TaxP.Amt").Cells
Dim i As Long
ReDim TaxPaid3_TDS(ColCount2)
For i = 1 To ColCount2
TaxPaid3_TDS(i) = rangecells.item(i).Value
If Len(TaxPaid3_TDS(i)) = 0 Then
End If

If Not chkCompulsory(TaxPaid3_TDS(i)) Then
   ' MsgBox_TDS2 = MsgBox_TDS2 + "* Tax Paid Amount is mandatory at Sr.NO " & i & " in Schedule IT" & Chr(13)
    MsgBox_TDS2 = MsgBox_TDS2 + "*Please enter the amount" & Chr(13)

    ValidateTaxPaid = False
    Exit Function
End If


If TaxPaid3_TDS(i) < 0 Then
    MsgBox_TDS2 = MsgBox_TDS2 + "* Tax Paid Amount cannot be negative at Sr.NO " & i & " in Schedule IT" & Chr(13)
    ValidateTaxPaid = False
End If

If Len(TaxPaid3_TDS(i)) > 14 Then
    MsgBox_TDS2 = MsgBox_TDS2 + "* Tax Paid Amount cannot cannot exceed 14 Digits at Sr.NO " & i & " in Schedule IT" & Chr(13)
    ValidateTaxPaid = False
End If

Next
End Function

Function ValidateTottaxPaid() As Boolean
ValidateTottaxPaid = True
Dim val As Variant
val = Sheet2.Range("TaxP.Sum").Value

If Len(val) > 14 Then
    MsgBox_TDS2 = MsgBox_TDS2 + "* Total Taxes paid this year cannot be more than 14 digit in Schedule IT"
    ValidateTottaxPaid = False
Exit Function
End If
End Function

Sub setTableInfo_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TaxP.BSRCode").count
    Set rangecells = Range("TaxP.BSRCode").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2 = ccount
End Sub

Sub setTableInfo1_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TaxP.DateDep").count
    Set rangecells = Range("TaxP.DateDep").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_1 = ccount
End Sub


Sub setTableInfo2_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TaxP.SrlNoOfChaln").count
    Set rangecells = Range("TaxP.SrlNoOfChaln").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_2 = ccount
End Sub


Sub setTableInfo3_Grid3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TaxP.Amt").count
    Set rangecells = Range("TaxP.Amt").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_3 = ccount
End Sub


Function ValidateTantype_text(strName As Variant) As Boolean
    ValidateTantype_text = True
    Dim len1, j As Long
    Dim s1 As String
    len1 = Len(strName)
    For j = 1 To len1
        s1 = Mid(strName, j, 1)
        If (((asc(s1) >= 65) And (asc(s1) <= 90)) Or (asc(s1) = 45)) Then
        Else
            ValidateTantype_text = False
        End If
    Next
End Function



 
Function CheckDateBefore1(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.


    CheckDateBefore1 = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBefore1 = False
           ' MsgBox ("Date of Credit into Govt Account in Sheet : TDS  must be a valid dd/mm/yyyy format")
            Exit Function
        End If
'        Change28.11.2022.102.06A
        'TAX DETAILS-E5
        'If EfilingCommon.checkFirstDateBefore(dob, "31/03/2022") Then
        
        'TAX DETAILS-C5
'        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2023") Then
        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2024") Then 'Changed by Ankita on 29/11/2024
        
'        ---End
           'MsgBox ("Date of of Credit into Govt Account in Sheet : TDS should be on or after 01/04/2017")
            CheckDateBefore1 = False
            Exit Function
            Else
            CreditDate = dob
       End If
    End If
 End Function
 
 
Sub HelpTDS_Click()
sPassword = EfilingCommon.getmsgstate
ActiveWorkbook.Unprotect Password:=sPassword
Sheet6.Activate
Sheet6.Visible = xlSheetVisible
ActiveWorkbook.Protect Password:=sPassword
End Sub

 Sub NextTDS_Click()
Sheet11.Activate
End Sub

Sub PreviousTDS_Click()
'Ankita_11/11/2024
'Ankita_14/05/2025
If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
Sheet201.Activate
'    ElseIf Sheet16.Visible = True Then
'        Sheet16.Activate
    ElseIf Sheet18.Visible = True Then
        Sheet18.Activate
    Else
       Sheet19.Activate

'Else:
''---------------------
'Sheet1.Activate
End If  'Ankita 11/11/2024
End Sub

Sub ValidateSchTDS_Click()
Dim vbMessgaeCaption As String
vbMessgaeCaption = "ITR 1: AY: 2026-27"                    'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
ValidateTDS_All
'MsgBox ("Sheet TDS is OK "), vbOKOnly, vbMessgaeCaption
fmsgboxoK ("Sheet TDS is OK ")
End Sub
Sub CommandButton4_Click()

printWorkSheet
End Sub



Sub AddRowSchTDS1_Click()
Dim vRows As Long
Dim sourceSheet As Worksheet

Set sourceSheet = ThisWorkbook.Sheets("TDS")
    sourceSheet.Activate
EfilingCommon.DefinedgridNameRange = "TDSal.TAN||TDSal.EmployerOrDeductorOrCollecterName||TDSal.IncChrgSalary||TDSal.TotalTDSSalary"
ActiveCellRange = EfilingCommon.searchLastRow("TDSal.TAN")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

 Sub AddRowsSchTDS2_Click()
Dim vRows As Long
Dim sourceSheet As Worksheet

Set sourceSheet = ThisWorkbook.Sheets("TDS")
sourceSheet.Activate
'EfilingCommon.DefinedgridNameRange = "TDSoth.TAN||TDSoth.EmployerOrDeductorOrCollecterName||TDSoth.AmountDeducted||TDSoth.DeductedYear||TDSoth.TotTDSOnAmtPaid||TDSoth.6income||TDsOthr.SectionTDS" 'Ankita_03/02/2025
'EfilingCommon.DefinedgridNameRange = "TDSoth.TAN||TDSoth.EmployerOrDeductorOrCollecterName||TDSoth.AmountDeducted||TDSoth.DeductedYear||TDSoth.TotTDSOnAmtPaid||TDSoth.6income" 'Removed by Ankita on 12/02/2025 as per V0.4
EfilingCommon.DefinedgridNameRange = "TDSoth.TAN||TDSoth.EmployerOrDeductorOrCollecterName||TDsOthr.SectionTDS||TDSoth.AmountDeducted||TDSoth.DeductedYear||TDSoth.TotTDSOnAmtPaid||TDSoth.6income"  'Jyoti2025-26
ActiveCellRange = EfilingCommon.searchLastRow("TDSoth.TAN")
vRows = EfilingCommon.insertRowUnderSectionWithFormula

End Sub


 Sub AddRowsSchTDS4_Click()
Dim vRows As Long
Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("TDS")
    sourceSheet.Activate
'EfilingCommon.DefinedgridNameRange = "TDS26QB.PAN||TDS26QB.Aadhar_Number||TDS26QB.EmployerOrDeductorName||TDS26QB.AmountDeducted||TDS26QB.DeductedYear||TDS26QB.TotTDSOnAmtPaid||TDS26QB.6income||TDsOthr.SectionTDS_ii"  'Ankita_03/02/2025
'EfilingCommon.DefinedgridNameRange = "TDS26QB.PAN||TDS26QB.Aadhar_Number||TDS26QB.EmployerOrDeductorName||TDS26QB.AmountDeducted||TDS26QB.DeductedYear||TDS26QB.TotTDSOnAmtPaid||TDS26QB.6income" 'Removed by Ankita on 12/02/2025 as per V0.4
EfilingCommon.DefinedgridNameRange = "TDS26QB.PAN||TDS26QB.Aadhar_Number||TDS26QB.EmployerOrDeductorName||TDsOthr2.SectionTDSDeducted||TDS26QB.AmountDeducted||TDS26QB.DeductedYear||TDS26QB.TotTDSOnAmtPaid||TDS26QB.6income" ''Jyoti2025-26

ActiveCellRange = EfilingCommon.searchLastRow("TDS26QB.PAN")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub





Sub AddRowsSchTDS3_Click()
Dim vRows As Long
Dim sourceSheet As Worksheet
Set sourceSheet = ThisWorkbook.Sheets("TDS")
sourceSheet.Activate
'EfilingCommon.DefinedgridNameRange = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.Amt||IT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfSATNew"  'Ankita_16/07/2025  'IT_DueDate
EfilingCommon.DefinedgridNameRange = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.Amt||`||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfSATNew||IT_DueDate" 'Ayush_DueDate_08/09/2025
ActiveCellRange = EfilingCommon.searchLastRow("TaxP.BSRCode")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub






'NEW TDS2 code


Function ValidateSheetTDS3() As Boolean
subProcCaption = "Validating TDS3"
ValidateSheetTDS3 = True
'Msgbox_TDS = "TDS3" & Chr(10)
    If Not ValidatePAN_TDS3 Then ValidateSheetTDS3 = False
    If Len((Range("TDS26QB.PAN").item(1).Value) > 0) Then
        If Not ValidateAadhar_TDS3 Then ValidateSheetTDS3 = False
        If Not ValidateDeductor_TDS3 Then ValidateSheetTDS3 = False
          If Not ValidateSectionTDSDeducted_TDS2ii Then ValidateSheetTDS3 = False 'Jyoti2025-26 19/04/2025
        If Not ValidateYear_TDS3 Then ValidateSheetTDS3 = False
        If Not ValidateTaxDeducted_TDS3 Then ValidateSheetTDS3 = False
'        If Not ValidateSectionTDS_Tds_TDS3 Then ValidateSheetTDS3 = False   'Ankita_03/02/2024 Commented on 12/02/2025 as per V0.4
        'If Not ValidateTotTaxDeducted_TDS3 Then ValidateSheetTDS3 = False
        'If Not ValidateClaimOutofTotAmount_TDS3 Then ValidateSheetTDS3 = False
        'If Not ValidateUniqueTDSCert_TDS2 Then ValidateSheetTDS2 = False
        
'        If Not ValidateAmtClaimedBySpouse3 Then ValidateSheetTDS3 = False
        If Not ValidateTotAmtClaimed3 Then ValidateSheetTDS3 = False
'        If (Sheet1.Range("sheet1.PortugeseCC5A").Value = "Yes") Then
'            If Not ValidatePortugeseCC5A_TDS3 Then ValidateSheetTDS3 = False
'        End If
'     End If
     
'
        If Not ValidateTotTaxDeducted_TDS3 Then ValidateSheetTDS3 = False
        If Not Validateinownhands6_TDS3 Then ValidateSheetTDS3 = False
'        If Not ValidateTDS6_TDS3 Then ValidateSheetTDS3 = False
'        'If Not ValidatePAN6other_TDS3 Then ValidateSheetTDS3 = False
'
'        If Not ValidateTotTaxDeducted7_TDS3 Then ValidateSheetTDS3 = False
'        If Not Validateinownhands8_TDS3 Then ValidateSheetTDS3 = False
'        If Not ValidateTDS8_TDS3 Then ValidateSheetTDS3 = False
'       ' If Not ValidatePAN8other_TDS3 Then ValidateSheetTDS3 = False
'        'If Not ValidateClaimOutofTotAmount_TDS Then ValidateSheetTDS2 = False
'        'If Not ValidateUniqueTDSCert_TDS2 Then ValidateSheetTDS2 = False
'
'        'If Not ValidateAmtClaimedBySpouse Then ValidateSheetTDS2 = False
'        If Not ValidateTotAmtClaimed3 Then ValidateSheetTDS3 = False
'        'If (Sheet1.Range("sheet1.PortugeseCC5A").Value = "Yes") Then
'        If Not ValidatePAN6_TDS3 Then ValidateSheetTDS3 = False
'        If Not ValidatePAN8_TDS3 Then ValidateSheetTDS3 = False
        'End If

     
    End If
     setTableInfo_Grid4
     setTableInfo_Adahar_Tenant
     setTableInfo1_Grid4
     setTableInfo2_Grid4
     setTableInfo3_Grid4
     setTableInfo4_Grid4
     'setTableInfo12_Grid4  'Ankita_03/02/2025 Commented on 12/02/2025 as per V0/4
     setTableInfo6_Grid4
     setTableInfo3b_Grid3  'Malli 23/04/2025
     
'     setTableInfo7_Grid4
'     setTableInfo8_Grid4
'     setTableInfo9_Grid4
'     setTableInfo10_Grid4
'     setTableInfo11_Grid4
     
     'setUTNtable
     
     
     'If (Sheet1.Range("sheet1.PortugeseCC5A").Value = "Yes") Then
          'setTableInfo5_Grid4
          'setTableInfo9_Grid4
    'End If
    
'    'If (Sheet1.Range("sheet1.PortugeseCC5A").Value = "Yes") Then
'        If ((ColCount3 <> ColCount3_1) Or (ColCount3 <> ColCount3_2) Or (ColCount3 <> ColCount3_3) Or (ColCount3 <> ColCount3_4) Or (ColCount3 <> ColCount3_6) Or (ColCount3 <> ColCount3_7) Or (ColCount3 <> ColCount3_8) Or (ColCount3 <> ColCount3_10) Or (ColCount3 <> ColCount3_11) Or (ColCount3 <> ColCount3_5) Or (ColCount3 <> ColCount3_9)) Then
'            Msgbox_TDS = Msgbox_TDS + "Enter all mandatory fields for Schedule TDS Other than Salary" & Chr(13)
'            ValidateSheetTDS3 = False
'        End If
             
    'If ((ColCount3 <> ColCount3_1) Or (ColCount3 <> ColCount3_2) Or (ColCount3 <> ColCount3_4) Or (ColCount3 <> ColCount3_6) Or (ColCount3 <> ColCount3_3)) Then
     If ((ColCount3 <> ColCount3_1) Or (ColCount3 <> ColCount3_2) Or (ColCount3 <> ColCount3_4) Or (ColCount3 <> ColCount3_6) Or (ColCount3 <> ColCount3_3) Or (ColCount3 <> ColCount3b_3)) Then
        MsgBox_TDS3 = MsgBox_TDS3 + "* Enter all mandatory fields in Schedule TDS 3." & Chr(13)
        ValidateSheetTDS3 = False
     End If
     
     If ColCountAadhar > ColCount3 Then
    ' If ((ColCount3 <> ColCountAadhar) Or (ColCountAadhar <> ColCount3_1) Or (ColCountAadhar <> ColCount3_2) Or (ColCountAadhar <> ColCount3_4) Or (ColCountAadhar <> ColCount3_6) Or (ColCountAadhar <> ColCount3_3)) Then
         MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the PAN of Tenant in Schedule TDS 3." & Chr(13)
        ValidateSheetTDS3 = False
     'End If
     End If
     
     
          
'   If (ColCount3_3 <> ColCount3_9 Or ColCount3_3 <> ColCount3_7) Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Enter all mandatory fields of Col(6) in Schedule TDS 26QC" & Chr(13)
'    ValidateSheetTDS3 = False
'    End If
'
'    If (ColCount3_10 <> ColCount3_5 Or ColCount3_10 <> ColCount3_11) Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Enter all mandatory fields of col(8) in Schedule TDS 26QC" & Chr(13)
'    ValidateSheetTDS3 = False
'    End If
End Function
     



Function ValidatePAN_TDS3() As Boolean
ValidatePAN_TDS3 = True
setTableInfo_Grid4
noOfProcessSub = ColCount3

Dim rangecells As Range
Set rangecells = Range("TDS26QB.PAN").Cells
Dim i As Long
ReDim PAN_TDS(ColCount3)

For i = 1 To ColCount3
    PAN_TDS(i) = rangecells.item(i).Value
    If Not Len(PAN_TDS(i)) = 0 Then
        If Not CheckIFSCTDS(Mid(PAN_TDS(i), 1, 10)) Then
        MsgBox_TDS3 = MsgBox_TDS3 + "* PAN is invalid. First 5 alphabets, next 4 digits, then alphabet at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
        ValidatePAN_TDS3 = False
        End If
        
'            If Not ValidateTantype_text(Mid(Tan2_TDS(i), 1, 4)) Then
'                Msgbox_TDS = Msgbox_TDS + "TAN at Sr. No  " & i & " in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
'                ValidatePAN_TDS = False
'                Exit Function
'            End If
'            If Not IsNumeric(Mid(Tan2_TDS(i), 5, 5)) Then
'                Msgbox_TDS = Msgbox_TDS + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
'                ValidatePAN_TDS = False
'                Exit Function
'            End If
'            If Not ValidateTantype_text(Right(Tan2_TDS(i), 1)) Then
'                Msgbox_TDS = Msgbox_TDS + "TAN at Sr. No  " & i & "  in Sheet TDS  is invalid. First 4 alphabets, next 5 digits, then alphabet" & Chr(13)
'                ValidatePAN_TDS = False
'               Exit Function
             'End If
    ElseIf Not chkCompulsory(PAN_TDS(i)) Then
    'Change.27.02.2023.102.IDS.14
'        MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the PAN of Tenant is mandatory at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
        MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the PAN of Tenant at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
    'End ChangeIDS
        ValidatePAN_TDS3 = False
        Exit Function
    End If
    UpdateProgressBar
    Next
End Function

'Function ValidateAadhar_TDS3() As Boolean
'ValidateAadhar_TDS3 = True
'setTableInfo_Grid4
'
'Dim rangecells As Range
'Set rangecells = Range("TDS26QB.Aadhar_Number").Cells
'Dim i As Long
'ReDim Tenant_Aadhar_TDS(ColCount3)
'
'For i = 1 To ColCount3
'    Tenant_Aadhar_TDS(i) = rangecells.Item(i).Value
'
'    If (Tenant_Aadhar_TDS(i)) <> "" Then
'    If (Tenant_Aadhar_TDS(i)) < 12 Then
'        MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the valid Aadhaar number at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
'        ValidateAadhar_TDS3 = False
'    End If
'    End If
'
'Next
'End Function

Function ValidateAadhar_TDS3() As Boolean
On Error Resume Next

   ValidateAadhar_TDS3 = True
   setTableInfo_Grid4

Dim rangecells As Range
Set rangecells = Range("TDS26QB.Aadhar_Number").Cells
Dim i As Long
ReDim Tenant_Aadhar_TDS(ColCount3)

For i = 1 To ColCount3
    Tenant_Aadhar_TDS(i) = rangecells.item(i).Value
   
    If Trim(Tenant_Aadhar_TDS(i)) <> "" Then
    
        If Not IsNumeric(Tenant_Aadhar_TDS(i)) Then
            'errmsgID = "is invalid"
            MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the valid Aadhaar number at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
            ValidateAadhar_TDS3 = False
            Exit Function
        End If
        
        If Len(Tenant_Aadhar_TDS(i)) <> 12 Then
        MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the valid Aadhaar number at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
            ValidateAadhar_TDS3 = False
            Exit Function
        End If
        
        If Tenant_Aadhar_TDS(i) = "000000000000" Then
        MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the valid Aadhaar number at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
            'errmsgID = "is invalid"
            ValidateAadhar_TDS3 = False
            Exit Function
        End If
        
        If Tenant_Aadhar_TDS(i) = "111111111111" Then
        MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the valid Aadhaar number at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
            'errmsgID = "is invalid"
            ValidateAadhar_TDS3 = False
            Exit Function
        End If
        
    End If
Next


End Function


Function ValidateDeductor_TDS3() As Boolean
ValidateDeductor_TDS3 = True
setTableInfo_Grid4
Dim rangecells As Range
Set rangecells = Range("TDS26QB.EmployerOrDeductorName").Cells
Dim i As Long
ReDim DeductName2_TDS3(ColCount3)
For i = 1 To ColCount3
DeductName2_TDS3(i) = rangecells.item(i).Value
If Len(DeductName2_TDS3(i)) = 0 Then
End If

If Not chkCompulsory(DeductName2_TDS3(i)) Then
'Change.27.02.2023.102.IDS.13
'    MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the Name of Tenant is mandatory at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
    MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter the Name of Tenant at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
'End ChangeIDS
    ValidateDeductor_TDS3 = False
    Exit Function
End If
Next
End Function

'Function ValidateUniqueTDSCert_TDS2() As Boolean
'ValidateUniqueTDSCert_TDS2 = True
'Dim rangecells As Range
'Dim i, UTNintCount As Long
'Set rangecells = Range("TDSoth.UTN").Cells
'UTNintCount = Range("TDSoth.UTN").Count'
'
'ReDim UTN_TDS2(UTNintCount)
'For i = 1 To UTNintCount
'If Not rangecells.Item(i).Value = "" Then
'UTN_TDS2(i) = rangecells.Item(i).Value
'  If Not checkfieldspecialcharacter(rangecells.Item(i).Value) Then
'        Msgbox_TDS = Msgbox_TDS + "Unique TDS Certificate Numbers Cannot contain Special Characters at Sr.No " & i & "in Sch TDS2" & Chr(10)
'        ValidateUniqueTDSCert_TDS2 = False
'    Exit Function
'    End If
'      If Len(UTN_TDS2(i)) < 6 Or Len(UTN_TDS2(i)) > 8 Then
'        Msgbox_TDS = Msgbox_TDS + "Unique TDS Certificate Numbers Should 8 digits at Sr.No " & i & "in Sch TDS2" & Chr(10)
'        ValidateUniqueTDSCert_TDS2 = False
'    Exit Function
'    End If
'End If
'Next
'End Function


Function ValidateYear_TDS3() As Boolean
ValidateYear_TDS3 = True
setTableInfo_Grid4
Dim rangecells As Range
Set rangecells = Range("TDS26QB.DeductedYear").Cells
Dim i As Long
ReDim Year2_TDS3(ColCount3)
For i = 1 To ColCount3
Year2_TDS3(i) = rangecells.item(i).Value
If Len(Year2_TDS3(i)) = 0 Then
End If

If (Year2_TDS3(i) = "" Or UCase(Year2_TDS3(i)) = "(SELECT)") Then
    'MsgBox_TDS3 = MsgBox_TDS3 + "* Deduction Year is mandatory at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
    MsgBox_TDS3 = MsgBox_TDS3 + "* Please select deduction year from dropdown" & Chr(13)

    ValidateYear_TDS3 = False
    Exit Function
End If

If Not checkfieldspecialcharacter_TDS_TCS(Year2_TDS3(i)) Then
     MsgBox_TDS3 = MsgBox_TDS3 + "* Deduction Year at Sr. No  " & i & "  in Schedule TDS 3, characters < > & ' " & Chr(34) & " are not allowed." & Chr(13)
     ValidateYear_TDS3 = False
     Exit Function
  End If

Next
End Function


Function ValidateTaxDeducted_TDS3() As Boolean
ValidateTaxDeducted_TDS3 = True
    setTableInfo_Grid4
    Dim rangecells As Range
    Set rangecells = Range("TDS26QB.AmountDeducted").Cells
    Dim i As Long
    ReDim TaxDeducted_TDS3(ColCount3)
    For i = 1 To ColCount3
        TaxDeducted_TDS3(i) = rangecells.item(i).Value
        If Len(TaxDeducted_TDS3(i)) = 0 Then
            TaxDeducted_TDS3(i) = ""
        End If
        
        
        If TaxDeducted_TDS3(i) < 0 Then
            MsgBox_TDS3 = MsgBox_TDS3 + "* Gross receipt which is subject to tax deduction cannot be negative at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
            ValidateTaxDeducted_TDS3 = False
            Exit Function
        End If
        
        If Len(TaxDeducted_TDS3(i)) > 14 Then
            MsgBox_TDS3 = MsgBox_TDS3 + "* Gross receipt which is subject to tax deduction cannot exceed 14 Digits at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
            ValidateTaxDeducted_TDS3 = False
            Exit Function
        End If
        
        If Not chkCompulsory(TaxDeducted_TDS3(i)) Then
           ' MsgBox_TDS3 = MsgBox_TDS3 + "* Gross receipt which is subject to tax deduction is mandatory at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
            MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter Gross receipt which is subject to tax deduction" & Chr(13)
            ValidateTaxDeducted_TDS3 = False
            Exit Function
        End If
        
       
        'Ankita_02/06/2025
    If Not IsNumeric(TaxDeducted_TDS3(i)) Then
            MsgBox_TDS3 = MsgBox_TDS3 & "* Gross Amount at Sr. No  " & i & "  in Schedule TDS 3 should be Numeric value." & Chr(13)
            ValidateTaxDeducted_TDS3 = False
            Exit Function
        End If
        '-------------

    Next
End Function





'Validation TDS3

 
Function ValidateTotTaxDeducted_TDS3() As Boolean
ValidateTotTaxDeducted_TDS3 = True
setTableInfo_Grid4
Dim rangecells As Range
Set rangecells = Range("TDS26QB.TotTDSOnAmtPaid").Cells
Dim i As Long
ReDim TotTaxDedct_TDS3(ColCount3)
For i = 1 To ColCount3
TotTaxDedct_TDS3(i) = rangecells.item(i).Value
If Len(TotTaxDedct_TDS3(i)) = 0 Then
End If


If TotTaxDedct_TDS3(i) < 0 Then
    MsgBox_TDS3 = MsgBox_TDS3 + "* Tax deducted in own hands cannot be negative at Sr. No  " & i & " in Schedule TDS 3." & Chr(13)
    ValidateTotTaxDeducted_TDS3 = False
    Exit Function
End If

If Len(TotTaxDedct_TDS3(i)) > 14 Then
    MsgBox_TDS3 = MsgBox_TDS3 + "* Tax deducted in own hands cannot exceed 14 Digits at Sr. No  " & i & " in Schedule TDS 3." & Chr(13)
    ValidateTotTaxDeducted_TDS3 = False
    Exit Function
End If

If Not chkCompulsory(TotTaxDedct_TDS3(i)) Then
    MsgBox_TDS3 = MsgBox_TDS3 + "*  Please enter Tax deducted at Sr. No  " & i & " in Schedule TDS3." & Chr(13)
    ValidateTotTaxDeducted_TDS3 = False
    Exit Function
End If
Next
End Function

Function Validateinownhands6_TDS3() As Boolean
Validateinownhands6_TDS3 = True
setTableInfo_Grid4
Dim rangecells As Range
Set rangecells = Range("TDS26QB.6income").Cells
Dim rangecells1 As Range
Set rangecells1 = Range("TDS26QB.TotTDSOnAmtPaid").Cells
Dim i As Long
ReDim inownhands6_TDS3(ColCount3)
ReDim Final_TDS2(ColCount3)
For i = 1 To ColCount3
inownhands6_TDS3(i) = rangecells.item(i).Value
Final_TDS2(i) = rangecells1.item(i).Value
If Len(inownhands6_TDS3(i)) = 0 Then
End If


If inownhands6_TDS3(i) < 0 Then
    MsgBox_TDS3 = MsgBox_TDS3 + "* TDS Credit out of (5) claimed this Year cannot be negative at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
    Validateinownhands6_TDS3 = False
    Exit Function
End If

If Len(inownhands6_TDS3(i)) > 14 Then
    MsgBox_TDS3 = MsgBox_TDS3 + "* TDS Credit out of (6) claimed this Year cannot exceed 14 Digits at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
    Validateinownhands6_TDS3 = False
    Exit Function
End If



If Not chkCompulsory(inownhands6_TDS3(i)) Then
    
    MsgBox_TDS3 = MsgBox_TDS3 + "* Please enter Amount out of (6) claimed this Year at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
    'MsgBox_TDS3 = MsgBox_TDS3 + "* TDS credit out of (6) claimed this Year is mandatory at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
    Validateinownhands6_TDS3 = False
    Exit Function
End If

If inownhands6_TDS3(i) > Final_TDS2(i) Then
    MsgBox_TDS3 = MsgBox_TDS3 + "* TDS credit claimed cannot be more than Tax deducted  at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)
    Validateinownhands6_TDS3 = False
    Exit Function
End If
Next
End Function


'Function ValidateTDS6_TDS3() As Boolean
'ValidateTDS6_TDS3 = True
'setTableInfo3_Grid4
'Dim rangecells As Range
'Set rangecells = Range("TDS26QB.6TDS").Cells
'Dim i As Long
'ReDim TDS6_TDS3(ColCount3_3)
'For i = 1 To ColCount3_3
'TDS6_TDS3(i) = rangecells.Item(i).Value
'If Len(TDS6_TDS3(i)) = 0 Then
'End If
'
'
'
'If TDS6_TDS3(i) < 0 Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter TDS amount col(6) in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (20 TDS3) cannot be negative" & Chr(13)
'    ValidateTDS6_TDS3 = False
'    Exit Function
'End If
'
'If Len(TDS6_TDS3(i)) > 14 Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter TDS amount col(6) in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (20 TDS3) cannot exceed 14 digits" & Chr(13)
'    ValidateTDS6_TDS3 = False
'    Exit Function
'End If
'
'If Not chkCompulsory(TDS6_TDS3(i)) Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter TDS amount col(6) in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (20 TDS3) is mandatory" & Chr(13)
'    ValidateTDS6_TDS3 = False
'    Exit Function
'End If
'Next
'End Function
'
'Function ValidatePAN6_TDS3() As Boolean
'ValidatePAN6_TDS3 = True
'setTableInfo3_Grid4
'noOfProcessSub = ColCount3_3
'
'Dim rangecells As Range
'Set rangecells = Range("TDS26QB.6panspouse").Cells
'Dim i As Long
'ReDim PAN6_TDS3(ColCount3_3)
'
'For i = 1 To ColCount3_3
'    PAN6_TDS3(i) = rangecells.Item(i).Value
'    If Not Len(PAN6_TDS3(i)) = 0 Then
'        If Not CheckIFSCTDS(Mid(PAN6_TDS3(i), 1, 10)) Then
'        MsgBox_TDS3 = MsgBox_TDS3 + "PAN of the spouse or other person at Sr. No  " & i & "  in Sheet TDS3 26QC is invalid. First 5 alphabets, next 4 digits, then alphabet  " & Chr(13)
'        ValidatePAN6_TDS3 = False
'        End If
'
'
'    ElseIf Not chkCompulsory(PAN6_TDS3(i)) Then
'        MsgBox_TDS3 = MsgBox_TDS3 + "PAN of the spouse or other person at Sr.No " & i & " in Sheet TDS3 26QC, (20 TDS3) is mandatory" & Chr(13)
'        ValidatePAN6_TDS3 = False
'        Exit Function
'    End If
'    UpdateProgressBar
'    Next
'End Function
'
'
'
'
''Function ValidatePAN6other_TDS3() As Boolean
''ValidatePAN6other_TDS3 = True
''setTableInfo_Grid4
''noOfProcessSub = ColCount3
''
''Dim rangecells As Range
''Set rangecells = Range("TDS26QB.6panperson").Cells
''Dim i As Long
''ReDim PAN6other_TDS3(ColCount3)
''
''For i = 1 To ColCount3
''    PAN6other_TDS3(i) = rangecells.Item(i).Value
''    If Not Len(PAN6other_TDS3(i)) = 0 Then
''        If Not mIncmDtls.CheckPAN(Mid(PAN6other_TDS3(i), 1, 10)) Then
''        MsgBox_TDS3 = MsgBox_TDS3 + "PAN of the other person at Sr. No  " & i & "  in Sheet TDS3 is invalid. First 5 alphabets, next 4 digits, then alphabet (4th alphabet must be ""P"") " & Chr(13)
''        ValidatePAN6other_TDS3 = False
''        End If
''
''   End If
''
''    UpdateProgressBar
''Next
''End Function
'
'' new
'
'
'Function ValidateTotTaxDeducted7_TDS3() As Boolean
'ValidateTotTaxDeducted7_TDS3 = True
'setTableInfo_Grid4
'Dim rangecells As Range
'Dim mintr As Variant
'Set rangecells = Range("TDS26QB.TotTDSOnAmtPaid7").Cells
''mintr = Range("TDS26QB.TotTDSOnAmtPaid7").Cells.Count
'Dim i As Long
'ReDim TotTaxDedct7_TDS3(ColCount3)
'For i = 1 To ColCount3
'TotTaxDedct7_TDS3(i) = rangecells.Item(i).Value
'If Len(TotTaxDedct7_TDS3(i)) = 0 Then
'End If
'
'If Not chkCompulsory(TotTaxDedct7_TDS3(i)) Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter TDS credit Amount out of (5) or (6) being claimed this Year  in own hands at Sr.NO " & i & " in Sheet TDS3, (20 TDS3) is mandatory" & Chr(13)
'    ValidateTotTaxDeducted7_TDS3 = False
'    Exit Function
'End If
'
'If TotTaxDedct7_TDS3(i) < 0 Then
'    MsgBox_TDS3 = MsgBox_TDS3 + " TDS credit Amount out of (5) or (6) being claimed this Year in own hands at Sr.NO " & i & " cannot be negative in Sheet TDS3, (20 TDS3 26QC) " & Chr(13)
'    ValidateTotTaxDeducted7_TDS3 = False
'    Exit Function
'End If
'
'If Len(TotTaxDedct7_TDS3(i)) > 14 Then
'    MsgBox_TDS3 = MsgBox_TDS3 + " TDS credit Amount out of (5) or (6) being claimed this Year in own hands at Sr.NO " & i & " cannot exceed 14 Digits in Sheet TDS3, (20 TDS3 26QC) " & Chr(13)
'    ValidateTotTaxDeducted7_TDS3 = False
'    Exit Function
'End If
'
'If (Range("TDS26QB.TotTDSOnAmtPaid7").Cells.Item(i).Value + Range("TDS26QB.8TDS").Cells.Item(i).Value > (Range("TDS26QB.6TDS").Cells.Item(i).Value + Range("TDS26QB.totTDSOnAmtPaid").Cells.Item(i).Value)) Then
'MsgBox_TDS3 = MsgBox_TDS3 + " Amount in field Col(7)and TDS at col(8) cannot be more than sum of field Col(5) and TDS at col(6)  in sch TDS3 26QC " & Chr(13)
'ValidateTotTaxDeducted7_TDS3 = False
'Exit Function
'End If
'Next
'End Function
'
'Function Validateinownhands8_TDS3() As Boolean
'Validateinownhands8_TDS3 = True
'setTableInfo10_Grid4
'Dim rangecells As Range
'Set rangecells = Range("TDS26QB.ClaimOutOfTotTDSOnAmtPaid37BA").Cells
'Dim i As Long
'ReDim inownhands8_TDS3(ColCount3_10)
'For i = 1 To ColCount3_10
'inownhands8_TDS3(i) = rangecells.Item(i).Value
'If Len(inownhands8_TDS3(i)) = 0 Then
'End If
'
'
'If inownhands8_TDS3(i) < 0 Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter Income in col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (20 TDS3) cannot be negative" & Chr(13)
'    Validateinownhands8_TDS3 = False
'    Exit Function
'End If
'
'If Len(inownhands8_TDS3(i)) > 14 Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter Income in col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (20 TDS3) cannot exceed 14 Digits" & Chr(13)
'    Validateinownhands8_TDS3 = False
'    Exit Function
'End If
'
'
'If Not chkCompulsory(inownhands8_TDS3(i)) Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter Income in col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (20 TDS3) is mandatory" & Chr(13)
'    Validateinownhands8_TDS3 = False
'    Exit Function
'End If
'Next
'End Function
'
'
'Function ValidateTDS8_TDS3() As Boolean
'ValidateTDS8_TDS3 = True
'setTableInfo10_Grid4
'Dim rangecells As Range
'Set rangecells = Range("TDS26QB.8TDS").Cells
'Dim i As Long
'ReDim TDS8_TDS3(ColCount3_10)
'For i = 1 To ColCount3_10
'TDS8_TDS3(i) = rangecells.Item(i).Value
'If Len(TDS8_TDS3(i)) = 0 Then
'End If
'
'If TDS8_TDS3(i) < 0 Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter TDS in col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (20 TDS3) cannot be negative" & Chr(13)
'    ValidateTDS8_TDS3 = False
'    Exit Function
'End If
'
'
'If Len(TDS8_TDS3(i)) > 14 Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter TDS in col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (20 TDS3) cannot exceed 14 Digits" & Chr(13)
'    ValidateTDS8_TDS3 = False
'    Exit Function
'End If
'
'If Not chkCompulsory(TDS8_TDS3(i)) Then
'    MsgBox_TDS3 = MsgBox_TDS3 + "Please enter TDS in col(8) which is subject to tax deduction in the hands of spouse as per section 5A or any other person as per rule 37BA(2) Sr.NO " & i & " in Sheet TDS, (20 TDS3) is mandatory" & Chr(13)
'    ValidateTDS8_TDS3 = False
'    Exit Function
'End If
'Next
'End Function
'
'Function ValidatePAN8_TDS3() As Boolean
'ValidatePAN8_TDS3 = True
'setTableInfo10_Grid4
'noOfProcessSub = ColCount3_10
'
'Dim rangecells As Range
'Set rangecells = Range("TDS26QB.AmtClaimedBySpouse").Cells
'Dim i As Long
'ReDim PAN8_TDS3(ColCount3_10)
'
'For i = 1 To ColCount3_10
'    PAN8_TDS3(i) = rangecells.Item(i).Value
'    If Not Len(PAN8_TDS3(i)) = 0 Then
'        If Not CheckIFSCTDS(Mid(PAN8_TDS3(i), 1, 10)) Then
'        MsgBox_TDS3 = MsgBox_TDS3 + "PAN of the spouse or other person  at Sr. No  " & i & "  in Sheet TDS3 26QC  is invalid. First 5 alphabets, next 4 digits, then alphabet  " & Chr(13)
'        ValidatePAN8_TDS3 = False
'        End If
'
'
'        ElseIf Not chkCompulsory(PAN8_TDS3(i)) Then
'        MsgBox_TDS3 = MsgBox_TDS3 + " PAN of the spouse or other person at Sr.No " & i & " in Sheet TDS3 26QC , (20 TDS3) is mandatory" & Chr(13)
'        ValidatePAN8_TDS3 = False
'        Exit Function
'    End If
'    UpdateProgressBar
'    Next
'End Function




'Function ValidatePAN8other_TDS3() As Boolean
'ValidatePAN8other_TDS3 = True
'setTableInfo_Grid2
'noOfProcessSub = ColCount3
'
'Dim rangecells As Range
'Set rangecells = Range("TDS26QB.8panperson").Cells
'Dim i As Long
'ReDim PAN8other_TDS3(ColCount3)
'
'For i = 1 To ColCount3
'    PAN8other_TDS3(i) = rangecells.Item(i).Value
'    If Not Len(PAN8other_TDS3(i)) = 0 Then
'        If Not mIncmDtls.CheckPAN(Mid(PAN8other_TDS3(i), 1, 10)) Then
'        MsgBox_TDS3 = MsgBox_TDS3 + "PAN of the other person at Sr. No  " & i & "  in Sheet TDS2  is invalid. First 5 alphabets, next 4 digits, then alphabet (4th alphabet must be ""P"") " & Chr(13)
'        ValidatePAN8other_TDS3 = False
'        End If
'    End If
'
'
'    UpdateProgressBar
'Next
'End Function

Function ValidateTotAmtClaimed3() As Boolean
ValidateTotAmtClaimed3 = True
Dim val As Variant
val = Sheet2.Range("TDS26QB.Sum").Value

If Len(val) > 14 Then
    MsgBox_TDS3 = MsgBox_TDS3 + "* Total Amount claimed this year cannot be more than 14 digit in Schedule TDS 3."
    ValidateTotAmtClaimed3 = False
Exit Function
End If
End Function






'
'Function ValidateClaimOutofTotAmount_TDS() As Boolean
'ValidateClaimOutofTotAmount_TDS = True
'setTableInfo_Grid2
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.ClaimOutOfTotTDSOnAmtPaid").Cells
'Dim i As Long
'ReDim ClaimAmt2_TDS(ColCount1)
'For i = 1 To ColCount1
'ClaimAmt2_TDS(i) = rangecells.Item(i).Value
'If Len(ClaimAmt2_TDS(i)) = 0 Then
'End If
'
'If Not chkCompulsory(ClaimAmt2_TDS(i)) Then
'    Msgbox_TDS = Msgbox_TDS + "Amount Claimed at Sr.NO " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'    ValidateClaimOutofTotAmount_TDS = False
'    Exit Function
'End If
'Next
'End Function


'Function ValidatePortugeseCC5A_TDS() As Boolean
'ValidatePortugeseCC5A_TDS = True
'setTableInfo_Grid2
'Dim rangecells As Range
'Set rangecells = Range("TDSoth.AmtClaimedBySpouse").Cells
'Dim i As Long
'ReDim Portuguese2_TDS(ColCount1)
'For i = 1 To ColCount1
'Portuguese2_TDS(i) = rangecells.Item(i).Value
'If Len(Portuguese2_TDS(i)) = 0 Then
'End If
'If Not chkCompulsory(Portuguese2_TDS(i)) Then
'    Msgbox_TDS = Msgbox_TDS + "Amount claimed in the hands of Spouse at Sr.NO " & i & " in Sheet TDS, (19 TDS2) is mandatory" & Chr(13)
'    ValidatePortugeseCC5A_TDS = False
'    Exit Function
'End If
'Next
'End Function



'Common Table count finder

Sub setTableInfo_Grid4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.PAN").count
    Set rangecells = Range("TDS26QB.PAN").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount3 = ccount
End Sub
Sub setTableInfo_Adahar_Tenant()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.Aadhar_Number").count
    Set rangecells = Range("TDS26QB.Aadhar_Number").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCountAadhar = ccount
End Sub
Sub setTableInfo1_Grid4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.DeductedYear").count
    Set rangecells = Range("TDS26QB.DeductedYear").Cells
    For mIntCtr = 1 To mIntCells
        If Not (rangecells.item(mIntCtr).Value = "" Or UCase(rangecells.item(mIntCtr).Value) = "(SELECT)") Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_1 = ccount
End Sub

Sub setTableInfo2_Grid4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.TotTDSOnAmtPaid").count
    Set rangecells = Range("TDS26QB.TotTDSOnAmtPaid").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_2 = ccount
End Sub

Sub setTableInfo3_Grid4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.6income").count
    Set rangecells = Range("TDS26QB.6income").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_3 = ccount
End Sub

Sub setTableInfo4_Grid4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.EmployerOrDeductorName").count
    Set rangecells = Range("TDS26QB.EmployerOrDeductorName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_4 = ccount
End Sub



Sub setTableInfo5_Grid4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.AmtClaimedBySpouse").count
    Set rangecells = Range("TDS26QB.AmtClaimedBySpouse").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_5 = ccount
End Sub

Sub setTableInfo6_Grid4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.AmountDeducted").count
    Set rangecells = Range("TDS26QB.AmountDeducted").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_6 = ccount
End Sub
'Malli---------
Sub setTableInfo3b_Grid3()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDsOthr2.SectionTDSDeducted").count
    Set rangecells = Range("TDsOthr2.SectionTDSDeducted").Cells
    For mIntCtr = 1 To mIntCells
         
         If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        ccount = ccount + 1
        End If
    Next
    ColCount3b_3 = ccount
End Sub
'Malli-----tds2
 
Sub setTableInfo2b_Grid2()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDsOthr.SectionTDS").count
    Set rangecells = Range("TDsOthr.SectionTDS").Cells
    For mIntCtr = 1 To mIntCells
         
         If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        ccount = ccount + 1
        End If
    Next
    ColCount2b_2 = ccount
End Sub

Sub setTableInfo7_Grid4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.6TDS").count
    Set rangecells = Range("TDS26QB.6TDS").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_7 = ccount
End Sub

Sub setTableInfo8_Grid4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.TotTDSOnAmtPaid7").count
    Set rangecells = Range("TDS26QB.TotTDSOnAmtPaid7").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_8 = ccount
End Sub

Sub setTableInfo9_Grid4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.6panspouse").count
    Set rangecells = Range("TDS26QB.6panspouse").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_9 = ccount
End Sub
Sub setTableInfo10_Grid4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.ClaimOutOfTotTDSOnAmtPaid37BA").count
    Set rangecells = Range("TDS26QB.ClaimOutOfTotTDSOnAmtPaid37BA").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_10 = ccount
End Sub

Sub setTableInfo11_Grid4()
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("TDS26QB.8TDS").count
    Set rangecells = Range("TDS26QB.8TDS").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCount3_11 = ccount
End Sub

'Ankita_03/02/2025
'Commented on 12/02/2025 by Ankita as V0.4
'Sub setTableInfo12_Grid4()
'Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("TDsOthr.SectionTDS_ii").count
'    Set rangecells = Range("TDsOthr.SectionTDS_ii").Cells
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    ColCount3_12 = ccount
'End Sub


Function CheckIFSCTDS(IFSC As String) As Boolean
On Error Resume Next
'IFSC Code should be exactly 11 characters,
'First 4 characters should be alphabets,
'5th character must be zero (0) and remaining 6 should be either numeric or alphabets
    CheckIFSCTDS = True
    Dim arr1 As Variant
    Dim j As Variant
  '  arr1 = Array("P", "C", "F", "L", "A", "B", "T", "J", "G")
    If Len(IFSC) > 0 Then
        If Not ChkAlphabet(Mid(IFSC, 1, 1)) Then
            CheckIFSCTDS = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(IFSC, 2, 1)) Then
            CheckIFSCTDS = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(IFSC, 3, 1)) Then
            CheckIFSCTDS = False
            Exit Function
        End If
        
'                For j = 0 To UBound(arr1)
'                    If Mid(IFSC, 4, 1) = arr1(j) Then
'                        CheckIFSCTDS = True
'                        Exit For
''                       Exit Function
'                    Else
'                    CheckIFSCTDS = False
'                    End If
'                Next
        If Not ChkAlphabet(Mid(IFSC, 4, 1)) Then
            CheckIFSCTDS = False
            Exit Function
        End If
    
        If Not ChkAlphabet(Mid(IFSC, 5, 1)) Then
            CheckIFSCTDS = False
            Exit Function
        End If
        
     
            If Not IsNumeric(Mid(IFSC, 6, 1)) Then
                CheckIFSCTDS = False
                Exit Function
            End If
            
            If Not IsNumeric(Mid(IFSC, 7, 1)) Then
                CheckIFSCTDS = False
                Exit Function
           End If

            If Not IsNumeric(Mid(IFSC, 8, 1)) Then
                CheckIFSCTDS = False
                Exit Function
            End If

            If Not IsNumeric(Mid(IFSC, 9, 1)) Then
                CheckIFSCTDS = False
                Exit Function
            End If

            If Not ChkAlphabet(Mid(IFSC, 10, 1)) Then
                CheckIFSCTDS = False
                 Exit Function
            End If
            
End If
End Function

Function ValidateTANCodes(TAN As String) As Boolean
    ValidateTANCodes = False
    Dim TANCode As Range

    
    Dim TAN1 As Range


''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet2.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
    
    Set TANCode = Sheet5.Range("TANCodes")
  
    
    For Each TAN1 In TANCode
        If Mid(TAN1.Value, 1, 3) = Mid(TAN, 1, 3) Then
        ValidateTANCodes = True
        Exit For
        End If
    Next
    
''----------------Unlock Password-------------------START---
'   sPassword = EfilingCommon.getmsgstate
'   Sheet2.Unprotect Password:=sPassword
''----------------Unlock Password-------------------END-----
End Function
'Malli
'Malli
Function CheckDateBefore2(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.


    CheckDateBefore2 = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBefore2 = False
           ' MsgBox ("Date of Credit into Govt Account in Sheet : TDS  must be a valid dd/mm/yyyy format")
            Exit Function
        End If
'        Change28.11.2022.102.06A
        'TAX DETAILS-E5
        'If EfilingCommon.checkFirstDateBefore(dob, "31/03/2022") Then
        
        'TAX DETAILS-C5
        If EfilingCommon.checkFirstDateBefore(dob, "01/04/2023") Then
        
'        ---End
           'MsgBox ("Date of of Credit into Govt Account in Sheet : TDS should be on or after 01/04/2017")
            CheckDateBefore2 = False
            Exit Function
            Else
            CreditDate = dob
       End If
    End If
 End Function

'Jyoti2025-26 19/04/2025
Function ValidateSectionTDSDeducted_TDS2i() As Boolean
ValidateSectionTDSDeducted_TDS2i = True
setTableInfo_Grid2

Dim rangecells As Range
Set rangecells = Range("TDsOthr.SectionTDS").Cells
Dim i As Long
ReDim Section_TDS2(ColCount1)
2
For i = 1 To ColCount1
Section_TDS2(i) = rangecells.item(i).Value

If Len(Section_TDS2(i)) = 0 Then
End If
 
'If Not chkCompulsory(Section_TDS2(i)) Then
'
' MsgBox_TDS2 = MsgBox_TDS2 + "Please select the section under which TDS is deducted from dropdown." & Chr(13)
'
'    ValidateSectionTDSDeducted_TDS2i = False
'    Exit Function
'End If
'
'If Section_TDS2(i) = "(Select)" Then
'         MsgBox_TDS2 = MsgBox_TDS2 + " Please select the section under which TDS is deducted from dropdown." & Chr(13)
'
'         'End ChangeIDS
'    ValidateSectionTDSDeducted_TDS2i = False
'    Exit Function
'End If
  If isdropdownblank(Section_TDS2(i)) Then
           Msgbox_TDS = Msgbox_TDS + "* ""Please select the section under which TDS is deducted from dropdown"" at Sr. No  " & i & "  in Schedule TDS 2." & Chr(13)
  ValidateSectionTDSDeducted_TDS2i = False
    Exit Function
  End If
  
  'Newly added by Bindu
  Dim TDS_section_error As Range
  Dim TDS_section_error_count, iCount_error As Long
      Set TDS_section_error = Sheet5.Range("TDS_Section_Error").Cells
       TDS_section_error_count = Sheet5.Range("TDS_Section_Error").count
       
        For iCount_error = 1 To TDS_section_error_count
       
          If UCase(Trim(Section_TDS2(i))) = UCase(Trim(TDS_section_error.item(iCount_error).Value)) Then
            Msgbox_TDS = Msgbox_TDS + "*TDS relating to salaries is required to be claimed in Schedule TDS 1 at Sr. No  " & i & "" & Chr(13)
            ValidateSectionTDSDeducted_TDS2i = False
          End If
       
        Next


Next
End Function


'Jyoti2025-26 19/04/2025

Function ValidateSectionTDSDeducted_TDS2ii() As Boolean
ValidateSectionTDSDeducted_TDS2ii = True
setTableInfo_Grid4
Dim rangecells As Range
Set rangecells = Range("TDsOthr2.SectionTDSDeducted").Cells
Dim i As Long
ReDim Section_TDS3(ColCount3)
For i = 1 To ColCount3
Section_TDS3(i) = rangecells.item(i).Value
If Len(Section_TDS3(i)) = 0 Then
End If
 
If Not chkCompulsory(Section_TDS3(i)) Then

 MsgBox_TDS3 = MsgBox_TDS3 + "* ""Please select the section under which TDS is deducted from dropdown"" at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)

    ValidateSectionTDSDeducted_TDS2ii = False
    Exit Function
End If

If Section_TDS3(i) = "(Select)" Then
         MsgBox_TDS3 = MsgBox_TDS3 + "* ""Please select the section under which TDS is deducted from dropdown"" at Sr. No  " & i & "  in Schedule TDS 3." & Chr(13)

    ValidateSectionTDSDeducted_TDS2ii = False
    Exit Function
End If

'newly added by Bindu

Dim TDS_section_error_1 As Range
  Dim TDS_section_error_count_1, iCount_error_1 As Long
      Set TDS_section_error_1 = Sheet5.Range("TDS_section_error").Cells
       TDS_section_error_count_1 = Sheet5.Range("TDS_section_error").count
       
        For iCount_error_1 = 1 To TDS_section_error_count_1
       
          If UCase(Trim(Section_TDS3(i))) = UCase(Trim(TDS_section_error_1.item(iCount_error_1).Value)) Then
            MsgBox_TDS3 = MsgBox_TDS3 + "*TDS relating to salaries is required to be claimed in Schedule TDS 1 at Sr. No  " & i & "" & Chr(13)
            ValidateSectionTDSDeducted_TDS2ii = False
          End If
       
        Next

Next
End Function


'Ankita_06/01/2026==========
Function CheckDateBefore3(dob As Variant) As Boolean
On Error Resume Next
    CheckDateBefore3 = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBefore3 = False
            Exit Function
        End If
        If EfilingCommon.checkFirstDateBefore(dob, Sheet5.Range("IT_DOD").Value) Then 'Changed by Ankita on 29/11/2024
            CheckDateBefore3 = False
            Exit Function
            Else
            CreditDate = dob
       End If
    End If
End Function
 

'===========================


