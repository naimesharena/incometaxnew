Attribute VB_Name = "SchTaxVerify"
Public Others_NOI, Others_NOI1, Others_NOI22, Others_Amt, end_OthersNOI, end_OthersAmt As Variant


Sub AddRows_Others()
    Dim vRows As Long
    Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
    sourceSheet.Activate
    EfilingCommon.DefinedgridNameRange = "Others.NOI||Nature_Others||Others.Amount||Sheet5.DescEI"   'Ankita_17/06/2026
    ActiveCellRange = EfilingCommon.searchLastRow("Others.NOI")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub



Function CheckIFSC(IFSC As String) As Boolean
On Error Resume Next
'IFSC Code should be exactly 11 characters,
'First 4 characters should be alphabets,
'5th character must be zero (0) and remaining 6 should be either numeric or alphabets
    CheckIFSC = True
    If Len(IFSC) > 0 Then
        If Not ChkAlphabet(Mid(IFSC, 1, 1)) Then
            CheckIFSC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(IFSC, 2, 1)) Then
            CheckIFSC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(IFSC, 3, 1)) Then
            CheckIFSC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(IFSC, 4, 1)) Then
            CheckIFSC = False
            Exit Function
        End If
        If Not Mid(IFSC, 5, 1) = 0 Then
            CheckIFSC = False
            Exit Function
        End If
        
       ' If IsNumeric(Mid(IFSC, 6, 6)) Then
       '     CheckIFSC = True
       '     Exit Function
       ' Else
           
            If Not IsNumeric(Mid(IFSC, 6, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 6, 1)) Then
            CheckIFSC = False
            Exit Function
            End If
            End If

            
            If Not IsNumeric(Mid(IFSC, 7, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 7, 1)) Then
            CheckIFSC = False
           Exit Function
            End If
            End If

            If Not IsNumeric(Mid(IFSC, 8, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 8, 1)) Then
            CheckIFSC = False
            Exit Function
            End If
            End If

            If Not IsNumeric(Mid(IFSC, 9, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 9, 1)) Then
            CheckIFSC = False
            Exit Function
            End If
            End If

            If Not IsNumeric(Mid(IFSC, 10, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 10, 1)) Then
            CheckIFSC = False
           Exit Function
            End If
            End If

            If Not IsNumeric(Mid(IFSC, 11, 1)) Then
            If Not ChkAlphabet(Mid(IFSC, 11, 1)) Then
            CheckIFSC = False
           Exit Function
            End If
            End If
If Not checkfieldSuperSpecialcharacterDot(IFSC) Then
CheckIFSC = False
End If
          

'End If
End If
        
    
End Function


Sub ValidatePartBTI_BTTI_Verification()
    subProcCaption = " Validating Part B-TTI"
    noOfProcessSub = 7
    vbMessgaeCaption = "Error"
    
    
    
    If Trim(Sheet3.Range("Ver.AssesseeVerName").Value) = "" Or IsEmpty((Sheet3.Range("Ver.AssesseeVerName").Value)) Then
       ' EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Assessee name is mandatory in Verification in sheet Tax Paid and Verification is  mandatory " & Chr(13)
       'ADDED BY AAVULA NARESH
       
        EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Assessee name is mandatory in Verification" & Chr(13)
    Else
        VerificationName = UCase(Sheet3.Range("Ver.AssesseeVerName").Value)
    End If
          
    UpdateProgressBar
    If Trim(Sheet3.Range("Ver.FatherName").Value) = "" Or IsEmpty((Sheet3.Range("Ver.FatherName").Value)) Then
        'EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Father's Name in Verification in sheet Tax Paid and Verification is  mandatory " & Chr(13)
        'ADDED BY AAVULA NARESH
        EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Father's name is mandatory in Verification" & Chr(13)

    Else
        FatherName = Sheet3.Range("Ver.FatherName").Value
    End If
    
    If Not checkfieldSuperSpecialcharactername(Sheet3.Range("Ver.FatherName").Value) Then
    EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Name of father should not Contain <, > " & Chr(34) & "& characters" & Chr(13)
    End If
    UpdateProgressBar
    
    
    If Trim(Sheet3.Range("Ver.capacity").Value) = "" Or isdropdownblank((Sheet3.Range("Ver.capacity").Value)) Then
        'EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Selection of capacity in verification Part is mandatory. Please select appropriate option from drop down in sheet Tax Paid and Verification" & Chr(13)
        'ADDED BY AAVULA NARESH
        EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Selection of capacity is mandatory. Please select appropriate option from drop down" & Chr(13)

'    Else
'        FatherName = Sheet3.Range("Ver.capacity").Value
    End If
    
    'Ankita_12/01/2026=========
    If (Range("Ver.Capacity").Value) = "Representative" And (Sheet1.Range("sheet1.RepAssessee").Value = "" Or UCase(Sheet1.Range("sheet1.RepAssessee").Value) = UCase("(Select)")) Then
         EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Details of representative assessee is mandatory as the return is being filed by representative in Sheet Income Details" & Chr(13)
    End If

'Ankita_10/02/2026========

    If (Range("Ver.Capacity").Value) = "Self" And (Sheet1.Range("sheet1.RepAssessee").Value) = "Yes" Then
        EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* ""Please select capacity as ""Representative"" as Yes is selected in sl no. A22 of Part A General Information."""
    End If
'
'    If (Range("Ver.Capacity").Value) = "Representative" And (Sheet1.Range("sheet1.RepAssessee").Value) = "No" Then
'        EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* ""Please select capacity as ""Representative"" as No is selected in sl no. A22 of Part A General Information."""
'    End If

    'Ankita_05/03/2026_SIT-113212
If Dformat(Trim(Range("sheet1.DOB").Value), "yyyy-mm-dd") >= Dformat(Trim("01/04/2008"), "yyyy-mm-dd") And Sheet3.Range("Ver.Capacity").Value = "Self" Then
    EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Since Date of Birth is on or after 01/04/2008 the return cannot be verified through ""Self"", please verify the return through ""Representative""." & Chr(13)
End If

    'Ankita_05/03/2026_SIT-113182
If Dformat(Trim(Range("sheet1.DOB").Value), "yyyy-mm-dd") >= Dformat(Trim("01/04/2008"), "yyyy-mm-dd") And Sheet3.Range("Ver.PAN").Value = Sheet1.Range("sheet1.PAN").Value Then
    EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* If date of birth is on or after 1 April 2008 then in verification PAN cannot be same as mentioned in general information." & Chr(13)
End If

'========================

    '==========================
'    If calculateAge(Sheet1.Range("sheet1.DOB").Value) < 18 Then
'    If Sheet3.Range("Ver.capacity").Value = "Self" Then
'    EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "*in varification drop down cannot be selected as self" & Chr(13)
'    End If
'    End If
    
    If Not (Trim(Sheet3.Range("Ver.capacity").Value) = "" Or isdropdownblank((Sheet3.Range("Ver.capacity").Value))) Then
    If Not (Trim(Sheet3.Range("Ver.capacity").Value) = "Self" Or Trim(Sheet3.Range("Ver.capacity").Value) = "Representative") Then
    EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Please select appropriate option from drop down in sheet Tax Paid and Verification" & Chr(13)

    End If
    End If
    
    
    
    UpdateProgressBar
    If Trim(Sheet3.Range("Ver.Place").Value) = "" Or IsEmpty((Sheet3.Range("Ver.Place").Value)) Then
        'EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Place in Verification  in sheet Tax Paid and Verification is  mandatory " & Chr(13)
    'SIT-66371 Change 'Malli
    EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "*Place is mandatory in Verification" & Chr(13)
    
    Else
        VerificationPlace = Sheet3.Range("Ver.Place").Value
    End If
   
    UpdateProgressBar
    If Trim(Sheet3.Range("Ver.Date").Value) = "" Or Trim((Sheet3.Range("Ver.Date").Value)) = "00/00/0000" Then
        EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Date in Verification in sheet Tax Paid and Verification is  mandatory " & Chr(13)
    Else
        If Len(Range("Ver.Date").Value) > 0 Then
            If Not FormatNCheckDate(Range("Ver.Date").Value) Then
                EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Date in Verification in Sheet Tax Paid and Verification  must be a valid dd/mm/yyyy format" & Chr(13)
            End If
            
            'Changing year from 2023 to 2024
            'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
            
          'Ankita_29/01/2026========
          Dim cuttoff13 As Date
          cuttoff13 = CDate(Sheet5.Range("DOB_extended").Value)
'            If EfilingCommon.checkFirstDateBefore(Range("Ver.Date").Value, "31/03/2025") Then
             If EfilingCommon.checkFirstDateBefore(Range("Ver.Date").Value, Sheet5.Range("Date_1").Value) Then
             EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Date in Verification in Sheet Tax Paid and Verification should be on or after " & Dformat1(cuttoff13, "dd/mm/yyyy") & ".""" & Chr(13)
            End If
        End If
        VerificationDate = Sheet3.Range("Ver.Date").Value
    End If
    
    UpdateProgressBar
    If Trim(Sheet3.Range("Ver.PAN").Value) = "" Or IsEmpty(Sheet3.Range("Ver.PAN").Value) Then
       'EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* PAN in Verification in sheet Tax Paid and Verification is  mandatory " & Chr(13)
    'SIT-66374 change'Malli
    EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* PAN is mandatory in Verification " & Chr(13)
    
    Else
       If Not mIncmDtls.CheckPAN(Sheet3.Range("Ver.PAN").Value) Then
             EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "*Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)         'Changed by Ankita on 16/12/2024
       Else
           VerificationPAN = Sheet3.Range("Ver.PAN").Value
       End If
       
'       If calculateAge(Trim(Sheet1.Range("sheet1.DOB").Value)) < 18 Then
'                  If Sheet3.Range("Ver.PAN").Value = Sheet1.Range("Sheet1.pan").Value Then
'                    EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* If date of birth is on or after 2 April 2002 then in verification PAN cannot be same as mentioned in general information." & Chr(13)
'                    Sheet3.Range("Ver.PAN").Value = ""
'                  End If
'        End If
    
    End If
     UpdateProgressBar
    If Trim(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value) <> "" Then
       If Mid(Range("Sheet2.IdentificationNoOfTRP").Value, 1, 1) = "T" Then
           If Not IsNumeric(Mid(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value, 2)) Or (Len(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value) <> 10) Then
                  EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Identification No Of TRP in sheet Tax Paid and Verification must be 10 character with T followed by 9 digits" & Chr(13)
         End If
    
    ElseIf IsNumeric(Mid(Range("Sheet2.IdentificationNoOfTRP").Value, 1, 1)) Then
        If Not IsNumeric(Mid(Range("Sheet2.IdentificationNoOfTRP").Value, 2)) Or (Len(Range("Sheet2.IdentificationNoOfTRP").Value) <> 6) Then
            EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Identification No Of TRP in sheet Tax Paid and Verification must be 6 digits" & Chr(13)
        End If
    Else
        EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Identification No Of TRP in sheet Tax Paid and Verification must begin with T or a digit" & Chr(13)
    End If
    End If
        
        
    UpdateProgressBar
    If Trim(Sheet3.Range("Sheet2.ReImbFrmGov").Value) <> "" Then
       If Trim(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value) = "" Or Trim(Sheet3.Range("Sheet2.NameOfTRP").Value) = "" Then
            EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Identification No Of TRP & Name of TRP in sheet Tax Paid and Verification mandatory" & Chr(13)
        End If
    End If
    

    If Trim(Sheet3.Range("Sheet2.NameOfTRP").Value) <> "" And Trim(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value) = "" Then
       EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Identification No Of TRP in sheet Tax Paid and Verification mandatory" & Chr(13)
    End If
    
    
      If Trim(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value) <> "" And Trim(Sheet3.Range("Sheet2.NameOfTRP").Value) = "" Then
           EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Name Of TRP in sheet Tax Paid and Verification mandatory" & Chr(13)
     End If
     
'    If Trim(Sheet3.Range("Sheet2.ReImbFrmGov").Value) = "" Then
'     If Trim(Sheet3.Range("Sheet2.NameOfTRP").Value) <> "" And Trim(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value) <> "" Then
'    EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "ReImbFrmGov is Mandatory."
'      End If
'      End If
     
        
    UpdateProgressBar
    If Len(Range("IncD.AdvanceTax").Value) > 14 Then EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Total Advance Tax in sheet Tax Paid and Verification should not be greater than 14 digits" & Chr(13)
    If Len(Range("IncD.TDS").Value) > 14 Then EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Total TDS Claimed in sheet Tax Paid and Verification should not be greater than 14 digits" & Chr(13)
    If Len(Range("IncD.TCS").Value) > 14 Then EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Total TCS Claimed in sheet Tax Paid and Verification should not be greater than 14 digits" & Chr(13)
    If Len(Range("IncD.SelfAssessmentTax").Value) > 14 Then EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Total SelfAssessment Tax in sheet Tax Paid and Verification should not be greater than 14 digits" & Chr(13)
    If Len(Range("IncD.TotalTaxesPaid").Value) > 14 Then EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Total Tax Paid in sheet Tax Paid and Verification should not be greater than 14 digits" & Chr(13)
    
    If (Range("ExcempIncome").Value) > 5000 Then EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* If agricultural income is more than Rs 5000/- then use ITR 2 or 3" & Chr(13)

    If (Range("EI.Sec10_34").Value) > 1000000 Then EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Sec 10(34) can not be more than Rs. 10,00,000/-" & Chr(13)
        

    
        
    If EfilingCommon.MsgPartBSheet = "" Then
    Else
        'MsgBox EfilingCommon.MsgPartBSheet, vbOKOnly, vbMessgaeCaption
        fmsgbox (EfilingCommon.MsgPartBSheet)
        Sheet3.Activate
        EfilingCommon.MsgAuditSheet = ""
        CloseMsg
    End If
        
End Sub

Function ValidateOthersEI() As Boolean
ValidateOthersEI = False
    
    If Len((Range("Others.NOI").item(1).Value) > 0) Then
        If Not ValidateNatureOfIncome Then ValidateOthersEI = False
        If Not ValidateSubCategory Then ValidateOthersEI = False
        If Not ValidateAmount Then ValidateOthersEI = False
    End If
     
    setTblinfo_OthersNOI
    setTblinfo_OthersSub
    setTblinfo_OthersAmt

'Ankita_06/03/2026=========
    If ((end_OthersNOI <> end_OthersAmt)) Then
         EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select category from the drop down in Exempt Income in Sheet Income Details." & Chr(13)
  End If

End Function

Sub setTblinfo_OthersNOI()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Others.NOI").count
    Set rangecells = Range("Others.NOI").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "" Or rangecells.item(mIntCtr).Value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_OthersNOI = ccount
 End Sub
 Sub setTblinfo_OthersSub()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Nature_Others").count
    Set rangecells = Range("Nature_Others").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "" Or rangecells.item(mIntCtr).Value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_OthersSub = ccount
 End Sub
 Sub setTblinfo_OthersAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Others.Amount").count
    Set rangecells = Range("Others.Amount").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_OthersAmt = ccount
End Sub
Function ValidateNatureOfIncome() As Boolean
ValidateNatureOfIncome = False

    setTblinfo_OthersNOI
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim cellrange As String
    Dim cellRange1 As String
    Dim cellRange2 As String
    
    Set rangecells = Sheet1.Range("Others.NOI").Cells
    Set rangecells1 = Sheet1.Range("Nature_Others").Cells
    Set rangecells2 = Sheet1.Range("Others.Amount").Cells
    Dim i As Long
    ReDim Others_NOI(end_OthersNOI)
    ReDim Others_NOI1(end_OthersNOI)
    ReDim Others_Amt(end_OthersNOI)
    For i = 1 To end_OthersNOI
        
        cellrange = GetMergedAddressCell(rangecells, i)
        cellRange1 = GetMergedAddressCell(rangecells1, i)
        cellRange2 = GetMergedAddressCell(rangecells2, i)
       
        Others_NOI(i) = Sheet1.Range(cellrange).Value
        Others_NOI1(i) = Sheet1.Range(cellRange1).Value
        Others_Amt(i) = Sheet1.Range(cellRange2).Value

    
'    If Sheet2.Range("TDS.IncSum").Value > 0 Then
'            If IIf(Sheet1.Range("IncD.TotalHeadSalaries").Value = "", 0, Sheet1.Range("IncD.TotalHeadSalaries").Value) < (Sheet2.Range("TDS.IncSum").Value - Round(Sheet2.Range("TDS.IncSum").Value * 0.1, 0)) Then
'             '  IIf(Sheet1.Range("IncD.IncomeFromSal").Value = "", 0, Sheet1.Range("IncD.IncomeFromSanl").Value) > (Sheet2.Range("TDS.IncSum").Value + Round(Sheet2.Range("TDS.IncSum").Value * 0.1, 0)) Then
'            If Not end_OthersNOI > 0 Then
'               EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "Since the amount disclosed in Income chargeable under the Head Salaries is less than 90% of Salary reported in TDS1, please ensure to fill details in " & Chr(34) & "Others" & Chr(34) & " in " & """Exempt Income""" & Chr(13)
'               ValidateNatureOfIncome = False
'             Exit Function
'           End If
'        End If
'    End If
    
'    If Len(Sheet1.Range("Total_Exempt").Value) > 14 Then
'             msgError = msgError & "* Total of Excempt Should not be exceed 14 digits in Sheet Taxes paid and verification" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'    End If

   '==============================
           'Ankita_06/03/2026====
'         If (Others_NOI(i) = "(Select)" Or Others_NOI(i) = "") Then
'             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Nature of Income at Sr. No  " & i & "  in Sheet Income Details  is mandatory" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'         End If
            
'         If Not chkCompulsory(Others_NOI(i)) Then
'             EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "Others: Nature of Income at Sr. No  " & i & "  in Sheet Taxes Paid and Verification  is mandatory" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'         End If
         
'         If (Others_NOI(i) = "Any Other") Then
'         If Others_NOI1(i) = "" Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) at Sr. No  " & i & "  in Sheet Income Details " & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'         End If
'         End If
'
'
'         If Len(Others_NOI1(i)) > 125 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Description at Sr. No  " & i & "  in Sheet Income Details cannot exceed 125 characters" & Chr(13)
'             ValidateNatureOfIncome = False
'             Exit Function
'         End If

   'Ankita_06/03/2026=============
    If (Others_NOI(i) = "(Select)" Or Others_NOI(i) = "") Then
             EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Please select category from the drop down in Exempt Income" & Chr(13)
             ValidateNatureOfIncome = False
             Exit Function
    End If

    If Others_NOI1(i) = "" Then
             EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* Please select sub-category from the drop down in Exempt Income" & Chr(13)
             ValidateNatureOfIncome = False
             Exit Function
    End If

'    If Mid(Others_NOI(i), 1, 4) = "Agri" Then  'Error message updated as per v0.7
'        If Others_Amt(i) > 5000 Then
'            EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* If sec 10(1) agricultural income is more than Rs 5000/- then use ITR 3 or 5" & Chr(13)
'            ValidateNatureOfIncome = False
'            Exit Function
'        End If
'    End If

If Mid(Others_NOI1(i), 1, 5) = "10(1)" Then
        If Others_Amt(i) > 5000 Then
            EfilingCommon.MsgPartBSheet = EfilingCommon.MsgPartBSheet + "* If sec 10(1) agricultural income is more than Rs 5000/- then use ITR 3 or 5" & Chr(13)
            ValidateNatureOfIncome = False
            Exit Function
        End If
    End If
    
    Next
End Function
Function ValidateAmount() As Boolean
ValidateAmount = False
    setTblinfo_OthersNOI
    Dim cellrange As String
    Dim cellRange1 As String
    Dim cellRange2 As String
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    
    Set rangecells = Range("Others.Amount").Cells
    Set rangecells1 = Range("Nature_Others").Cells
    Set rangecells2 = Range("Sheet5.DescEI").Cells
    
    Dim i As Long
    ReDim Others_Amt(end_OthersNOI)
    ReDim Others_NOI1(end_OthersNOI)
    ReDim Others_NOI22(end_OthersNOI)
    For i = 1 To end_OthersNOI

       cellrange = GetMergedAddressCell(rangecells, i)
       cellRange1 = GetMergedAddressCell(rangecells1, i)
       cellRange2 = GetMergedAddressCell(rangecells2, i)
        Others_Amt(i) = Sheet1.Range(cellrange).Value
        Others_NOI1(i) = Sheet1.Range(cellRange1).Value
        Others_NOI22(i) = Sheet1.Range(cellRange2).Value
        
      'Ankita_06/03/2026=========
        If Not chkCompulsory(Others_Amt(i)) Then
'           EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Exempt income (For reporting Purposes) - Please enter Amount at Sr. No  " & i & "  in Sheet Income Details is mandatory" & Chr(13)
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter Amount of exempt income at Sr. No  " & i & "  in Sheet Income Details is mandatory" & Chr(13)
            ValidateAmount = False
            Exit Function
        End If
        If Not IsNumeric(Others_Amt(i)) Then
           EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Others Amount at Sr. No  " & i & "  in Sheet Income Details should be Numeric value" & Chr(13)
            ValidateAmount = False
            Exit Function
        End If
        If Others_Amt(i) > 99999999999999# Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Others Amount at Sr. No  " & i & "  in Sheet Income Details cannot exceed 14 digits" & Chr(13)
            ValidateAmount = False
            Exit Function
        End If
          If Others_NOI1(i) = "Income exempt as per CBDT Circular" Then
             If Others_NOI22(i) = "" Then
                EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + """Please enter the relevant Circular for sub-category “Income exempt as per CBDT Circular” at Sr. No " & i & " in Exempt Income""" & Chr(13)
                ValidateAmount = False
             End If
        End If
          If Others_NOI1(i) = "Income exempt as per CBDT Notification" Then
             If Others_NOI22(i) = "" Then
                EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + """Please enter the relevant Notification No. for other exempt income for sub-category “Income exempt as per CBDT Notification” at Sr. No " & i & " in Exempt Income & Chr(13)"
                ValidateAmount = False
             End If
        End If
        If Others_NOI1(i) = "Receipts not in the nature of income" Then
               If Others_NOI22(i) = "" Then
                  EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + """Please enter the relevant Act/Section reference for sub-category “Receipts not in the nature of income” at Sr. No " & i & " in Exempt Income""" & Chr(13)
                  ValidateAmount = False
               End If
        End If
       
    Next
End Function
  
Function ValidateSubCategory() As Boolean
ValidateSubCategory = False

    setTblinfo_OthersNOI
    Dim cellrange As String
    Dim rangecells As Range
    Set rangecells = Range("Nature_Others").Cells
    Dim i As Long
    ReDim Others_subcategory(end_OthersNOI)
    For i = 1 To end_OthersNOI
    
        cellrange = GetMergedAddressCell(rangecells, i)
        Others_subcategory(i) = Sheet1.Range(cellrange).Value
        
      'Ankita_06/03/2026=========
      If (Others_subcategory(i) = "(Select)" Or Others_subcategory(i) = "") Then
 '       If Not chkCompulsory(Others_subcategory(i)) Then
'           EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Exempt income (For reporting Purposes) - Please enter Amount at Sr. No  " & i & "  in Sheet Income Details is mandatory" & Chr(13)
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select sub-category from the drop down in Exempt Income at Sr. No " & i & " in Sheet Income Details" & Chr(13)
            ValidateSubCategory = False
            Exit Function
        End If
        
        
'        If Not IsNumeric(Others_subcategory(i)) Then
'           EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* dferouuvtrt Sr. No  " & i & "  in Sheet Income Details should be Numeric value" & Chr(13)
'            ValidateSubCategory = False
'            Exit Function
'        End If
'
'        If Others_subcategory(i) > 99999999999999# Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* edsbfhhbdsbfa at Sr. No  " & i & "  in Sheet Income Details cannot exceed 14 digits" & Chr(13)
'            ValidateSubCategory = False
'            Exit Function
'        End If

    If Mid(Others_subcategory(i), 1, 5) = "10(1)" Then
        If Others_Amt(i) > 5000 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*If sec 10(1) agricultural income is more than Rs 5000/- then use ITR 3 or 5" & Chr(13)
            ValidateSubCategory = False
            Exit Function
        End If
    End If
 
'Ankita_07/05/2026=================

        If Sheet5.Range("bacvalue") = 1 Then
         If Mid(Others_subcategory(i), 1, 6) = "10(32)" Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Sub category ""10(32)-Minor child's income—small exemption"" is not applicable for new tax regime" & Chr(13)
            Others_subcategory(i) = "(Select)"
            ValidateSub_category = False
            Exit Function
           End If
        End If
          Next
End Function

Sub ValidateSchTaxPaid_Verify_Click()
    Dim vbMessgaeCaption As String
    EfilingCommon.MsgPartBSheet = ""
    vbMessgaeCaption = "ITR 1: AY: 2026-27"         'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
    ValidateBA
    ValidatePartBTI_BTTI_Verification
   
    If EfilingCommon.MsgPartBSheet = "" Then
        'MsgBox "Tax paid and Verification Sheet is OK", vbOKOnly, vbMessgaeCaption
        fmsgboxoK ("Tax paid and Verification Sheet is OK")
   '     FMsgBox "Tax paid and Verification Sheet is OK"
        Sheet3.Activate
    End If
End Sub
Private Sub CommandButton1_Click()

printWorkSheet
End Sub

 Sub HelpTPV_Click()
 sPassword = EfilingCommon.getmsgstate
ActiveWorkbook.Unprotect Password:=sPassword
Sheet6.Activate
Sheet6.Visible = xlSheetVisible
ActiveWorkbook.Protect Password:=sPassword
End Sub

Sub NextTPV_Click()
'Ankita 11/11/2024
If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
Sheet202.Activate

'Ankita_29/01/2026========
'Else:
''-------------------------
'
'If Sheet5.Range("BacValue").Value = 2 Then
'Sheet9.Activate
'ElseIf Sheet16.Visible = True Then
'Sheet16.Activate
'Else
'MsgBox "can not go to next sheet as you have selected 115BAC = Y "
'End If
End If 'Ankita 11/11/2024
End Sub

 Sub PrevTPV_Click()

Dim a As Worksheet
Set a = ThisWorkbook.Sheets("TCS")
a.Activate


End Sub

Function checkfieldspecialcharacter_Bank(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_Bank = True
    Dim arr As Variant
    arr = Array("@", "*", "!", ".", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_Bank = False
            Exit Function
        End If
        Next
    Next
End Function

