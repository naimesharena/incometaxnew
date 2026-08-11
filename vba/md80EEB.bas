Attribute VB_Name = "md80EEB"
Public MsgBox_80EEB As Variant
Dim end_80EEB As Variant
Dim end_80EEBLoanfrm As Variant
Dim end_80EEBIFSC As Variant
Dim end_80EEBbankName As Variant
Dim end_80EEBPAN As Variant
Dim end_80EEBAccntNum As Variant
Dim end_80EEBLoanDate As Variant
Dim end_80EEBLoanAmt As Variant
Dim end_80EEBLoanOutstanding As Variant
Dim end_80EEBVehicleValue As Variant
Dim end_80EEBVehicleReg As Variant
Dim end_80EEBIntrst As Variant

Sub ValidateSheet80EEB_Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 1: AY: 2026-27"
Validate80EEB_All

'fmsgboxStatus "Schedule 80EE is OK"
fmsgboxStatus "Schedule 80EEB is OK"
End Sub
Sub Validate80EEB_All()
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80E_80EE_80EEA_80EEB")
If Not Validate_80EEB Then
    sourceSheet.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    fmsgboxStatus MsgBox_80EEB
    CloseMsg
End If
End Sub

Function Validate_80EEB()
Validate_80EEB = True
MsgBox_80EEB = ""
end_80EEB = 0

setTblinfo_80EEBLoanfrm
'setTblinfo_80EEBIFSC
'setTblinfo_80EEBPAN
setTblinfo_80EEBBankName
setTblinfo_80EEBAccntNum
setTblinfo_80EEBLoanDate
setTblinfo_80EEBLoanAmt
setTblinfo_80EEBLoanOutstanding
'setTblinfo_80EEBVehicleValue
setTblinfo_80EEBVehicleReg
setTblinfo_80EEBIntrst

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'end_80EEB = WorksheetFunction.Max(0, end_80EEB, end_80EEBLoanfrm, end_80EEBbankName, end_80EEBAccntNum, end_80EEBLoanDate, end_80EEBLoanAmt, end_80EEBLoanOutstanding, end_80EEBVehicleValue, end_80EEBVehicleReg, end_80EEBIntrst)
end_80EEB = WorksheetFunction.Max(0, end_80EEBLoanfrm, end_80EEBbankName, end_80EEBAccntNum, end_80EEBLoanDate, end_80EEBLoanAmt, end_80EEBLoanOutstanding, end_80EEBVehicleReg, end_80EEBIntrst)

If Not ValidateLoanfrm_80EEB Then Validate_80EEB = False
'If Not ValidateIFSC_80EEB Then Validate_80EEB = False
'If Not ValidatePAN_80EEB Then Validate_80EEB = False
If Not ValidateBankName_80EEB Then Validate_80EEB = False
If Not ValidateAccntNum_80EEB Then Validate_80EEB = False
If Not ValidateLoanDate_80EEB Then Validate_80EEB = False
If Not ValidateLoanAmt_80EEB Then Validate_80EEB = False

'    If (Sheet20.Range("MaxLoan_80EEB").value) > 3500000 Then
'        MsgBox_80EEB = MsgBox_80EEB & "* ""Total loan taken"" shall not be more than Rs. 35 Lakhs in schedule 80EE" & Chr(13)
'        Validate_80EEB = False
'    End If

If Not ValidateLoanOutstanding_80EEB Then Validate_80EEB = False
'If Not ValidateVehicleValue_80EEB Then Validate_80EEB = False
If Not ValidateVehicleReg_80EEB Then Validate_80EEB = False
If Not ValidateIntrst_80EEB Then Validate_80EEB = False
'If Not ValidateLoanfrm_80EEB Then Validate_80EEB = False

'ValidateIFSC_80EEB

'    If (Sheet20.Range("TotAmt.80EEB").Value) > 2000000 Then
'        MsgBox_80EEB = MsgBox_80EEB & " To claim ""Interest payable on borrowed capital more than Rs.20 lakhs"", you may please consider filing ITR 3. Refer Rule 12 for further details" & Chr(13)
'        Validate_80EEB = False
'    End If
    
    If Len(Sheet17.Range("TotAmt.80EEB").Value) > 14 Then
        MsgBox_80EEB = MsgBox_80EEB & "* Total of Interest u/s 80EEB in schedule 80EEB cannot exceed 14 digits" & Chr(13)
        Validate_80EEB = False
    End If

End Function

Sub setTblinfo_80EEBLoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("LoanfrmBankOrInstitute.80EEB").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "" Or rangecells.item(mIntCtr).Value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBLoanfrm = ccount
End Sub

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Sub setTblinfo_80EEBIFSC()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("IFSC.80EEB").count
'    Set rangecells = Range("IFSC.80EEB").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EEBIFSC = ccount
'End Sub


Sub setTblinfo_80EEBBankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("bankName.80EEB").count
    Set rangecells = Range("bankName.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBbankName = ccount
End Sub


'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Sub setTblinfo_80EEBPAN()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("PAN.80EEB").count
'    Set rangecells = Range("PAN.80EEB").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EEBPAN = ccount
'End Sub

Sub setTblinfo_80EEBAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAccNum.80EEB").count
    Set rangecells = Range("loanAccNum.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBAccntNum = ccount
End Sub
Sub setTblinfo_80EEBLoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanDate.80EEB").count
    Set rangecells = Range("loanDate.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBLoanDate = ccount
End Sub
Sub setTblinfo_80EEBLoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAmt.80EEB").count
    Set rangecells = Range("loanAmt.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBLoanAmt = ccount
End Sub

Sub setTblinfo_80EEBLoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanOutstanding.80EEB").count
    Set rangecells = Range("loanOutstanding.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBLoanOutstanding = ccount
End Sub
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Sub setTblinfo_80EEBVehicleValue()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Vehicle_value.80EEB").count
'    Set rangecells = Range("Vehicle_value.80EEB").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EEBVehicleValue = ccount
'End Sub
Sub setTblinfo_80EEBVehicleReg()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("VehicleRegNum.80EEB").count
    Set rangecells = Range("VehicleRegNum.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBVehicleReg = ccount
End Sub

Sub setTblinfo_80EEBIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Intrst.80EEB").count
    Set rangecells = Range("Intrst.80EEB").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEBIntrst = ccount
End Sub

Function ValidateLoanfrm_80EEB() As Boolean
    ValidateLoanfrm_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Cells
    ReDim Loanfrm_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        Loanfrm_80EEB(i) = rangecells.item(i).Value
'        If Not chkCompulsory(Loanfrm_80EEB(i)) Then
        If isdropdownblank(Loanfrm_80EEB(i)) Then
        'Ankita_25/04/2025
'             MsgBox_80EEB = MsgBox_80EEB + "* Please select dropdown ""Loan taken from"" in schedule 80EEB at Sr. No " & i & "" & Chr(13)
              MsgBox_80EEB = MsgBox_80EEB + "* Please select dropdown from ""Loan taken from"" in schedule 80EEB at Sr. No " & i & "" & Chr(13)  'Changed as per DE sheet v0.8 by Ankita
            '"Please select dropdown  "Loan taken from" in schedule 80EEB
            ValidateLoanfrm_80EEB = False
            Exit Function
        End If
'         If Loanfrm_80EEB(i) = "(Select)" Then
'         'Ankita_25/04/2025
'          MsgBox_80EEB = MsgBox_80EEB + "* Please select dropdown ""Loan taken from"" in schedule 80EEB at Sr. No " & i & "" & Chr(13)
'            ValidateLoanfrm_80EEB = False
'            Exit Function
'        End If
         
Next
End Function

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Function ValidateIFSC_80EEB() As Boolean
'ValidateIFSC_80EEB = True
'
'   ' setTblinfo_80EEBIFSC
'   setTblinfo_80EEBLoanfrm
'    Dim rangecells As Range
''    Dim rangecells1 As Range  'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    Set rangecells = Range("LoanfrmBankOrInstitute.80EEB").Cells
''    Set rangecells1 = Range("IFSC.80EEB").Cells  'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    Dim cellrange As String
'    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80EEBLoanfrm)
'    ReDim Others_IFSC(end_80EEBLoanfrm)
'
'
'    For i = 1 To end_80EEBLoanfrm
'    cellrange = GetMergedAddressCell(rangecells, i)
'    cellRange1 = GetMergedAddressCell(rangecells1, i)
'
'        Others_Loan(i) = Sheet17.Range(cellrange).Value
'        Others_IFSC(i) = Sheet17.Range(cellRange1).Value
'
'         If (Others_Loan(i) = "Bank") Then
'         If Others_IFSC(i) = "" Then
'         'Change.03.03.2023.102.IDS.113
'             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
'             MsgBox_80EEB = MsgBox_80EEB & "* Please provide IFSC of the Bank from which loan is taken in schedule 80EEB at Sr. No " & i & "" & Chr(13)
'         'End Change.IDS113
'             ValidateIFSC_80EEB = False
'             Exit Function
'         End If
'         End If
'
'         If Len(Others_IFSC(i)) > 11 Then
'          MsgBox_80EEB = MsgBox_80EEB + "* IFSC of the Bank at Sr. No " & i & " in schedule 80EEB should be 11 characters." & Chr(13)
'            ValidateIFSC_80EEB = False
'            Exit Function
'        End If
'
'    Next
'
'End Function

'Function ValidatePAN_80EEB() As Boolean
'ValidatePAN_80EEB = True
'
'   ' setTblinfo_80EEBIFSC
'   setTblinfo_80EEBLoanfrm
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Set rangecells = Range("LoanfrmBankOrInstitute.80EEB").Cells
'    Set rangecells1 = Range("PAN.80EEB").Cells
'    Dim cellrange As String
'    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80EEBLoanfrm)
'    ReDim Others_PAN(end_80EEBLoanfrm)
'
'        For i = 1 To end_80EEBLoanfrm
'    cellrange = GetMergedAddressCell(rangecells, i)
'    cellRange1 = GetMergedAddressCell(rangecells1, i)
'
'        Others_Loan(i) = Sheet17.Range(cellrange).Value
'        Others_PAN(i) = Sheet17.Range(cellRange1).Value
'
'
'
'         If (Others_Loan(i) = "Institution") Then
'         If Others_PAN(i) = "" Then
'         'Change.03.03.2023.102.IDS.113
'             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
'             MsgBox_80EEB = MsgBox_80EEB & "* Please provide PAN of the institution from which loan is taken in schedule 80EEB at Sr. No " & i & "" & Chr(13)
'         'End Change.IDS113
'             ValidatePAN_80EEB = False
'             Exit Function
'         End If
'         End If
'
'         If Len(Others_PAN(i)) > 10 Then
'          MsgBox_80EEB = MsgBox_80EEB + "* PAN of the institution at Sr. No " & i & " in schedule 80EEB should be 10 characters." & Chr(13)
'            ValidatePAN_80EEB = False
'            Exit Function
'        End If
'
'        'Added by Malli(24/04/2025)NewDev
'        If Others_PAN(i) <> "" Then
'            If Not EfilingCommon.CheckPAN(UCase(Trim(Others_PAN(i)))) Then
'            MsgBox_80EEB = MsgBox_80EEB + "* Invalid PAN. PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet in schedule 80EEB at Sr. No " & i & "" & Chr(13)
'            ValidatePAN_80EEB = False
'            Exit Function
'          End If
'          End If
'
'    Next
'
'End Function


Function ValidateBankName_80EEB() As Boolean
    ValidateBankName_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("bankName.80EEB").Cells
    ReDim BankName_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        BankName_80EEB(i) = rangecells.item(i).Value
        If Not chkCompulsory(BankName_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Name of the Bank / Institution from which the loan is taken in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateBankName_80EEB = False
            Exit Function
        End If
         If Len(BankName_80EEB(i)) > 125 Then
          MsgBox_80EEB = MsgBox_80EEB + "* Name of the Bank / Institution at Sr. No " & i & " in Sheet 80EEB less than 125 characters." & Chr(13)
            ValidateBankName_80EEB = False
            Exit Function
        End If
         If Not checkfieldSuperSpecialcharacter(BankName_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* Name of the Bank / Institution from which the loan is taken in schedule 80EEB at Sr.no. " & i & " should not Contain <, >, characters." & Chr(13)
            ValidateBankName_80EEB = False
            Exit Function
        End If
Next
End Function

Function ValidateAccntNum_80EEB() As Boolean
    ValidateAccntNum_80EEB = True

    Dim rangecells, rangecells1 As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanAccNum.80EEB").Cells
    Set rangecells1 = Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Cells
    
    ReDim AccntNum_80EEB(end_80EEB)
     ReDim Bankorothebank_80EEB(end_80EEB)
     
    For i = 1 To end_80EEB
        AccntNum_80EEB(i) = rangecells.item(i).Value
        Bankorothebank_80EEB(i) = rangecells1.item(i).Value
         
        If Not chkCompulsory(AccntNum_80EEB(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'           MsgBox_80EEB = MsgBox_80EEB + "* Please provide Loan Account number of the Bank / Reference number of the Institution from which loan is taken in schedule 80EEB at Sr. No " & i & "" & Chr(13)
            MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Loan Account number of the Bank / Institution from which loan is taken in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateAccntNum_80EEB = False
            Exit Function
        End If
         If Len(AccntNum_80EEB(i)) > 20 Then
            MsgBox_80EEB = MsgBox_80EEB + "* Loan Account number in Sheet 80EEB should be less than or equal to 20 characters at Sr. No " & i & "." & Chr(13)
            ValidateAccntNum_80EEB = False
            Exit Function
        End If
        
        If Bankorothebank_80EEB(i) = "Bank" Then
        If Not ValidateBankAccountNumber_80EEB(AccntNum_80EEB(i), i) Then
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntNum_80EEB = False
        Exit Function
    End If
    End If
    
    'Doubt_05/05/2025
    
    If Bankorothebank_80EEB(i) = "Institution" Then
     If Not checkfieldspecialcharacter_BthacntReferencenumber(AccntNum_80EEB(i)) Then
            MsgBox_80EEB = MsgBox_80EEB + "* Loan Account number of the Bank / Institution at Sr.No " & i & " is invalid in schedule 80EEB, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateAccntNum_80EEB = False
            Exit Function
     End If
    End If
         
Next
End Function

Function ValidateVehicleReg_80EEB() As Boolean
    ValidateVehicleReg_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("VehicleRegNum.80EEB").Cells
    ReDim VehicleReg_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        VehicleReg_80EEB(i) = rangecells.item(i).Value
        If Not chkCompulsory(VehicleReg_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please enter vehicle Registration number in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateVehicleReg_80EEB = False
            Exit Function
        End If
         If Len(VehicleReg_80EEB(i)) > 11 Then
          MsgBox_80EEB = MsgBox_80EEB + "*  ""Vehicle Registration number cannot exceed 11 characters in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateVehicleReg_80EEB = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacter(VehicleReg_80EEB(i)) Then
         MsgBox_80EEB = MsgBox_80EEB + "* Vehicle Registration number should be alphanumeic at Sr. No " & i & " in Sheet 80EEB " & Chr(13)
            ValidateVehicleReg_80EEB = False
            Exit Function
        End If
        
         
Next
End Function

Function ValidateLoanDate_80EEB() As Boolean
    ValidateLoanDate_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanDate.80EEB").Cells
    ReDim LoanDate_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        LoanDate_80EEB(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanDate_80EEB(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Date of sanction of loan in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanDate_80EEB = False
            Exit Function
        End If
         If Len(LoanDate_80EEB(i)) > 10 Then
         'Ankita_05/05/2025_Commented as per DESheet_v0.7
          MsgBox_80EEB = MsgBox_80EEB + "* Date of sanction of loan at Sr. No " & i & " in schedule 80EEB less than 10 characters." & Chr(13)
            ValidateLoanDate_80EEB = False
            Exit Function
        End If
        'CheckDateddmmyyyy
         If Not CheckDateddmmyyyy(LoanDate_80EEB(i)) Then
         MsgBox_80EEB = MsgBox_80EEB + "* ""Please enter date in valid format"" in schedule 80EEB at Sr. No " & i & "." & Chr(13)
        ValidateLoanDate_80EEB = False
                 

          Exit Function
        End If
        
        If Not ChkMaxDate_80EEB(Trim(LoanDate_80EEB(i)), "31-03-2023") Then
         'MsgBox_80EEB = MsgBox_80EEB + " Date of taking of Loan shall be between 01/04/2019 to 31/03/2023. at Sr. No " & i & "." & Chr(13)
          MsgBox_80EEB = MsgBox_80EEB + "* ""Date of sanction of Loan shall be between 01/04/2019 to 31/03/2023"" in schedule 80EEB at Sr. No " & i & "." & Chr(13)
         ValidateLoanDate_80EEB = False
          Exit Function
        End If
                 
      If Not ChkMinInclusiveDate(Trim(Dformat(LoanDate_80EEB(i), "yyyy-mm-dd")), "2019-04-01") Then
           'MsgBox_80EEB = MsgBox_80EEB + " Date of taking of Loan shall be between 01/04/2019 to 31/03/2023. at Sr. No " & i & "." & Chr(13)
             MsgBox_80EEB = MsgBox_80EEB + "* ""Date of sanction of Loan shall be between 01/04/2019 to 31/03/2023"" in schedule 80EEB at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EEB = False
                 

          Exit Function
        End If
   
Next
End Function

Function ValidateLoanAmt_80EEB() As Boolean
    ValidateLoanAmt_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanAmt.80EEB").Cells
    ReDim LoanAmt_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        LoanAmt_80EEB(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanAmt_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide ""Total amount of Loan"" in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanAmt_80EEB = False
            Exit Function
        End If
'         If Len(LoanAmt_80EEB(i)) > 14 Then
'          MsgBox_80EEB = MsgBox_80EEB + "* Loan Amount  at Sr. No " & i & " in Sheet 80EEB less than 15 characters." & Chr(13)
'            ValidateLoanAmt_80EEB = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanAmt_80EEB(i)) Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan amount at Sr. No  " & i & " in schedule 80EEB should be Numeric value" & Chr(13)
            ValidateLoanAmt_80EEB = False
            Exit Function
        End If
        
        If LoanAmt_80EEB(i) > 99999999999999# Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan amount at Sr. No  " & i & " in schedule 80EEB cannot exceed 14 digits" & Chr(13)
            ValidateLoanAmt_80EEB = False
            Exit Function
        End If
        
        If LoanAmt_80EEB(i) < 0 Or LoanAmt_80EEB(i) = 0 Then
            MsgBox_80EEB = MsgBox_80EEB & "* Total amount of Loan should be more than 0 in schedule 80EEB at Sr. No  " & i & "" & Chr(13)
            ValidateLoanAmt_80EEB = False
            Exit Function
        End If
         
Next
End Function
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Function ValidateVehicleValue_80EEB() As Boolean
'    ValidateVehicleValue_80EEB = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet17.Range("Vehicle_value.80EEB").Cells
'    ReDim VehicleValue_80EEB(end_80EEB)
'    For i = 1 To end_80EEB
'        VehicleValue_80EEB(i) = rangecells.item(i).Value
'        If Not chkCompulsory(VehicleValue_80EEB(i)) Then
'             MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Vehicle value in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
'            ValidateVehicleValue_80EEB = False
'            Exit Function
'        End If
''         If Len(LoanAmt_80EEB(i)) > 14 Then
''          MsgBox_80EEB = MsgBox_80EEB + "* Loan Amount  at Sr. No " & i & " in Sheet 80EEB less than 15 characters." & Chr(13)
''            ValidateLoanAmt_80EEB = False
''            Exit Function
''        End If
'
'        If Not IsNumeric(VehicleValue_80EEB(i)) Then
'            MsgBox_80EEB = MsgBox_80EEB & "* Vehicle value at Sr. No  " & i & "  in schedule 80EEB should be Numeric value" & Chr(13)
'            ValidateVehicleValue_80EEB = False
'            Exit Function
'        End If
'
'        If VehicleValue_80EEB(i) > 99999999999999# Then
'            MsgBox_80EEB = MsgBox_80EEB & "* Vehicle value at Sr. No  " & i & "  in schedule 80EEB cannot exceed 14 digits" & Chr(13)
'            ValidateVehicleValue_80EEB = False
'            Exit Function
'        End If
'
'        If VehicleValue_80EEB(i) < 0 Or VehicleValue_80EEB(i) = 0 Then
'            MsgBox_80EEB = MsgBox_80EEB & "* Vehicle value should be more than 0 in schedule 80EEB at Sr. No  " & i & "" & Chr(13)
'            ValidateVehicleValue_80EEB = False
'            Exit Function
'        End If
'
'Next
'End Function

Function ValidateLoanOutstanding_80EEB() As Boolean
    ValidateLoanOutstanding_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanOutstanding.80EEB").Cells
    ReDim LoanOutStanding_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        LoanOutStanding_80EEB(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanOutStanding_80EEB(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
             MsgBox_80EEB = MsgBox_80EEB + "* ""Loan outstanding as on last date of financial year is mandatory in schedule 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanOutstanding_80EEB = False
            Exit Function
        End If
'         If Len(LoanOutstanding_80EEB(i)) > 14 Then
'          MsgBox_80EEB = MsgBox_80EEB + "* Loan Outstanding  at Sr. No " & i & " in Sheet 80EEB less than 15 characters." & Chr(13)
'            ValidateLoanOutstanding_80EEB = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanOutStanding_80EEB(i)) Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan outstanding at Sr. No  " & i & " in schedule 80EEB should be Numeric value" & Chr(13)
            ValidateLoanOutstanding_80EEB = False
            Exit Function
        End If
        
        If LoanOutStanding_80EEB(i) > 99999999999999# Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan outstanding at Sr. No  " & i & " in schedule 80EEB cannot exceed 14 digits" & Chr(13)
            ValidateLoanOutstanding_80EEB = False
            Exit Function
        End If
        
        If LoanOutStanding_80EEB(i) < 0 Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan outstanding as on last date of financial year can't be less than 0 in schedule 80EEB at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
            ValidateLoanOutstanding_80EEB = False
            Exit Function
        End If
         
Next
End Function

Function ValidateIntrst_80EEB() As Boolean
    ValidateIntrst_80EEB = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("Intrst.80EEB").Cells
    ReDim Intrst_80EEB(end_80EEB)
    For i = 1 To end_80EEB
        Intrst_80EEB(i) = rangecells.item(i).Value
        If Not chkCompulsory(Intrst_80EEB(i)) Then
             MsgBox_80EEB = MsgBox_80EEB + "* ""Please provide Interest u/s 80EEB"" at Sr. No " & i & "" & Chr(13)
            ValidateIntrst_80EEB = False
            Exit Function
        End If
'         If Len(Intrst_80EEB(i)) > 15 Then
'          MsgBox_80EEB = MsgBox_80EEB + "* Interest  at Sr. No " & i & " in Sheet 80EEB less than 15 characters." & Chr(13)
'            ValidateIntrst_80EEB = False
'            Exit Function
'        End If

        If Not IsNumeric(Intrst_80EEB(i)) Then
            MsgBox_80EEB = MsgBox_80EEB & "* Interest at Sr. No  " & i & " in schedule 80EEB should be Numeric value" & Chr(13)
            ValidateIntrst_80EEB = False
            Exit Function
        End If
        
        If Intrst_80EEB(i) > 99999999999999# Then
            MsgBox_80EEB = MsgBox_80EEB & "* Interest at Sr. No  " & i & " in schedule 80EEB cannot exceed 14 digits" & Chr(13)
            ValidateIntrst_80EEB = False
            Exit Function
        End If
        
        'Interest u/s 80EEB should be more than 0 in schedule 80EEB
        
        If Intrst_80EEB(i) < 0 Or Intrst_80EEB(i) = 0 Then
            MsgBox_80EEB = MsgBox_80EEB & "* Interest u/s 80EEB should be more than 0 in schedule 80EEB at Sr. No  " & i & "" & Chr(13)
            ValidateIntrst_80EEB = False
            Exit Function
        End If
         
Next
End Function


Function ValidateBankAccountNumber_80EEB(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_80EEB = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB(), ValidateIFSC As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
    
   ' BankAccountNumber = Sheet5.Range("IncD.BankAccountNumber")

           
    If Len(BankAccountNumber) > 0 Then
        'If Not checkfieldspecialcharacter1(BankAccountNumber) Then
        'Malli
        If Not checkfieldspecialcharacter80E_EE_EEA_EEB(BankAccountNumber) Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80EEB = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB " & Chr(13)
            ValidateBankAccountNumber_80EEB = False
            Exit Function
        End If
       'Malli
'       If IsNumeric(BankAccountNumber) Then
'            MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEB " & Chr(13)
'            ValidateBankAccountNumber_80EEB = False
'            Exit Function
'        End If
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80EEB = MsgBox_80EEB & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is mandatory in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEB" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If

    ReDim myB(Len(BankAccountNumber) - 1)
    For i = 1 To Len(BankAccountNumber)
        myB(i - 1) = Mid(BankAccountNumber, i, 1)
    Next

    For i = LBound(myB) To UBound(myB)
        If IsNumeric(myB(i)) Then
            countnum = countnum + 1
        End If

        If i > LBound(myB) And i < UBound(myB) Then
            If myB(i) = 0 Then
                If myB(i - 1) = 0 Then
                    zeroCount = zeroCount + 1
                    AfterZero = IIf(Not IsNumeric(myB(i + 1)), myB(i + 1), "")
                Else
                    BeforeZero = IIf(Not IsNumeric(myB(i - 1)), myB(i - 1), "")
                End If
            End If
        End If

    Next

    If BeforeZero <> "" And AfterZero <> "" Then
        If zeroCount > 1 Then
            MsgBox_80EEB = MsgBox_80EEB & "* Loan Account number at Sr.No " & cc + 1 & " is invalid in schedule 80EEB" & Chr(13)
            ValidateBankAccountNumber_80EEB = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc + 1 & "  in schedule 80EEB is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80EEB = False
        Exit Function
    End If
End Function


Public Function ChkMaxDate_80EEB(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_80EEB = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ankita_25-26
'        If Year > 2024 Then
        If Year > 2023 Then
            ChkMaxDate_80EEB = False
            Exit Function
        Else
'            If Year = 2024 Then
                If Year = 2023 Then
                If month > 4 Then
                    ChkMaxDate_80EEB = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_80EEB = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_80EEB = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function




