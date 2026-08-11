Attribute VB_Name = "md80E"
Public MsgBox_80E As Variant
Dim end_80E As Variant
Dim end_80ELoanfrm As Variant
Dim end_80EIFSC As Variant
Dim end_80EbankName As Variant
Dim end_80EPAN As Variant
Dim end_80EAccntNum As Variant
Dim end_80ELoanDate As Variant
Dim end_80ELoanAmt As Variant
Dim end_80ELoanOutstanding As Variant
Dim end_80EIntrst As Variant


Sub ValidateSheet80E_Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 1: AY: 2026-27"
Validate80E_All
fmsgboxStatus "Schedule 80E is OK"
End Sub

Sub Validate80E_All()
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80E_80EE_80EEA_80EEB")
If Not Validate_80E Then
    sourceSheet.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    fmsgboxStatus MsgBox_80E
    CloseMsg
End If
End Sub

Function Validate_80E()
Validate_80E = True
end_80E = 0

setTblinfo_80ELoanfrm
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'setTblinfo_80EIFSC
'setTblinfo_80EPAN
setTblinfo_80EBankName
setTblinfo_80EAccntNum
setTblinfo_80ELoanDate
setTblinfo_80ELoanAmt
setTblinfo_80ELoanOutstanding
setTblinfo_80EIntrst

'
'end_80E = WorksheetFunction.Max(0, end_80E, end_80ELoanfrm, end_80EbankName, , end_80EAccntNum, end_80ELoanDate, end_80ELoanAmt, end_80ELoanOutstanding, end_80EIntrst)

'Malli------22/04/2025
'end_80E = WorksheetFunction.Max(0, end_80ELoanfrm, end_80EPAN, end_80EbankName, end_80EAccntNum, end_80ELoanDate, end_80ELoanAmt, end_80ELoanOutstanding, end_80EIntrst)
end_80E = WorksheetFunction.Max(0, end_80ELoanfrm, end_80EbankName, end_80EAccntNum, end_80ELoanDate, end_80ELoanAmt, end_80ELoanOutstanding, end_80EIntrst)

If Not ValidateLoanfrm_80E Then Validate_80E = False
'If Not ValidateIFSC_80E Then Validate_80E = False
'If Not ValidatePAN_80E Then Validate_80E = False
If Not ValidateBankName_80E Then Validate_80E = False
If Not ValidateAccntNum_80E Then Validate_80E = False
If Not ValidateLoanDate_80E Then Validate_80E = False
If Not ValidateLoanAmt_80E Then Validate_80E = False
If Not ValidateLoanOutstanding_80E Then Validate_80E = False
If Not ValidateIntrst_80E Then Validate_80E = False
'If Not ValidateLoanfrm_80E Then Validate_80E = False

'ValidateIFSC_80E

'If (Sheet17.Range("TotAmt.80E").Value) > 2000000 Then
'        MsgBox_80E = MsgBox_80E & " To claim ""Interest payable on borrowed capital more than Rs.20 lakhs"", you may please consider filing ITR 3. Refer Rule 12 for further details" & Chr(13)
'        Validate_80E = False
'    End If
If Len(Sheet17.Range("TotAmt.80E").Value) > 14 Then
        MsgBox_80E = MsgBox_80E & "* Total of Interest u/s 80E in schedule 80E cannot exceed 14 digits" & Chr(13)
        Validate_80E = False
    End If


End Function

Sub setTblinfo_80ELoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("LoanfrmBankOrInstitute.80E").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "" Or rangecells.item(mIntCtr).Value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    end_80ELoanfrm = ccount
End Sub
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Sub setTblinfo_80EIFSC()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("IFSC.80E").count
'    Set rangecells = Range("IFSC.80E").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EIFSC = ccount
'End Sub


Sub setTblinfo_80EBankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("bankName.80E").count
    Set rangecells = Range("bankName.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EbankName = ccount
End Sub

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Sub setTblinfo_80EPAN()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("PAN.80E").count
'    Set rangecells = Range("PAN.80E").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EPAN = ccount
'End Sub

Sub setTblinfo_80EAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAccNum.80E").count
    Set rangecells = Range("loanAccNum.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EAccntNum = ccount
End Sub
Sub setTblinfo_80ELoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanDate.80E").count
    Set rangecells = Range("loanDate.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80ELoanDate = ccount
End Sub
Sub setTblinfo_80ELoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAmt.80E").count
    Set rangecells = Range("loanAmt.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80ELoanAmt = ccount
End Sub

Sub setTblinfo_80ELoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanOutstanding.80E").count
    Set rangecells = Range("loanOutstanding.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80ELoanOutstanding = ccount
End Sub
Sub setTblinfo_80EIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Intrst.80E").count
    Set rangecells = Range("Intrst.80E").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EIntrst = ccount
End Sub



Function ValidateLoanfrm_80E() As Boolean
    ValidateLoanfrm_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("LoanfrmBankOrInstitute.80E").Cells
    ReDim Loanfrm_80E(end_80E)
    For i = 1 To end_80E
        Loanfrm_80E(i) = rangecells.item(i).Value
         'Ankita_12/05/2025
'        If Not chkCompulsory(Loanfrm_80E(i)) Then
        If isdropdownblank(Loanfrm_80E(i)) Then
             MsgBox_80E = MsgBox_80E + "* Please select dropdown ""Loan taken from"" in schedule 80E at Sr. No " & i & "" & Chr(13)
            ValidateLoanfrm_80E = False
            Exit Function
        End If
'         If Loanfrm_80E(i) = "(Select)" Then
'         'Ankita_06/05/2025_Commented as per DESheet_v0.7
'          MsgBox_80E = MsgBox_80E + "* Please select dropdown from ""Loan taken from"" in schedule 80E at Sr. No " & i & "" & Chr(13)
'            ValidateLoanfrm_80E = False
'            Exit Function
'        End If
         
Next
End Function

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Function ValidateIFSC_80E() As Boolean
'ValidateIFSC_80E = True
'
'   ' setTblinfo_80EIFSC
'   setTblinfo_80ELoanfrm
'    Dim rangecells As Range
''    Dim rangecells1 As Range
'    Set rangecells = Range("LoanfrmBankOrInstitute.80E").Cells
''    Set rangecells1 = Range("IFSC.80E").Cells
'    Dim cellrange As String
''    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80ELoanfrm)
''    ReDim Others_IFSC(end_80ELoanfrm)
'
'    For i = 1 To end_80ELoanfrm
'    cellrange = GetMergedAddressCell(rangecells, i)
'    cellRange1 = GetMergedAddressCell(rangecells1, i)
'
'        Others_Loan(i) = Sheet17.Range(cellrange).Value
'        Others_IFSC(i) = Sheet17.Range(cellRange1).Value
'
'
'
'         If (Others_Loan(i) = "Bank") Then
'         If Others_IFSC(i) = "" Then
'         'Change.03.03.2023.102.IDS.113
'             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
'             MsgBox_80E = MsgBox_80E & "* Please provide  IFSC of the Bank from which loan is taken in schedule 80E at Sr. No " & i & "" & Chr(13)
'         'End Change.IDS113
'             ValidateIFSC_80E = False
'             Exit Function
'         End If
'         End If
'
'         If Len(Others_IFSC(i)) > 11 Then
'          MsgBox_80E = MsgBox_80E + "* IFSC of the Bank at Sr. No " & i & " in Sheet 80E should be 11 characters." & Chr(13)
'            ValidateIFSC_80E = False
'            Exit Function
'        End If
'
'    Next
'
'End Function

'Function ValidatePAN_80E() As Boolean
'ValidatePAN_80E = True
'
'   ' setTblinfo_80EIFSC
'   setTblinfo_80ELoanfrm
'    Dim rangecells As Range
''    Dim rangecells1 As Range
'    Set rangecells = Range("LoanfrmBankOrInstitute.80E").Cells
'    Set rangecells1 = Range("PAN.80E").Cells
'    Dim cellrange As String
'    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80ELoanfrm)
'    ReDim Others_PAN(end_80ELoanfrm)
'
'
'    For i = 1 To end_80ELoanfrm
'    cellrange = GetMergedAddressCell(rangecells, i)
'    cellRange1 = GetMergedAddressCell(rangecells1, i)
'
'        Others_Loan(i) = Sheet17.Range(cellrange).Value
'        Others_PAN(i) = Sheet17.Range(cellRange1).Value
'
'
'         'malli------22/04/2025
'         'If (Others_Loan(i) = "Institution") Then
''         If (Others_Loan(i) = "Instituition") Then
'         '------------------------------
'
'         'Ankita_29/04/2025
'         If (Others_Loan(i) = "Institution") Then
'         If Others_PAN(i) = "" Then
'         'Change.03.03.2023.102.IDS.113
'             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
''             MsgBox_80E = MsgBox_80E & "* Please provide  PAN of the institution/person from which loan is taken in schedule 80E at Sr. No " & i & "" & Chr(13)
'
'             'Malli-
'              MsgBox_80E = MsgBox_80E & " ""Please provide PAN of the institution from which loan is taken in schedule 80E"" at Sr. No " & i & "" & Chr(13)
'              '-----------
'
'         'End Change.IDS113
'             ValidatePAN_80E = False
'             Exit Function
'         End If
'         End If
'
'         If Len(Others_PAN(i)) > 10 Then
'          MsgBox_80E = MsgBox_80E + "* PAN of the institution at Sr. No " & i & " in Sheet 80E should be 10 characters." & Chr(13)
'            ValidatePAN_80E = False
'            Exit Function
'        End If
'
'        'Added by Malli(24/04/2025)NewDev
'        If Others_PAN(i) <> "" Then
'          If Not EfilingCommon.CheckPAN(UCase(Trim(Others_PAN(i)))) Then
'            MsgBox_80E = MsgBox_80E + "* Invalid PAN. PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet in schedule 80E at Sr. No " & i & "" & Chr(13)
'            ValidatePAN_80E = False
'            Exit Function
'        End If
'        End If
'
'    Next
'
'End Function


Function ValidateBankName_80E() As Boolean
    ValidateBankName_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("bankName.80E").Cells
    ReDim BankName_80E(end_80E)
    For i = 1 To end_80E
        BankName_80E(i) = rangecells.item(i).Value
        If Not chkCompulsory(BankName_80E(i)) Then
              'MsgBox_80E = MsgBox_80E + "* Please provide Name of the Bank/ Institution/Person from which the loan is taken in schedule 80E at Sr. No " & i & "" & Chr(13)
              'Malli-------
              MsgBox_80E = MsgBox_80E + " ""Please provide Name of the Bank / Institution from which the loan is taken in schedule 80E"" at Sr. No " & i & "" & Chr(13)
              '-----------------
            ValidateBankName_80E = False
            Exit Function
        End If
         If Len(BankName_80E(i)) > 125 Then
            MsgBox_80E = MsgBox_80E + "* Name of the Bank / Institution in Sheet 80E should be less than or equal to 125 characters at Sr. No " & i & " ." & Chr(13)
            ValidateBankName_80E = False
            Exit Function
        End If
         
         If Not checkfieldSuperSpecialcharacter(BankName_80E(i)) Then
             MsgBox_80E = MsgBox_80E + "* Name of the Bank / Institution from which the loan is taken in schedule 80E at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
            ValidateBankName_80E = False
            Exit Function
        End If
Next
End Function

Function ValidateAccntNum_80E() As Boolean
    ValidateAccntNum_80E = True

    Dim rangecells, rangecells1 As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanAccNum.80E").Cells
    Set rangecells1 = Sheet17.Range("LoanfrmBankOrInstitute.80E").Cells  'Malli
    
    ReDim AccntNum_80E(end_80E)
    ReDim Bankorothebank_80E(end_80E)
   
    For i = 1 To end_80E
        AccntNum_80E(i) = rangecells.item(i).Value
        Bankorothebank_80E(i) = rangecells1.item(i).Value
        
        If Not chkCompulsory(AccntNum_80E(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'             MsgBox_80E = MsgBox_80E + " ""Please provide Loan Account number of the Bank / Reference number of the Institution from which loan is taken in schedule 80E"" at Sr. No " & i & "" & Chr(13)
             MsgBox_80E = MsgBox_80E + " ""Please provide Loan Account number of the Bank / Institution from which loan is taken in schedule 80E"" at Sr. No " & i & "" & Chr(13)
            ValidateAccntNum_80E = False
            Exit Function
        End If
         If Len(AccntNum_80E(i)) > 20 Then
            MsgBox_80E = MsgBox_80E + "* Loan Account number in Sheet 80E should be less than or equal to 20 characters at Sr. No " & i & "." & Chr(13)
            ValidateAccntNum_80E = False
            Exit Function
        End If
        
        If Bankorothebank_80E(i) = "Bank" Then
        If Not ValidateBankAccountNumber_80E(AccntNum_80E(i), i) Then
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntNum_80E = False
        Exit Function
        End If
        End If
        
        If Bankorothebank_80E(i) = "Institution" Then
        If Not checkfieldspecialcharacter_BthacntReferencenumber(AccntNum_80E(i)) Then
             'MsgBox_24b = MsgBox_24b & "* Refrence number of the Institution at Sr.No " & i & " is invalid  in schedule 24(b), Only "" / "" and "" - "" special characters are allowed." & Chr(13)
              MsgBox_80E = MsgBox_80E & " Loan Account number of the Bank / Institution at Sr.No " & i & " is invalid  in schedule 80E, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateAccntNum_80E = False
        Exit Function
        End If
        End If
         
Next
End Function

'Ankita_23/01/2026=========

Function checkfieldspecialcharacter_BthacntReferencenumber(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_BthacntReferencenumber = True
    Dim arr As Variant
    arr = Array("@", "|", "_", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_BthacntReferencenumber = False
            Exit Function
        End If
        Next
    Next
End Function

Function ValidateLoanDate_80E() As Boolean
    ValidateLoanDate_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanDate.80E").Cells
    ReDim LoanDate_80E(end_80E)
    For i = 1 To end_80E
        LoanDate_80E(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanDate_80E(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'            MsgBox_80E = MsgBox_80E + "* ""Please provide Date of taking loan in schedule 80E"" at Sr. No " & i & "" & Chr(13)
             MsgBox_80E = MsgBox_80E + "* ""Please provide Date of sanction of loan in schedule 80E"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanDate_80E = False
            Exit Function
        End If
         If Len(LoanDate_80E(i)) > 10 Then
          MsgBox_80E = MsgBox_80E + "* Date of sanction of loan at Sr. No " & i & " in Sheet 80E less than 10 characters." & Chr(13)
            ValidateLoanDate_80E = False
            Exit Function
        End If
        'CheckDateddmmyyyy
         If Not CheckDateddmmyyyy(LoanDate_80E(i)) Then
         MsgBox_80E = MsgBox_80E + "* ""Please enter date in valid format"" at Sr. No " & i & "." & Chr(13)
        ValidateLoanDate_80E = False
                 

          Exit Function
        End If
        
'Ankita_23/12/2025===========
        Dim d As Date, cutoff As Date
        cutoff = CDate(Sheet5.Range("Date_1").Value)
        
        If Not ChkMaxDate_80E(Trim(LoanDate_80E(i)), (Sheet5.Range("Date_1").Value)) Then
            MsgBox_80E = MsgBox_80E + "* Date can not be after " & Format$(cutoff, "dd/mm/yyyy") & " at Sr. No " & i & "." & Chr$(13)
            ValidateLoanDate_80E = False
            Exit Function
        End If
'=========================
         
Next
End Function

Function ValidateLoanAmt_80E() As Boolean
    ValidateLoanAmt_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanAmt.80E").Cells
    ReDim LoanAmt_80E(end_80E)
    For i = 1 To end_80E
        LoanAmt_80E(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanAmt_80E(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'            MsgBox_80E = MsgBox_80E + "* ""Please provide Total Loan taken in schedule 80E"" at Sr. No " & i & "" & Chr(13)
            MsgBox_80E = MsgBox_80E + "* ""Please provide Total amount of loan in schedule 80E"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanAmt_80E = False
            Exit Function
        End If
'         If Len(LoanAmt_80E(i)) > 14 Then
'          MsgBox_80E = MsgBox_80E + "* Loan Amount  at Sr. No " & i & " in Sheet 80E less than 15 characters." & Chr(13)
'            ValidateLoanAmt_80E = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanAmt_80E(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'            MsgBox_80E = MsgBox_80E & "* Loan amount at Sr. No  " & i & "  in Sheet 80E should be Numeric value" & Chr(13)
            MsgBox_80E = MsgBox_80E & "* Total amount of loan at Sr. No  " & i & "  in Sheet 80E should be Numeric value" & Chr(13)
            ValidateLoanAmt_80E = False
            Exit Function
        End If
        
        If LoanAmt_80E(i) > 99999999999999# Then
          'Ankita_05/05/2025_Commented as per DESheet_v0.7
            MsgBox_80E = MsgBox_80E & "* Total amount of loan at Sr. No  " & i & "  in Sheet 80E cannot exceed 14 digits" & Chr(13)
            ValidateLoanAmt_80E = False
            Exit Function
        End If
        
        If LoanAmt_80E(i) < 0 Or LoanAmt_80E(i) = 0 Then
          'Ankita_05/05/2025_Commented as per DESheet_v0.7
            MsgBox_80E = MsgBox_80E & "* Total amount of loan should be more than 0 in schedule 80E at Sr. No " & i & "" & Chr(13)
            ValidateLoanAmt_80E = False
            Exit Function
        End If
         
Next
End Function

Function ValidateLoanOutstanding_80E() As Boolean
    ValidateLoanOutstanding_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanOutstanding.80E").Cells
    ReDim LoanOutStanding_80E(end_80E)
    For i = 1 To end_80E
        LoanOutStanding_80E(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanOutStanding_80E(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'             MsgBox_80E = MsgBox_80E + "* ""Loan outstanding as on 31-03-2025 is mandatory in schedule 80E"" at Sr. No " & i & "" & Chr(13)
              MsgBox_80E = MsgBox_80E + "* ""Loan outstanding as on last date of financial year is mandatory in schedule 80E"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanOutstanding_80E = False
            Exit Function
        End If
'         If Len(LoanOutstanding_80E(i)) > 14 Then
'          MsgBox_80E = MsgBox_80E + "* Loan Outstanding  at Sr. No " & i & " in Sheet 80E less than 15 characters." & Chr(13)
'            ValidateLoanOutstanding_80E = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanOutStanding_80E(i)) Then
            MsgBox_80E = MsgBox_80E & "* Loan outstanding at Sr. No  " & i & "  in Sheet 80E should be Numeric value" & Chr(13)
            ValidateLoanOutstanding_80E = False
            Exit Function
        End If
        
        If LoanOutStanding_80E(i) > 99999999999999# Then
            MsgBox_80E = MsgBox_80E & "* Loan outstanding at Sr. No  " & i & "  in Sheet 80E cannot exceed 14 digits" & Chr(13)
            ValidateLoanOutstanding_80E = False
            Exit Function
        End If
        
        If LoanOutStanding_80E(i) < 0 Then
         'Ankita_05/05/2025_Commented as per DESheet_v0.7
            MsgBox_80E = MsgBox_80E & "* Loan outstanding as on last date of financial year can't be less than 0 in schedule 80E. You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
            ValidateLoanOutstanding_80E = False
            Exit Function
        End If
         
Next
End Function

Function ValidateIntrst_80E() As Boolean
    ValidateIntrst_80E = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("Intrst.80E").Cells
    ReDim Intrst_80E(end_80E)
    For i = 1 To end_80E
        Intrst_80E(i) = rangecells.item(i).Value
        If Not chkCompulsory(Intrst_80E(i)) Then
             MsgBox_80E = MsgBox_80E + "* ""Please provide Interest u/s 80E"" at Sr. No " & i & "" & Chr(13)
            ValidateIntrst_80E = False
            Exit Function
        End If
'         If Len(Intrst_80E(i)) > 15 Then
'          MsgBox_80E = MsgBox_80E + "* Interest  at Sr. No " & i & " in Sheet 80E less than 15 characters." & Chr(13)
'            ValidateIntrst_80E = False
'            Exit Function
'        End If

        If Not IsNumeric(Intrst_80E(i)) Then
            MsgBox_80E = MsgBox_80E & "* Interest at Sr. No  " & i & "  in Sheet 80E should be Numeric value" & Chr(13)
            ValidateIntrst_80E = False
            Exit Function
        End If
        
        If Intrst_80E(i) > 99999999999999# Then
            MsgBox_80E = MsgBox_80E & "* Interest at Sr. No  " & i & "  in Sheet 80E cannot exceed 14 digits" & Chr(13)
            ValidateIntrst_80E = False
            Exit Function
        End If
        
        'Interest u/s 80E should be more than 0 in schedule 80E
        
        If Intrst_80E(i) < 0 Or Intrst_80E(i) = 0 Then
            MsgBox_80E = MsgBox_80E & "* Interest u/s 80E should be more than 0 in schedule 80E at Sr. No  " & i & "" & Chr(13)
            ValidateIntrst_80E = False
            Exit Function
        End If
         
Next
End Function


Function ValidateBankAccountNumber_80E(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_80E = True
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
        'Malli--25/04/2025
          If Not checkfieldspecialcharacter80E_EE_EEA_EEB(BankAccountNumber) Then
            MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80E, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80E = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80E " & Chr(13)
            ValidateBankAccountNumber_80E = False
            Exit Function
        End If
        
        'Malli
'         If IsNumeric(BankAccountNumber) Then
'            MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80E " & Chr(13)
'            ValidateBankAccountNumber_80E = False
'            Exit Function
'        End If
    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80E = MsgBox_80E & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc & " is mandatory in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80E" & Chr(13)
        ValidateBankAccountNumber_80E = False
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
            MsgBox_80E = MsgBox_80E & "* Loan Account number at Sr.No " & cc + 1 & " is invalid  in schedule 80E" & Chr(13)
            ValidateBankAccountNumber_80E = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc + 1 & "  in schedule 80E is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80E = False
        Exit Function
    End If
End Function


Public Function ChkMaxDate_80E(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_80E = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ankita_25-26
'        If Year > 2024 Then
        If Year > CInt(Sheet5.Range("DOB_Year").Value) Then
            ChkMaxDate_80E = False
            Exit Function
        Else
'            If Year = 2024 Then
                If Year = CInt(Sheet5.Range("DOB_Year").Value) Then
                If month > 4 Then
                    ChkMaxDate_80E = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_80E = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_80E = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function

Sub Next_80e_ee_eea_eeb_Click()
Dim sourceSheet As Worksheet
Set sourceSheet = ThisWorkbook.Sheets("Schedule 24(b)")
sourceSheet.Activate
End Sub
Sub Prev_80e_ee_eea_eeb_Click()
Dim sourceSheet As Worksheet
Set sourceSheet = ThisWorkbook.Sheets("80C")
sourceSheet.Activate
End Sub
'Malli --25 /04 / 2025
Function checkfieldspecialcharacter80E_EE_EEA_EEB(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter80E_EE_EEA_EEB = True
    Dim arr As Variant
    arr = Array(".", "@", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "_", "|")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter80E_EE_EEA_EEB = False
            Exit Function
        End If
        Next
    Next
End Function

