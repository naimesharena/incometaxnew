Attribute VB_Name = "mdInt24b"
'Option Explicit
'
'
'Public MsgBox_24b As Variant
'Public end_24b As Variant
'Public end_24bLoanfrm As Variant
'Public end_24bIFSC As Variant
'Public end_24bbankName As Variant
'Public end_24bPAN As Variant
'Public end_24bAccntNum As Variant
'Public end_24bLoanDate As Variant
'Public end_24bLoanAmt As Variant
'Public end_24bLoanOutstanding As Variant
'Public end_24bIntrst As Variant
'
'Public Bankorothebank_24b As Variant
'
'
'Sub ValidateSheet24b_Click()
'Dim vbMessgaeCaption As String
'
''vbMessgaeCaption = "ITR 1: AY: 2024-25"
'vbMessgaeCaption = "ITR 1: AY: 2026-27"
'
'Validate24b_All
'fmsgboxStatus "Schedule 24b is OK"
'
''fmsgboxStatus "Sheet 24b is OK"
'End Sub
'Sub Validate24b_All()
'Dim vbMessgaeCaption As String
'vbMessgaeCaption = "Error"
'   Dim sourceSheet As Worksheet
'    Set sourceSheet = ThisWorkbook.Sheets("Schedule 24(b)")
'If Not Validate_24b Then
'    sourceSheet.Activate
'   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
'    fmsgboxStatus MsgBox_24b
'    CloseMsg
'End If
'End Sub
'
''Ankita_02/05/2025
''Ankita_08/05/2025
'Sub cmd_Next24b_click()
'Sheet2.Activate
'
'End Sub
'
''Ankita_02/05/2025
''Ankita_08/05/2025
'Sub cmd_Prev24b_click()
'Sheet18.Activate
'End Sub
'Function Validate_24b()
'Validate_24b = True
'end_24b = 0
'
'setTblinfo_24bLoanfrm
''setTblinfo_24bIFSC
'setTblinfo_24bBankName
''setTblinfo_24bPAN
'setTblinfo_24bAccntNum
'setTblinfo_24bLoanDate
'setTblinfo_24bLoanAmt
'setTblinfo_24bLoanOutstanding
'setTblinfo_24bIntrst
'
'
''end_24b = WorksheetFunction.Max(0, end_24b, end_24bLoanfrm, end_24bbankName, , end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
'
''malli
''Ankita_12/05/2025
''end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bIFSC, end_24bPAN, end_24bbankName, , end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
'end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, , end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
'
'If Not ValidateLoanfrm_24b Then Validate_24b = False
'
''If Not ValidateIFSC_24b Then Validate_24b = False
''If Not ValidatePAN_24b Then Validate_24b = False
'If Not ValidateBankName_24b Then Validate_24b = False
'If Not ValidateAccntNum_24b Then Validate_24b = False
'If Not ValidateLoanDate_24b Then Validate_24b = False
'If Not ValidateLoanAmt_24b Then Validate_24b = False
'If Not ValidateLoanOutstanding_24b Then Validate_24b = False
'If Not ValidateIntrst_24b Then Validate_24b = False
''If Not ValidateLoanfrm_24b Then Validate_24b = False
'If Not Validategreater_24b Then Validate_24b = False 'ValidateIFSC_24b   'Ankita_06/05/2025
'
''ValidateIFSC_24b
''Ankita_05/05/2025_Commented as per DESheet_v0.7
''If (Sheet16.Range("TotAmt.24b").Value) > 2000000 Then
''        MsgBox_24b = MsgBox_24b & "* To claim ""Interest payable on borrowed capital more than Rs.20 lakhs"", you may please consider filing ITR 3. Refer Rule 12 for further details" & Chr(13)
''        Validate_24b = False
''    End If
'
'End Function
'
'Sub setTblinfo_24bLoanfrm()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    Sheet16.Activate
'    mIntCells = Range("LoanfrmBankOrInstitute.24b").count
'    Set rangecells = Range("LoanfrmBankOrInstitute.24b").Cells
'    For mIntCtr = 1 To mIntCells
'           ' If Not rangecells.item(mIntCtr).Value = "" Then
'           'Malli
'             If Not rangecells.item(mIntCtr).Value = "" And Not rangecells.item(mIntCtr).Value = "(Select)" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_24bLoanfrm = ccount
'End Sub
'
''Ankita_05/05/2025_Commented as per DESheet_v0.7
''Sub setTblinfo_24bIFSC()
''    Dim rangecells As Range
''    Dim mIntCells As Long
''    Dim mIntCtr As Long
''    Dim ccount As Long
''    ccount = 0
''    mIntCells = Range("IFSC.24b").count
''    Set rangecells = Range("IFSC.24b").Cells
''    For mIntCtr = 1 To mIntCells
''            If Not rangecells.item(mIntCtr).Value = "" Then
''                ccount = ccount + 1
''            End If
''    Next
''    end_24bIFSC = ccount
''End Sub
'
'
'Sub setTblinfo_24bBankName()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("bankName.24b").count
'    Set rangecells = Range("bankName.24b").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_24bbankName = ccount
'End Sub
'
''Ankita_05/05/2025_Commented as per DESheet_v0.7
''Sub setTblinfo_24bPAN()
''    Dim rangecells As Range
''    Dim mIntCells As Long
''    Dim mIntCtr As Long
''    Dim ccount As Long
''    ccount = 0
''    mIntCells = Range("PAN.24b").count
''    Set rangecells = Range("PAN.24b").Cells
''    For mIntCtr = 1 To mIntCells
''            If Not rangecells.item(mIntCtr).Value = "" Then
''                ccount = ccount + 1
''            End If
''    Next
''    end_24bPAN = ccount
''End Sub
'
'Sub setTblinfo_24bAccntNum()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("loanAccNum.24b").count
'    Set rangecells = Range("loanAccNum.24b").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_24bAccntNum = ccount
'End Sub
'Sub setTblinfo_24bLoanDate()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("loanDate.24b").count
'    Set rangecells = Range("loanDate.24b").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_24bLoanDate = ccount
'End Sub
'Sub setTblinfo_24bLoanAmt()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("loanAmt.24b").count
'    Set rangecells = Range("loanAmt.24b").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_24bLoanAmt = ccount
'End Sub
'
'Sub setTblinfo_24bLoanOutstanding()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("loanOutstanding.24b").count
'    Set rangecells = Range("loanOutstanding.24b").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_24bLoanOutstanding = ccount
'End Sub
'Sub setTblinfo_24bIntrst()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Intrst.24b").count
'    Set rangecells = Range("Intrst.24b").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_24bIntrst = ccount
'End Sub
'
'
'
'Function ValidateLoanfrm_24b() As Boolean
'    ValidateLoanfrm_24b = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet16.Range("LoanfrmBankOrInstitute.24b").Cells
'    ReDim Loanfrm_24b(end_24b)
'    For i = 1 To end_24b
'        Loanfrm_24b(i) = rangecells.item(i).Value
'
'        'Malli comented-----------
''        If Not chkCompulsory(Loanfrm_24b(i)) Then
''             MsgBox_24b = MsgBox_24b + "* Please select dropdown from ""Loan taken from"" in 24(b) schedule at Sr. No " & i & "" & Chr(13)
''            ValidateLoanfrm_24b = False
''            Exit Function
''        End If
''         If Loanfrm_24b(i) = "(Select)" Then
''          MsgBox_24b = MsgBox_24b + "* Please select dropdown from ""Loan taken from"" at Sr. No " & i & " in 24(b) schedule at Sr. No " & i & "" & Chr(13)
''            ValidateLoanfrm_24b = False
''            Exit Function
''        End If
''         '-----------------------------
'
'         'Malli-----
'         If isdropdownblank(Loanfrm_24b(i)) Then
'            MsgBox_24b = MsgBox_24b + "* Please select dropdown from ""Loan taken from"" in schedule 24(b) at Sr. No " & i & "" & Chr(13)
'            ValidateLoanfrm_24b = False
'            Exit Function
'        End If
'         '-----------
'Next
'End Function
'
'
''Function ValidateIFSC_24b() As Boolean
''    ValidateIFSC_24b = True
''
'''    Dim rangecells As Range
'''    Dim rangecells1 As Range
'''    Dim i As Long
'''    Set rangecells = Sheet21.Range("IFSC.24b").Cells
'''    Set rangecells1 = Sheet21.Range("LoanfrmBankOrInstitute.24b").Cells
'''    ReDim IFSC_24b(end_24b)
'''    ReDim Loanfrom_24b(end_24b)
'''    For i = 1 To end_24b
'''        IFSC_24b(i) = rangecells.item(i).value
'''       ' Loanfrom_24b(i) = rangecells.item(i).value
'''     '   If Loanfrom_24b(i) = "Bank" Then
'''        If Not chkCompulsory(IFSC_24b(i)) Then
'''             MsgBox_24b = MsgBox_24b + "* Please select dropdown from ""Loan taken from"" at Sr. No " & i & " in 24(b) schedule." & Chr(13)
'''            ValidateIFSC_24b = False
'''            Exit Function
'''        End If
'''      '  End If
'''
'''         If Len(IFSC_24b(i)) > 11 Then
'''          MsgBox_24b = MsgBox_24b + "* IFSC at Sr. No " & i & " in Sheet 24b should be 11 characters." & Chr(13)
'''            ValidateIFSC_24b = False
'''            Exit Function
'''        End If
'''
'''Next
''
''If (Sheet21.Range("LoanfrmBankOrInstitute.24b").value = "Bank") Then
''        If Sheet21.Range("IFSC.24b").value = "" Then
'''            msgError = msgError & "*If A23a is ""Yes"" Date of filing of Form 10IEA is mandatory. " & Chr(13)
'''Ankita_10/02
'''            msgError = msgError & "*""Date of filing of Form 10IEA is mandatory.""" & Chr(13) 'Modified by sai on 27/01/2025
''            MsgBox_24b = MsgBox_24b & "*""IFSC""" & Chr(13)
''            ValidateIFSC_24b = False
''        End If
''        End If
''
''End Function
'
'
''Ankita_12/05/2025
''Function ValidateIFSC_24b() As Boolean
''ValidateIFSC_24b = True
''
''   ' setTblinfo_24bIFSC
''   setTblinfo_24bLoanfrm
''    Dim rangecells As Range
'''    Dim rangecells1 As Range 'Ankita_05/05/2025_Commented as per DESheet_v0.7
''    Set rangecells = Range("LoanfrmBankOrInstitute.24b").Cells
'''    Set rangecells1 = Range("IFSC.24b").Cells  'Ankita_05/05/2025_Commented as per DESheet_v0.7
''    Dim cellrange As String
''    Dim cellRange1 As String
''    Dim i As Long
''    ReDim Others_Loan(end_24bLoanfrm)
''    ReDim Others_IFSC(end_24bLoanfrm)
''
''
''
''
''
''
''
''    For i = 1 To end_24bLoanfrm
''    cellrange = GetMergedAddressCell(rangecells, i)
'''    cellRange1 = GetMergedAddressCell(rangecells1, i)
''
''        Others_Loan(i) = Sheet16.Range(cellrange).Value
'''        Others_IFSC(i) = Sheet16.Range(cellRange1).Value  'Ankita_05/05/2025_Commented as per DESheet_v0.7
''
''
'''Ankita_05/05/2025_Commented as per DESheet_v0.7
''
'''         If (Others_Loan(i) = "Bank") Then
'''         If Others_IFSC(i) = "" Then
'''         'Change.03.03.2023.102.IDS.113
'''             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
'''             MsgBox_24b = MsgBox_24b & " Please provide  IFSC of the Bank from which loan is taken in schedule 24(b). at Sr. No " & i & "" & Chr(13)
'''         'End Change.IDS113
'''             ValidateIFSC_24b = False
'''             Exit Function
'''         End If
'''         End If
''
'''         If Len(Others_IFSC(i)) > 11 Then
'''            MsgBox_24b = MsgBox_24b + " IFSC of the Bank at Sr. No " & i & " in Sheet 24b should be 11 characters." & Chr(13)
'''            ValidateIFSC_24b = False
'''            Exit Function
'''        End If
'''
''    Next
''
''End Function
'
'Function ValidateAccntNum_24b() As Boolean
'    ValidateAccntNum_24b = True
'
'    Dim rangecells, rangecells1 As Range
'    Dim i As Long
'    Set rangecells = Sheet16.Range("loanAccNum.24b").Cells
'    Set rangecells1 = Sheet16.Range("LoanfrmBankOrInstitute.24b").Cells
'
'    ReDim AccntNum_24b(end_24b)
'
'    ReDim Bankorothebank_24b(end_24b)
'
'
'    For i = 1 To end_24b
'        AccntNum_24b(i) = rangecells.item(i).Value
'        Bankorothebank_24b(i) = rangecells1.item(i).Value
'
'        If Not chkCompulsory(AccntNum_24b(i)) Then
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
''             MsgBox_24b = MsgBox_24b + "* ""Please provide Loan Account number of the Bank / Reference number of the Institution from which loan is taken in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
'              MsgBox_24b = MsgBox_24b + "* ""Please provide Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
'          ValidateAccntNum_24b = False
'            Exit Function
'        End If
'
'         If Len(AccntNum_24b(i)) > 20 Then
'          MsgBox_24b = MsgBox_24b + "* Loan Account number  at Sr. No " & i & " in Sheet 24b less than 20 characters." & Chr(13)
'            ValidateAccntNum_24b = False
'            Exit Function
'        End If
'
'
'    If Bankorothebank_24b(i) = "Bank" Then
'
'    If Not ValidateBankAccountNumber_24b(AccntNum_24b(i), i) Then
'        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
'        ValidateAccntNum_24b = False
'        Exit Function
'    End If
'
'    End If
'
'
'    If Bankorothebank_24b(i) = "Other than bank" Then
'          If Not checkfieldspecialcharacter_BthacntReferencenumber(AccntNum_24b(i)) Then
''            MsgBox_24b = MsgBox_24b & "* Loan Account number of the Bank / Institution at Sr.No " & i & " is invalid  in schedule 24(b), Only "" / "" and "" - "" special characters are allowed." & Chr(13)
'             MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & i & "" & Chr(13)
'           ValidateAccntNum_24b = False
'            Exit Function
'        End If
'    End If
'
'Next
'End Function
'
'
''Ankita_12/05/2025
''Function ValidatePAN_24b() As Boolean
''ValidatePAN_24b = True
''
''   ' setTblinfo_24bIFSC
''   setTblinfo_24bLoanfrm
''    Dim rangecells As Range
'''    Dim rangecells1 As Range  'Ankita_05/05/2025_Commented as per DESheet_v0.7
''    Set rangecells = Range("LoanfrmBankOrInstitute.24b").Cells
'''    Set rangecells1 = Range("PAN.24b").Cells  'Ankita_05/05/2025_Commented as per DESheet_v0.7
''    Dim cellrange As String
''    Dim cellRange1 As String
''    Dim i As Long
''    ReDim Others_Loan(end_24bLoanfrm)
''    ReDim Others_PAN(end_24bLoanfrm)
''
''
''    For i = 1 To end_24bLoanfrm
''    cellrange = GetMergedAddressCell(rangecells, i)
'''    cellRange1 = GetMergedAddressCell(rangecells1, i)  'Ankita_05/05/2025_Commented as per DESheet_v0.7
''
''        Others_Loan(i) = Sheet16.Range(cellrange).Value
'''        Others_PAN(i) = Sheet16.Range(cellRange1).Value  'Ankita_05/05/2025_Commented as per DESheet_v0.7
''
''
''
'''         If (Others_Loan(i) = "Other than bank") Then
'''         If Others_PAN(i) = "" Then
'''         'Change.03.03.2023.102.IDS.113
'''             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
'''             MsgBox_24b = MsgBox_24b & " Please provide  PAN of the institution/person from which loan is taken in schedule 24(b) at Sr. No " & i & "" & Chr(13)
'''         'End Change.IDS113
'''             ValidatePAN_24b = False
'''             Exit Function
'''         End If
'''         End If
'''
'''         If Len(Others_PAN(i)) > 10 Then
'''          MsgBox_24b = MsgBox_24b + "* PAN of the institution/person at Sr. No " & i & " in Sheet 24b should be 10 characters." & Chr(13)
'''            ValidatePAN_24b = False
'''            Exit Function
'''        End If
'''
'''        'Added by Malli(24/04/2025)NewDev
'''        If Others_PAN(i) <> "" Then
'''          If Not EfilingCommon.CheckPAN(UCase(Trim(Others_PAN(i)))) Then
'''            MsgBox_24b = MsgBox_24b + "* Invalid PAN. PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet in schedule 24(b) at Sr. No " & i & "" & Chr(13)
'''            ValidatePAN_24b = False
'''            Exit Function
'''          End If
'''        End If
''
''    Next
''
''End Function
'
'
'Function ValidateBankName_24b() As Boolean
'    ValidateBankName_24b = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet16.Range("bankName.24b").Cells
'    ReDim BankName_24B(end_24b)
'    For i = 1 To end_24b
'        BankName_24B(i) = rangecells.item(i).Value
'        If Not chkCompulsory(BankName_24B(i)) Then
'             MsgBox_24b = MsgBox_24b + "* ""Please provide Name of the Bank/ Institution/Person from which the loan is taken in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
'            ValidateBankName_24b = False
'            Exit Function
'        End If
'         If Len(BankName_24B(i)) > 125 Then
'          MsgBox_24b = MsgBox_24b + "* Name of the Bank/ Institution/Person at Sr. No " & i & " in Sheet 24(b) should be less than or equal to 125." & Chr(13)
'            ValidateBankName_24b = False
'            Exit Function
'        End If
'
'        If Not checkfieldSuperSpecialcharacter(BankName_24B(i)) Then
'             MsgBox_24b = MsgBox_24b + "* Name of the Bank/ Institution/Person in schedule 24b at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
'            ValidateBankName_24b = False
'            Exit Function
'        End If
'
'
'Next
'End Function
'
'
'
'
'Function ValidateLoanDate_24b() As Boolean
'    ValidateLoanDate_24b = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet16.Range("loanDate.24b").Cells
'    ReDim LoanDate_24B(end_24b)
'    For i = 1 To end_24b
'        LoanDate_24B(i) = rangecells.item(i).Value
'        If Not chkCompulsory(LoanDate_24B(i)) Then
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
''             MsgBox_24b = MsgBox_24b + "* ""Please provide Date of taking loan in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
'             MsgBox_24b = MsgBox_24b + "* ""Please provide Date of sanction of loan in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
'            ValidateLoanDate_24b = False
'            Exit Function
'        End If
'         If Len(LoanDate_24B(i)) > 10 Then
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
''         MsgBox_24b = MsgBox_24b + "* Date of taking loan  at Sr. No " & i & " in schedule 24(b) less than 10 characters" & Chr(13)
'          MsgBox_24b = MsgBox_24b + "* Date of sanction of loan at Sr. No " & i & " in schedule 24(b) less than 10 characters" & Chr(13)
'            ValidateLoanDate_24b = False
'            Exit Function
'        End If
'        'CheckDateddmmyyyy
'         If Not CheckDateddmmyyyy(LoanDate_24B(i)) Then
'         MsgBox_24b = MsgBox_24b + "* ""Please enter date in valid format"" in schedule 24(b) at Sr. No " & i & "" & Chr(13)
'        ValidateLoanDate_24b = False
'          Exit Function
'        End If
'
''        'Ankita_23/12/2025=============
'        Dim cutoff1 As Date
'        cutoff1 = CDate(Sheet5.Range("Date_1").Value)
'        If Not ChkMaxDate_24b(Trim(LoanDate_24B(i)), Sheet5.Range("Date_1").Value) Then
'             MsgBox_24b = MsgBox_24b + "* Date can not be after " & Format(cutoff1, "dd/mm/yyyy") & " at Sr. No " & i & "." & Chr$(13)
'             ValidateLoanDate_24b = False
'             Exit Function
'        End If
'          '================================
'Next
'End Function
'
'Function ValidateLoanAmt_24b() As Boolean
'    ValidateLoanAmt_24b = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet16.Range("loanAmt.24b").Cells
'    ReDim LoanAmt_24b(end_24b)
'    For i = 1 To end_24b
'        LoanAmt_24b(i) = rangecells.item(i).Value
'        If Not chkCompulsory(LoanAmt_24b(i)) Then
'             'MsgBox_24b = MsgBox_24b + "* ""Please provide Date of taking loan in schedule 24(b) at Sr. No " & i & "" & Chr(13)
'              'malli
'              MsgBox_24b = MsgBox_24b + "* ""Please provide Total amount of loan in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
'            ValidateLoanAmt_24b = False
'            Exit Function
'        End If
''         If Len(LoanAmt_24b(i)) > 14 Then
''          MsgBox_24b = MsgBox_24b + "* Loan Amount  at Sr. No " & i & " in Sheet 24b less than 15 characters." & Chr(13)
''            ValidateLoanAmt_24b = False
''            Exit Function
''        End If
'
'        If Not IsNumeric(LoanAmt_24b(i)) Then
'            MsgBox_24b = MsgBox_24b & "* Loan amount at Sr. No  " & i & "  in schedule 24(b) should be Numeric value" & Chr(13)
'            ValidateLoanAmt_24b = False
'            Exit Function
'        End If
'
'        If LoanAmt_24b(i) > 99999999999999# Then
'            MsgBox_24b = MsgBox_24b & "* Loan amount at Sr. No  " & i & "  in schedule 24(b) cannot exceed 14 digits" & Chr(13)
'            ValidateLoanAmt_24b = False
'            Exit Function
'        End If
'
'        If LoanAmt_24b(i) < 0 Or LoanAmt_24b(i) = 0 Then
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
''           MsgBox_24b = MsgBox_24b & "* Total loan taken should be more than 0 in schedule 24(b) at Sr. No  " & i & "" & Chr(13)
'            MsgBox_24b = MsgBox_24b & "* Total amount of loan should be more than 0 in schedule 24(b) at Sr. No  " & i & "" & Chr(13)
'            ValidateLoanAmt_24b = False
'            Exit Function
'        End If
'
'Next
'End Function
'
'Function ValidateLoanOutstanding_24b() As Boolean
'    ValidateLoanOutstanding_24b = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet16.Range("loanOutstanding.24b").Cells
'    ReDim LoanOutStanding_24B(end_24b)
'    For i = 1 To end_24b
'        LoanOutStanding_24B(i) = rangecells.item(i).Value
'        If Not chkCompulsory(LoanOutStanding_24B(i)) Then
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
''             MsgBox_24b = MsgBox_24b + "* ""Loan outstanding as on 31-03-2025 is mandatory in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
'             MsgBox_24b = MsgBox_24b + "* ""Loan outstanding as on last date of financial year is mandatory in schedule 24(b)"" at Sr. No " & i & "" & Chr(13)
'           ValidateLoanOutstanding_24b = False
'            Exit Function
'        End If
''         If Len(LoanOutstanding_24b(i)) > 14 Then
''          MsgBox_24b = MsgBox_24b + "* Loan Outstanding  at Sr. No " & i & " in Sheet 24b less than 15 characters." & Chr(13)
''            ValidateLoanOutstanding_24b = False
''            Exit Function
''        End If
'
'        If Not IsNumeric(LoanOutStanding_24B(i)) Then
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
''           MsgBox_24b = MsgBox_24b & "* Loan outstanding at Sr. No  " & i & "  in Sheet 24b should be Numeric value" & Chr(13)
'            MsgBox_24b = MsgBox_24b & "* Loan outstanding as on last date of financial year at Sr. No  " & i & "  in Sheet 24b should be Numeric value" & Chr(13)
'            ValidateLoanOutstanding_24b = False
'            Exit Function
'        End If
'
'        If LoanOutStanding_24B(i) > 99999999999999# Then
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
''           MsgBox_24b = MsgBox_24b & "* Loan outstanding at Sr. No  " & i & "  in Sheet 24b should be Numeric value" & Chr(13)
'            MsgBox_24b = MsgBox_24b & "* Loan outstanding as on last date of financial year at Sr. No  " & i & "  in Sheet 24b cannot exceed 14 digits" & Chr(13)
'            ValidateLoanOutstanding_24b = False
'            Exit Function
'        End If
'
'        If LoanOutStanding_24B(i) < 0 Then
'         'Ankita_05/05/2025_Commented as per DESheet_v0.7
''           MsgBox_24b = MsgBox_24b & "* Loan outstanding as on 31.03.2025 can't be less than 0 in schedule 24(b) at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
'            MsgBox_24b = MsgBox_24b & "* Loan outstanding as on last date of financial year can't be less than 0 in schedule 24(b) at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
'            ValidateLoanOutstanding_24b = False
'            Exit Function
'        End If
'
'Next
'End Function
'
'Function ValidateIntrst_24b() As Boolean
'    ValidateIntrst_24b = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet16.Range("Intrst.24b").Cells
'    ReDim Intrst_24B(end_24b)
'    For i = 1 To end_24b
'        Intrst_24B(i) = rangecells.item(i).Value
'        If Not chkCompulsory(Intrst_24B(i)) Then
'        'Ankita_25/04/2025
''             MsgBox_24b = MsgBox_24b + "* ""Please provide Interest u//s 24(b)"" at Sr. No " & i & "" & Chr(13)
'              MsgBox_24b = MsgBox_24b + "* ""Please provide Interest u/s 24(b)"" at Sr. No " & i & "" & Chr(13)
'            ValidateIntrst_24b = False
'            Exit Function
'        End If
''         If Len(Intrst_24b(i)) > 15 Then
''          MsgBox_24b = MsgBox_24b + "* Interest  at Sr. No " & i & " in Sheet 24b less than 15 characters." & Chr(13)
''            ValidateIntrst_24b = False
''            Exit Function
''        End If
'
'        If Not IsNumeric(Intrst_24B(i)) Then
'            MsgBox_24b = MsgBox_24b & "* Interest at Sr. No  " & i & "  in Sheet 24b should be Numeric value" & Chr(13)
'            ValidateIntrst_24b = False
'            Exit Function
'        End If
'
'        If Intrst_24B(i) > 99999999999999# Then
'            MsgBox_24b = MsgBox_24b & "* Interest at Sr. No  " & i & "  in Sheet 24b cannot exceed 14 digits" & Chr(13)
'            ValidateIntrst_24b = False
'            Exit Function
'        End If
'
'        'Interest u/s 24(b) should be more than 0 in schedule 24(b)
'
'        If Intrst_24B(i) < 0 Or Intrst_24B(i) = 0 Then
'            MsgBox_24b = MsgBox_24b & "* Interest u/s 24(b) should be more than 0 in schedule 24(b) at Sr. No  " & i & "" & Chr(13)
'            ValidateIntrst_24b = False
'            Exit Function
'        End If
'
'Next
'End Function
'
'
'Function ValidateBankAccountNumber_24b(BankAccountNumber As Variant, cc As Long) As Boolean
'    ValidateBankAccountNumber_24b = True
'    Dim numfound As Boolean
'    Dim countnum As Long
'    Dim myB(), ValidateIFSC As Variant
'    Dim i As Long
'    Dim zeroCount As Long
'    Dim BeforeZero, AfterZero As String
'    errmsgVerification = ""
'    numfound = False
'    countnum = 0
'    BeforeZero = ""
'    AfterZero = ""
'    zeroCount = 1
'
'   ' BankAccountNumber = Sheet5.Range("IncD.BankAccountNumber")
'
'
'    If Len(BankAccountNumber) > 0 Then
''        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
''Malli
'        If Not checkfieldspecialcharacter80E_EE_EEA_EEB(BankAccountNumber) Then
''            MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b), Only "" / "" and "" - "" special characters are allowed." & Chr(13)
'            MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'            ValidateBankAccountNumber_24b = False
'            Exit Function
'        End If
'
'        If BankAccountNumber = 0 Then
''            MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b) " & Chr(13)
'             MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'            ValidateBankAccountNumber_24b = False
'            Exit Function
'        End If
'
'    End If
'  '----------------------------------------------------------------
'    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
''        MsgBox_24b = MsgBox_24b & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 24(b)" & Chr(13)
'      MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'        ValidateBankAccountNumber_24b = False
'        Exit Function
'    End If
'
'
'    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
'        MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is mandatory in schedule 24(b)" & Chr(13)
'        ValidateBankAccountNumber_24b = False
'        Exit Function
'    End If
''--------------------------------------------------------------------
'
'    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
''        MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
'         MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'        ValidateBankAccountNumber_24b = False
'        Exit Function
'    End If
'
'    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
''        MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
'         MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'        ValidateBankAccountNumber_24b = False
'        Exit Function
'    End If
'
'    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
''        MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
'        MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'        ValidateBankAccountNumber_24b = False
'        Exit Function
'    End If
'
''malli-----
'
'If (Not checkfieldspecialcharacter_BthacntReferencenumber(BankAccountNumber)) Then
''    MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
'MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'        ValidateBankAccountNumber_24b = False
'        Exit Function
'    End If
'
'''-----------
'
'
'
'    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
''        MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13)
'         MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'        ValidateBankAccountNumber_24b = False
'        Exit Function
'    End If
'
'    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
''        MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 24(b)" & Chr(13
'         MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'        ValidateBankAccountNumber_24b = False
'        Exit Function
'    End If
'
'    ReDim myB(Len(BankAccountNumber) - 1)
'    For i = 1 To Len(BankAccountNumber)
'        myB(i - 1) = Mid(BankAccountNumber, i, 1)
'    Next
'
'    For i = LBound(myB) To UBound(myB)
'        If IsNumeric(myB(i)) Then
'            countnum = countnum + 1
'        End If
'
'        If i > LBound(myB) And i < UBound(myB) Then
'            If myB(i) = 0 Then
'                If myB(i - 1) = 0 Then
'                    zeroCount = zeroCount + 1
'                    AfterZero = IIf(Not IsNumeric(myB(i + 1)), myB(i + 1), "")
'                Else
'                    BeforeZero = IIf(Not IsNumeric(myB(i - 1)), myB(i - 1), "")
'                End If
'            End If
'        End If
'
'    Next
'
'    If BeforeZero <> "" And AfterZero <> "" Then
'        If zeroCount > 1 Then
''            MsgBox_24b = MsgBox_24b & "* Loan Account number at Sr.No " & cc + 1 & " is invalid  in schedule 24(b)" & Chr(13)
'            MsgBox_24b = MsgBox_24b & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'            ValidateBankAccountNumber_24b = False
'            Exit Function
'        End If
'    End If
'
'
'    If countnum < 1 Then
''        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc + 1 & "  in schedule 24(b) is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
'        errmsgVerification = errmsgVerification & "* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & cc & "" & Chr(13)
'        ValidateBankAccountNumber_24b = False
'        Exit Function
'    End If
'End Function
'
'
'Public Function ChkMaxDate_24b(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
'On Error Resume Next
'Dim Year, month, dat As Variant
'
'     ChkMaxDate_24b = True
'     If Len(dateEntered) > 0 Then
'
'     Year = val(Mid(dateEntered, 7, 4))
'     month = val(Mid(dateEntered, 4, 2))
'     dat = val(Mid(dateEntered, 1, 2))
'
''        If Year > 2023 Then
'        'PAG_C1 AY 2024-25 Change
'        'Ankita_25-26
''        If Year > 2024 Then
'        If Year > CInt(Sheet5.Range("DOB_Year").Value) Then 'Ankita_25/12/2025====
'            ChkMaxDate_24b = False
'            Exit Function
'        Else
''            If Year = 2024 Then
'                If Year = CInt(Sheet5.Range("DOB_Year").Value) Then  'Ankita_25/12/2025====
'                If month > 4 Then
'                    ChkMaxDate_24b = False
'                    Exit Function
'                Else
'                    If month = 4 Then
'                        If dat > 1 Then
'                           ChkMaxDate_24b = False
'                            Exit Function
'                        Else
'                            If dat = 1 Then
'                               ChkMaxDate_24b = False
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'     End If
'End Function
'
''Malli----------
'Function checkfieldspecialcharacter_BthacntReferencenumber(field As Variant) As Boolean
'    Dim i, j As Long
'    checkfieldspecialcharacter_BthacntReferencenumber = True
'    Dim arr As Variant
'    arr = Array("@", "|", "_", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
'    For i = 1 To Len(field)
'        For j = 0 To UBound(arr)
'        If Mid(field, i, 1) = arr(j) Then
'            checkfieldspecialcharacter_BthacntReferencenumber = False
'            Exit Function
'        End If
'        Next
'    Next
'End Function
'
''Ankita on 12/05/2025
'
''Function ValidateRefferenceNumber_24b(BankAccountNumber As Variant, cc As Long) As Boolean
''    ValidateRefferenceNumber_24b = True
''    Dim numfound As Boolean
''    Dim countnum As Long
''    Dim myB(), ValidateIFSC As Variant
''    Dim i As Long
''    Dim zeroCount As Long
''    Dim BeforeZero, AfterZero As String
''    errmsgVerification = ""
''    numfound = False
''    countnum = 0
''    BeforeZero = ""
''    AfterZero = ""
''    zeroCount = 1
''
''
''
''
''    If Len(BankAccountNumber) > 0 Then
''        If Not checkfieldspecialcharacter_BthacntReferencenumber(BankAccountNumber) Then
''            MsgBox_24b = MsgBox_24b & "* Reference number of the Institution at Sr.No " & cc & " is invalid  in schedule 24(b), Only "" / "" and "" - "" special characters are allowed." & Chr(13)
''           ValidateRefferenceNumber_24b = False
''            Exit Function
''        End If
''
''    End If
''
''    If Len(BankAccountNumber) > 20 Then
''        MsgBox_24b = MsgBox_24b & "* Reference number of the Institution should be less than 20 characters at Sr.No " & cc & "in schedule 24(b)" & Chr(13)
''           ValidateRefferenceNumber_24b = False
''            Exit Function
''
''    End If
''
''
''
''End Function
'
'
'
''by Ankita on 06/05/2025
'Function Validategreater_24b() As Boolean
'    Validategreater_24b = True
'         If (Len(Sheet16.Range("TotAmt.24b").Value) > 14) Then
'            MsgBox_24b = MsgBox_24b + "*  Total of interest on borrowed capital u/s 24(b) cannot exceed 14 Digits." & Chr(13)
'            Validategreater_24b = False
'            Exit Function
'         End If
'End Function
'
