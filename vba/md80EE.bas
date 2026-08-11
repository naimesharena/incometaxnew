Attribute VB_Name = "md80EE"
Public MsgBox_80EE As Variant
Dim end_80EE As Variant
Dim end_80EELoanfrm As Variant
Dim end_80EEIFSC As Variant
Dim end_80EEbankName As Variant
Dim end_80EEPAN As Variant
Dim end_80EEAccntNum As Variant
Dim end_80EELoanDate As Variant
Dim end_80EELoanAmt As Variant
Dim end_80EELoanOutstanding As Variant
Dim end_80EEIntrst As Variant
'Dim end_80EELoanfrm As Variant


Sub ValidateSheet80EE_Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 1: AY: 2026-27"
'Validate80EE_All  'malli comented

If Validate80EE_All Then
Sheet17.Activate


fmsgboxStatus "Schedule 80EE is OK"
CloseMsg
End If
End Sub


'Sub Validate80EE_All()
Function Validate80EE_All()
 Validate80EE_All = True

Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80E_80EE_80EEA_80EEB")
If Not Validate_80EE Then
    sourceSheet.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    fmsgboxStatus MsgBox_80EE
    CloseMsg
End If
''Ankita_18/04
'Dim Data24b, rowcount_80EE, i, j, rowcount_Int24B As Long
'
'rowcount_80EE = getRowNo(Sheet17.Range("LoanfrmBankOrInstitute.80EE").name)
'
'If end_80EE = 1 Then
'   end_80EE = rowcount_80EE
'Else
'   end_80EE = (end_80EE + rowcount_80EE)
'End If
'
'
'setTblinfo_24bLoanfrm
''setTblinfo_24bIFSC
'
'setTblinfo_24bPAN
'setTblinfo_24bBankName
'setTblinfo_24bAccntNum
'setTblinfo_24bLoanDate
'setTblinfo_24bLoanAmt
'setTblinfo_24bLoanOutstanding
'setTblinfo_24bIntrst
'
'Data24b = end_24bLoanfrm > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
'
'end_24b = WorksheetFunction.Max(0, end_24b, end_24bLoanfrm, end_24bbankName, , end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
'
'rowcount_Int24B = getRowNo(Sheet16.Range("LoanfrmBankOrInstitute.24b").name)
'
'If end_24b = 1 Then
'   end_24b = rowcount_Int24B
'Else
'   end_24b = (end_24b + rowcount_Int24B)
'End If
'
'
'Dim IFSC_80EE, IFSC_24B
'
'IFSC_80EE = Sheet17.Range("Combination_80EE").Column
'IFSC_24B = Sheet16.Range("Combination_24B").Column
'
'
'If Data24b = True Then
'    For i = rowcount_80EE To end_80EE
'        For j = rowcount_Int24B To end_24b
'            If Sheet17.Cells(i, IFSC_80EE).Value = Sheet16.Cells(j, IFSC_24B).Value Then
'               Exit For
'            End If
'        Next
'
'    Next
'Else
'fmsgboxStatus ("Deduction u/s 80EE can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EE is being claimed should be the same as in schedule 24(b)")
'End If
'
'_____________________

'Malli______________________
'Ankita_----------------
Dim rowcount_80EE, i, j, rowcount_Int24B   As Long


'for80EE
setTblinfo_80EELoanfrm
'setTblinfo_80EEIFSC   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'setTblinfo_80EEPAN    'Ankita_05/05/2025_Commented as per DESheet_v0.7
setTblinfo_80EEBankName
setTblinfo_80EEAccntNum
setTblinfo_80EELoanDate
setTblinfo_80EELoanAmt
setTblinfo_80EELoanOutstanding
setTblinfo_80EEIntrst

Dim Data80EE, Data24b, Combination_80EE_24B As Boolean

'Data80EE = end_80EELoanfrm > 0 And end_80EEIFSC > 0 Or end_80EEPAN > 0 And end_80EEbankName > 0 And end_80EEAccntNum > 0 And end_80EELoanDate > 0 And end_80EELoanAmt > 0 And end_80EELoanOutstanding > 0 And end_80EEIntrst > 0
'end_80EE = WorksheetFunction.Max(0, end_80EELoanfrm, end_80EEbankName, end_80EEIFSC, end_80EEPAN, end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)
'Ankita_09/05/2025_Commented as per DESheet_v0.7
Data80EE = end_80EELoanfrm > 0 And end_80EEbankName > 0 And end_80EEAccntNum > 0 And end_80EELoanDate > 0 And end_80EELoanAmt > 0 And end_80EELoanOutstanding > 0 And end_80EEIntrst > 0
end_80EE = WorksheetFunction.Max(0, end_80EELoanfrm, end_80EEbankName, end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)
If Data80EE = True Then

    setTblinfo_24bLoanfrm
'    setTblinfo_24bIFSC   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    setTblinfo_24bPAN    'Ankita_05/05/2025_Commented as per DESheet_v0.7
    setTblinfo_24bBankName
    setTblinfo_24bAccntNum
    setTblinfo_24bLoanDate
    setTblinfo_24bLoanAmt
    setTblinfo_24bLoanOutstanding
    setTblinfo_24bIntrst
    
    'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    Data24b = end_24bLoanfrm > 0 And end_24bIFSC > 0 Or end_24bPAN > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
     Data24b = end_24bLoanfrm > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
'    end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bIFSC, end_24bPAN, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
     end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)

'Ankita_23/01/2026====
'    rowcount_Int24B = getRowNo(Sheet16.Range("LoanfrmBankOrInstitute.24b").name)
'    If end_24b = 1 Then
'       end_24b = rowcount_Int24B
'    Else
'       end_24b = ((end_24b - 1) + rowcount_Int24B) 'Newly Update on 02/05/2025
'    End If

    rowcount_80EE = getRowNo(Sheet17.Range("LoanfrmBankOrInstitute.80EE").name)
    If end_80EE = 1 Then
       end_80EE = rowcount_80EE
    Else
       end_80EE = ((end_80EE - 1) + rowcount_80EE)
    End If
    
    Dim Dedn_80EE, Dedn_24B
    Dedn_80EE = Sheet17.Range("Combination_80EE").Column
'   Dedn_24B = Sheet16.Range("Combination_24B").Column
    
    
    If Data24b = True Then
       For HpCount = 1 To Sheet19.Range("PropertySectionCount").Value
           Dedn_24B = Sheet19.Range("Combination_24B" & HpCount).Column
           
           rowcount_Int24B = getRowNo(Sheet19.Range("LoanfrmBankOrInstitute.24b" & HpCount).name)
           end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)

            If end_24b = 1 Then
                end_24b = rowcount_Int24B
            Else
                end_24b = ((end_24b - 1) + rowcount_Int24B)
            End If

            Combination_80EE_24B = False
            For i = rowcount_80EE To end_80EE
                For j = rowcount_Int24B To end_24b
                    If UCase(Trim(Sheet17.Cells(i, Dedn_80EE).Value)) = UCase(Trim(Sheet19.Cells(j, Dedn_24B).Value)) Then
                       Combination_80EE_24B = True
                       GoTo checking
'                      Exit For

                    End If
                Next
                
            Next
          Next
    Else
        Combination_80EE_24B = False
    End If

checking:

If Combination_80EE_24B = False Then
         sourceSheet.Activate  'Ankita_13/05/2025
'        MsgBox_80EE = MsgBox_80EE & ("""Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EE is being claimed should be the same as in schedule 24(b)""") & Chr(13)
         MsgBox_80EE = MsgBox_80EE & ("""Deduction u/s 80EE can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EE is being claimed should be the same as in schedule 24(b)""") & Chr(13)
         Validate80EE_All = False   ' Ankita_09/05/2025
         fmsgboxStatus MsgBox_80EE
         CloseMsg
'        Validate80EE_All = False
        Exit Function
End If

End If

'Ankita_26/01/2026========

Dim k, Total As Long

    For k = 1 To Sheet19.Range("PropertySectionCount").Value
       Total = Total + Sheet19.Range("TotAmt.24b" & k).Value
    
    Next
    
 If (Sheet17.Range("TotAmt.80EE").Value > 0 And Not Total > 0) Then
        sourceSheet.Activate  'Ankita_13/05/2025
        MsgBox_80EE = MsgBox_80EE & " ""Deduction u/s 80EE can be claimed only if the limit u/s 24(b) is exhausted.""" & Chr(13)
        fmsgboxStatus MsgBox_80EE
        CloseMsg
    End If
    
    
'Ankita_21/04
 If MsgBox_80EE <> "" Then
  sourceSheet.Activate  'Ankita_13/05/2025
  Validate80EE_All = False
  Else
  sourceSheet.Activate  'Ankita_13/05/2025
  Validate80EE_All = True
  End If
'
'___________________________


End Function

Function Validate_80EE()
Validate_80EE = True
MsgBox_80EE = ""
setTblinfo_80EELoanfrm
'setTblinfo_80EEIFSC
end_80EE = 0

'setTblinfo_80EEPAN
setTblinfo_80EEBankName
setTblinfo_80EEAccntNum
setTblinfo_80EELoanDate
setTblinfo_80EELoanAmt
setTblinfo_80EELoanOutstanding
setTblinfo_80EEIntrst


'end_80EE = WorksheetFunction.Max(0, end_80EE, end_80EELoanfrm, end_80EEbankName, , end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)

'end_80EE = WorksheetFunction.Max(0, end_80EELoanfrm, end_80EEbankName, end_80EEPAN, end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)
end_80EE = WorksheetFunction.Max(0, end_80EELoanfrm, end_80EEbankName, end_80EEAccntNum, end_80EELoanDate, end_80EELoanAmt, end_80EELoanOutstanding, end_80EEIntrst)

'Ankita_12/05/2025_Commented as per DESheet_v0.7
'If end_80EE > 0 Then
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    If (Sheet17.Range("ResidentialHP.80E").Value) = "" Then
'        MsgBox_80EE = MsgBox_80EE & " ""Value of residential House property is mandatory in schedule 80EE""" & Chr(13)
'        Validate_80EE = False
'    End If
'
'     If (Sheet17.Range("ResidentialHP.80E").Value) > 5000000 Then
'        MsgBox_80EE = MsgBox_80EE & " ""Value of residential house property"" shall not be more than Rs. 50 Lakhs in schedule 80EE" & Chr(13)
'        Validate_80EE = False
'    End If
'End If

If Not ValidateLoanfrm_80EE Then Validate_80EE = False
'If Not ValidateIFSC_80EE Then Validate_80EE = False
'If Not ValidatePAN_80EE Then Validate_80EE = False
If Not ValidateBankName_80EE Then Validate_80EE = False
If Not ValidateAccntNum_80EE Then Validate_80EE = False
If Not ValidateLoanDate_80EE Then Validate_80EE = False
If Not ValidateLoanAmt_80EE Then Validate_80EE = False

    If (Sheet17.Range("MaxLoan_80EE").Value) > 3500000 Then
'        MsgBox_80EE = MsgBox_80EE & "* ""Total loan taken"" shall not be more than Rs. 35 Lakhs in schedule 80EE" & Chr(13)
        MsgBox_80EE = MsgBox_80EE & "* ""Total amount of Loan"" shall not be more than Rs. 35 Lakhs in schedule 80EE" & Chr(13)
   
        Validate_80EE = False
    End If

If Not ValidateLoanOutstanding_80EE Then Validate_80EE = False
If Not ValidateIntrst_80EE Then Validate_80EE = False
'If Not ValidateLoanfrm_80EE Then Validate_80EE = False

'ValidateIFSC_80EE

'    If (Sheet17.Range("TotAmt.80EE").Value) > 2000000 Then
'        MsgBox_80EE = MsgBox_80EE & " To claim ""Interest payable on borrowed capital more than Rs.20 lakhs"", you may please consider filing ITR 3. Refer Rule 12 for further details" & Chr(13)
'        Validate_80EE = False
'    End If
    'Ankita_06/05/2025_Commented as per DESheet_v0.7
        If Len(Sheet17.Range("TotAmt.80EE").Value) > 14 Then
        MsgBox_80EE = MsgBox_80EE & " Total of Interest u/s 80EE in schedule 80EE cannot exceed 14 digits" & Chr(13)
        Validate_80EE = False
    End If


End Function

Sub setTblinfo_80EELoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Sheet17.Activate
    mIntCells = Range("LoanfrmBankOrInstitute.80EE").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "" Or rangecells.item(mIntCtr).Value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    end_80EELoanfrm = ccount
    
End Sub

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Sub setTblinfo_80EEIFSC()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("IFSC.80EE").count
'    Set rangecells = Range("IFSC.80EE").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EEIFSC = ccount
'End Sub


Sub setTblinfo_80EEBankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("bankName.80EE").count
    Set rangecells = Range("bankName.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEbankName = ccount
    
End Sub

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Sub setTblinfo_80EEPAN()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("PAN.80EE").count
'    Set rangecells = Range("PAN.80EE").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EEPAN = ccount
'
'End Sub

Sub setTblinfo_80EEAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAccNum.80EE").count
    Set rangecells = Range("loanAccNum.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEAccntNum = ccount
    
End Sub
Sub setTblinfo_80EELoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanDate.80EE").count
    Set rangecells = Range("loanDate.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EELoanDate = ccount
    
End Sub
Sub setTblinfo_80EELoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanAmt.80EE").count
    Set rangecells = Range("loanAmt.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EELoanAmt = ccount
     
End Sub

Sub setTblinfo_80EELoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("loanOutstanding.80EE").count
    Set rangecells = Range("loanOutstanding.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EELoanOutstanding = ccount
     
End Sub
Sub setTblinfo_80EEIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Intrst.80EE").count
    Set rangecells = Range("Intrst.80EE").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEIntrst = ccount
    
End Sub



Function ValidateLoanfrm_80EE() As Boolean
    ValidateLoanfrm_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("LoanfrmBankOrInstitute.80EE").Cells
    ReDim Loanfrm_80EE(end_80EE)
    For i = 1 To end_80EE
        Loanfrm_80EE(i) = rangecells.item(i).Value
        'Ankita_25/04/2025
'        If Not chkCompulsory(Loanfrm_80EE(i)) Then
        If isdropdownblank(Loanfrm_80EE(i)) Then
             MsgBox_80EE = MsgBox_80EE + "* Please select dropdown from ""Loan taken from"" in schedule 80EE at Sr. No " & i & "" & Chr(13)  'Changed as per DE sheet v0.8 by Ankita
            ValidateLoanfrm_80EE = False
            Exit Function
        End If
'         If Loanfrm_80EE(i) = "(Select)" Then
'          MsgBox_80EE = MsgBox_80EE + " Please select dropdown ""Loan taken from"" in schedule 80EE at Sr. No " & i & "" & Chr(13)
'            ValidateLoanfrm_80EE = False
'            Exit Function
'        End If
         
Next
End Function


'Function ValidateIFSC_80EE() As Boolean
'ValidateIFSC_80EE = True
''Ankita_05/05/2025_Commented as per DESheet_v0.7
'   ' setTblinfo_80EEIFSC
'    setTblinfo_80EELoanfrm
'    Dim rangecells As Range
''    Dim rangecells1 As Range
'    Set rangecells = Range("LoanfrmBankOrInstitute.80EE").Cells
''    Set rangecells1 = Range("IFSC.80EE").Cells
'    Dim cellrange As String
'    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80EELoanfrm)
''    ReDim Others_IFSC(end_80EELoanfrm)
'
'
'    For i = 1 To end_80EELoanfrm
'    cellrange = GetMergedAddressCell(rangecells, i)
''    cellRange1 = GetMergedAddressCell(rangecells1, i)
'
'        Others_Loan(i) = Sheet17.Range(cellrange).Value
''        Others_IFSC(i) = Sheet17.Range(cellRange1).Value
'
'
'
''         If (Others_Loan(i) = "Bank") Then
''         If Others_IFSC(i) = "" Then
''         'Change.03.03.2023.102.IDS.113
''             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
''             MsgBox_80EE = MsgBox_80EE & "* Please provide IFSC of the Bank from which loan is taken in schedule 80EE at Sr. No " & i & "" & Chr(13)
''         'End Change.IDS113
''             ValidateIFSC_80EE = False
''             Exit Function
''         End If
''         End If
'
''         If Len(Others_IFSC(i)) > 11 Then
''          MsgBox_80EE = MsgBox_80EE + "* IFSC of the Bank at Sr. No " & i & " in Sheet 80EE should be 11 characters." & Chr(13)
''            ValidateIFSC_80EE = False
''            Exit Function
''        End If
'
'    Next
'
'End Function


'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Function ValidatePAN_80EE() As Boolean
'ValidatePAN_80EE = True
'
'   ' setTblinfo_80EEIFSC
'   setTblinfo_80EELoanfrm
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Set rangecells = Range("LoanfrmBankOrInstitute.80EE").Cells
'    Set rangecells1 = Range("PAN.80EE").Cells
'    Dim cellrange As String
'    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80EELoanfrm)
'    ReDim Others_PAN(end_80EELoanfrm)
'
'
'    For i = 1 To end_80EELoanfrm
'    cellrange = GetMergedAddressCell(rangecells, i)
'    cellRange1 = GetMergedAddressCell(rangecells1, i)
'
'        Others_Loan(i) = Sheet17.Range(cellrange).Value
'        Others_PAN(i) = Sheet17.Range(cellRange1).Value
'
'
'
'         If (Others_Loan(i) = "Institution") Then
'
'
'         If Others_PAN(i) = "" Then
'         'Change.03.03.2023.102.IDS.113
'             'msgError = msgError & "* Please enter description in Nature of Income (Exempt income (For reporting Purposes)) in Sheet Income Details" & Chr(13)
'             MsgBox_80EE = MsgBox_80EE & "* ""Please provide PAN of the institution from which loan is taken in schedule 80EE"" at Sr. No " & i & "" & Chr(13)
'         'End Change.IDS113
'             ValidatePAN_80EE = False
'             Exit Function
'         End If
'         End If
'
'         If Len(Others_PAN(i)) > 10 Then
'          MsgBox_80EE = MsgBox_80EE + "* PAN of the institution at Sr. No " & i & " in Sheet 80EE should be 10 characters." & Chr(13)
'            ValidatePAN_80EE = False
'            Exit Function
'        End If
'
'        'Added by Malli(24/04/2025)NewDev
'        If Others_PAN(i) <> "" Then
'           If Not EfilingCommon.CheckPAN(UCase(Trim(Others_PAN(i)))) Then
'              MsgBox_80EE = MsgBox_80EE + "* Invalid PAN. PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet in schedule 80EE at Sr. No " & i & "" & Chr(13)
'            ValidatePAN_80EE = False
'            Exit Function
'           End If
'         End If
'
'    Next
'
'End Function


Function ValidateBankName_80EE() As Boolean
    ValidateBankName_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("bankName.80EE").Cells
    ReDim BankName_80EE(end_80EE)
    For i = 1 To end_80EE
        BankName_80EE(i) = rangecells.item(i).Value
        If Not chkCompulsory(BankName_80EE(i)) Then
             'MsgBox_80EE = MsgBox_80EE + "* Please provide Name of the Bank/ Institution/Person from which the loan is taken in schedule 80EE at Sr. No " & i & "" & Chr(13)
              MsgBox_80EE = MsgBox_80EE + "* Please provide ""Name of the Bank / Institution from which the loan is taken"" in schedule 80EE at Sr. No " & i & "" & Chr(13)
            ValidateBankName_80EE = False
            Exit Function
        End If
        'Ankita_06/05/2025_Commented as per DESheet_v0.7
         If Len(BankName_80EE(i)) > 125 Then
         ' MsgBox_80EE = MsgBox_80EE + "* Name of the Bank/ Institution/Person at Sr. No " & i & " in Sheet 80EE less than 75 characters." & Chr(13)
            MsgBox_80EE = MsgBox_80EE + "* Name of the Bank / Institution in Sheet 80EE should be less than or equal to 125 characters at Sr. No " & i & " ." & Chr(13)
            ValidateBankName_80EE = False
            Exit Function
        End If
         If Not checkfieldSuperSpecialcharacter(BankName_80EE(i)) Then
             MsgBox_80EE = MsgBox_80EE + "* Name of the Bank / Institution from which the loan is taken in schedule 80EE at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
            ValidateBankName_80EE = False
            Exit Function
        End If
         
Next
End Function

'Function ValidateIFSC_80EE() As Boolean
'    ValidateIFSC_80EE = True
'
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Dim i As Long
'    Set rangecells = Sheet17.Range("IFSC.80EE").Cells
'    Set rangecells1 = Sheet17.Range("LoanfrmBankOrInstitute.80EE").Cells
'    ReDim IFSC_80EE(end_80EE)
'    ReDim Loanfrom_80EE(end_80EE)
'    For i = 1 To end_80EE
'        IFSC_80EE(i) = rangecells.item(i).value
'       ' Loanfrom_80EE(i) = rangecells.item(i).value
'     '   If Loanfrom_80EE(i) = "Bank" Then
'        If Not chkCompulsory(IFSC_80EE(i)) Then
'             MsgBox_80EE = MsgBox_80EE + "* Please select dropdown from ""Loan taken from"" at Sr. No " & i & " in 80EE schedule." & Chr(13)
'            ValidateIFSC_80EE = False
'            Exit Function
'        End If
'      '  End If
'
'         If Len(IFSC_80EE(i)) > 11 Then
'          MsgBox_80EE = MsgBox_80EE + "* IFSC at Sr. No " & i & " in Sheet 80EE should be 11 characters." & Chr(13)
'            ValidateIFSC_80EE = False
'            Exit Function
'        End If
'
'Next
''
''If (Sheet17.Range("LoanfrmBankOrInstitute.80EE").value = "Bank") Then
''        If Sheet17.Range("IFSC.80EE").value = "" Then
'''            msgError = msgError & "*If A23a is ""Yes"" Date of filing of Form 10IEA is mandatory. " & Chr(13)
'''Ankita_10/02
'''            msgError = msgError & "*""Date of filing of Form 10IEA is mandatory.""" & Chr(13) 'Modified by sai on 27/01/2025
''            MsgBox_80EE = MsgBox_80EE & "*""IFSC""" & Chr(13)
''            ValidateIFSC_80EE = False
''        End If
''        End If
'
'End Function
Function ValidateAccntNum_80EE() As Boolean
    ValidateAccntNum_80EE = True

    Dim rangecells, rangecells1 As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanAccNum.80EE").Cells
    Set rangecells1 = Sheet17.Range("LoanfrmBankOrInstitute.80EE").Cells
    
    ReDim AccntNum_80EE(end_80EE)
    ReDim Bankorothebank_80EE(end_80EE)
    
    For i = 1 To end_80EE
        AccntNum_80EE(i) = rangecells.item(i).Value
        Bankorothebank_80EE(i) = rangecells1.item(i).Value
        
        If Not chkCompulsory(AccntNum_80EE(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'            MsgBox_80EE = MsgBox_80EE + "* Please provide ""Loan Account number of the Bank / Reference number of the Institution"" from which loan is taken in schedule 80EE at Sr. No " & i & "" & Chr(13)
             MsgBox_80EE = MsgBox_80EE + "* Please provide ""Loan Account number of the Bank / Institution"" from which loan is taken in schedule 80EE at Sr. No " & i & "" & Chr(13)
            ValidateAccntNum_80EE = False
            Exit Function
        End If
         If Len(AccntNum_80EE(i)) > 20 Then
'          MsgBox_80EE = MsgBox_80EE + "* Loan Account number  at Sr. No " & i & " in Sheet 80EE less than 20 characters." & Chr(13)
            MsgBox_80EE = MsgBox_80EE + "* Loan Account number in Sheet 80EE should be less than or equal to 20 characters at Sr. No " & i & "." & Chr(13)
            ValidateAccntNum_80EE = False
            Exit Function
        End If
       
       If Bankorothebank_80EE(i) = "Bank" Then
        If Not ValidateBankAccountNumber_80EE(AccntNum_80EE(i), i) Then
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntNum_80EE = False
        Exit Function
    End If
    End If
     
     If Bankorothebank_80EE(i) = "Institution" Then
     If Not checkfieldspecialcharacter_BthacntReferencenumber(AccntNum_80EE(i)) Then
        MsgBox_80EE = MsgBox_80EE + "* Loan Account number of the Bank / Institution at Sr.No " & i & " is invalid  in schedule 80EE, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
        ValidateAccntNum_80EE = False
        Exit Function
      End If
    End If
Next
End Function

Function ValidateLoanDate_80EE() As Boolean
    ValidateLoanDate_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanDate.80EE").Cells
    ReDim LoanDate_80EE(end_80EE)
    For i = 1 To end_80EE
        LoanDate_80EE(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanDate_80EE(i)) Then
             'MsgBox_80EE = MsgBox_80EE + "* Please provide Date of taking loan in schedule 80EE at Sr. No " & i & "" & Chr(13)
               MsgBox_80EE = MsgBox_80EE + "* Please provide ""Date of sanction of loan"" in schedule 80EE at Sr. No " & i & "" & Chr(13)
            ValidateLoanDate_80EE = False
            Exit Function
        End If
        
         If Len(LoanDate_80EE(i)) > 10 Then
          'MsgBox_80EE = MsgBox_80EE + "* Date of taking loan  at Sr. No " & i & " in Sheet 80EE less than 10 characters." & Chr(13)
          MsgBox_80EE = MsgBox_80EE + "* Date of sanction of loan  at Sr. No " & i & " in Sheet 80EE less than 10 characters." & Chr(13)
            ValidateLoanDate_80EE = False
            Exit Function
        End If
        
        'CheckDateddmmyyyy
         If Not CheckDateddmmyyyy(LoanDate_80EE(i)) Then
         MsgBox_80EE = MsgBox_80EE + "* ""Please enter date in valid format"" at Sr. No " & i & "." & Chr(13)
        ValidateLoanDate_80EE = False
          Exit Function
        End If
        
        If Not ChkMaxDate_80EE(Trim(LoanDate_80EE(i)), "31-03-2017") Then
            MsgBox_80EE = MsgBox_80EE + " ""Date of sanction of Loan"" shall be between 01/04/2016 to 31/03/2017 in schedule 80EE at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EE = False
          Exit Function
        End If
        
       If Not ChkMinInclusiveDate(Trim(Dformat(LoanDate_80EE(i), "yyyy-mm-dd")), "2016-04-01") Then
           MsgBox_80EE = MsgBox_80EE + " ""Date of sanction of Loan"" shall be between 01/04/2016 to 31/03/2017 in schedule 80EE at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EE = False
          Exit Function
        End If
                 
Next
End Function

Function ValidateLoanAmt_80EE() As Boolean
    ValidateLoanAmt_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanAmt.80EE").Cells
    ReDim LoanAmt_80EE(end_80EE)
    For i = 1 To end_80EE
        LoanAmt_80EE(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanAmt_80EE(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
             MsgBox_80EE = MsgBox_80EE + "* ""Please provide ""Total amount of Loan"" in schedule 80EE"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanAmt_80EE = False
            Exit Function
        End If
'         If Len(LoanAmt_80EE(i)) > 14 Then
'          MsgBox_80EE = MsgBox_80EE + "* Loan Amount  at Sr. No " & i & " in Sheet 80EE less than 15 characters." & Chr(13)
'            ValidateLoanAmt_80EE = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanAmt_80EE(i)) Then
            MsgBox_80EE = MsgBox_80EE & "* Loan amount at Sr. No  " & i & "  in Sheet 80EE should be Numeric value" & Chr(13)
            ValidateLoanAmt_80EE = False
            Exit Function
        End If
        
        If LoanAmt_80EE(i) > 99999999999999# Then
            MsgBox_80EE = MsgBox_80EE & "* Loan amount at Sr. No  " & i & "  in Sheet 80EE cannot exceed 14 digits" & Chr(13)
            ValidateLoanAmt_80EE = False
            Exit Function
        End If
        
        If LoanAmt_80EE(i) < 0 Or LoanAmt_80EE(i) = 0 Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
            MsgBox_80EE = MsgBox_80EE & "* ""Total amount of Loan"" should be more than 0 in schedule 80EE" & Chr(13)
            ValidateLoanAmt_80EE = False
            Exit Function
        End If
         
Next
End Function

Function ValidateLoanOutstanding_80EE() As Boolean
    ValidateLoanOutstanding_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanOutstanding.80EE").Cells
    ReDim LoanOutStanding_80EE(end_80EE)
    For i = 1 To end_80EE
        LoanOutStanding_80EE(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanOutStanding_80EE(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'             MsgBox_80EE = MsgBox_80EE + "* ""Loan outstanding as on 31-03-2025 is mandatory in schedule 80EE"" at Sr. No " & i & "" & Chr(13)
             MsgBox_80EE = MsgBox_80EE + "* ""Loan outstanding as on last date of financial year is mandatory in schedule 80EE"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanOutstanding_80EE = False
            Exit Function
        End If
'         If Len(LoanOutstanding_80EE(i)) > 14 Then
'          MsgBox_80EE = MsgBox_80EE + "* Loan Outstanding  at Sr. No " & i & " in Sheet 80EE less than 15 characters." & Chr(13)
'            ValidateLoanOutstanding_80EE = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanOutStanding_80EE(i)) Then
            MsgBox_80EE = MsgBox_80EE & "* Loan outstanding at Sr. No  " & i & "  in Sheet 80EE should be Numeric value" & Chr(13)
            ValidateLoanOutstanding_80EE = False
            Exit Function
        End If
        
        If LoanOutStanding_80EE(i) > 99999999999999# Then
            MsgBox_80EE = MsgBox_80EE & "* Loan outstanding at Sr. No  " & i & "  in Sheet 80EE cannot exceed 14 digits" & Chr(13)
            ValidateLoanOutstanding_80EE = False
            Exit Function
        End If
        
        If LoanOutStanding_80EE(i) < 0 Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'            MsgBox_80EE = MsgBox_80EE & "* ""Loan outstanding as on 31.03.2025"" can't be less than 0 in schedule 80EE at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment"" & Chr(13)"
            MsgBox_80EE = MsgBox_80EE & "* ""Loan outstanding as on last date of financial year"" can't be less than 0 in schedule 80EE at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment"" & Chr(13)"
            ValidateLoanOutstanding_80EE = False
            Exit Function
        End If
         
Next
End Function

Function ValidateIntrst_80EE() As Boolean
    ValidateIntrst_80EE = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("Intrst.80EE").Cells
    ReDim Intrst_80EE(end_80EE)
    For i = 1 To end_80EE
        Intrst_80EE(i) = rangecells.item(i).Value
        If Not chkCompulsory(Intrst_80EE(i)) Then
             MsgBox_80EE = MsgBox_80EE + "* ""Please provide Interest u/s 80EE"" at Sr. No " & i & "" & Chr(13)
            ValidateIntrst_80EE = False
            Exit Function
        End If
'         If Len(Intrst_80EE(i)) > 15 Then
'          MsgBox_80EE = MsgBox_80EE + "* Interest  at Sr. No " & i & " in Sheet 80EE less than 15 characters." & Chr(13)
'            ValidateIntrst_80EE = False
'            Exit Function
'        End If

        If Not IsNumeric(Intrst_80EE(i)) Then
            MsgBox_80EE = MsgBox_80EE & "* Interest at Sr. No  " & i & "  in Sheet 80EE should be Numeric value" & Chr(13)
            ValidateIntrst_80EE = False
            Exit Function
        End If
        
        If Intrst_80EE(i) > 99999999999999# Then
            MsgBox_80EE = MsgBox_80EE & "* Interest at Sr. No  " & i & "  in Sheet 80EE cannot exceed 14 digits" & Chr(13)
            ValidateIntrst_80EE = False
            Exit Function
        End If
        
        'Interest u/s 80EE should be more than 0 in schedule 80EE
        
        If Intrst_80EE(i) < 0 Or Intrst_80EE(i) = 0 Then
            MsgBox_80EE = MsgBox_80EE & "* ""Interest u/s 80EE"" should be more than 0 in schedule 80EE at Sr. No  " & i & "" & Chr(13)
            ValidateIntrst_80EE = False
            Exit Function
        End If
         
Next
End Function


Function ValidateBankAccountNumber_80EE(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_80EE = True
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
       ' If Not checkfieldspecialcharacter1(BankAccountNumber) Then
       'Malli
        If Not checkfieldspecialcharacter80E_EE_EEA_EEB(BankAccountNumber) Then
            MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80EE = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE " & Chr(13)
            ValidateBankAccountNumber_80EE = False
            Exit Function
        End If
        
        'Malli
'        If IsNumeric(BankAccountNumber) Then
'            MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EE " & Chr(13)
'            ValidateBankAccountNumber_80EE = False
'            Exit Function
'        End If
    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80EE = MsgBox_80EE & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is mandatory in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EE" & Chr(13)
        ValidateBankAccountNumber_80EE = False
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
            MsgBox_80EE = MsgBox_80EE & "* Loan Account number at Sr.No " & cc + 1 & " is invalid in schedule 80EE" & Chr(13)
            ValidateBankAccountNumber_80EE = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc + 1 & "  in schedule 80EE is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80EE = False
        Exit Function
    End If
End Function


Public Function ChkMaxDate_80EE(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_80EE = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ankita_25-26
'        If Year > 2024 Then
        If Year > 2017 Then
            ChkMaxDate_80EE = False
            Exit Function
        Else
'            If Year = 2024 Then
                If Year = 2017 Then
                If month > 4 Then
                    ChkMaxDate_80EE = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_80EE = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_80EE = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function



