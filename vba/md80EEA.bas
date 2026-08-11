Attribute VB_Name = "md80EEA"
Public MsgBox_80EEA As Variant
Dim end_80EEA As Variant
Dim end_80EEALoanfrm As Variant
Dim end_80EEAIFSC As Variant
Dim end_80EEAbankName As Variant
Dim end_80EEAPAN As Variant
Dim end_80EEAAccntNum As Variant
Dim end_80EEALoanDate As Variant
Dim end_80EEALoanAmt As Variant
Dim end_80EEALoanOutstanding As Variant
Dim end_80EEAIntrst As Variant

Sub ValidateSheet80EEA_Click()
Dim vbMessgaeCaption As String

vbMessgaeCaption = "ITR 1: AY: 2026-27"
Validate80EEA_All

fmsgboxStatus "Schedule 80EEA is OK"
End Sub


'Sub Validate80EEA_All()
Function Validate80EEA_All()
 Validate80EEA_All = True
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
   Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("80E_80EE_80EEA_80EEB")
    
''    Public lock_80EEA_flag, lock_80EE_flag As Boolean
''    If lock_80EEA_flag = False Then
''     fmsgboxStatus ("* ""Deduction u/s 80EE and 80EEA can't be claimed together""")
''    CloseMsg
''    End If
    
If Not Validate_80EEA Then
    sourceSheet.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    fmsgboxStatus MsgBox_80EEA
    CloseMsg
End If




''Ankita_18/04
'Dim Data24b, rowcount_80EEA, i, j, rowcount_Int24B As Long
'
'rowcount_80EEA = getRowNo(Sheet17.Range("LoanfrmBankOrInstitute.80EEA").name)
'
'If end_80EEA = 1 Then
'   end_80EEA = rowcount_80EEA
'Else
'   end_80EEA = (end_80EEA + rowcount_80EEA)
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
'Dim IFSC_80EEA, IFSC_24B
'
'IFSC_80EEA = Sheet17.Range("Combination_80EEA").Column
'IFSC_24B = Sheet16.Range("Combination_24B").Column
'
'
'If Data24b = True Then
'    For i = rowcount_80EEA To end_80EEA
'        For j = rowcount_Int24B To end_24b
'            If Sheet17.Cells(i, IFSC_80EEA).Value = Sheet16.Cells(j, IFSC_24B).Value Then
'               Exit For
'            End If
'        Next
'
'    Next
'Else
'fmsgboxStatus ("Deductions")
'End If
'
'Malli_______________________23/04/2025

'Ankita_21/04----------------------
Dim rowcount_80EEA, i, j, rowcount_Int24B   As Long


'for80EEA
setTblinfo_80EEALoanfrm
'setTblinfo_80EEAIFSC 'Ankita_05/05/2025_Commented as per DESheet_v0.7
'setTblinfo_80EEAPAN  'Ankita_05/05/2025_Commented as per DESheet_v0.7
setTblinfo_80EEABankName
setTblinfo_80EEAAccntNum
setTblinfo_80EEALoanDate
setTblinfo_80EEALoanAmt
setTblinfo_80EEALoanOutstanding
setTblinfo_80EEAIntrst

Dim Data80EEA, Data24b, Combination_80EEA_24B As Boolean
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Data80EEA = end_80EEALoanfrm > 0 And end_80EEAIFSC > 0 Or end_80EEAPAN > 0 And end_80EEAbankName > 0 And end_80EEAAccntNum > 0 And end_80EEALoanDate > 0 And end_80EEALoanAmt > 0 And end_80EEALoanOutstanding > 0 And end_80EEAIntrst > 0
Data80EEA = end_80EEALoanfrm > 0 And end_80EEAbankName > 0 And end_80EEAAccntNum > 0 And end_80EEALoanDate > 0 And end_80EEALoanAmt > 0 And end_80EEALoanOutstanding > 0 And end_80EEAIntrst > 0
'end_80EEA = WorksheetFunction.Max(0, end_80EEALoanfrm, end_80EEAbankName, end_80EEAIFSC, end_80EEAPAN, end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)
end_80EEA = WorksheetFunction.Max(0, end_80EEALoanfrm, end_80EEAbankName, end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)
If Data80EEA = True Then

    setTblinfo_24bLoanfrm
   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'   setTblinfo_24bIFSC
'   setTblinfo_24bPAN
    setTblinfo_24bBankName
    setTblinfo_24bAccntNum
    setTblinfo_24bLoanDate
    setTblinfo_24bLoanAmt
    setTblinfo_24bLoanOutstanding
    setTblinfo_24bIntrst
    
    'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    Data24b = end_24bLoanfrm > 0 And end_24bIFSC > 0 Or end_24bPAN > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
'    end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bIFSC, end_24bPAN, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
    
     Data24b = end_24bLoanfrm > 0 And end_24bbankName > 0 And end_24bAccntNum > 0 And end_24bLoanDate > 0 And end_24bLoanAmt > 0 And end_24bLoanOutstanding > 0 And end_24bIntrst > 0
     end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
   
'    rowcount_Int24B = getRowNo(Sheet16.Range("LoanfrmBankOrInstitute.24b").name)
'
'    If end_24b = 1 Then
'       end_24b = rowcount_Int24B
'    Else
'       end_24b = ((end_24b - 1) + rowcount_Int24B)  'Newly Update on 02/05/2025
'    End If
    
    'Ankita_24/01/2026===========
    
    rowcount_80EEA = getRowNo(Sheet17.Range("LoanfrmBankOrInstitute.80EEA").name)
    
    If end_80EEA = 1 Then
       end_80EEA = rowcount_80EEA
    Else
       end_80EEA = ((end_80EEA - 1) + rowcount_80EEA)
    End If
    
    
    
    Dim Dedn_80EEA, Dedn_24B
    
    Dedn_80EEA = Sheet17.Range("Combination_80EEA").Column
'    Dedn_24B = Sheet16.Range("Combination_24B").Column
    
    
'    If Data24b = True Then
'        Combination_80EEA_24B = False
'        For i = rowcount_80EEA To end_80EEA
'            For j = rowcount_Int24B To end_24b
''            Debug.Print Sheet17.Cells(i, Dedn_80EEA).Value
''            Debug.Print Sheet16.Cells(j, Dedn_24B).Value
'
'                If UCase(Trim(Sheet17.Cells(i, Dedn_80EEA).Value)) = UCase(Trim(Sheet16.Cells(j, Dedn_24B).Value)) Then
'                   Combination_80EEA_24B = True
'                   Exit For
'
'                End If
'            Next
'
'        Next
'    Else
'        Combination_80EEA_24B = False
'        'MsgBox_80EEA = MsgBox_80EEA & ("Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EE is being claimed should be the same as in schedule 24(b)") & Chr(13)
'        'fmsgboxStatus MsgBox_80EEA
'    End If
''End If

    'Ankita_24/01/2026===========

    If Data24b = True Then
        Combination_80EEA_24B = False
         For HpCount = 1 To Sheet19.Range("PropertySectionCOunt").Value
               Dedn_24B = Sheet19.Range("Combination_24B" & HpCount).Column
               
               rowcount_Int24B = getRowNo(Sheet19.Range("LoanfrmBankOrInstitute.24b" & HpCount).name)
               end_24b = WorksheetFunction.Max(0, end_24bLoanfrm, end_24bbankName, end_24bAccntNum, end_24bLoanDate, end_24bLoanAmt, end_24bLoanOutstanding, end_24bIntrst)
    
                If end_24b = 1 Then
                    end_24b = rowcount_Int24B
                Else
                    end_24b = ((end_24b - 1) + rowcount_Int24B)
                End If
                For i = rowcount_80EEA To end_80EEA
                    For j = rowcount_Int24B To end_24b
                        If UCase(Trim(Sheet17.Cells(i, Dedn_80EEA).Value)) = UCase(Trim(Sheet19.Cells(j, Dedn_24B).Value)) Then
                           Combination_80EEA_24B = True
                           GoTo checking
                        End If
                    Next
                Next
           Next
    Else
        Combination_80EEA_24B = False
    End If
'End If

checking:

If Combination_80EEA_24B = False Then
        sourceSheet.Activate  'Ankita_13/05/2025
        MsgBox_80EEA = MsgBox_80EEA & ("""Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted. So please note that the loan details of respective property under which 80EEA is being claimed should be the same as in schedule 24(b)""") & Chr(13)
        fmsgboxStatus MsgBox_80EEA
        Validate80EEA_All = False
        CloseMsg
        'Exit Function
End If

End If

'Ankita_26/01/2026========

Dim k, Total As Long

    For k = 1 To Sheet19.Range("PropertySectionCount").Value
       Total = Total + Sheet19.Range("TotAmt.24b" & k).Value
    
    Next

'Ankita_21/04
            
'  If (Sheet17.Range("TotAmt.80EEA").Value > 0 And Not Sheet1.Range("IncD.InterestBorrowedCapital").Value > 0) Then
   If (Sheet17.Range("TotAmt.80EEA").Value > 0 And Not Total > 0) Then
         sourceSheet.Activate  'Ankita_13/05/2025
        MsgBox_80EEA = MsgBox_80EEA & " ""Deduction u/s 80EEA can be claimed only if the limit u/s 24(b) is exhausted.""" & Chr(13)
        fmsgboxStatus MsgBox_80EEA
       CloseMsg
    End If
    
    
'Ayush_21/04
 If MsgBox_80EEA <> "" Then
  sourceSheet.Activate  'Ankita_13/05/2025
  Validate80EEA_All = False
  Else
  sourceSheet.Activate  'Ankita_13/05/2025
  Validate80EEA_All = True
 End If
  
  '-------------------------------
            
End Function



Function Validate_80EEA()
Validate_80EEA = True
end_80EEA = 0
MsgBox_80EEA = ""

setTblinfo_80EEALoanfrm
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'setTblinfo_80EEAIFSC
'setTblinfo_80EEAPAN
setTblinfo_80EEABankName
setTblinfo_80EEAAccntNum
setTblinfo_80EEALoanDate
setTblinfo_80EEALoanAmt
setTblinfo_80EEALoanOutstanding
setTblinfo_80EEAIntrst

'end_80EEAPAN
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'end_80EEA = WorksheetFunction.Max(0, end_80EEA, end_80EEALoanfrm, end_80EEAbankName, , end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)
'end_80EEA = Application.WorksheetFunction.Max(0, end_80EEALoanfrm, end_80EEAbankName, end_80EEAPAN, end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)
end_80EEA = Application.WorksheetFunction.Max(0, end_80EEALoanfrm, end_80EEAbankName, end_80EEAAccntNum, end_80EEALoanDate, end_80EEALoanAmt, end_80EEALoanOutstanding, end_80EEAIntrst)



'Ankita_15/05/2025
If Sheet17.Range("Stampduty.80EEA").Value > 0 And end_80EEA = 0 Then
     MsgBox_80EEA = MsgBox_80EEA & " ""Please provide details in respect of interest on loan taken for house property""" & Chr(13)
        Validate_80EEA = False
End If

If end_80EEA > 0 Then
    If (Sheet17.Range("Stampduty.80EEA").Value) = "" Then
    'Ankita_05/05/2025_Commented as per DESheet_v0.7
'       MsgBox_80EEA = MsgBox_80EEA & " ""The Stamp duty Value of property"" is mandatory in schedule 80EEA" & Chr(13)
        MsgBox_80EEA = MsgBox_80EEA & " ""Stamp value of residential house property"" is mandatory in schedule 80EEA" & Chr(13)
        Validate_80EEA = False
    End If

     If (Sheet17.Range("Stampduty.80EEA").Value) > 4500000 Then
       'Ankita_05/05/2025_Commented as per DESheet_v0.7
'       MsgBox_80EEA = MsgBox_80EEA & " ""Stamp duty Value of property"" shall not be more than Rs. 45 Lakhs in schedule 80EEA" & Chr(13)
        MsgBox_80EEA = MsgBox_80EEA & " ""Stamp value of residential house property"" shall not be more than Rs. 45 Lakhs in schedule 80EEA" & Chr(13)
        Validate_80EEA = False
     End If
 End If
If Not ValidateLoanfrm_80EEA Then Validate_80EEA = False
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'If Not ValidateIFSC_80EEA Then Validate_80EEA = False
'If Not ValidatePAN_80EEA Then Validate_80EEA = False
If Not ValidateBankName_80EEA Then Validate_80EEA = False
If Not ValidateAccntNum_80EEA Then Validate_80EEA = False
If Not ValidateLoanDate_80EEA Then Validate_80EEA = False
If Not ValidateLoanAmt_80EEA Then Validate_80EEA = False

'    If (Sheet17.Range("MaxLoan_80EEA").value) > 3500000 Then
'        MsgBox_80EEA = MsgBox_80EEA & "* ""Total loan taken"" shall not be more than Rs. 35 Lakhs in schedule 80EE" & Chr(13)
'        Validate_80EEA = False
'    End If

If Not ValidateLoanOutstanding_80EEA Then Validate_80EEA = False
If Not ValidateIntrst_80EEA Then Validate_80EEA = False
'If Not ValidateLoanfrm_80EEA Then Validate_80EEA = False

'ValidateIFSC_80EEA

'    If (Sheet17.Range("TotAmt.80EEA").Value) > 2000000 Then
'        MsgBox_80EEA = MsgBox_80EEA & " To claim ""Interest payable on borrowed capital more than Rs.20 lakhs"", you may please consider filing ITR 3. Refer Rule 12 for further details" & Chr(13)
'        Validate_80EEA = False
'    End If
    
     If Len(Sheet17.Range("TotAmt.80EEA").Value) > 14 Then
        MsgBox_80EEA = MsgBox_80EEA & "* Total of Interest u/s 80EEA in schedule 80EEA cannot exceed 14 digits" & Chr(13)
        Validate_80EEA = False
    End If
    
End Function


Sub setTblinfo_80EEALoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Sheet17.Activate
    mIntCells = Range("LoanfrmBankOrInstitute.80EEA").count
    Set rangecells = Range("LoanfrmBankOrInstitute.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "" Or rangecells.item(mIntCtr).Value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    end_80EEALoanfrm = ccount
       
End Sub

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Sub setTblinfo_80EEAIFSC()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'     Sheet17.Activate
'    mIntCells = Range("IFSC.80EEA").count
'    Set rangecells = Range("IFSC.80EEA").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EEAIFSC = ccount
'     end_80EEAIFSC_1 = ccount
'End Sub


Sub setTblinfo_80EEABankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
     Sheet17.Activate
    mIntCells = Range("bankName.80EEa").count
    Set rangecells = Range("bankName.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEAbankName = ccount
      
End Sub

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Sub setTblinfo_80EEAPAN()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'     Sheet17.Activate
'    mIntCells = Range("PAN.80EEA").count
'    Set rangecells = Range("PAN.80EEA").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80EEAPAN = ccount
'
'End Sub

Sub setTblinfo_80EEAAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
     Sheet17.Activate
    mIntCells = Range("loanAccNum.80EEA").count
    Set rangecells = Range("loanAccNum.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEAAccntNum = ccount
    
End Sub
Sub setTblinfo_80EEALoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
     Sheet17.Activate
    mIntCells = Range("loanDate.80EEA").count
    Set rangecells = Range("loanDate.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEALoanDate = ccount
   
End Sub
Sub setTblinfo_80EEALoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
     Sheet17.Activate
    mIntCells = Range("loanAmt.80EEA").count
    Set rangecells = Range("loanAmt.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEALoanAmt = ccount
    
End Sub

Sub setTblinfo_80EEALoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
     Sheet17.Activate
    mIntCells = Range("loanOutstanding.80EEA").count
    Set rangecells = Range("loanOutstanding.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEALoanOutstanding = ccount
     
End Sub
Sub setTblinfo_80EEAIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
     Sheet17.Activate
    mIntCells = Range("Intrst.80EEA").count
    Set rangecells = Range("Intrst.80EEA").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80EEAIntrst = ccount
     
End Sub



Function ValidateLoanfrm_80EEA() As Boolean
    ValidateLoanfrm_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Cells
    ReDim Loanfrm_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        Loanfrm_80EEA(i) = rangecells.item(i).Value
        'Ankita_12/05/2025
'        If Not chkCompulsory(Loanfrm_80EEA(i)) Then
        If isdropdownblank(Loanfrm_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* Please select dropdown from ""Loan taken from"" in schedule 80EEA at Sr. No " & i & "" & Chr(13)   'Ankita_28/04/2025  'Changed as per DE sheet v0.8 by Ankita
            ValidateLoanfrm_80EEA = False
            Exit Function
        End If
'         If Loanfrm_80EEA(i) = "(Select)" Then
'          MsgBox_80EEA = MsgBox_80EEA + "* ""Please select dropdown ""Loan taken from"" in schedule 80EEA"" at Sr. No " & i & "" & Chr(13)  'Ankita_28/04/2025
'            ValidateLoanfrm_80EEA = False
'            Exit Function
'        End If
         
Next
End Function

'Ankita_05/05/2025_Commented as per DESheet_v0.7
'Function ValidateIFSC_80EEA() As Boolean
'ValidateIFSC_80EEA = True
'
'   ' setTblinfo_80EEAIFSC
'   setTblinfo_80EEALoanfrm
'    Dim rangecells As Range
''    Dim rangecells1 As Range
'    Set rangecells = Range("LoanfrmBankOrInstitute.80EEA").Cells
''    Set rangecells1 = Range("IFSC.80EEA").Cells
'    Dim cellrange As String
'    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80EEALoanfrm)
'    ReDim Others_IFSC(end_80EEALoanfrm)
'
'
'    For i = 1 To end_80EEALoanfrm
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
'             MsgBox_80EEA = MsgBox_80EEA & "* Please provide IFSC of the Bank from which loan is taken in schedule 80EEA at Sr. No " & i & "" & Chr(13)
'         'End Change.IDS113
'             ValidateIFSC_80EEA = False
'             Exit Function
'         End If
'         End If
'
'         If Len(Others_IFSC(i)) > 11 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* IFSC of the Bank at Sr. No " & i & " in Sheet 80EEA should be 11 characters." & Chr(13)
'            ValidateIFSC_80EEA = False
'            Exit Function
'        End If
'
'    Next
'
'End Function

'Ankita_05/05/2025_Commented as per DESheet_v0.7
''
'Function ValidatePAN_80EEA() As Boolean
'ValidatePAN_80EEA = True
'
'   ' setTblinfo_80EEAIFSC
'   setTblinfo_80EEALoanfrm
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Set rangecells = Range("LoanfrmBankOrInstitute.80EEA").Cells
'    Set rangecells1 = Range("PAN.80EEA").Cells
'    Dim cellrange As String
'    Dim cellRange1 As String
'    Dim i As Long
'    ReDim Others_Loan(end_80EEALoanfrm)
'    ReDim Others_PAN(end_80EEALoanfrm)
'
'
'
'    For i = 1 To end_80EEALoanfrm
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
'             MsgBox_80EEA = MsgBox_80EEA & "* Please provide PAN of the institution from which loan is taken in schedule 80EEA at Sr. No " & i & "" & Chr(13)
'         'End Change.IDS113
'             ValidatePAN_80EEA = False
'             Exit Function
'         End If
'         End If
'
'         If Len(Others_PAN(i)) > 10 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* PAN of the institution at Sr. No " & i & " in Sheet 80EEA should be 10 characters." & Chr(13)
'            ValidatePAN_80EEA = False
'            Exit Function
'        End If
'
'        'Added by Malli(24/04/2025)NewDev
'        If Others_PAN(i) <> "" Then
'          If Not EfilingCommon.CheckPAN(UCase(Trim(Others_PAN(i)))) Then
'          MsgBox_80EEA = MsgBox_80EEA + "* Invalid PAN. PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet in schedule 80EEA at Sr. No " & i & "" & Chr(13)
'            ValidatePAN_80EEA = False
'            Exit Function
'          End If
'         End If
'
'    Next
'
'End Function


Function ValidateBankName_80EEA() As Boolean
    ValidateBankName_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("bankName.80EEA").Cells
    ReDim BankName_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        BankName_80EEA(i) = rangecells.item(i).Value
        If Not chkCompulsory(BankName_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* Please provide ""Name of the Bank / Institution from which the loan is taken"" in schedule 80EEA at Sr. No " & i & "" & Chr(13)
            ValidateBankName_80EEA = False
            Exit Function
        End If
         If Len(BankName_80EEA(i)) > 125 Then
            MsgBox_80EEA = MsgBox_80EEA + "* Name of the Bank / Institution in Sheet 80EEA should be less than or equal to 125 characters at Sr. No " & i & " ." & Chr(13)
            ValidateBankName_80EEA = False
            Exit Function
        End If
         If Not checkfieldSuperSpecialcharacter(BankName_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* Name of the Bank / Institution from which the loan is taken in schedule 80EEA at Sl.no. " & i & " should not Contain <, >, characters." & Chr(13)
            ValidateBankName_80EEA = False
            Exit Function
        End If
Next
End Function

'Function ValidateIFSC_80EEA() As Boolean
'    ValidateIFSC_80EEA = True
'
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Dim i As Long
'    Set rangecells = Sheet17.Range("IFSC.80EEA").Cells
'    Set rangecells1 = Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Cells
'    ReDim IFSC_80EEA(end_80EEA)
'    ReDim Loanfrom_80EEA(end_80EEA)
'    For i = 1 To end_80EEA
'        IFSC_80EEA(i) = rangecells.item(i).value
'       ' Loanfrom_80EEA(i) = rangecells.item(i).value
'     '   If Loanfrom_80EEA(i) = "Bank" Then
'        If Not chkCompulsory(IFSC_80EEA(i)) Then
'             MsgBox_80EEA = MsgBox_80EEA + "* Please select dropdown from ""Loan taken from"" at Sr. No " & i & " in 80EEA schedule." & Chr(13)
'            ValidateIFSC_80EEA = False
'            Exit Function
'        End If
'      '  End If
'
'         If Len(IFSC_80EEA(i)) > 11 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* IFSC at Sr. No " & i & " in Sheet 80EEA should be 11 characters." & Chr(13)
'            ValidateIFSC_80EEA = False
'            Exit Function
'        End If
'
'Next
''
''If (Sheet20.Range("LoanfrmBankOrInstitute.80EE").value = "Bank") Then
''        If Sheet20.Range("IFSC.80EE").value = "" Then
'''            msgError = msgError & "*If A23a is ""Yes"" Date of filing of Form 10IEA is mandatory. " & Chr(13)
'''Ankita_10/02
'''            msgError = msgError & "*""Date of filing of Form 10IEA is mandatory.""" & Chr(13) 'Modified by sai on 27/01/2025
''            MsgBox_80EEA = MsgBox_80EEA & "*""IFSC""" & Chr(13)
''            ValidateIFSC_80EEA = False
''        End If
''        End If
'
'End Function
Function ValidateAccntNum_80EEA() As Boolean
    ValidateAccntNum_80EEA = True

    Dim rangecells, rangecells1 As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanAccNum.80EEA").Cells
    Set rangecells1 = Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Cells
    
    ReDim AccntNum_80EEA(end_80EEA)
    ReDim Bankorothebank_80EEA(end_80EEA)
    
    
    For i = 1 To end_80EEA
        AccntNum_80EEA(i) = rangecells.item(i).Value
        Bankorothebank_80EEA(i) = rangecells1.item(i).Value
        
        If Not chkCompulsory(AccntNum_80EEA(i)) Then
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'            MsgBox_80EEA = MsgBox_80EEA + "* Please provide ""Loan Account number of the Bank / Reference number of the Institution"" from which loan is taken in schedule 80EEA at Sr. No " & i & "" & Chr(13)
             MsgBox_80EEA = MsgBox_80EEA + "* Please provide ""Loan Account number of the Bank /Institution"" from which loan is taken in schedule 80EEA at Sr. No " & i & "" & Chr(13)
            ValidateAccntNum_80EEA = False
            Exit Function
        End If
         If Len(AccntNum_80EEA(i)) > 20 Then
            MsgBox_80EEA = MsgBox_80EEA + "* Loan Account number in Sheet 80EEA should be less than or equal to 20 characters at Sr. No " & i & "." & Chr(13)
            ValidateAccntNum_80EEA = False
            Exit Function
        End If
        
        If Bankorothebank_80EEA(i) = "Bank" Then
        If Not ValidateBankAccountNumber_80EEA(AccntNum_80EEA(i), i) Then
        'Msgbox_BA = Msgbox_BA + "Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntNum_80EEA = False
        Exit Function
        End If
        End If
        
        'Doubt_05/05/2025
        If Bankorothebank_80EEA(i) = "Institution" Then
         If Not checkfieldspecialcharacter_BthacntReferencenumber(AccntNum_80EEA(i)) Then
            MsgBox_80EEA = MsgBox_80EEA + "* Loan Account number of the Bank / Institution at Sr.No " & i & " is invalid  in schedule 80EEA, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ' MsgBox_80E = MsgBox_80E & " Refrence number of the Institution at Sr.No " & i & " is invalid  in schedule 80E, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateAccntNum_80EEA = False
           
        Exit Function
        End If
        End If
         
Next
End Function

Function ValidateLoanDate_80EEA() As Boolean
    ValidateLoanDate_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanDate.80EEA").Cells
    ReDim LoanDate_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        LoanDate_80EEA(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanDate_80EEA(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
             MsgBox_80EEA = MsgBox_80EEA + "* ""Please provide ""Date of sanction of loan"" in schedule 80EEA at Sr. No " & i & """" & Chr(13)
            ValidateLoanDate_80EEA = False
            Exit Function
        End If
        
         If Len(LoanDate_80EEA(i)) > 10 Then
         'Ankita_05/05/2025_Commented as per DESheet_v0.7
          MsgBox_80EEA = MsgBox_80EEA + "* Date of sanction loan at Sr. No " & i & " in Sheet 80EEA less than 10 characters." & Chr(13)
            ValidateLoanDate_80EEA = False
            Exit Function
        End If
        
        'CheckDateddmmyyyy
         If Not CheckDateddmmyyyy(LoanDate_80EEA(i)) Then
         MsgBox_80EEA = MsgBox_80EEA + "* ""Please enter date in valid format"" at Sr. No " & i & "." & Chr(13)
        ValidateLoanDate_80EEA = False
        Exit Function
        End If
        'ChkMaxDate_80EEA_1
       ' If Not ChkMaxDate_80EE(Trim(LoanDate_80EEA(i)), "31-03-2022") Then
         If Not ChkMaxDate_80EEA_1(Trim(LoanDate_80EEA(i)), "31-03-2022") Then
            MsgBox_80EEA = MsgBox_80EEA + " ""Date of sanction of Loan"" shall be between 01/04/2019 to 31/03/2022 in schedule 80EEA at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EEA = False
                 

          Exit Function
        End If
        
        If Not ChkMinInclusiveDate(Trim(Dformat(LoanDate_80EEA(i), "YYYY-MM-DD")), "2019-04-01") Then
           MsgBox_80EEA = MsgBox_80EEA + " ""Date of sanction of Loan"" shall be between 01/04/2019 to 31/03/2022 in schedule 80EEA at Sr. No " & i & "." & Chr(13)
            ValidateLoanDate_80EEA = False
                 

          Exit Function
        End If
                 
                 
Next
End Function

Function ValidateLoanAmt_80EEA() As Boolean
    ValidateLoanAmt_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanAmt.80EEA").Cells
    ReDim LoanAmt_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        LoanAmt_80EEA(i) = rangecells.item(i).Value
        'Ankita_28/04/2025
        If Not chkCompulsory(LoanAmt_80EEA(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
             'MsgBox_80EEA = MsgBox_80EEA + "* Please provide Total Loan taken in schedule 80EEA at Sr. No " & i & "" & Chr(13)
              MsgBox_80EEA = MsgBox_80EEA + "* ""Please provide ""Total amount of Loan"" in schedule 80EEA"" at Sr. No " & i & "" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If
'         If Len(LoanAmt_80EEA(i)) > 14 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* Loan Amount  at Sr. No " & i & " in Sheet 80EEA less than 15 characters." & Chr(13)
'            ValidateLoanAmt_80EEA = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanAmt_80EEA(i)) Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan amount at Sr. No  " & i & " in schedule 80EEA should be Numeric value" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If
        
        If LoanAmt_80EEA(i) > 99999999999999# Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan amount at Sr. No  " & i & " in schedule 80EEA cannot exceed 14 digits" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
        If LoanAmt_80EEA(i) < 0 Or LoanAmt_80EEA(i) = 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* ""Total amount of Loan"" should be more than 0 in schedule 80EEA at Sr. No  " & i & "" & Chr(13)
            ValidateLoanAmt_80EEA = False
            Exit Function
        End If
         
Next
End Function

Function ValidateLoanOutstanding_80EEA() As Boolean
    ValidateLoanOutstanding_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("loanOutstanding.80EEA").Cells
    ReDim LoanOutStanding_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        LoanOutStanding_80EEA(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanOutStanding_80EEA(i)) Then
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
             MsgBox_80EEA = MsgBox_80EEA + "* ""Loan outstanding as on last date of financial year"" is mandatory in schedule 80EEA at Sr. No " & i & "" & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
'         If Len(LoanOutstanding_80EEA(i)) > 14 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* Loan Outstanding  at Sr. No " & i & " in Sheet 80EEA less than 15 characters." & Chr(13)
'            ValidateLoanOutstanding_80EEA = False
'            Exit Function
'        End If
        
        If Not IsNumeric(LoanOutStanding_80EEA(i)) Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan outstanding at Sr. No  " & i & "  in schedule 80EEA should be Numeric value" & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
        
        If LoanOutStanding_80EEA(i) > 99999999999999# Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan outstanding at Sr. No  " & i & "  in schedule 80EEA cannot exceed 14 digits" & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
        'Ankita_05/05/2025_Commented as per DESheet_v0.7

        If LoanOutStanding_80EEA(i) < 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan outstanding as on last date of financial year can't be less than 0 in schedule 80EEA at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment." & Chr(13)
            ValidateLoanOutstanding_80EEA = False
            Exit Function
        End If
         
Next
End Function

Function ValidateIntrst_80EEA() As Boolean
    ValidateIntrst_80EEA = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet17.Range("Intrst.80EEA").Cells
    ReDim Intrst_80EEA(end_80EEA)
    For i = 1 To end_80EEA
        Intrst_80EEA(i) = rangecells.item(i).Value
        If Not chkCompulsory(Intrst_80EEA(i)) Then
             MsgBox_80EEA = MsgBox_80EEA + "* ""Please provide Interest u/s 80EEA"" at Sr. No " & i & "" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If
'         If Len(Intrst_80EEA(i)) > 15 Then
'          MsgBox_80EEA = MsgBox_80EEA + "* Interest  at Sr. No " & i & " in Sheet 80EEA less than 15 characters." & Chr(13)
'            ValidateIntrst_80EEA = False
'            Exit Function
'        End If

        If Not IsNumeric(Intrst_80EEA(i)) Then
            MsgBox_80EEA = MsgBox_80EEA & "* Interest at Sr. No  " & i & " in schedule 80EEA should be Numeric value" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If
        
        If Intrst_80EEA(i) > 99999999999999# Then
            MsgBox_80EEA = MsgBox_80EEA & "* Interest at Sr. No  " & i & " in schedule 80EEA cannot exceed 14 digits" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If
        
        'Interest u/s 80EEA should be more than 0 in schedule 80EEA
        
        If Intrst_80EEA(i) < 0 Or Intrst_80EEA(i) = 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* ""Interest u/s 80EEA"" should be more than 0 in schedule 80EEA at Sr. No  " & i & "" & Chr(13)
            ValidateIntrst_80EEA = False
            Exit Function
        End If
         
Next
End Function

'Need to check
Function ValidateBankAccountNumber_80EEA(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_80EEA = True
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
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80EEA = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEA " & Chr(13)
            ValidateBankAccountNumber_80EEA = False
            Exit Function
        End If
    
'        If IsNumeric(BankAccountNumber) Then
'            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid  in schedule 80EEA " & Chr(13)
'            ValidateBankAccountNumber_80EEA = False
'            Exit Function
'        End If
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80EEA = MsgBox_80EEA & "*  Please enter the Loan Account number in Bank Details at Sr.No " & cc & " in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is mandatory in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc & " is invalid in schedule 80EEA" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
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
            MsgBox_80EEA = MsgBox_80EEA & "* Loan Account number at Sr.No " & cc + 1 & " is invalid in schedule 80EEA" & Chr(13)
            ValidateBankAccountNumber_80EEA = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Loan Account number at Sr.No " & cc + 1 & "  in schedule 80EEA is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80EEA = False
        Exit Function
    End If
End Function


Public Function ChkMaxDate_80EEA(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_80EEA = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ankita_25-26
'        If Year > 2024 Then
        If Year > 2025 Then
            ChkMaxDate_80EEA = False
            Exit Function
        Else
'            If Year = 2024 Then
                If Year = 2025 Then
                If month > 4 Then
                    ChkMaxDate_80EEA = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_80EEA = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_80EEA = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function



'Malli-----
Public Function ChkMaxDate_80EEA_1(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_80EEA_1 = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ayush_25-26
'        If Year > 2024 Then
        If Year > 2022 Then
            ChkMaxDate_80EEA_1 = False
            Exit Function
        Else
'            If Year = 2024 Then
                If Year = 2022 Then
                If month > 4 Then
                   ChkMaxDate_80EEA_1 = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                          ChkMaxDate_80EEA_1 = False
                            Exit Function
                        Else
                            If dat = 1 Then
                              ChkMaxDate_80EEA_1 = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function
