Attribute VB_Name = "mdATI"
Public msgbox_1398AATI As String
Public ColCount2_IT1 As Long
Public ColCount2_1_IT1 As Long
Public ColCount2_2_IT1 As Long
Public ColCount2_3_IT1 As Long
Public ColCount2_IT2 As Long
Public ColCount2_1_IT2 As Long
Public ColCount2_2_IT2 As Long
Public ColCount2_3_IT2 As Long
Function ChkMinInclusiveDate_1398A(Mininclusive As Variant, Mininclusivedate As Variant) As Boolean
'Note:
'' both date must be in format yyyy-mm-dd
    ChkMinInclusiveDate_1398A = True
    If Len(Mininclusive) > 0 Then
        If Mid(Mininclusive, 1, 4) < Mid(Mininclusivedate, 1, 4) Then
            ChkMinInclusiveDate_1398A = False
            Exit Function
        Else
            If Mid(Mininclusive, 1, 4) = Mid(Mininclusivedate, 1, 4) Then
                If (Mid(Mininclusive, 6, 2) < Mid(Mininclusivedate, 6, 2)) Then
                    ChkMinInclusiveDate_1398A = False
                    Exit Function
                ElseIf ((Mid(Mininclusive, 6, 2) = Mid(Mininclusivedate, 6, 2))) Then
                    If (Mid(Mininclusive, 9, 2) < Mid(Mininclusivedate, 9, 2)) Then
                        ChkMinInclusiveDate_1398A = False
                        Exit Function
                   End If
                End If
            End If
        End If
    End If
  
End Function
Function checkfieldspecialcharacter_BsRCode_1398A(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_BsRCode_1398A = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "_", "|", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_BsRCode_1398A = False
            Exit Function
        End If
        Next
    Next
End Function
Function chkCompulsory_1398A(field As Variant) As Boolean
    chkCompulsory_1398A = True
    If Len(Trim(field)) <= 0 Or IsEmpty(field) Then
        chkCompulsory_1398A = False
    End If
End Function
Sub Cmd_Validate_ATI_Click()
ValidateATI
fmsgboxStatus_1398A "Sheet Part B ATI is ok"

End Sub
Sub ValidateATI()

subProcCaption = "Validating Part B ATI"

If Not Validatesheet_ATI Then
fmsgboxsmall_1398A (msgbox_1398AATI)
CloseMsg
End If

If Not ValidateSheetATI_IT_1 Then
fmsgboxsmall_1398A (msgbox_1398AATI)
CloseMsg
End If

If Not ValidateSheetATI_IT_2 Then
fmsgboxsmall_1398A (msgbox_1398AATI)
CloseMsg
End If

End Sub
Function Validatesheet_ATI() As Boolean
Validatesheet_ATI = True
'msgbox_1398AATI = "Part B ATI Ok " & Chr(10)
If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then

    If Sheet202.Range("U_TaxDue10_11").Value > 0 Then
        Validatesheet_ATI = False
        msgbox_1398AATI = msgbox_1398AATI + "* Make sure you have no taxes due before filing the updated return" & Chr(13)
    End If
    
    'Ankita_UR
        If Sheet201.Range("U_PreviouslyFiledForThisAY").Value = "Yes" Then
        If Sheet202.Range("U_LatestTotInc").Value = "" Then
           Sheet202.Unprotect Password:=getmsgstate
           msgbox_1398AATI = msgbox_1398AATI + """Since ""Yes"" is selected for A5 ""Whether return previously filed for this assessment year?"" in Schedule Part A Gen_139(8A), ""Total income as per last Valid return"" is required to be filled manadtorily.""" & Chr(13)
           Validatesheet_ATI = False
        
        End If
     End If


  If Sheet202.Range("U_NetPayable").Value <= 0 Then
        msgbox_1398AATI = msgbox_1398AATI + "* ""Net Amount Payable at sl.no.11 should be more than ""0"" to file ITR u/s 139(8A)""" & Chr(13)
        Validatesheet_ATI = False
    End If
    
    If Sheet202.Range("U_TaxUS140B").Value <= 0 Then
        msgbox_1398AATI = msgbox_1398AATI + "* ""Tax Paid u/s 140B at sl.no.12 should be more than ""0"" to file ITR u/s 139(8A)""" & Chr(13)
        Validatesheet_ATI = False
    End If
    
    
    If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
 
        If Sheet202.Range("U_LastAmtPayable").Value > 0 Then
           msgbox_1398AATI = msgbox_1398AATI + "If A5 in Part A gen 139(8A)  is ""No"",then ""Amount payable as per prior ITR can’t be more than 0""" & Chr(13)
           Validatesheet_ATI = False
        End If
 
     End If
 

     
     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
        If Sheet202.Range("U_Refund").Value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           msgbox_1398AATI = msgbox_1398AATI + "If A5 in Part A gen 139(8A)  is ""No"",then ""Refund claimed can’t be more than 0""" & Chr(13)
           Validatesheet_ATI = False
        End If
 
     End If
      
 
     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
        If Sheet202.Range("U_TotRefund").Value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           msgbox_1398AATI = msgbox_1398AATI + "If A5 in Part A gen 139(8A)  is ""No"", then ""Refund issued can’t be more than 0""" & Chr(13)
           Validatesheet_ATI = False
        End If
 
     End If
       
 
     If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
        If Sheet202.Range("U_RegAssessementTAX").Value > 0 Then
           Sheet202.Unprotect Password:=getmsgstate
           msgbox_1398AATI = msgbox_1398AATI + "If A5 in Part A gen 139(8A)  is ""No"", then ""Regular Assessment Tax"" can't be claimed""" & Chr(13)
           Validatesheet_ATI = False
        End If
 
     End If
     
    '=======================================
    If Len(Sheet202.Range("U_Total")) > 14 Then
        Validatesheet_ATI = False
        msgbox_1398AATI = msgbox_1398AATI + "* Length of 1Af in Schedule Part B-ATI should not be more than 14 characters."
    End If
     If Sheet202.Range("U_AmtPayable").Value < Sheet202.Range("U_LastAmtPayable").Value Then
        msgbox_1398AATI = msgbox_1398AATI + "* Amount at Sr No 3 cannot be less than Amount at Sr No 5"
        Validatesheet_ATI = False
    End If
     If Sheet202.Range("U_AmtRefundable").Value > Sheet202.Range("U_Refund").Value Then
        msgbox_1398AATI = msgbox_1398AATI + "* Amount at Sr No 4 cannot be greater than Amount at Sr No 6i"
        Validatesheet_ATI = False
    End If
End If
End Function

'SUB FOR ATI
Sub Cmd_AddRows_IT1_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("U_BSRCode1||U_DateDep1||U_SrlNoOfChaln1||U_Amt1||U_DateOfAmount1")
ActiveCellRange = EfilingCommon.searchLastRow("U_BSRCode1")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub


Sub Cmd_AddRows_IT2_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("U_BSRCode2||U_DateDep2||U_SrlNoOfChaln2||U_Amt2")
ActiveCellRange = EfilingCommon.searchLastRow("U_BSRCode2")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub


Function ValidateSheetATI_IT_1() As Boolean
ValidateSheetATI_IT_1 = True
    setTableInfo_Grid3_IT1
    setTableInfo1_Grid3_IT1
    setTableInfo2_Grid3_IT1
    setTableInfo3_Grid3_IT1
    'Adding personal function for validation
    '----
    
    If Len((Sheet202.Range("U_BSRCode1").item(1).Value) > 0) Then
        If Not ValidateBSR_IT1 Then ValidateSheetATI_IT_1 = False
        If Not ValidateDate_IT1 Then ValidateSheetATI_IT_1 = False
        If Not ValidateSerialNum_IT1 Then ValidateSheetATI_IT_1 = False
        If Not ValidateTaxPaid_IT1 Then ValidateSheetATI_IT_1 = False
        If Not ValidateMandatoryShIT1 Then ValidateSheetATI_IT_1 = False
    End If
    
    
    If ((ColCount2_IT1 <> ColCount2_1_IT1) Or (ColCount2_IT1 <> ColCount2_2_IT1) Or (ColCount2_IT1 <> ColCount2_3_IT1)) Then
    msgbox_1398AATI = msgbox_1398AATI + "* Enter All Mandatory Details in PartB ATI"
    ValidateSheetATI_IT_1 = False
    End If
     
End Function
Function ValidateMandatoryShIT1() As Boolean
ValidateMandatoryShIT1 = True
Dim i As Long
Dim flag As Boolean
flag = True
For i = 1 To Sheet202.Range("U_BSRCode1").Rows.count
'    If Sheet202.Range("U_BSRCode1").item(i).Value <> "" Or Sheet202.Range("U_DateDep1").item(i).Value <> "" Or _
'        Sheet202.Range("U_SrlNoOfChaln1").item(i).Value <> "" Or Sheet202.Range("U_Amt1").item(i).Value <> "" Then
        
      'Ankita_16/12/2025========
        If Sheet202.Range("U_BSRCode1").item(i).Value <> "" Or Sheet202.Range("U_DateDep1").item(i).Value <> "" Or _
        Sheet202.Range("U_SrlNoOfChaln1").item(i, 0).Value <> "" Or Sheet202.Range("U_Amt1").item(i).Value <> "" Then
  
        If Sheet202.Range("U_BSRCode1").item(i).Value = "" Then
            flag = False
            msgbox_1398AATI = msgbox_1398AATI + "* Please enter the BSR code in Schedule IT-1" & Chr(13)
        End If
        
'        'Ankita_UR
'        If Sheet202.Range("U_DateDep1").Cells(i, 0).Value = "" Then
'            flag = False
'            msgbox_1398AATI = msgbox_1398AATI + "* Please enter date of credit into Govt account in schedule IT" & Chr(13)
'        End If
'
'        If Sheet202.Range("U_SrlNoOfChaln1").Cells(i, 0).Value = "" Then
'            flag = False
'            msgbox_1398AATI = msgbox_1398AATI + "* Please enter serial number of challan in schedule IT"
'        End If
'
'        If Sheet202.Range("U_Amt1").Cells(i, 0).Value = "" Then
'            flag = False
'            msgbox_1398AATI = msgbox_1398AATI + "* Please enter amount in schedule IT"
'        End If
        
        If flag = False Then
            ValidateMandatoryShIT1 = False
            Exit Function
        End If
        
    End If
Next i
End Function

'Ankita 11/11/2024
Sub NextPARTBATI_Click()
If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" And Sheet5.Range("BacValue").Value = 2 Then
Sheet9.Activate
End If
End Sub

Sub PreviousPARTBATI_Click()
'Ankita  11/11/2024
If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
Sheet3.Activate
End If  'Ankita 11/11/2024
End Sub



Sub setTableInfo_Grid3_IT1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet202.Range("U_BSRCode1").count
    Set rangecells = Sheet202.Range("U_BSRCode1").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_IT1 = ccount
End Sub

Sub setTableInfo1_Grid3_IT1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet202.Range("U_DateDep1").count
    Set rangecells = Sheet202.Range("U_DateDep1").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_1_IT1 = ccount
End Sub


Sub setTableInfo2_Grid3_IT1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet202.Range("U_SrlNoOfChaln1").count
    Set rangecells = Sheet202.Range("U_SrlNoOfChaln1").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_2_IT1 = ccount
End Sub


Sub setTableInfo3_Grid3_IT1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet202.Range("U_Amt1").count
    Set rangecells = Sheet202.Range("U_Amt1").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_3_IT1 = ccount
End Sub
Function ValidateBSR_IT1() As Boolean
ValidateBSR_IT1 = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_BSRCode1").Cells
Dim i As Long
ReDim BSR_IT1(ColCount2_IT1)
noOfProcessSub = ColCount2_IT1
For i = 1 To ColCount2_IT1
BSR_IT1(i) = rangecells.item(i).Value
If Len(BSR_IT1(i)) = 0 Then
End If

'If Not chkCompulsory_1398A(BSR_IT1(i)) Then
'    msgbox_1398AATI = msgbox_1398AATI + "* Please enter the BSR Code at Sr.NO " & i & "in PartB ATI" & Chr(13)
'    ValidateBSR_IT1 = False
'    Exit Function
'End If

 If Not Validate_BSRCODE1(BSR_IT1(i)) Then
         msgbox_1398AATI = msgbox_1398AATI + "Invalid BSR code at Sr. No  " & i & "  in Sheet PartB ATI. First three digits from the left stand are for district code; next three digits stand for revenue centre code within the district and last single digit stands for population range code" & Chr(13)
         ValidateBSR_IT1 = False
        Exit Function
     End If
UpdateProgressBar
Next
End Function
Function Validate_BSRCODE1(BsrCode As Variant) As Boolean
'On Error Resume Next


    Validate_BSRCODE1 = True
    If Len(BsrCode) > 0 Then
        If Not IsNumeric(Mid(BsrCode, 1, 1)) Then
            Validate_BSRCODE1 = False
            Exit Function
        End If
        If Not IsNumeric(Mid(BsrCode, 2, 1)) Then
            Validate_BSRCODE1 = False
            Exit Function
        End If
        If Not IsNumeric(Mid(BsrCode, 3, 1)) Then
            Validate_BSRCODE1 = False
            Exit Function
        End If

        If Not checkfieldspecialcharacter_BsRCode_1398A(Mid(BsrCode, 4, 4)) Then
            Validate_BSRCODE1 = False
            Exit Function
        End If

    End If
End Function
Function ValidateDate_IT1() As Boolean
ValidateDate_IT1 = True
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_DateDep1").Cells
Dim i As Long
ReDim DateCredit_IT1(ColCount2_IT1)
For i = 1 To ColCount2_IT1
DateCredit_IT1(i) = rangecells.item(i).Value
If Len(DateCredit_IT1(i)) = 0 Then
End If
'Ankita_16/12/2025=================================

If Not FormatNCheckDate(DateCredit_IT1(i)) Then
    ValidateDate_IT1 = False
   msgbox_1398AATI = msgbox_1398AATI + "* Date of Deposit into Govt Account must be a valid dd/mm/yyyy format at Sr.NO " & i & " in schedule IT-1" & Chr(13)
    Exit Function
End If

If Not chkCompulsory_1398A(DateCredit_IT1(i)) Then
    msgbox_1398AATI = msgbox_1398AATI + "* Please enter the Date of Deposit into Govt Account at Sr.NO " & i & " in schedule IT-1" & Chr(13)
    ValidateDate_IT1 = False
    Exit Function
End If

If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT1(i), ""), "2023-01-01") Then ''
    msgbox_1398AATI = msgbox_1398AATI + "*Date of credit to Central government at Sr. No  " & i & " cannot be prior to 01/01/2023 in schedule IT-1" & Chr(13)
    ValidateDate_IT1 = False
    Exit Function
End If

Next
End Function
Function ValidateSerialNum_IT1() As Boolean
ValidateSerialNum_IT1 = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_SrlNoOfChaln1").Rows
Dim i As Long
ReDim SerialNum_IT1(ColCount2_IT1)
For i = 1 To ColCount2_IT1
SerialNum_IT1(i) = rangecells.item(i).Cells(1).Value
If Len(SerialNum_IT1(i)) = 0 Then
End If

'Ankita_04/11/2025
If Not chkCompulsory_1398A(SerialNum_IT1(i)) Then
'   msgbox_1398AATI = msgbox_1398AATI + "* Please enter the Serial Number of Challan at Sr.NO " & i & "in PartB ATI" & Chr(13)
    msgbox_1398AATI = msgbox_1398AATI + "* ""Please enter serial number of challan at Sr.NO " & i & " in schedule IT-1.""" & Chr(13)
    ValidateSerialNum_IT1 = False
    Exit Function
End If


If Not IsNumeric(SerialNum_IT1(i)) Then
'   msgbox_1398AATI = msgbox_1398AATI + "* Serial Number of Challan at Sr.NO " & i & " can allow only Numeric Values in PartB ATI " & Chr(13)
    msgbox_1398AATI = msgbox_1398AATI + "* Serial Number of Challan at Sr.NO " & i & " can allow only Numeric Values in schedule IT-1. " & Chr(13)
    ValidateSerialNum_IT1 = False
    Exit Function
End If
Next
End Function

Function ValidateTaxPaid_IT1() As Boolean
ValidateTaxPaid_IT1 = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_Amt1").Cells
Dim i As Long
ReDim TaxPaid3_IT1(ColCount2_IT1)
For i = 1 To ColCount2_IT1
TaxPaid3_IT1(i) = rangecells.item(i).Value
If Len(TaxPaid3_IT1(i)) = 0 Then
End If

'Ankita_04/11/2025===================
If Not chkCompulsory_1398A(TaxPaid3_IT1(i)) Then
'   msgbox_1398AATI = msgbox_1398AATI + "* Please enter the Tax Paid Amount at Sr.NO " & i & " in PartB ATI" & Chr(13)
    msgbox_1398AATI = msgbox_1398AATI + "* ""Please enter amount at Sr.NO " & i & " in schedule IT-1""" & Chr(13)
    ValidateTaxPaid_IT1 = False
    Exit Function
End If
Next
End Function


Function ValidateSheetATI_IT_2() As Boolean
ValidateSheetATI_IT_2 = True
    setTableInfo_Grid3_IT2
    setTableInfo1_Grid3_IT2
    setTableInfo2_Grid3_IT2
    setTableInfo3_Grid3_IT2
  
    If Len((Sheet202.Range("U_BSRCode2").item(1).Value) > 0) Then
        If Not ValidateBSR_IT2 Then ValidateSheetATI_IT_2 = False
        If Not ValidateDate_IT2 Then ValidateSheetATI_IT_2 = False
        If Not ValidateSerialNum_IT2 Then ValidateSheetATI_IT_2 = False
        If Not ValidateTaxPaid_IT2 Then ValidateSheetATI_IT_2 = False
        If Not ValidateMandatoryShIT_2 Then ValidateSheetATI_IT_2 = False
    End If
    
    
    If ((ColCount2_IT2 <> ColCount2_1_IT2) Or (ColCount2_IT2 <> ColCount2_2_IT2) Or (ColCount2_IT2 <> ColCount2_3_IT2)) Then
    msgbox_1398AATI = msgbox_1398AATI + "* Enter All Mandatory Details in PartB ATI"
    ValidateSheetATI_IT_2 = False
    End If
     
End Function

'Change 02, Satya, 13.09.2022
Function ValidateMandatoryShIT_2() As Boolean
ValidateMandatoryShIT_2 = True
Dim i As Long
Dim flag As Boolean
flag = True
For i = 1 To Sheet202.Range("U_BSRCode2").Rows.count
'Ankita_16/12/2025=====================
    If Sheet202.Range("U_BSRCode2").item(i).Value <> "" Or Sheet202.Range("U_DateDep2").item(i).Value <> "" Or _
        Sheet202.Range("U_SrlNoOfChaln2").item(i, 0).Value <> "" Or Sheet202.Range("U_Amt2").item(i).Value <> "" Then
        
        If Sheet202.Range("U_BSRCode2").item(i).Value = "" Then
            flag = False
            msgbox_1398AATI = msgbox_1398AATI + "* Please enter the BSR code in Schedule IT-2" & Chr(13)
        End If
        
        If flag = False Then
            ValidateMandatoryShIT_2 = False
            Exit Function
        End If
        
    End If
Next i
End Function
'End change

Sub setTableInfo_Grid3_IT2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet202.Range("U_BSRCode2").count
    Set rangecells = Sheet202.Range("U_BSRCode2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_IT2 = ccount
End Sub

Sub setTableInfo1_Grid3_IT2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet202.Range("U_DateDep2").count
    Set rangecells = Sheet202.Range("U_DateDep2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_1_IT2 = ccount
End Sub


Sub setTableInfo2_Grid3_IT2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet202.Range("U_SrlNoOfChaln2").count
    Set rangecells = Sheet202.Range("U_SrlNoOfChaln2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_2_IT2 = ccount
End Sub


Sub setTableInfo3_Grid3_IT2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet202.Range("U_Amt2").count
    Set rangecells = Sheet202.Range("U_Amt2").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount2_3_IT2 = ccount
End Sub

Function ValidateBSR_IT2() As Boolean
ValidateBSR_IT2 = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_BSRCode2").Cells
Dim i As Long
ReDim BSR_IT2(ColCount2_IT2)
noOfProcessSub = ColCount2_IT2
For i = 1 To ColCount2_IT2
BSR_IT2(i) = rangecells.item(i).Value
If Len(BSR_IT2(i)) = 0 Then
End If

'If Not chkCompulsory_1398A(BSR_IT2(i)) Then
'    msgbox_1398AATI = msgbox_1398AATI + "* Please enter the BSR Code at Sr.NO " & i & "in PartB ATI" & Chr(13)
'    ValidateBSR_IT2 = False
'    Exit Function
'End If

 If Not Validate_BSRCODE1(BSR_IT2(i)) Then
         msgbox_1398AATI = msgbox_1398AATI + "Invalid BSR code at Sr. No  " & i & "  in Sheet PartB ATI. First three digits from the left stand are for district code; next three digits stand for revenue centre code within the district and last single digit stands for population range code" & Chr(13)
         ValidateBSR_IT2 = False
        Exit Function
     End If
UpdateProgressBar
Next
End Function

Function ValidateDate_IT2() As Boolean
ValidateDate_IT2 = True
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_DateDep2").Cells
Dim i As Long
ReDim DateCredit_IT2(ColCount2_IT2)
For i = 1 To ColCount2_IT2
DateCredit_IT2(i) = rangecells.item(i).Value
If Len(DateCredit_IT2(i)) = 0 Then
End If
'Ankita_16/12/2025=================================

If Not FormatNCheckDate_1398A(DateCredit_IT2(i)) Then
    ValidateDate_IT2 = False
   msgbox_1398AATI = msgbox_1398AATI + "* Date of Deposit into Govt Account must be a valid dd/mm/yyyy format at Sr.NO " & i & " in schedule IT-2" & Chr(13)
    Exit Function
End If

If Not chkCompulsory_1398A(DateCredit_IT2(i)) Then
    msgbox_1398AATI = msgbox_1398AATI + "* Please enter the Date of Deposit into Govt Account at Sr.NO " & i & " in schedule IT-2" & Chr(13)
    ValidateDate_IT2 = False
    Exit Function
End If

If Not ChkMinInclusiveDate_1398A(Dformat(DateCredit_IT2(i), ""), "2022-04-01") Then ''
    msgbox_1398AATI = msgbox_1398AATI + "*Date of credit to Central government at Sr. No  " & i & " cannot be prior to 01/04/2022 in schedule IT-2" & Chr(13)
    ValidateDate_IT2 = False
    Exit Function
End If
Next
End Function

Function ValidateSerialNum_IT2() As Boolean
ValidateSerialNum_IT2 = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_SrlNoOfChaln2").Rows
Dim i As Long
ReDim SerialNum_IT2(ColCount2_IT2)
For i = 1 To ColCount2_IT2
SerialNum_IT2(i) = rangecells.item(i).Cells(1).Value
If Len(SerialNum_IT2(i)) = 0 Then
End If

'Ankita_04/11/2025====================================
If Not chkCompulsory_1398A(SerialNum_IT2(i)) Then
'    msgbox_1398AATI = msgbox_1398AATI + "* Please enter the Serial Number of Challan at Sr.NO " & i & " in PartB ATI" & Chr(13)
     msgbox_1398AATI = msgbox_1398AATI + "* Please enter serial number of challan at Sr.NO " & i & " in schedule IT-2." & Chr(13)
    ValidateSerialNum_IT2 = False
    Exit Function
End If


If Not IsNumeric(SerialNum_IT2(i)) Then
'   msgbox_1398AATI = msgbox_1398AATI + "* Serial Number of Challan at Sr.NO " & i & " can allow only Numeric Values in PartB ATI" & Chr(13)
    msgbox_1398AATI = msgbox_1398AATI + "* Serial Number of Challan at Sr.NO " & i & " can allow only Numeric Values schedule IT-2." & Chr(13)
    ValidateSerialNum_IT2 = False
    Exit Function
End If
Next
End Function

Function ValidateTaxPaid_IT2() As Boolean
ValidateTaxPaid_IT2 = True
'setTableInfo_Grid3
Dim rangecells As Range
Set rangecells = Sheet202.Range("U_Amt2").Cells
Dim i As Long
ReDim TaxPaid3_IT2(ColCount2_IT2)
For i = 1 To ColCount2_IT2
TaxPaid3_IT2(i) = rangecells.item(i).Value
If Len(TaxPaid3_IT2(i)) = 0 Then
End If

'Ankita_04/11/2025===============================
If Not chkCompulsory_1398A(TaxPaid3_IT2(i)) Then
'    msgbox_1398AATI = msgbox_1398AATI + "* Please enter the Tax Paid Amount at Sr.NO " & i & " in PartB ATI" & Chr(13)
    msgbox_1398AATI = msgbox_1398AATI + "* ""Please enter amount at Sr.NO " & i & " in schedule IT-2.""" & Chr(13)
    ValidateTaxPaid_IT2 = False
    Exit Function
End If
Next
End Function

Sub temp()
'Sheet3.Unprotect Password:=getmsgstate
Application.EnableEvents = True

End Sub




'Ankita_UR---------------------------------------11/10/2024

Sub UNLOCKRANGEIT2()

Application.EnableEvents = False

Dim i As Long

Dim chcell As Range

Dim rangenamestring_IT2 As Variant

  Sheet202.Unprotect Password:=getmsgstate

rangenamestring_IT2 = ("U_BSRCode2||U_DateDep2||U_SrlNoOfChaln2||U_Amt2")

    rangenamestring = Split(rangenamestring_IT2, "||")

        For i = 0 To UBound(rangenamestring)

            For Each chcell In Sheet202.Range(rangenamestring(i))

            'Sheet202.Unprotect Password:=getmsgstate

            chcell.MergeArea.Locked = False

            chcell.MergeArea.Interior.Color = (&HCCFFCC)

            chcell.MergeArea.ClearContents

            Next

        Next

Application.EnableEvents = True

End Sub



Sub LOCKRANGEIT2()

Application.EnableEvents = False

Dim i As Long

Dim chcell As Range

Dim rangenamestring_IT2 As Variant

  Sheet202.Unprotect Password:=getmsgstate

rangenamestring_IT2 = ("U_BSRCode2||U_DateDep2||U_SrlNoOfChaln2||U_Amt2")

    rangenamestring = Split(rangenamestring_IT2, "||")

        For i = 0 To UBound(rangenamestring)

            For Each chcell In Sheet202.Range(rangenamestring(i))

            chcell.MergeArea.Locked = True

            chcell.MergeArea.Interior.Color = (&HD8D8D8)

            chcell.MergeArea.ClearContents

           ' chcell.Item(1,1)

            Next

        Next

Application.EnableEvents = True

End Sub

