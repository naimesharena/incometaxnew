Attribute VB_Name = "SchTCS"
'----TCS----
Public TAN_TCS As Variant
Public EmployerOrDeductorOrCollecterName_TCS As Variant
Public DeductedYear_TCS As Variant
Public Year2_TCS As Variant
Public AmtTaxCollected_TCS As Variant

Public BroughtFwdTCSAmt_TCS As Variant
Public TotTCSOnAmtPaidCY_TCS As Variant
Public AmtClaimedOnOwnHands_TCS As Variant
Public AmtCarriedFwd_TCS As Variant
Public AmtClaimedBySpouse_TCS As Variant

Dim rngname_TCS1 As Variant
Dim end_TCS1 As Variant
Dim rngname_TCS As Variant
Public end_TCS As Variant
Dim incBy_TCS As Variant

Dim rngname_TCS_7 As Variant
Dim end_TCS_7 As Variant
Public end_CollectorName_TCS, end_CollectionYear_TCS, end_AmtCollection_TCS, end_TaxCollected_TCS, end_AmtClaim_TCS, end_AmtPortuguese_TCS As Long

Sub ValidateSchTCS_Click()
vbMessageCaption = "ITR 1 : 2025-26"         'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
validateSchTCS
'MsgBox "Sheet TCS is Ok", vbOKOnly, vbMessageCaption
fmsgboxoK "Sheet TCS is Ok"
End Sub

Sub validateSchTCS()
MsgTCS = ""
    If Not ValidatesheetTCS Then
        Sheet11.Activate
        'MsgBox (MsgTCS), vbOKOnly, "Error"
        fmsgbox (MsgTCS)
        CloseMsg
    End If
End Sub

Sub PrevTCS_Click()
Dim a As Worksheet
    Set a = ThisWorkbook.Sheets("TDS")
    a.Activate
End Sub

Sub NextTCS_Click()
Dim a As Worksheet
    Set a = ThisWorkbook.Sheets("Taxes Paid and Verification")
    a.Activate
End Sub

Function ValidatesheetTCS() As Boolean
On Error Resume Next
Dim rangename, intcntr, name, cellcount As Variant
Dim k  As Long
    ValidatesheetTCS = True
    MsgTCS = "* TCS." & Chr(13)
    setTblinfo_TCS
    
    noOfProcessSub = end_TCS
    subProcCaption = "Validating TCS"
For k = 1 To end_TCS
    If Not (Sheet11.Range("TCS.TAN").Cells.item(k) <> "" Or Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").Cells.item(k) <> "" Or Sheet11.Range("TCS.AmountCollected").Cells.item(k) <> "" Or Sheet11.Range("TCS.CollectionYear").Cells.item(k) <> "" Or Sheet11.Range("TCS.TotalTCS").Cells.item(k) <> "" Or Sheet11.Range("TCS.AmtTCSClaimedThisYear").Cells.item(k) <> "") Then
        MsgTCS = MsgTCS & "* At Sr. No " & k & " Please fill all the Mandatory Fields in Schedule TCS" & Chr(13)
        ValidatesheetTCS = False
        Exit Function
    End If
    UpdateProgressBar
Next
    
    If end_TCS > 0 Then GoTo xyz4
    cellcount = Sheet11.Range("TCS.TAN").count
'If Mid(Range("sheet1.PortugeseCC5A").Value, 1, 3) <> "Yes" Then
        rangename = "TCS.TAN|TCS.EmployerOrDeductorOrCollecterName|TCS.TotalTCS|TCS.AmtTCSClaimedThisYear|TCS.AmountCollected"
'Else:
        'rangename = "TCS.TAN|TCS.EmployerOrDeductorOrCollecterName|TCS.TotalTCS|TCS.AmtTCSClaimedThisYear|TCS.AmtClaimedBySpouse|TCS.AmountCollected"
'End If
    rangename = Split(rangename, "|")
    
'If Mid(Range("sheet1.PortugeseCC5A").Value, 1, 3) <> "Yes" Then
    For intcntr = 1 To cellcount
        If Sheet11.Range("TCS.TAN").Cells.item(intcntr) <> "" Or Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").Cells.item(intcntr) <> "" Or Sheet11.Range("TCS.AmountCollected").Cells.item(intcntr) <> "" Or Sheet11.Range("TCS.TotalTCS").Cells.item(intcntr) <> "" Or Sheet11.Range("TCS.AmtTCSClaimedThisYear").Cells.item(intcntr) <> "" Then
            For Each name In rangename
                If Sheet11.Range(name).Cells.item(intcntr) = "" Then ValidatesheetTCS = False
            Next
            If ValidatesheetTCS = False Then
                MsgTCS = MsgTCS & "* At Sr. No " & intcntr & " Please fill all the Mandatory Fields in Schedule TCS" & Chr(13)
                GoTo xyz4
            End If
        End If
    Next
    
'Else:
'    For intcntr = 1 To cellcount
'        If Sheet11.Range("TCS.TAN").Cells.Item(intcntr) <> "" Or Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").Cells.Item(intcntr) <> "" Or Sheet11.Range("TCS.AmountCollected").Cells.Item(intcntr) <> "" Or Sheet11.Range("TCS.TotalTCS").Cells.Item(intcntr) <> "" Or Sheet11.Range("TCS.AmtTCSClaimedThisYear").Cells.Item(intcntr) <> "" Then
'            For Each name In rangename
'                If Sheet11.Range(name).Cells.Item(intcntr) = "" Then ValidatesheetTCS = False
'            Next
'            If ValidatesheetTCS = False Then
'                MsgTCS = MsgTCS & "At Sr. No " & intcntr & " Please fill all the Mandatory Fields " & Chr(13)
'                GoTo xyz4
'            End If
'        End If
'    Next
'End If

xyz4:
    If Not ValidateTAN_TCS() Then ValidatesheetTCS = False
    If Not ValidateEmployerOrDeductorOrCollecterName_TCS() Then ValidatesheetTCS = False
    If Not ValidateYear_TCS() Then ValidatesheetTCS = False
    If Not ValidateTaxCollected_TCS() Then ValidatesheetTCS = False
    If Not ValidateClaimOutOfTotTCSOnAmtPaid_TCS() Then ValidatesheetTCS = False

        setTableInfo_CollectorName_TCS
        setTableInfo_TaxCollect_TCS
        setTableInfo_AmtClain_TCS
        setTableInfo_AmtPortugese_TCS
        setTableInfo_CollectionYear_TCS
        setTableInfo_AmtCollection_TCS
 
'If Sheet1.Range("sheet1.PortugeseCC5A").Value = "Yes" Then
'    If ((end_TCS <> end_CollectorName_TCS) Or (end_TCS <> end_AmtCollection_TCS) Or (end_TCS <> end_CollectionYear_TCS) Or (end_TCS <> end_TaxCollected_TCS) Or (end_TCS <> end_AmtClaim_TCS) Or (end_TCS <> end_AmtPortuguese_TCS)) Then
'        MsgTCS = MsgTCS & "Please fill all the Mandatory Fields " & Chr(13)
'        ValidatesheetTCS = False
'    End If
    
'Else:
    If ((end_TCS <> end_CollectorName_TCS) Or (end_TCS <> end_AmtCollection_TCS) Or (end_TCS <> end_CollectionYear_TCS) Or (end_TCS <> end_TaxCollected_TCS) Or (end_TCS <> end_AmtClaim_TCS)) Then
        MsgTCS = MsgTCS & "* Please fill all the Mandatory Fields in Schedule TCS" & Chr(13)
        ValidatesheetTCS = False
    End If
'End If

    If (Len(Range("TCS.Sum")) > 14) Then
        MsgTCS = MsgTCS & "* Total of TCS cannot be more than 14 digits in Schedule TCS" & Chr(13)
        ValidatesheetTCS = False
    End If
End Function

Function ValidateTAN_TCS() As Boolean
    ValidateTAN_TCS = True
    setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet11.Range("TCS.TAN").Cells
    ReDim TAN_TCS(end_TCS)
       For i = 1 To end_TCS
        TAN_TCS(i) = rangecells.item(i).Value
            
            If Not chkCompulsory(TAN_TCS(i)) Then
               ' MsgTCS = MsgTCS + "* TAN at Sr. No  " & i & "  is Mandatory" & Chr(13)
               'Added by Aavula Naresh
                MsgTCS = MsgTCS + "* Please enter TAN of collector at Sr. No  " & i & " in Schedule TCS." & Chr(13)

                ValidateTAN_TCS = False
                Exit Function
            End If

            
            If Not ValidateTantype_text(Mid(TAN_TCS(i), 1, 4)) Then
                MsgTCS = MsgTCS + "* TAN is invalid. First 4 alphabets, next 5 digits, then alphabet at Sr. No  " & i & " in Schedule TCS." & Chr(13)
                ValidateTAN_TCS = False
                Exit Function
            End If
            If Not IsNumeric(Mid(TAN_TCS(i), 5, 5)) Then
                MsgTCS = MsgTCS + "* TAN is invalid. First 4 alphabets, next 5 digits, then alphabet  at Sr. No  " & i & "  in Schedule TCS." & Chr(13)
                ValidateTAN_TCS = False
                Exit Function
            End If
            If Not ValidateTantype_text(Right(TAN_TCS(i), 1)) Then
                MsgTCS = MsgTCS + "* TAN is invalid. First 4 alphabets, next 5 digits, then alphabet  at Sr. No  " & i & "  in Schedule TCS." & Chr(13)
                ValidateTAN_TCS = False
                Exit Function
            End If
            If Not ValidateTANCodes(UCase(TAN_TCS(i))) Then
            
                MsgTCS = MsgTCS + "* Invalid TAN at Sr.No " & i & "  Schedule TCS.Please enter valid TAN" & Chr(13)
                ValidateTAN_TCS = False
                Exit Function
            End If

    Next
End Function

Function ValidateEmployerOrDeductorOrCollecterName_TCS() As Boolean
 
    ValidateEmployerOrDeductorOrCollecterName_TCS = True
    setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").Cells
    ReDim EmployerOrDeductorOrCollecterName_TCS(end_TCS)
    For i = 1 To end_TCS
        EmployerOrDeductorOrCollecterName_TCS(i) = rangecells.item(i).Value
        If Not chkCompulsory(EmployerOrDeductorOrCollecterName_TCS(i)) Then
             'MsgTCS = MsgTCS + "* Name of Collector is Mandatory at Sr. No  " & i & " in Schedule TCS." & Chr(13)
             'Added by Aavula Naresh
              MsgTCS = MsgTCS + "* Please enter name of Collector at Sr. No  " & i & " in Schedule TCS." & Chr(13)

            ValidateEmployerOrDeductorOrCollecterName_TCS = False
            Exit Function
        End If


         If Len(EmployerOrDeductorOrCollecterName_TCS(i)) > 125 Then
            MsgTCS = MsgTCS + "* Name Of Collector cannot be more than 125 charachters at Sr. No  " & i & " in Schedule TCS." & Chr(13)
            ValidateEmployerOrDeductorOrCollecterName_TCS = False
            Exit Function
         End If
 Next
End Function

Function ValidateYear_TCS() As Boolean
ValidateYear_TCS = True
    setTblinfo_TCS
    Dim rangecells As Range
    Set rangecells = Range("TCS.CollectionYear").Cells
    Dim i As Long
    ReDim Year2_TCS(end_TCS)

    For i = 1 To end_TCS
        Year2_TCS(i) = rangecells.item(i).Value
        If Len(Year2_TCS(i)) = 0 Then
        End If
        
        If (Year2_TCS(i) = "" Or UCase(Year2_TCS(i)) = "(SELECT)") Then
            'MsgTCS = MsgTCS + "* Collection Year is mandatory at Sr. No  " & i & " in Schedule TCS." & Chr(13)
            'Added by Aavula Naresh
             MsgTCS = MsgTCS + "* Please enter collection year at Sr. No  " & i & " in Schedule TCS." & Chr(13)

            ValidateYear_TCS = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter_TDS_TCS(Year2_TCS(i)) Then
             MsgTCS = MsgTCS + "* Collection Year at Sr. No  " & i & " in Schedule TCS, characters < > & ' " & Chr(34) & " are not allowed" & Chr(13)
             ValidateYear_TCS = False
             Exit Function
        End If
    Next
End Function


Function ValidateTaxCollected_TCS() As Boolean
ValidateTaxCollected_TCS = True
    setTblinfo_TCS
    Dim rangecells As Range
    Set rangecells = Range("TCS.AmountCollected").Cells
    Dim i As Long
    ReDim AmtTaxCollected_TCS(end_TCS)
    For i = 1 To end_TCS
        AmtTaxCollected_TCS(i) = rangecells.item(i).Value
        If Len(AmtTaxCollected_TCS(i)) = 0 Then
            AmtTaxCollected_TCS(i) = ""
        End If
        
        
'        If AmtTaxCollected_TCS(i) < 0 Then
'            MsgTCS = MsgTCS + "Amount which is subject to tax collection at Sr.NO " & i & " in Sheet TCS cannot be negative" & Chr(13)
'            ValidateTaxCollected_TCS = False
'            Exit Function
'        End If
        
        
        
        If Not chkCompulsory(AmtTaxCollected_TCS(i)) Then
            'MsgTCS = MsgTCS + "* Gross payment which is subject to tax collection is mandatory at Sr. No  " & i & " in Schedule TCS." & Chr(13)
            MsgTCS = MsgTCS + "* Please enter the Gross payment which is subject to tax collection at Sr. No  " & i & " in Schedule TCS." & Chr(13)

            ValidateTaxCollected_TCS = False
            Exit Function
        End If
    Next
End Function


Function ValidateClaimOutOfTotTCSOnAmtPaid_TCS() As Boolean
    ValidateClaimOutOfTotTCSOnAmtPaid_TCS = True
    setTblinfo_TCS
    
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    
    Dim i As Long
    
    Set rangecells1 = Sheet11.Range("TCS.TotalTCS").Cells
    Set rangecells3 = Sheet11.Range("TCS.AmtTCSClaimedThisYear").Cells
   ' Set rangecells2 = Sheet11.Range("TCS.AmtClaimedBySpouse").Cells
    Set rangecells4 = Sheet11.Range("TCS.AmountCollected").Cells
    
    
    ReDim BroughtFwdTCSAmt_TCS(end_TCS)
    ReDim AmtClaimedOnOwnHands_TCS(end_TCS)
    'ReDim AmtClaimedBySpouse_TCS(end_TCS)
    ReDim AmtTaxCollected_TCS(end_TCS)
    
    For i = 1 To end_TCS
      BroughtFwdTCSAmt_TCS(i) = rangecells1.item(i).Value
      AmtClaimedOnOwnHands_TCS(i) = rangecells3.item(i).Value
      'AmtClaimedBySpouse_TCS(i) = rangecells2.Item(i).Value
      AmtTaxCollected_TCS(i) = rangecells4.item(i).Value
    
    If (BroughtFwdTCSAmt_TCS(i) = "") Then
'    change.27.02.2023.102.IDS.17
        'MsgTCS = MsgTCS + "* Please enter the Tax Collected is Mandatory at Sr. No " & i & " in Schedule TCS." & Chr(13)
        MsgTCS = MsgTCS + "* Please enter the Tax Collected at Sr. No " & i & " in Schedule TCS." & Chr(13)
'   End Change
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
    End If
    
'
'    If BroughtFwdTCSAmt_TCS(i) < 0 Then
'        MsgTCS = MsgTCS + "Total Tax Collecteted at Sr. No " & i & " cannot be negative" & Chr(13)
'        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
'    End If
    
    
    If IsNumeric(AmtTaxCollected_TCS(i)) Then
       If AmtTaxCollected_TCS(i) < 0 Then
            MsgTCS = MsgTCS + "* TCS Credit in col.no 3 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & " in Schedule TCS" & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        End If
        If Len("" & AmtTaxCollected_TCS(i)) > 14 Then
            MsgTCS = MsgTCS + "* TCS Credit in col.no 3 cannot exceed 99999999999999 in Row at Sr. No  " & i & " in Schedule TCS." & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        End If
    Else
        MsgTCS = MsgTCS + "* TCS Credit in col.no 3 is invalid at Sr. No  " & i & " in Schedule TCS." & Chr(13)
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
    End If
    
    
    
    If IsNumeric(BroughtFwdTCSAmt_TCS(i)) Then
       If BroughtFwdTCSAmt_TCS(i) < 0 Then
            MsgTCS = MsgTCS + "* Amount in col.no 5 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & " in Schedule TCS" & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        End If
        If Len("" & BroughtFwdTCSAmt_TCS(i)) > 14 Then
            MsgTCS = MsgTCS + "* Amount in col.no 5 cannot exceed 99999999999999 at Sr. No  " & i & " in Schedule TCS." & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        End If
    Else
        MsgTCS = MsgTCS + "* Amount in col.no 5 is invalid at Sr. No  " & i & " in Schedule TCS." & Chr(13)
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
    End If
    
    
    If (AmtClaimedOnOwnHands_TCS(i) = "") Then
        MsgTCS = MsgTCS + "* Please enter Amount claimed for this year at Sr. No " & i & " in Schedule TCS" & Chr(13)
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
    End If
    
'    If (AmtClaimedOnOwnHands_TCS(i) < 0) Then
'        MsgTCS = MsgTCS + "Amount Claimed for this Year at Sr. No " & i & " Cannot be negative" & Chr(13)
'        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
'    End If
    
'    If Sheet1.Range("sheet1.PortugeseCC5A").Value = "Yes" Then
'        If (AmtClaimedBySpouse_TCS(i) = "") Then
'            MsgTCS = MsgTCS + "Amount claimed in the hands of spouse is mandatory as the assessee is governed by Portuguese Civil Code under Sec 5A. In case of nil amount, please enter zero" & Chr(13)
'            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
'        End If
'    End If
    
    If IsNumeric(AmtClaimedOnOwnHands_TCS(i)) Then
      If AmtClaimedOnOwnHands_TCS(i) < 0 Then
            MsgTCS = MsgTCS + "* Amount in col.no 6 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & " in Schedule TCS" & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
      End If
     If Len("" & AmtClaimedOnOwnHands_TCS(i)) > 14 Then
            MsgTCS = MsgTCS + "* Amount in col.no 6 cannot exceed 99999999999999 at Sr. No  " & i & " in Schedule TCS." & Chr(13)
            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
            Exit Function
     End If
Else
        MsgTCS = MsgTCS + "* Amount in col.no 4 is invalid at Sr. No  " & i & " in Schedule TCS." & Chr(13)
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        Exit Function
End If
    
'If IsNumeric(AmtClaimedBySpouse_TCS(i)) Then
'        If AmtClaimedBySpouse_TCS(i) < 0 Then
'            MsgTCS = MsgTCS + "Amount in col.no 7 should be Numeric, Non Negative, not exceeding 14 digits in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
'            Exit Function
'        End If
'
'    If Len("" & AmtClaimedBySpouse_TCS(i)) > 14 Then
'            MsgTCS = MsgTCS + "Amount in col.no 7 cannot exceed 99999999999999 in Row  " & i & Chr(13)
'            ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
'            Exit Function
'     End If
'Else
'        MsgTCS = MsgTCS + "Amount in col.no 5 is invalid in Row  " & i & Chr(13)
'        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
'        Exit Function
'End If
        
   
    If AmtClaimedOnOwnHands_TCS(i) > (BroughtFwdTCSAmt_TCS(i)) Then
        MsgTCS = MsgTCS + "* TCS credit claimed for this year cannot be more than tax collected at Sr. No  " & i & " in Schedule TCS." & Chr(13)
        ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
        Exit Function
    End If
    
    'If AmtTaxCollected_TCS(i) < (2 * (BroughtFwdTCSAmt_TCS(i))) Then
    '    MsgTCS = MsgTCS + "Amount which is subject to tax collection could not be less than twice the amount of tax collected in Col.No 5 in Row " & i & Chr(13)
    '    ValidateClaimOutOfTotTCSOnAmtPaid_TCS = False
    '    Exit Function
    'End If

    If (AmtClaimedOnOwnHands_TCS(i) = "") Then
        AmtClaimedOnOwnHands_TCS(i) = 0
    End If


 Next
End Function

Sub setTblinfo_TCS()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet11.Range("TCS.TAN").count
 Set rangecells = Sheet11.Range("TCS.TAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TCS = ccount
 DefinedgridNameRange = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.AmountCollected||TCS.CollectionYear||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear"
 End Sub

Sub setTblinfo_TCS_7()
 Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet11.Range("TCS.TAN").count
 Set rangecells = Sheet11.Range("TCS.TAN").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TCS_7 = ccount
 rngname_TCS_7 = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear||TCS.AmtClaimedBySpouse||TCS.AmountCollected||TCS.CollectionYear"
 End Sub
Sub setTableInfo_CollectorName_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").count
 Set rangecells = Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_CollectorName_TCS = ccount
End Sub

Sub setTableInfo_TaxCollect_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet11.Range("TCS.TotalTCS").count
 Set rangecells = Sheet11.Range("TCS.TotalTCS").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_TaxCollected_TCS = ccount
End Sub

Sub setTableInfo_AmtClain_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet11.Range("TCS.AmtTCSClaimedThisYear").count
 Set rangecells = Sheet11.Range("TCS.AmtTCSClaimedThisYear").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_AmtClaim_TCS = ccount
End Sub


Sub setTableInfo_AmtPortugese_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet11.Range("TCS.AmtClaimedBySpouse").count
 Set rangecells = Sheet11.Range("TCS.AmtClaimedBySpouse").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_AmtPortuguese_TCS = ccount
End Sub

Sub setTableInfo_CollectionYear_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet11.Range("TCS.CollectionYear").count
 Set rangecells = Sheet11.Range("TCS.CollectionYear").Cells
 For mIntCtr = 1 To mIntCells
     If Not (rangecells.item(mIntCtr).Value = "" Or UCase(rangecells.item(mIntCtr).Value) = "(SELECT)") Then
         ccount = ccount + 1
     End If
 Next
 end_CollectionYear_TCS = ccount
End Sub

Sub setTableInfo_AmtCollection_TCS()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet11.Range("TCS.AmountCollected").count
 Set rangecells = Sheet11.Range("TCS.AmountCollected").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_AmtCollection_TCS = ccount
End Sub


Sub AddRowSchTCS_Click()
Dim vRows As Long
Dim sourceSheet As Worksheet

Set sourceSheet = ThisWorkbook.Sheets("TCS")
    sourceSheet.Activate
EfilingCommon.DefinedgridNameRange = "TCS.TAN||TCS.EmployerOrDeductorOrCollecterName||TCS.TotalTCS||TCS.AmtTCSClaimedThisYear||TCS.AmountCollected||TCS.CollectionYear"
ActiveCellRange = EfilingCommon.searchLastRow("TCS.TAN")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub


