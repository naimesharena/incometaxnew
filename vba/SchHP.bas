Attribute VB_Name = "SchHP"
Option Explicit
Dim sPassword As Variant

Public rngname_hpco As Variant
Public rngname_hpco2 As Variant
Public end_hpco As Variant
Public end_hpco2 As Variant
Public end_hpco_24a As Variant
Public end_hpco_24b As Variant
Public rngname_hpco_24b As Variant
Public end_hpco_24c As Variant
Public end_hpco_24e As Variant
Public end_hpco_24f As Variant
Public end_hpco_24k As Variant
Public end_hpco_24s As Variant
Public end_hpco_24r As Variant
'-----------------------------

Public AddrDetail_HP As Variant
Public CityOrTownOrDistrict_HP As Variant
Public CountryCode_HP As Variant
Public OwnerProperty_HP As Variant
Public OwnerPropertyDescription_HP As Variant
Public StateCode_HP As Variant
Public ZipCode_HP As Variant
Public PinCode_HP As Variant
Public ifLetOut_HP As Variant
Public NameofTenant_HP As Variant
Public PANofTenant_HP As Variant
Public TANofTenant_HP As Variant
Public CoName_HP As Variant
Public CoPAN_HP As Variant
Public CoAadhar_HP As Variant
Public CoShare_HP As Variant

Public AnnualLetableValue_HP As Variant
Public RentNotRealized_HP As Variant
Public LocalTaxes_HP As Variant
Public TotalUnrealizedAndTax_HP As Variant
Public BalanceALV_HP As Variant
Public ThirtyPercentOfBalance_HP As Variant
Public IntOnBorwCap_HP As Variant
Public TotalDeduct_HP As Variant
Public Arrears_HP As Variant
Public IncomeOfHP_HP As Variant
Public IncomeOfHPInOwnHand_HP As Variant
Public RentOfEarlierYrSec25AandAA_HP As Variant
Public RentArearsSec25BAfter30pcDeduct_HP As Variant
Public TotalIncomeChargeableUnHP_HP As Variant

Public CoOwnerAadhar_HP As Variant
Public AadharofTenant_HP As Variant


Public CoOwnerName1_HP As Variant
Public CoOwnerName2_HP As Variant
Public CoOwnerName3_HP As Variant
Public CoOwnerName4_HP As Variant
Public CoOwnerName5_HP As Variant

Public CoOwnerPAN1_HP As Variant
Public CoOwnerPAN2_HP As Variant
Public CoOwnerPAN3_HP As Variant
Public CoOwnerPAN4_HP As Variant
Public CoOwnerPAN5_HP As Variant

Public CoOwnerSharePer1_HP As Variant
Public CoOwnerSharePer2_HP As Variant
Public CoOwnerSharePer3_HP As Variant
Public CoOwnerSharePer4_HP As Variant
Public CoOwnerSharePer5_HP As Variant

Public CoOwnedYN_HP As Variant
Public CoOwnedShare_HP As Variant

Public frmsize_hprptfrm As Variant
Public end_hprptfrm As Variant
Public cntrRng_hprptfrm As Variant
Public frmRngname_hprptfrm As Variant
Public rngname_hprptfrm As Variant

Dim msgValidateSheethprptfrm As String
Dim msgValidateSheetHP As String

Public end_24b As Variant
Public end_24bLoanfrm As Variant
Public end_24bbankName As Variant
Public end_24bPAN As Variant
Public end_24bAccntNum As Variant
Public end_24bLoanDate As Variant
Public end_24bLoanAmt As Variant
Public end_24bLoanOutstanding As Variant
Public end_24bIntrst As Variant

Public BankorInst_24b As Variant
Public BankName_24B As Variant
Public AccntNum_24b As Variant
Public LoanDate_24B As Variant
Public LoanAmt_24b As Variant
Public LoanOutStanding_24B As Variant
Public Intrst_24B As Variant

Sub cmdNext_Click_HP()
Sheet18.Activate
End Sub
Sub cmdPrev_Click_HP()
Sheet1.Activate
End Sub



'Sub Next24b_Click()
'Dim sourceSheet As Worksheet
'    If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
'       Sheet201.Activate
'    Else:
'  Set sourceSheet = ThisWorkbook.Sheets("HP")
'    sourceSheet.Activate
'        End If
'End Sub

'Ankita_21/01/2026====

Sub AddHPCoowner(Optional blockcount As Long = 0)
On Error GoTo endline
    
Dim vRows  As Long
Dim blocknum As Variant
Dim blockvalue As Variant
Dim newrngname As Variant
Dim numberofrows As Long
Dim newfrmsize, i, te  As Long
Application.EnableEvents = False
    Sheets("HP").Activate
    Sheet19.Unprotect Password:=getmsgstate
If blockcount = 0 Then
    blocknum = ActiveSheet.Buttons(Application.Caller).TopLeftCell.name
    blockvalue = Range(blocknum).Offset(0, -1).Value
Else
    blockvalue = blockcount
End If
    DefinedgridNameRange = "HP.Co.Name" & blockvalue & "||HP.Co.Pan" & blockvalue & "||HP.Co.Aadhaar" & blockvalue & "||HP.Co.Share" & blockvalue
    ActiveCellRange = searchLastRow("HP.Co.Name" & blockvalue)
    vRows = insertRowUnderSectionOneRow
    Sheet19.Unprotect Password:=getmsgstate
    Application.EnableEvents = False
endline:
Sheet19.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub

Sub AddRows24b_HP_click(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    Application.EnableEvents = False
    Sheet19.Activate
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    numberofrows = iRows

    For te = 1 To Sheet19.Range("PropertySectionCOunt").Value
        EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b||Combination_24B"
        ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.24b" & te)
        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    Application.EnableEvents = False
    For i = 1 To Sheet19.Range("PropertySectionCount").Value
        newfrmsize = Sheet19.Range("NumRowsPropertyBlock").Value
        newfrmsize = newfrmsize + numberofrows
        Sheet19.Range("NumRowsPropertyBlock").Value = newfrmsize
    Next
endline:
'----------------Lock Password-------------------START---
   Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True

End Sub

 
'=====================
Function ValidateTenantPan(tenantPan As String) As Boolean
On Error Resume Next
    ValidateTenantPan = True
    If Len(tenantPan) > 0 Then
        If Not ChkAlphabet(Mid(tenantPan, 1, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(tenantPan, 2, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(tenantPan, 3, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If

        If Not ChkAlphabet(Mid(tenantPan, 4, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If

        If Not ChkAlphabet(Mid(tenantPan, 5, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If
        If Not IsNumeric(Mid(tenantPan, 6, 4)) Then
            ValidateTenantPan = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(tenantPan, 10, 1)) Then
            ValidateTenantPan = False
            Exit Function
        End If
    End If
End Function

Sub AddRows_hpco()
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim te As Long
    Dim i As Long
    Dim frmsize_hprptfrmnew As Long
    Dim newfrmsize_hprptfrm As Long
    
    newrngname = rngname_hpco
    For te = 1 To Sheet19.Range("cntr.hprptfrm").Value
        Application.EnableEvents = False
        setTblinfo_hpco
        SelectLastRow ("hp.co.srno" & te)
        
        If te > 1 Then
            numberofrows = InsertRowsAndFillFormulasBPA(numberofrows)
        Else
            numberofrows = InsertRowsAndFillFormulasBPA()
        End If
        
        Call ExendRangeNameToTablehp(CLng(numberofrows), rngname_hpco, te)
        Application.EnableEvents = True
    Next

    sPassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=sPassword
    
    Application.EnableEvents = False
    For i = 1 To Sheet19.Range("cntr.hprptfrm").Value
        newfrmsize_hprptfrm = Sheet19.Range("hprptfrm.size").Value
        newfrmsize_hprptfrm = newfrmsize_hprptfrm + numberofrows
        Sheet19.Range("hprptfrm.size").Value = newfrmsize_hprptfrm
    Next
    Application.EnableEvents = True
    Sheet19.Protect Password:=sPassword
End Sub

Sub ValidateSheetHPClick()
ValidateSheetHouseProperty
If msgValidateSheetHP = "" Then MsgBox "Sheet House Property (HP) is OK" ', vbOKOnly,
End Sub

Sub ValidateSheetHouseProperty()
subProcCaption = "Validating HP"
    If Not ValidatesheetHP Then
        Sheet19.Activate
        fmsgbox msgValidateSheetHP ', vbOKOnly, "Error(s)!"
        CloseMsg
    End If
End Sub
Function GetLetOut() As Variant
    If ValidateifLetOut_HP Then
        GetLetOut = ifLetOut_HP
    End If
End Function


Function msgbox_hprptfrm(strmsg As String) As String
     msgValidateSheetHP = msgValidateSheetHP & strmsg & Chr(13)
End Function
Function msgbox_HP(strmsg As String) As String
     msgValidateSheetHP = msgValidateSheetHP & strmsg & Chr(13)
End Function

Function ValidatesheetHP() As Boolean
    ValidatesheetHP = True
    
    If Not ValidateAddrDetail_HP() Then ValidatesheetHP = False
    If Not ValidateCoName_HP() Then ValidatesheetHP = False
        If Not Validate24B_Table() Then ValidatesheetHP = False

    
    Dim te As Long
    Dim de As Long
    For te = 1 To Sheet19.Range("PropertySectionCOunt").Value
        If Sheet19.Range("HP.RentOfEarlierYrSec_AandAA" & te).Value > 0 Then
          If Sheet19.Range("HP.AddrDetail" & te).Value = "" Then
          'Error message changed as per V0.4_Ankita_27/02/2026===========
'          msgbox_hprptfrm ("* ""Address of property " & te & " is mandatory in schedule HP.""")
           msgbox_hprptfrm ("* ""Address is mandatory in tab: Sl no B2""")
            ValidatesheetHP = False
            Exit Function
        End If
        End If
    Next
    For te = 1 To Sheet19.Range("PropertySectionCOunt").Value
        If Sheet19.Range("HP.TotalUnrealizedAndTax" & te).Value > 0 Or Sheet19.Range("HP.BalanceALV" & te).Value > 0 Or Sheet19.Range("HP.IncomeOfHPInOwnHand" & te).Value > 0 Or Sheet19.Range("HP.TotalDeduct" & te).Value > 0 Or Sheet19.Range("HP.IncomeOfHP" & te).Value > 0 Then
          If Sheet19.Range("HP.AddrDetail" & te).Value = "" Then
         'Error message changed as per V0.4_Ankita_27/02/2026===========
'          msgbox_hprptfrm ("* ""Address of property " & te & " is mandatory in schedule HP.""")
           msgbox_hprptfrm ("* ""Address is mandatory in tab: Sl no B2""")
            ValidatesheetHP = False
            Exit Function
        End If
        End If
    Next
     For de = 1 To Sheet19.Range("PropertySectionCOunt").Value
       
        If Sheet19.Range("HP.TotalUnrealizedAndTax" & de) > 99999999999999# Then
            msgbox_hprptfrm "* Total (1b + 1c) in schedule House Property cannot exceed 14 digits" & Chr(13)
            ValidatesheetHP = False
            Exit Function
        End If
        Next
        
'    For te = 1 To Sheet19.Range("PropertySectionCOunt").Value
'    If Not isdropdownblank(Sheet19.Range("HP.StateCode" & te).Value) _
'    Or Not isdropdownblank(Sheet19.Range("HP.CountryCode" & te).Value) _
'    Or Not isdropdownblank(Sheet19.Range("HP.OwnerProperty" & te).Value) _
'    Or Not isdropdownblank(Sheet19.Range("HP.CoOwnedYN" & te).Value) _
'    Or Sheet19.Range("HP.AddrDetail" & te).Value <> "" Then


   For te = 1 To Sheet19.Range("PropertySectionCOunt").Value
    If Not Sheet19.Range("HP.StateCode" & te).Value = "" Or Sheet19.Range("HP.StateCode" & te).Value = "(Select)" Or Sheet19.Range("HP.OwnerProperty" & te).Value = "" Or Sheet19.Range("HP.CoOwnedYN" & te).Value = "" Or Sheet19.Range("HP.AddrDetail" & te).Value = "" Then
    If Not isdropdownblank(Sheet19.Range("HP.StateCode" & te).Value) _
    Or Not isdropdownblank(Sheet19.Range("HP.OwnerProperty" & te).Value) _
    Or Not isdropdownblank(Sheet19.Range("HP.CoOwnedYN" & te).Value) _
    Or Sheet19.Range("HP.AddrDetail" & te).Value <> "" Then

'Ankita_27/02/2026==========
         If Sheet19.Range("HP.AddrDetail" & te).Value = "" Then
'            msgbox_hprptfrm ("* ""Address of property " & te & " is mandatory in schedule HP.""")
             msgbox_hprptfrm ("* ""Address is mandatory in tab: Sl no B2""")
            ValidatesheetHP = False
            Exit Function
        End If
        End If
        End If
    Next
    If (Len(Sheet19.Range("HP.AddrDetail1")) > 0) Then
        If Not ValidateCityOrTownOrDistrict_HP() Then ValidatesheetHP = False
        If Not ValidateCountryCode_HP() Then ValidatesheetHP = False
        If Not ValidateStateCode_HP() Then ValidatesheetHP = False
        If Not ValidateZipCode_HP() Then ValidatesheetHP = False
        If Not ValidatePinCode_HP() Then ValidatesheetHP = False
        If Not ValidateOwnerProperty_HP() Then ValidatesheetHP = False
        If Not ValidateCountryStateCode_HP() Then ValidatesheetHP = False
        If Not ValidateifLetOut_HP() Then ValidatesheetHP = False
        If Not ValidateNameofTenant_HP() Then ValidatesheetHP = False
        If Not ValidatePANofTenant_HP() Then ValidatesheetHP = False
        If Not ValidateAadharofTenant_HP() Then ValidatesheetHP = False
        If Not ValidateTANofTenant_HP() Then ValidatesheetHP = False
        If Not ValidateAnnualLetableValue_HP() Then ValidatesheetHP = False
        If Not ValidateRentNotRealized_HP() Then ValidatesheetHP = False
        If Not ValidateLocalTaxes_HP() Then ValidatesheetHP = False
        If Not ValidateTotalUnrealizedAndTax_HP() Then ValidatesheetHP = False
        If Not ValidateBalanceALV_HP() Then ValidatesheetHP = False
        If Not ValidateThirtyPercentOfBalance_HP() Then ValidatesheetHP = False
        If Not ValidateIntOnBorwCap_HP() Then ValidatesheetHP = False
        If Not ValidateTotalDeduct_HP() Then ValidatesheetHP = False
        If Not ValidateArrears_HP() Then ValidatesheetHP = False
        If Not ValidateIncomeOfHP_HP() Then ValidatesheetHP = False
        If Not ValidateRentArearsSec25BAfter30pcDeduct_HP() Then ValidatesheetHP = False
        If Not ValidateTotalIncomeChargeableUnHP_HP() Then ValidatesheetHP = False
        If Not ValidateCoownerName1_HP() Then ValidatesheetHP = False
        If Not ValidatePercenShare1_HP() Then ValidatesheetHP = False
        If Not ValidateCoOwnerPAN1_HP() Then ValidatesheetHP = False
        If Not ValidateCoOwnerAadhar_HP() Then ValidatesheetHP = False
        If Not ValidateSharePercent_HP() Then ValidatesheetHP = False
        If Not ValidateCoOwnedYN_HP() Then ValidatesheetHP = False
        If Not ValidateCoownerRules_HP() Then ValidatesheetHP = False
        If Not ValidatesharesRules_HP() Then ValidatesheetHP = False

            If ValidatesheetHP Then
            Dim noofsop, temp1, hpi, SPI, ind As Long
            Dim letoutarr As Variant
            letoutarr = GetLetOut
            noofsop = 0
            For hpi = 1 To UBound(letoutarr)
                If Mid(Sheet19.Range("HP.ifLetOut" & hpi).Value, 1, 1) = "S" Then
                 noofsop = noofsop + 1
                    If (Sheet19.Range("HP.AnnualLetableValue" & hpi).Value > 0) Then
                    Sheet19.Activate
                        MsgBox "Property " & hpi & " is deemed to be letout as its A.L.V. > 0" ', vbOKOnly, "Error!"
                        CloseMsg
                    End If
                    
                     If Mid(Sheet19.Range("HP.ifLetOut" & hpi).Value, 1, 1) = "S" Then
                          If noofsop > 2 Then
                               msgbox_hprptfrm "* Self occupied property cannot be selected more than twice from house property is declared in Schedule HP." & Chr(13)
                                ValidatesheetHP = False
                          End If
                          temp1 = Sheet19.Range("HP.IntOnBorwCap" & hpi) + temp1
                          If temp1 > 200000 Then
                            msgbox_hprptfrm "* F for category self-occupied in Schedule HP." & Chr(13)
                                ValidatesheetHP = False
                          End If
                        
                        End If
                Else
                    If letoutarr(hpi) = "" Then
                        If (Sheet19.Range("HP.AnnualLetableValue" & hpi).Value = 0) Then
                            noofsop = noofsop + 1
                        Else
                        End If
                    Else
                    End If
                End If
            Next
    
            letoutarr = GetLetOut
            For SPI = 1 To UBound(letoutarr)
                If (letoutarr(SPI) = "S" Or letoutarr(SPI) = "" Or _
                    Sheet19.Range("HP.AnnualLetableValue" & SPI).Value = 0) Then
                    ind = SPI
                    Dim count As Variant
                    Dim tempval As Range
                    count = 0
                If letoutarr(SPI) = "S" Then
                  For Each tempval In Sheet19.Range("HP.NameofTenant" & ind).Cells
                    If Not tempval.Value = "" Then
                        count = count + 1
                    End If
                  Next
                  If count > 0 Then
                        msgbox_hprptfrm ("* If Self Occupied then Name of Tenant in schedule HP at Block " & ind & " is not required")
                        ValidatesheetHP = False
                    End If
                    
                  count = 0
                  For Each tempval In Sheet19.Range("HP.PANofTenant" & ind).Cells
                    If Not tempval.Value = "" Then
                        count = count + 1
                    End If
                  Next
                  
                    If count > 0 Then
                         msgbox_hprptfrm ("* If Self Occupied then PAN of Tenant in schedule HP at Block " & ind & " is not required")
                         ValidatesheetHP = False
                        End If
                        
                         count = 0
                  For Each tempval In Sheet19.Range("HP.AadharofTenant" & ind).Cells
                    If Not tempval.Value = "" Then
                        count = count + 1
                    End If
                  Next
                  
                    If count > 0 Then
                         msgbox_hprptfrm ("* If Self Occupied then Aadhaar of Tenant in schedule HP at Block " & ind & " is not required")
                         ValidatesheetHP = False
                        End If
                        
                        
                    count = 0
                  For Each tempval In Sheet19.Range("HP.TANofTenant" & ind).Cells
                    If Not tempval.Value = "" Then
                        count = count + 1
                    End If
                  Next
                  
                    If count > 0 Then
                         msgbox_hprptfrm ("* If Self Occupied then TAN of Tenant in schedule HP at Block " & ind & " is not required")
                         ValidatesheetHP = False
                        End If
                        count = 0
                
                    End If
                    If val(Sheet19.Range("HP.IntOnBorwCap" & ind)) > 200000 Then
                        If Mid(Sheet19.Range("HP.ifLetOut" & ind), 1, 1) = "D" Or Mid(Sheet19.Range("HP.ifLetOut" & ind), 1, 1) = "L" Then
                        
                        Else
                            Sheet19.Activate
                            MsgBox "If Not Letout (OR ALV=0) then Interest Payable on borrowed capital should not exceed 2 Lacs" ' , vbOKOnly, "Error!"
                            CloseMsg
                        End If
                    End If
                
                    If Sheet19.Range("HP.LocalTaxes" & ind).Value > 0 Then
                    Sheet19.Activate
                        MsgBox "Tax paid to local authorities can be claimed only if income from house property is declared" ', vbOKOnly, "Error!"
                        CloseMsg
                    End If
                
                    If val(Sheet19.Range("HP.RentNotRealized" & ind)) > 0 Then
                    Sheet19.Activate
                        MsgBox """Rent not realized cannot exceed Gross rent received or receivable or letable value in schedule HP""" ', vbOKOnly, "Error!"
                        CloseMsg
                    End If
                End If
                            
                If (letoutarr(SPI) <> "S") Then
                    ind = SPI
                    If val(Sheet19.Range("HP.RentNotRealized" & ind)) > Sheet19.Range("HP.AnnualLetableValue" & ind).Value Then
                    Sheet19.Activate
                        MsgBox """Rent not realized cannot exceed Gross Rent received or receivable or letable value in schedule HP""" ', vbOKOnly, "Error!"
                        CloseMsg
                    End If
                    If letoutarr(SPI) = "L" Then
                        If Sheet19.Range("HP.AnnualLetableValue" & ind) <= 0 Then
                        Sheet19.Activate
                            MsgBox "If Type of House Property is Let Out/ Deemed Let Out then Gross Rent received or receivable or lettable value field cannot be zero/ blank"
                            
                            CloseMsg
                        End If
                    End If
                    
                    If letoutarr(ind) = "D" Then
                        For Each tempval In Sheet19.Range("HP.NameofTenant" & ind).Cells
                            If Not tempval.Value = "" Then
                            count = count + 1
                            End If
                        Next
                        If count > 0 Then
                            msgbox_hprptfrm ("* If Deemed let out then Name of Tenant in schedule HP at Block " & ind & " is not required")
                            ValidatesheetHP = False
                            End If
                        
                        count = 0
                        For Each tempval In Sheet19.Range("HP.PANofTenant" & ind).Cells
                            If Not tempval.Value = "" Then
                            count = count + 1
                            End If
                        Next
                        
                        If count > 0 Then
                            msgbox_hprptfrm ("* If Deemed let out then PAN of Tenant in schedule HP at Block " & ind & " is not required")
                            ValidatesheetHP = False
                        End If
                        count = 0
                        For Each tempval In Sheet19.Range("HP.TANofTenant" & ind).Cells
                            If Not tempval.Value = "" Then
                            count = count + 1
                            End If
                        Next
                        
                        If count > 0 Then
                            msgbox_hprptfrm ("* If Deemed let out then TAN of Tenant in schedule HP at Block " & ind & " is not required")
                            ValidatesheetHP = False
                        End If
                        count = 0
                        End If
                    End If
            Next
        End If
    Else
        If Not ValidateRentArearsSec25BAfter30pcDeduct_HP() Then ValidatesheetHP = False
        If Not ValidateTotalIncomeChargeableUnHP_HP() Then ValidatesheetHP = False
    End If


'Ankita_10/03/2026=============
'If Sheet19.Range("HP.AddrDetail2").Value = "" Or Sheet19.Range("HP.CityOrTownOrDistrict2").Value = "" Or Sheet19.Range("HP.StateCode2").Value = "(Select)" Or Sheet19.Range("HP.StateCode2").Value = "" _
'Or Sheet19.Range("HP.CountryCode2").Value = "(Select)" Or Sheet19.Range("HP.CountryCode2").Value = "" Or Sheet19.Range("HP.ifLetOut2").Value = "(Select)" Or Sheet19.Range("HP.ifLetOut2").Value = "" Then
'    If Sheet19.Range("HP.AddrDetail2").Value <> "" Or Sheet19.Range("HP.CityOrTownOrDistrict2").Value <> "" Then
'        msgbox_hprptfrm "* Enter all mandatory fields of Block 1B in Schedule HP."
'        ValidatesheetHP = False
'    End If
'End If
 
 If Not ValidateMandatorytable1 Then ValidatesheetHP = False
 
 If ValidatesheetHP = True Then
 'Ayush_20/03/2026
 '   If Sheet19.Range("HP.AddrDetail1").Value = "" Or Sheet19.Range("HP.CityOrTownOrDistrict1").Value = "" Or Sheet19.Range("HP.StateCode1").Value = "(Select)" Or Sheet19.Range("HP.StateCode1").Value = "" Or Sheet19.Range("HP.CountryCode1").Value = "91-INDIA" Or Sheet19.Range("HP.CountryCode1").Value = "" Or Sheet19.Range("HP.ifLetOut1").Value = "(Select)" Or Sheet19.Range("HP.ifLetOut1").Value = "" Then
     If Sheet19.Range("HP.AddrDetail1").Value = "" Or Sheet19.Range("HP.CityOrTownOrDistrict1").Value = "" Or (Sheet19.Range("HP.StateCode1").Value = "(Select)" Or Sheet19.Range("HP.StateCode1").Value = "") Or (Sheet19.Range("HP.CountryCode1").Value = "" Or Sheet19.Range("HP.CountryCode1").Value = "(Select)") Or Sheet19.Range("HP.ifLetOut1").Value = "(Select)" Or Sheet19.Range("HP.ifLetOut1").Value = "" Then
        If Sheet19.Range("Block_1B").Value < 8 Then
            If Sheet19.Range("Final_1B") = "Data" Then
                msgbox_hprptfrm ("* Enter all mandatory fields of Block 1B in Schedule HP.")
                ValidatesheetHP = False
                Exit Function
            End If
        End If
    End If
End If

End Function


'========
Sub setTblinfo_24bankname(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet19.Range("bankName.24b" & myindex).count
    Set rangecells = Sheet19.Range("bankName.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).Value
        
        If scode = "" Or IsEmpty(scode) Or UCase(scode) = UCase("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24b = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub

Sub setTblinfo_24bBankorInst(ByVal myindex As Long)
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet19.Range("LoanfrmBankOrInstitute.24b" & myindex).count
    Set rangecells = Sheet19.Range("LoanfrmBankOrInstitute.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).Value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
    
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24a = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub
Sub setTblinfo_24bLoan(ByVal myindex As Long)
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet19.Range("loanAccNum.24b" & myindex).count
    Set rangecells = Sheet19.Range("loanAccNum.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).Value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24e = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub

'Ankita_23/01/2026======

Sub setTblinfo_24bLoanfrm()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet19.Range("PropertySectionCOunt").Value
    mIntCells = Range("LoanfrmBankOrInstitute.24b" & i).count
    Set rangecells = Range("LoanfrmBankOrInstitute.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "" Or rangecells.item(mIntCtr).Value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanfrm = ccount
End Sub

Sub setTblinfo_24bBankName()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet19.Range("PropertySectionCOunt").Value
    mIntCells = Range("bankName.24b" & i).count
    Set rangecells = Range("bankName.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bbankName = ccount
End Sub

Sub setTblinfo_24bAccntNum()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet19.Range("PropertySectionCOunt").Value
    mIntCells = Range("loanAccNum.24b" & i).count
    Set rangecells = Range("loanAccNum.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bAccntNum = ccount
End Sub

Sub setTblinfo_24bLoanDate()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet19.Range("PropertySectionCOunt").Value
    mIntCells = Range("loanDate.24b" & i).count
    Set rangecells = Range("loanDate.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanDate = ccount
End Sub

Sub setTblinfo_24bLoanAmt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet19.Range("PropertySectionCOunt").Value
    mIntCells = Range("loanAmt.24b" & i).count
    Set rangecells = Range("loanAmt.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanAmt = ccount
End Sub

Sub setTblinfo_24bLoanOutstanding()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet19.Range("PropertySectionCOunt").Value
    mIntCells = Range("loanOutstanding.24b" & i).count
    Set rangecells = Range("loanOutstanding.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bLoanOutstanding = ccount
End Sub

Sub setTblinfo_24bIntrst()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr, i As Long
    Dim ccount As Long
    ccount = 0
    For i = 1 To Sheet19.Range("PropertySectionCOunt").Value
    mIntCells = Range("Intrst.24b" & i).count
    Set rangecells = Range("Intrst.24b" & i).Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    Next i
    end_24bIntrst = ccount
End Sub


Sub setTblinfo_24bdate(ByVal myindex As Long)
 Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet19.Range("loanDate.24b" & myindex).count
    Set rangecells = Sheet19.Range("loanDate.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).Value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24f = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub
Sub setTblinfo_24bAmount(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet19.Range("loanAmt.24b" & myindex).count
    Set rangecells = Sheet19.Range("loanAmt.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).Value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24k = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"
End Sub
Sub setTblinfo_24bloanout(ByVal myindex As Long)
Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet19.Range("loanOutstanding.24b" & myindex).count
    Set rangecells = Sheet19.Range("loanOutstanding.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).Value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24s = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"

End Sub
Sub setTblinfo_24bInt(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet19.Range("Intrst.24b" & myindex).count
    Set rangecells = Sheet19.Range("Intrst.24b" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).Value
        
        If scode = "" Or IsEmpty(scode) Or (scode) = ("(Select)") Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco_24r = ccount
    rngname_hpco_24b = "LoanfrmBankOrInstitute.24b;bankName.24b;loanAccNum.24b;loanDate.24b;loanAmt.24b;loanOutstanding.24b;Intrst.24b"

End Sub

'========
    Sub setTblinfo_hprptfrm()
    Dim te As Long
    Dim ccount As Long
    
    frmsize_hprptfrm = Sheet19.Range("NumRowsPropertyBlock").Value
    For te = 1 To Sheet19.Range("PropertySectionCOunt").Value
        If Not Sheet19.Range("HP.AddrDetail" & te).Value = "" Then
            ccount = ccount + 1
        Else
            Exit For
        End If
    Next
    end_hprptfrm = ccount
    cntrRng_hprptfrm = "PropertySectionCOunt"
    frmRngname_hprptfrm = "hprptfrm"
    rngname_hprptfrm = "HP.AddrDetail1;HP.CityOrTownOrDistrict1;HP.CountryCode1;HP.StateCode1;HP.ZipCode1;HP.PinCode1;HP.OwnerProperty1;HP.OwnerPropertyDescription1;HP.CoOwnedYN1;HP.SharePercent1;HP.Co.Srno1;HP.Co.Name1;HP.Co.PAN1;HP.Co.Aadhaar1;HP.Co.Share1;HP.ifLetOut1;HP.NameofTenant1;HP.PANofTenant1;HP.AadharofTenant1;HP.TANofTenant1;HP.AnnualLetableValue1;HP.RentNotRealized1;HP.LocalTaxes1;HP.TotalUnrealizedAndTax1;HP.BalanceALV1;HP.ThirtyPercentOfBalance1;HP.IntOnBorwCap1;HP.TotalDeduct1;HP.RentOfEarlierYrSec_AandAA1;HP.IncomeOfHP1;HP.IncomeOfHPInOwnHand1;Co_OwnnerBlock1;LoanfrmBankOrInstitute.24b1;bankName.24b1;loanAccNum.24b1;loanDate.24b1;loanAmt.24b1;loanOutstanding.24b1;Intrst.24b1;TotAmt.24b1;Combination_24B1"
 End Sub

    
 Function ValidateAddrDetail_HP() As Boolean
    Dim i As Long
    ValidateAddrDetail_HP = True
    setTblinfo_hprptfrm
    
    ReDim AddrDetail_HP(end_hprptfrm)
    noOfProcessSub = end_hprptfrm
    For i = 1 To end_hprptfrm
        AddrDetail_HP(i) = Sheet19.Range("HP.AddrDetail" & i).Value

'Ankita_27/02/2026========
        If Not checkfieldspecialcharacter3(AddrDetail_HP(i)) Then
'            msgbox_hprptfrm ("* Address of Property " & i & " characters < > & ' " & Chr(34) & " are not allowed schedule HP ")
             msgbox_hprptfrm ("* Address " & i & " characters < > & ' " & Chr(34) & " are not allowed schedule HP ")
            ValidateAddrDetail_HP = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Sub setTblinfo_hpcoindex(ByVal myindex As Long)
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    mIntCells = Sheet19.Range("hp.co.name" & myindex).count
    Set rangecells = Sheet19.Range("hp.co.name" & myindex).Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).Value
        
        If scode = "" Or IsEmpty(scode) Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco = ccount
    rngname_hpco = "HP.Co.Srno;HP.Co.Name;HP.Co.Pan;HP.Co.Aadhaar;HP.Co.Share;"
End Sub

   
Function ValidateCoName_HP() As Boolean
    Dim i As Long
    Dim j As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    
    Dim state_1 As Variant
    Dim sharepercent As Variant
    Dim cosharepercent As Double
    
    ValidateCoName_HP = True
    

    sharepercent = 0
    cosharepercent = 0
    setTblinfo_hprptfrm
    For i = 1 To end_hprptfrm
        sharepercent = Sheet19.Range("HP.SharePercent" & i).text
        Set rangecells1 = Sheet19.Range("HP.Co.Name" & i).Cells
        Set rangecells2 = Sheet19.Range("HP.Co.PAN" & i).Cells
        Set rangecells3 = Sheet19.Range("HP.Co.Share" & i).Cells
        Set rangecells4 = Sheet19.Range("HP.StateCode" & i).Cells
        Set rangecells5 = Sheet19.Range("HP.Co.Aadhaar" & i).Cells
        setTblinfo_hpcoindex (i)
        
        ReDim CoName_HP(end_hprptfrm, end_hpco)
        ReDim CoPAN_HP(end_hprptfrm, end_hpco)
        ReDim CoAadhar_HP(end_hprptfrm, end_hpco)
        ReDim CoShare_HP(end_hprptfrm, end_hpco)
        
        state_1 = rangecells4.Value
        state_1 = Mid(state_1, 1, 2)
        
        cosharepercent = 0
        For j = 1 To end_hpco
            CoName_HP(i, j) = rangecells1.item(j).Value
            CoPAN_HP(i, j) = rangecells2.item(j).Value
            CoAadhar_HP(i, j) = rangecells5.item(j).Value
            CoShare_HP(i, j) = rangecells3.item(j).text
            If CoShare_HP(i, j) <> "" Then
                CoShare_HP(i, j) = Round(CoShare_HP(i, j), 2)
            
                cosharepercent = cosharepercent + IIf(CoShare_HP(i, j) = "", 0, CoShare_HP(i, j))
            End If
            
            If Not ValidatePAN(CStr(CoPAN_HP(i, j))) Then
                msgbox_hprptfrm ("* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in schedule HP")
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not ValidateAadhaar(CStr(CoAadhar_HP(i, j))) Then
                msgbox_hprptfrm ("* Enter Valid Aadhaar of CO Owner " & j & "  in Property " & i & "  with 12 digits schedule HP")
                ValidateCoName_HP = False
                Exit Function
            End If
        Next
        
         If sharepercent < 0.01 Then
            msgbox_hprptfrm ("* ""Your percentage of share in co-owned property cannot be zero""")
            ValidateCoName_HP = False
            Exit Function
        End If
    Next
endlin1:
End Function
    
Function ValidateCoName_HPold1() As Boolean
    Dim i As Long
    Dim j As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    
    Dim state_1 As Variant
    Dim sharepercent As Variant
    Dim cosharepercent As Double
    
    ValidateCoName_HP = True
    sharepercent = 0
    cosharepercent = 0
    setTblinfo_hprptfrm
    For i = 1 To end_hprptfrm
        sharepercent = Sheet19.Range("HP.SharePercent" & i).text
        Set rangecells1 = Sheet19.Range("HP.Co.Name" & i).Cells
        Set rangecells2 = Sheet19.Range("HP.Co.PAN" & i).Cells
        Set rangecells3 = Sheet19.Range("HP.Co.Share" & i).Cells
        Set rangecells4 = Sheet19.Range("HP.StateCode" & i).Cells
        Set rangecells5 = Sheet19.Range("HP.Co.Aadhaar" & i).Cells
        
        
        setTblinfo_hpcoindex (i)
        setTblinfo_24bankname (i)
        ReDim CoName_HP(end_hprptfrm, end_hpco)
        ReDim CoPAN_HP(end_hprptfrm, end_hpco)
        ReDim CoAadhar_HP(end_hprptfrm, end_hpco)
        ReDim CoShare_HP(end_hprptfrm, end_hpco)
        
        state_1 = rangecells4.Value
        state_1 = Mid(state_1, 1, 2)
        
        cosharepercent = 0
        For j = 1 To end_hpco
            CoName_HP(i, j) = rangecells1.item(j).Value
            CoPAN_HP(i, j) = rangecells2.item(j).Value
            CoAadhar_HP(i, j) = rangecells5.item(j).Value
            CoShare_HP(i, j) = rangecells3.item(j).text
            
            CoShare_HP(i, j) = Round(CoShare_HP(i, j), 2)
            
            cosharepercent = cosharepercent + IIf(CoShare_HP(i, j) = "", 0, CoShare_HP(i, j))
    
            If Not chkCompulsory(CoShare_HP(i, j)) Then
                msgbox_hprptfrm ("* Share of other CO-Owner " & j & "  in Property " & i & " is Mandatory schedule HP")
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not checkfieldspecialcharacter(CoName_HP(i, j)) Then
                msgbox_hprptfrm ("* Name of other CO-Owner " & j & "  in Property " & i & "  cannot contain special characters schedule HP")
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not ValidatePAN(CStr(CoPAN_HP(i, j))) Then
                msgbox_hprptfrm ("* Enter Valid PAN of other CO-Owner " & j & "  in Property " & i & "  with 1st 5 alphabets, next 4 digits and last alphabet schedule HP")
                ValidateCoName_HP = False
                Exit Function
            End If
            
            If Not ValidateAadhaar(CStr(CoAadhar_HP(i, j))) Then
                msgbox_hprptfrm ("* Enter Valid Aadhaar of other CO-Owner " & j & "  in Property " & i & "  with 12 digits schedule HP")
                ValidateCoName_HP = False
                Exit Function
            End If
        Next
        
         If sharepercent < 0.01 Then
            msgbox_hprptfrm ("*Share Percentage of Owner and CO Owner must add to 100 in Property " & i & " schedule HP")
            ValidateCoName_HP = False
            Exit Function
        End If
    Next
endlin1:
End Function
 
  Function ValidateCityOrTownOrDistrict_HP() As Boolean
    Dim i As Long
    
    ValidateCityOrTownOrDistrict_HP = True
    setTblinfo_hprptfrm
    ReDim CityOrTownOrDistrict_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        
        CityOrTownOrDistrict_HP(i) = Sheet19.Range("HP.CityOrTownOrDistrict" & i).Value
        If Not chkCompulsory(CityOrTownOrDistrict_HP(i)) Then
            msgbox_hprptfrm ("* ""Town/City is mandatory in tab: Sl no B2""")
            
            ValidateCityOrTownOrDistrict_HP = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter3(CityOrTownOrDistrict_HP(i)) Then
            msgbox_hprptfrm ("* CityOrTownOrDistrict" & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidateCityOrTownOrDistrict_HP = False
            Exit Function
        End If
    Next
End Function
  
 Function ValidateStateCode_HP() As Boolean
    Dim i As Long
    ValidateStateCode_HP = True
    setTblinfo_hprptfrm
    ReDim StateCode_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
           
        StateCode_HP(i) = Sheet19.Range("HP.StateCode" & i).Value
        StateCode_HP(i) = Mid(StateCode_HP(i), 1, 2)
        If isdropdownblank(StateCode_HP(i)) Then
            StateCode_HP(i) = ""
        End If
        
        If Not chkCompulsory(StateCode_HP(i)) Then
            msgbox_hprptfrm ("* ""State of property is mandatory in schedule HP.""")
            
            ValidateStateCode_HP = False
            Exit Function
        End If
    Next
End Function

Function ValidateCountryCode_HP() As Boolean
    Dim i As Long
    Dim Country1 As Variant
    ValidateCountryCode_HP = True
    setTblinfo_hprptfrm
    ReDim CountryCode_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
           
        CountryCode_HP(i) = Sheet19.Range("HP.CountryCode" & i).Value
        Country1 = Sheet19.Range("HP.CountryCode" & i).Value
        If isdropdownblank(CountryCode_HP(i)) Then
            CountryCode_HP(i) = ""
        End If
        
        If Not chkCompulsory(CountryCode_HP(i)) Then
            msgbox_hprptfrm ("* ""Country of property is mandatory in schedule HP.""")
            ValidateCountryCode_HP = False
            Exit Function
        End If
        CountryCode_HP(i) = Mid(Country1, 1, WorksheetFunction.Search("-", Country1) - 1)
                
    Next
End Function

Function ValidateCountryStateCode_HP() As Boolean
    Dim i As Long
    Dim Country1, state1 As Variant
    ValidateCountryStateCode_HP = True
    setTblinfo_hprptfrm
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If

        Country1 = Sheet19.Range("HP.CountryCode" & i).Value
        state1 = Sheet19.Range("HP.StateCode" & i).Value
        
        If isdropdownblank(Country1) Then
            Country1 = ""
        End If
        
        If isdropdownblank(state1) Then
            state1 = ""
        End If
        
        If chkCompulsory(Country1) And chkCompulsory(state1) Then
            Country1 = Mid(Country1, 1, WorksheetFunction.Search("-", Country1) - 1)
            state1 = Mid(state1, 1, WorksheetFunction.Search("-", state1) - 1)
                        
            If (Country1 <> "91" And state1 <> "99") Then
                msgbox_hprptfrm ("* ""Country cannot be other than India as you have selected an Indian state " & i & " in schedule HP""")
                ValidateCountryStateCode_HP = False
                Exit Function
            End If
            If (Country1 = "91" And state1 = "99") Then
                msgbox_hprptfrm ("* ""Country cannot be India as you have selected a Foreign state " & i & " in schedule HP""")
                ValidateCountryStateCode_HP = False
                Exit Function
            End If
        End If
    Next
End Function

Function ValidateZipCode_HP() As Boolean
    Dim i As Long
    
    ValidateZipCode_HP = True
    setTblinfo_hprptfrm
    ReDim ZipCode_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        ZipCode_HP(i) = Sheet19.Range("HP.ZipCode" & i).Value
        If CountryCode_HP(i) <> "91" Then
            If Not chkCompulsory(ZipCode_HP(i)) Then
                msgbox_hprptfrm ("* ""Please enter ZIP code, if ZIP code is not available, then enter XXXXXX in schedule HP.""")
                
                ValidateZipCode_HP = False
                Exit Function
            End If
            If Not checkfieldspecialcharacter1(ZipCode_HP(i)) Then
                msgbox_hprptfrm ("* ZipCode" & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
                ValidateZipCode_HP = False
                Exit Function
            End If
        Else
            ZipCode_HP(i) = ""
        End If
    Next
End Function


Function ValidatePinCode_HP() As Boolean
    Dim i As Long
    
    ValidatePinCode_HP = True
    setTblinfo_hprptfrm
    ReDim PinCode_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        PinCode_HP(i) = Sheet19.Range("HP.PinCode" & i).Value
        If CountryCode_HP(i) = "91" Then
            If Not chkCompulsory(PinCode_HP(i)) Then
               msgbox_hprptfrm ("* ""Pin code of Property is mandatory in schedule HP """)
                ValidatePinCode_HP = False
                Exit Function
            End If
            If Not chkNumeric(PinCode_HP(i)) Then
                msgbox_hprptfrm ("* PinCode" & i & " only digits 0 to 9 allowed in schedule HP")
                ValidatePinCode_HP = False
                Exit Function
            End If
        Else
            PinCode_HP(i) = ""
        End If
         If PinCode_HP(i) <> "" Then
         Dim PIN_targetadd1, state_targetadd1 As String
         
     
     PIN_targetadd1 = Replace(Sheet19.Range("HP.PinCode" & i).Address, "$", "")
     state_targetadd1 = Replace(PIN_targetadd1, "J", "H")
      
                             Dim ws1 As Worksheet
                             Set ws1 = Worksheets("HP")
         If Not PINstate_ModualValidation(ws1, PIN_targetadd1, state_targetadd1) Then
                Sheet19.Range("HP.PinCode" & i).Value = ""
                msgbox_hprptfrm ("* ""Pin code of Property is mandatory in schedule HP """) & Chr(13)
                ValidatePinCode_HP = False
                Exit Function
         End If

    
    End If
    Next
End Function


Function ValidateOwnerProperty_HP() As Boolean
    Dim i As Long

    ValidateOwnerProperty_HP = True
    setTblinfo_hprptfrm
    ReDim OwnerProperty_HP(end_hprptfrm)
    ReDim OwnerPropertyDescription_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        OwnerProperty_HP(i) = Sheet19.Range("HP.OwnerProperty" & i).Value
        OwnerPropertyDescription_HP(i) = Sheet19.Range("HP.OwnerPropertyDescription" & i).Value
        
            If isdropdownblank(OwnerProperty_HP(i)) Then
                msgbox_hprptfrm ("* ""Please select dropdown from owner of the Property in schedule HP.""")
                
                ValidateOwnerProperty_HP = False
                Exit Function
            End If
        If Mid(OwnerProperty_HP(i), 1, 1) = "O" Then
                  If Not chkCompulsory(OwnerPropertyDescription_HP(i)) Then
                  msgbox_hprptfrm ("* ""In case of others, mandatory to enter owner in textbox in schedule HP""")
                  ValidateOwnerProperty_HP = False
                 Exit Function
                End If
              
              If Len(OwnerPropertyDescription_HP(i)) > 50 Then
                 msgbox_hprptfrm ("* Description for Owner of the Property" & i & "  Cannot be more than 50 characters in schedule HP")
                 ValidateOwnerProperty_HP = False
                 Exit Function
                End If
            
             If Not checkfieldspecialcharacter(OwnerPropertyDescription_HP(i)) Then
                msgbox_hprptfrm ("* Description for Owner of the Property" & i & "  characters < > & ' " & Chr(34) & " are not allowed in schedule HP") & Chr(13)
                ValidateOwnerProperty_HP = False
                Exit Function
             End If
        End If
        
    Next
End Function

Function ValidateifLetOut_HP() As Boolean
    Dim i As Long
    ValidateifLetOut_HP = True
    
    setTblinfo_hprptfrm
    ReDim ifLetOut_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        ifLetOut_HP(i) = Sheet19.Range("HP.ifLetOut" & i).Value
        ifLetOut_HP(i) = Mid(ifLetOut_HP(i), 1, 1)
        
        If isdropdownblank(ifLetOut_HP(i)) Then
            ifLetOut_HP(i) = ""
        End If
        
        If Not chkCompulsory(ifLetOut_HP(i)) Then
            msgbox_hprptfrm ("* ""Please select type of House Property in schedule HP """)
            ValidateifLetOut_HP = False
            Exit Function
        End If
    Next
End Function

Function ValidateNameofTenant_HP() As Boolean
    Dim i, count As Long
    Dim tempPan As String
    Dim rangecells As Range
    ValidateNameofTenant_HP = True
    setTblinfo_hprptfrm
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        
        Set rangecells = Sheet19.Range("HP.NameofTenant" & i).Cells
        ReDim NameofTenant_HP(rangecells.Cells.count)
        Dim j As Variant
        count = 0
        For j = 1 To rangecells.Cells.count
        
        NameofTenant_HP(j) = rangecells.item(j).Value
        If Len(NameofTenant_HP(j)) > 0 Then
            If Not checkfieldspecialcharacter1(NameofTenant_HP(j)) Then
                msgbox_hprptfrm ("* NameofTenant" & j & "  characters < > & ' " & Chr(34) & " are not allowed in Sheet HOUSE_PROPERTY")
                ValidateNameofTenant_HP = False
                Exit Function
            End If
        End If
        If Len(NameofTenant_HP(j)) = 0 And (Mid(ifLetOut_HP(i), 1, 1) <> "S" And Mid(ifLetOut_HP(i), 1, 1) <> "(") Then
            If (Mid(ifLetOut_HP(i), 1, 1) <> "D") Then
            count = count + 1
            End If
        End If
       
        Next
          If count = UBound(NameofTenant_HP) Then
             msgbox_hprptfrm ("* ""Please enter name of Tenant in schedule HP""")
            ValidateNameofTenant_HP = False
            Exit Function
        End If
    Next
End Function
 
Function ValidatePANofTenant_HP() As Boolean
    Dim i As Long
    Dim tempPan As String
    Dim rangecells As Range
    
    ValidatePANofTenant_HP = True
    setTblinfo_hprptfrm
    
    
    
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        Set rangecells = Sheet19.Range("HP.PANofTenant" & i).Cells
        ReDim PANofTenant_HP(rangecells.Cells.count)
        Dim j As Variant
        
        For j = 1 To rangecells.Cells.count
        PANofTenant_HP(j) = rangecells.item(j).Value
        
        tempPan = PANofTenant_HP(j)
        If Len(PANofTenant_HP(j)) > 0 Then
            If Not ValidatePAN(tempPan) Then
                  msgbox_hprptfrm ("* ""Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in schedule HP""")
                ValidatePANofTenant_HP = False
                Exit Function
            End If
        End If
        
        Next
        
    Next
End Function
Function CheckAtoZ(chr1) As Boolean
    CheckAtoZ = True
    If ((asc(chr1) < 65) Or (asc(chr1) > 90)) Then
        CheckAtoZ = False
    End If
End Function


Function ValidatePAN(panentry As String) As Boolean
    ValidatePAN = True

    If Len(panentry) > 0 Then
        If Not IsNumeric(Mid(panentry, 6, 4)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 1, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 2, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 3, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 4, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 5, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
        
        If Not CheckAtoZ(Mid(panentry, 10, 1)) Then
            ValidatePAN = False
            Exit Function
        End If
    End If
End Function
Function ValidateLoanfrm_24b() As Boolean
    ValidateLoanfrm_24b = True

    Dim rangecells As Range
    Dim i, j As Long
For j = 1 To Sheet19.Range("PropertySectionCOunt").count
    Set rangecells = Sheet19.Range("LoanfrmBankOrInstitute.24b" & j).Cells
    ReDim Loanfrm_24b(end_24b)
    For i = 1 To end_24b
        Loanfrm_24b(i) = rangecells.item(i).Value
        If isdropdownblank(Loanfrm_24b(i)) Then
             msgbox_hprptfrm ("* Please select dropdown from ""Loan taken from"" in schedule 24(b) at Sr. No " & i & "") & Chr(13)
            ValidateLoanfrm_24b = False
            Exit Function
        End If
         If Loanfrm_24b(i) = "(Select)" Then
          msgbox_hprptfrm ("* Please select dropdown from ""Loan taken from"" at Sr. No " & i & " in 24(b) schedule at Sr. No " & i & "") & Chr(13)
            ValidateLoanfrm_24b = False
            Exit Function
        End If

Next i
Next j
End Function

Function ValidateBankName_24b() As Boolean
    ValidateBankName_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet19.Range("PropertySectionCOunt").count
    Set rangecells = Sheet19.Range("bankName.24b" & j).Cells
    ReDim BankName_24B(end_24b)
    For i = 1 To end_24b
        BankName_24B(i) = rangecells.item(i).Value
        If Not chkCompulsory(BankName_24B(i)) Then
            msgbox_hprptfrm ("*""Please provide Name of the Bank/ Institution/Person from which the loan is taken in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateBankName_24b = False
            Exit Function
        End If
        
         If Len(BankName_24B(i)) > 125 Then
           msgbox_hprptfrm ("* Name of the Bank/ Institution/Person at Sr. No " & i & " in Sheet 24(b) should be less than or equal to 125 characters.") & Chr(13)
            ValidateBankName_24b = False
            Exit Function
        End If
        
        
        If Not checkfieldSuperSpecialcharacter(BankName_24B(i)) Then
            msgbox_hprptfrm ("* Name of the Bank/ Institution/Person in schedule 24b at Sl.no. " & i & " should not Contain <, >, characters.") & Chr(13)
            ValidateBankName_24b = False
            Exit Function
        End If
         
Next i
Next j
End Function

Function ValidateAccntNum_24b() As Boolean
    ValidateAccntNum_24b = True

    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim i, j As Long
    For j = 1 To Sheet19.Range("PropertySectionCOunt").count
    Set rangecells = Sheet19.Range("loanAccNum.24b" & j).Cells
   
    Set rangecells1 = Sheet19.Range("LoanfrmBankOrInstitute.24b" & j).Cells
    ReDim AccntNum_24b(end_24b)
    ReDim BankorInst_24b(end_24b)
    For i = 1 To end_24b
        AccntNum_24b(i) = rangecells.item(i).Value
        
        BankorInst_24b(i) = rangecells1.item(i).Value
        If Not chkCompulsory(AccntNum_24b(i)) Then
            msgbox_hprptfrm ("* ""Please provide Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateAccntNum_24b = False
            Exit Function
        End If
        
         If Len(AccntNum_24b(i)) > 20 Then
            msgbox_hprptfrm ("* Loan Account number  at Sr. No " & i & " in schedule 24b less than 20 characters.") & Chr(13)
            ValidateAccntNum_24b = False
            Exit Function
        End If
        
        
        If BankorInst_24b(i) = "Bank" Then
        If Not ValidateBankAccountNumber_24b(AccntNum_24b(i)) Then
        ValidateAccntNum_24b = False
        Exit Function
         End If
        End If
    
    If BankorInst_24b(i) = "Other than bank" Then
        If Not checkfieldspecialcharacter1(AccntNum_24b(i)) Then
            msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr.No " & i & "") & Chr(13)
           ValidateAccntNum_24b = False
            Exit Function
        End If
        End If
                 
Next i
Next j
End Function
   Function ValidateLoanDate_24b() As Boolean
    ValidateLoanDate_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet19.Range("PropertySectionCOunt").count
    Set rangecells = Sheet19.Range("loanDate.24b" & j).Cells
    ReDim LoanDate_24B(end_24b)
    For i = 1 To end_24b
        LoanDate_24B(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanDate_24B(i)) Then
             msgbox_hprptfrm ("* ""Please provide Date of sanction of Loan in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateLoanDate_24b = False
            Exit Function
        End If
         If Len(LoanDate_24B(i)) > 10 Then
          msgbox_hprptfrm ("* Date of sanction of Loan  at Sr. No " & i & " in schedule 24b less than 10 characters.") & Chr(13)
         ValidateLoanDate_24b = False
            Exit Function
        End If
         If Not CheckDateddmmyyyy(LoanDate_24B(i)) Then
            msgbox_hprptfrm ("* ""Please enter date in valid format"" at Sr. No " & i & ".") & Chr(13)
            ValidateLoanDate_24b = False
            Exit Function
        End If
        
        If Not ChkMaxDate_24b(Trim(LoanDate_24B(i)), "31-03-2026") Then
            msgbox_hprptfrm ("* Date can not be after 31/03/2026 at Sr. No " & i & ".") & Chr(13)
            ValidateLoanDate_24b = False
          Exit Function
        End If
                 
                                 
Next i
Next j
End Function

Function ValidateLoanAmt_24b() As Boolean
    ValidateLoanAmt_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet19.Range("PropertySectionCOunt").count
    Set rangecells = Sheet19.Range("loanAmt.24b" & j).Cells
    ReDim LoanAmt_24b(end_24b)
    For i = 1 To end_24b
        LoanAmt_24b(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanAmt_24b(i)) Then
                msgbox_hprptfrm ("* ""Please provide Total amount of loan in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateLoanAmt_24b = False
            Exit Function
        End If
        If Not IsNumeric(LoanAmt_24b(i)) Then
            msgbox_hprptfrm ("* Loan amount at Sr. No  " & i & "  in schedule 24b should be Numeric value") & Chr(13)
            ValidateLoanAmt_24b = False
            Exit Function
        End If
        
        If LoanAmt_24b(i) > 99999999999999# Then
            msgbox_hprptfrm ("* Loan amount at Sr. No  " & i & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
            ValidateLoanAmt_24b = False
            Exit Function
        End If
        
        If LoanAmt_24b(i) < 0 Or LoanAmt_24b(i) = 0 Then
           msgbox_hprptfrm ("* Total amount of loan should be more than 0 in schedule 24(b) at Sr. No  " & i & "") & Chr(13)
            ValidateLoanAmt_24b = False
            Exit Function
        End If
         
Next i
Next j
End Function
Function ValidateLoanOutstanding_24b() As Boolean
    ValidateLoanOutstanding_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet19.Range("PropertySectionCOunt").count
    Set rangecells = Sheet19.Range("loanOutstanding.24b" & j).Cells
    ReDim LoanOutStanding_24B(end_24b)
    For i = 1 To end_24b
        LoanOutStanding_24B(i) = rangecells.item(i).Value
        If Not chkCompulsory(LoanOutStanding_24B(i)) Then
                msgbox_hprptfrm ("* ""Loan outstanding as on last date of financial year is mandatory in schedule 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateLoanOutstanding_24b = False
            Exit Function
        End If
        If Not IsNumeric(LoanOutStanding_24B(i)) Then
            msgbox_hprptfrm ("* Loan outstanding at Sr. No  " & i & "  in schedule 24b should be Numeric value") & Chr(13)
            ValidateLoanOutstanding_24b = False
            Exit Function
        End If
        
        If LoanOutStanding_24B(i) > 99999999999999# Then
            msgbox_hprptfrm ("* Loan outstanding at Sr. No  " & i & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
            ValidateLoanOutstanding_24b = False
            Exit Function
        End If
        
        If LoanOutStanding_24B(i) < 0 Then
             msgbox_hprptfrm ("* Loan outstanding as on last date of finacial year can't be less than 0 in schedule 24(b) at Sr. No  " & i & ". You may please enter as 0 if it become negative as result of excess payment.") & Chr(13)
            ValidateLoanOutstanding_24b = False
            Exit Function
        End If
         
Next i
Next j
End Function


Function ValidateIntrst_24b() As Boolean
    ValidateIntrst_24b = True

    Dim rangecells As Range
    Dim i, j As Long
    For j = 1 To Sheet19.Range("PropertySectionCOunt").count
    Set rangecells = Sheet19.Range("Intrst.24b" & j).Cells
    ReDim Intrst_24B(end_24b)
    For i = 1 To end_24b
        Intrst_24B(i) = rangecells.item(i).Value
        If Not chkCompulsory(Intrst_24B(i)) Then
             msgbox_hprptfrm ("* ""Please provide Interest u/s 24(b)"" at Sr. No " & i & "") & Chr(13)
            ValidateIntrst_24b = False
            Exit Function
        End If
        If Not IsNumeric(Intrst_24B(i)) Then
           msgbox_hprptfrm ("* Interest at Sr. No  " & i & "  in schedule 24b should be Numeric value") & Chr(13)
            
            ValidateIntrst_24b = False
            Exit Function
        End If
        
        If Intrst_24B(i) > 99999999999999# Then
            msgbox_hprptfrm ("* Interest at Sr. No  " & i & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
            ValidateIntrst_24b = False
            Exit Function
        End If
        If Intrst_24B(i) < 0 Or Intrst_24B(i) = 0 Then
            msgbox_hprptfrm ("* Interest u/s 24(b) should be more than 0 in schedule 24(b) at Sr. No  " & i & "") & Chr(13)
            ValidateIntrst_24b = False
            Exit Function
        End If
         
Next i
Next j
End Function

Function ValidateBankAccountNumber_24b(BankAccountNumber As Variant) As Boolean
    ValidateBankAccountNumber_24b = True
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
    Dim cc
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter1(BankAccountNumber) Then
            msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr.No " & cc & "") & Chr(13) 'updated by Shrutika(SIT-93563)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            msgbox_hprptfrm ("*""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If
    
    End If
  '----------------------------------------------------------------
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
    
    
    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        msgbox_hprptfrm ("* Loan Account number at Sr.No " & cc & " is mandatory in schedule 24(b)") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
'--------------------------------------------------------------------

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0) Or (InStr(BankAccountNumber, "-/") > 0) Or (InStr(BankAccountNumber, "/-") > 0)) Then
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
      msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
    
    If (Not checkfieldspecialcharacter(Mid(BankAccountNumber, 1, 1))) Then
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If


    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
        ValidateBankAccountNumber_24b = False
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
           msgbox_hprptfrm ("*""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr.No " & cc & "") & Chr(13)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If
    End If

    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr.No " & cc & "" & Chr(13)
        ValidateBankAccountNumber_24b = False
        Exit Function
    End If
    
    
       'Ankita_27/02/2026=================
        Dim has19 As Boolean, ch As String
        For i = 1 To Len(BankAccountNumber)
            ch = Mid$(BankAccountNumber, i, 1)
            If ch Like "[1-9]" Then
                has19 = True
                Exit For
            End If
        Next i
        If Not has19 Then
            msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)""") & Chr(13)
            ValidateBankAccountNumber_24b = False
            Exit Function
        End If
    '----------
End Function

Function Validategreater_24b() As Boolean
    Validategreater_24b = True
         If (Len(Sheet19.Range("TotAmt.24b1").Value) > 14) Then
            msgbox_hprptfrm ("*  Total of interest on borrowed capital u/s 24(b) cannot exceed 14 Digits.") & Chr(13)
            Validategreater_24b = False
            Exit Function
         End If
End Function


Function ValidateAadharofTenant_HP() As Boolean
    Dim i As Long
    Dim tempPan As String
    Dim rangecells As Range

    ValidateAadharofTenant_HP = True
    setTblinfo_hprptfrm

    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        Set rangecells = Sheet19.Range("HP.AadharofTenant" & i).Cells
        ReDim AadharofTenant_HP(rangecells.Cells.count)
        Dim j As Variant

        For j = 1 To rangecells.Cells.count
        AadharofTenant_HP(j) = rangecells.item(j).Value

          If Len(AadharofTenant_HP(j)) > 0 Then
            If Len(AadharofTenant_HP(j)) <> 12 Then
                msgbox_hprptfrm ("*Aadhaar no of other Tenant " & j & " shall not exceed or less than 12 Digits in schedule HP")
                ValidateAadharofTenant_HP = False
                Exit Function
                End If

                If AadharofTenant_HP(j) = "000000000000" Then
                msgbox_hprptfrm ("*Invalid Aadhaar of Tenant should be 12 digits " & j & " shall not exceed or less than 12 Digits in schedule HP")
                ValidateAadharofTenant_HP = False
                Exit Function
                End If

                If AadharofTenant_HP(j) = "111111111111" Then
                msgbox_hprptfrm ("*Invalid Aadhaar of Tenant should be 12 digits " & j & " shall not exceed or less than 12 Digits in schedule HP")
                ValidateAadharofTenant_HP = False
                Exit Function
                End If

                If Not checkallfieldspecialcharacter(AadharofTenant_HP(j)) Then
                msgbox_hprptfrm ("* Aadhar of Tenant" & j & "  Special characters are not allowed in schedule HP")
                ValidateAadharofTenant_HP = False
                Exit Function
                End If
            End If
        Next

    Next
End Function


Function checkallfieldspecialcharacter(field As Variant) As Boolean
    Dim i, j As Long
    checkallfieldspecialcharacter = True
    Dim arr As Variant
    arr = Array("@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", "&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkallfieldspecialcharacter = False
            Exit Function
        End If
        Next
    Next
End Function

Function ValidateAadhaar(Aadhaar As String) As Boolean
    ValidateAadhaar = True

   If Aadhaar <> "" Then
        If Not IsNumeric(Aadhaar) Then
            ValidateAadhaar = False
            Exit Function
        End If
    
        If Aadhaar = "000000000000" Then
            ValidateAadhaar = False
            Exit Function
        End If
        
        If Aadhaar = "111111111111" Then
            ValidateAadhaar = False
            Exit Function
        End If
            
        If Len(Aadhaar) <> 12 Then
            ValidateAadhaar = False
            Exit Function
        End If
        Dim i, j As Long
        Dim arr As Variant
        arr = Array(".", ",")
        For i = 1 To Len(Aadhaar)
        For j = 0 To UBound(arr)
        If Mid(Aadhaar, i, 1) = arr(j) Then
        ValidateAadhaar = False
            Exit Function
        End If
        Next
        Next
    End If
End Function


Function ValidateTANofTenant_HP() As Boolean
    Dim i As Long
    Dim tempTan As String
    Dim rangecells As Range
    
    ValidateTANofTenant_HP = True
    setTblinfo_hprptfrm
    For i = 1 To end_hprptfrm
        If Len(AddrDetail_HP(i)) = 0 Then
            Exit For
        End If
        Set rangecells = Sheet19.Range("HP.TANofTenant" & i).Cells
        ReDim TANofTenant_HP(rangecells.Cells.count)
        Dim j As Variant
        
        For j = 1 To rangecells.Cells.count
        TANofTenant_HP(j) = rangecells.item(j).Value
        
        tempTan = TANofTenant_HP(j)
        If Len(TANofTenant_HP(j)) > 0 Then
        If Not (CheckTAN(tempTan) Or ValidateTenantPan(tempTan)) Then
'                msgbox_hprptfrm ("* ""Invalid PAN/TAN. PAN/TAN format should be First Four Alphabets, next 5 digits, then 1 Alphabet or First 4 Alphabets, next 5 digits, then 1 Alphabet in schedule HP""")
                 msgbox_hprptfrm ("""Invalid PAN/TAN. PAN/TAN format should be First Four Alphabets, next 5 digits, then 1 Alphabet in case of  TAN OR First Five Alphabets, next 4 digits, then 1 Alphabet in schedule HP in case of PAN""")
                ValidateTANofTenant_HP = False
                Exit Function
            End If
        End If
        Next
    Next
End Function

Function CheckTAN(TAN As Variant) As Boolean
On Error Resume Next
TAN = UCase(TAN)
    CheckTAN = True
    If Len(TAN) > 0 Then
        If Not ChkAlphabet(Mid(TAN, 1, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(TAN, 2, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(TAN, 3, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        
      If Not ChkAlphabet(Mid(TAN, 4, 1)) Then
            CheckTAN = False
            Exit Function
        End If
            
        If Not IsNumeric(Mid(TAN, 5, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(TAN, 6, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(TAN, 7, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(TAN, 8, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        
        If Not IsNumeric(Mid(TAN, 9, 1)) Then
            CheckTAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(TAN, 10, 1)) Then
            CheckTAN = False
            Exit Function
        End If
    End If
End Function

Function CheckHousePropertyIncome(HP As Variant) As Boolean
On Error Resume Next
 CheckHousePropertyIncome = True
 If (HP > 200000) Then
 fmsgboxsmall "* Interest cannot be more than Rs. 2,00,000 for category self-occupied in schedule HP." ', vbOKOnly, "Error(s)!"
  CheckHousePropertyIncome = False
 Exit Function
 End If
End Function


Function ValidateAnnualLetableValue_HP() As Boolean
    Dim i As Long
    ValidateAnnualLetableValue_HP = True
    setTblinfo_hprptfrm
    
    ReDim AnnualLetableValue_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        AnnualLetableValue_HP(i) = Sheet19.Range("HP.AnnualLetableValue" & i).Value
          If AnnualLetableValue_HP(i) = 0 And (Mid(ifLetOut_HP(i), 1, 1) = "D" Or Mid(ifLetOut_HP(i), 1, 1) = "L") Then
            msgbox_hprptfrm ("* If ""Type of House Property"" is ""Let Out/ Deemed Let Out"" then Gross Rent received or receivable or lettable value field cannot be zero/ blank")
            ValidateAnnualLetableValue_HP = False
            Exit Function
        End If
    Next
End Function

Function ValidateRentNotRealized_HP() As Boolean
    Dim i As Long
    
    ValidateRentNotRealized_HP = True
    setTblinfo_hprptfrm
    ReDim RentNotRealized_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        RentNotRealized_HP(i) = Sheet19.Range("HP.RentNotRealized" & i).Value
    Next
End Function

Function ValidateLocalTaxes_HP() As Boolean
    Dim i As Long
    
    ValidateLocalTaxes_HP = True
    setTblinfo_hprptfrm
    ReDim LocalTaxes_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        LocalTaxes_HP(i) = Sheet19.Range("HP.LocalTaxes" & i).Value
    Next
End Function

Function ValidateTotalUnrealizedAndTax_HP() As Boolean
    Dim i As Long
    
    ValidateTotalUnrealizedAndTax_HP = True
    setTblinfo_hprptfrm
    ReDim TotalUnrealizedAndTax_HP(end_hprptfrm)
    
    For i = 1 To end_hprptfrm
        TotalUnrealizedAndTax_HP(i) = Sheet19.Range("HP.TotalUnrealizedAndTax" & i).Value
    Next
End Function

Function ValidateBalanceALV_HP() As Boolean
    Dim i As Long
    
    ValidateBalanceALV_HP = True
    setTblinfo_hprptfrm
    ReDim BalanceALV_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        BalanceALV_HP(i) = Sheet19.Range("HP.BalanceALV" & i).Value
    Next
End Function


Function ValidateThirtyPercentOfBalance_HP() As Boolean
    Dim i As Long
    
    ValidateThirtyPercentOfBalance_HP = True
    setTblinfo_hprptfrm
    ReDim ThirtyPercentOfBalance_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        ThirtyPercentOfBalance_HP(i) = Sheet19.Range("HP.ThirtyPercentOfBalance" & i).Value
    Next
End Function

Function ValidateIntOnBorwCap_HP() As Boolean
    Dim i As Long
    
    ValidateIntOnBorwCap_HP = True
    setTblinfo_hprptfrm
    ReDim IntOnBorwCap_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        IntOnBorwCap_HP(i) = Sheet19.Range("HP.IntOnBorwCap" & i).Value
    Next
End Function

Function ValidateTotalDeduct_HP() As Boolean
    Dim i As Long
    
    ValidateTotalDeduct_HP = True
    setTblinfo_hprptfrm
    ReDim TotalDeduct_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        TotalDeduct_HP(i) = Sheet19.Range("HP.TotalDeduct" & i).Value
    Next
End Function

Function ValidateArrears_HP() As Boolean
    Dim i As Long
    
    ValidateArrears_HP = True
    setTblinfo_hprptfrm
    ReDim Arrears_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        Arrears_HP(i) = Sheet19.Range("HP.RentOfEarlierYrSec_AandAA" & i).Value
    Next
End Function


Function ValidateIncomeOfHP_HP() As Boolean
    Dim i As Long
    
    ValidateIncomeOfHP_HP = True
    setTblinfo_hprptfrm
    
    ReDim IncomeOfHP_HP(end_hprptfrm)
    ReDim IncomeOfHPInOwnHand_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        IncomeOfHP_HP(i) = Sheet19.Range("HP.IncomeOfHP" & i).Value
        IncomeOfHPInOwnHand_HP(i) = Sheet19.Range("HP.IncomeOfHPInOwnHand" & i).Value
    Next
End Function

Function ValidateRentArearsSec25BAfter30pcDeduct_HP() As Boolean
    ValidateRentArearsSec25BAfter30pcDeduct_HP = True
End Function

Function ValidateTotalIncomeChargeableUnHP_HP() As Boolean
    ValidateTotalIncomeChargeableUnHP_HP = True
    TotalIncomeChargeableUnHP_HP = Sheet19.Range("HP.TotalIncomeChargeableUnHP").Value
    
    If Len(TotalIncomeChargeableUnHP_HP) > 14 Then
        msgbox_hprptfrm ("* Total income chargeable cannot be greater than 14 digits")
        ValidateTotalIncomeChargeableUnHP_HP = False
    End If
    
End Function

Function ValidateCoownerName1_HP() As Boolean
    Dim i, j As Long
    Dim rangecells As Range
    ValidateCoownerName1_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnerName1_HP(end_hprptfrm)
    ReDim CowOwned1_HP(end_hprptfrm)
    
    For j = 1 To end_hprptfrm
    Set rangecells = Sheet19.Range("HP.Co.Name" & j).Cells
    For i = 1 To end_hprptfrm
        CoOwnerName1_HP(i) = rangecells.item(i).Value
        
        'Ankita_05/02/2026===========
        CowOwned1_HP(i) = Sheet19.Range("HP.CoOwnedYN" & i).Value
        If Mid(CowOwned1_HP(i), 1, 1) = "Y" Then
        If rangecells.item(i).Offset(0, 1).Value <> "" Or rangecells.item(i).Offset(0, 2).Value <> "" Or rangecells.item(i).Offset(0, 3).Value <> "" Then
        If Not chkCompulsory(CoOwnerName1_HP(i)) Then
             msgbox_hprptfrm ("* ""Please enter the name of the other co-owner.""") & Chr(13)
            ValidateCoownerName1_HP = False
            Exit Function
        End If
        End If
        End If

    Next
    Next
End Function

Function ValidatePercenShare1_HP() As Boolean
    Dim i, j As Long
    Dim rangecells As Range
    ValidatePercenShare1_HP = True
    setTblinfo_hprptfrm
    
    
    ReDim CoOwnerSharePer1_HP(end_hprptfrm)
    For j = 1 To end_hprptfrm
    Set rangecells = Sheet19.Range("HP.Co.Share" & j).Cells
    For i = 1 To end_hprptfrm
        CoOwnerSharePer1_HP(i) = rangecells.item(i).Value
        If Not checkfieldspecialcharacter(CoOwnerSharePer1_HP(i)) Then
            msgbox_hprptfrm ("* Percentage Share in Property " & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidatePercenShare1_HP = False
            Exit Function
        End If
    Next
    Next
End Function

Function ValidateCoOwnerPAN1_HP() As Boolean
    Dim i, j As Long
    Dim tempPan As String
    Dim rangecells As Range
 
    ValidateCoOwnerPAN1_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnerPAN1_HP(end_hprptfrm)
    For j = 1 To end_hprptfrm
    Set rangecells = Sheet19.Range("HP.Co.Pan" & j).Cells
    
    For i = 1 To end_hprptfrm
        CoOwnerPAN1_HP(i) = rangecells.item(i).Value
        If Not checkfieldspecialcharacter(CoOwnerPAN1_HP(i)) Then
            msgbox_hprptfrm ("* PAN of CoownerA" & i & "   characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidateCoOwnerPAN1_HP = False
            Exit Function
        End If
        If UCase(CoOwnerPAN1_HP(i)) <> "" Then
        If UCase(CoOwnerPAN1_HP(i)) = Sheet1.Range("sheet1.PAN").Value Then
                msgbox_hprptfrm ("* Enter Valid PAN of Other co-owner(s) at Sr. No " & i & " in Property " & i & " schedule HP which should not match pan with Sheet Part A - General" & Chr(13))
                ValidateCoOwnerPAN1_HP = False
            Exit Function
        End If
        End If
    Next
    Next
End Function
Function ValidateCoOwnerAadhar_HP() As Boolean
    Dim i, j As Long
    Dim tempPan As String
    Dim rangecells As Range
 
    ValidateCoOwnerAadhar_HP = True
    setTblinfo_hprptfrm
    
    For j = 1 To end_hprptfrm
    Set rangecells = Sheet19.Range("HP.Co.Aadhaar" & j).Cells
    ReDim CoOwnerAadhar_HP(rangecells.count)
    
    For i = 1 To rangecells.count
        CoOwnerAadhar_HP(i) = rangecells.item(i).Value
        
         If Len(CoOwnerAadhar_HP(i)) > 0 Then
                If Len(CoOwnerAadhar_HP(i)) > 12 Then
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
                
                If CoOwnerAadhar_HP(i) = "000000000000" Then
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
                
                If CoOwnerAadhar_HP(i) = "111111111111" Then
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
            
                If Not checkallfieldspecialcharacter(CoOwnerAadhar_HP(i)) Then
                msgbox_hprptfrm "*""Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                ValidateCoOwnerAadhar_HP = False
                Exit Function
                End If
         End If
    Next
    Next
End Function

Function ValidateSharePercent_HP() As Boolean
    Dim i As Long
    
    ValidateSharePercent_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnedShare_HP(end_hprptfrm)
    ReDim CoOwnedYN_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        CoOwnedShare_HP(i) = Sheet19.Range("HP.SharePercent" & i).Value
        CoOwnedYN_HP(i) = Sheet19.Range("HP.CoOwnedYN" & i).Value
        
        CoOwnedShare_HP(i) = Round(CoOwnedShare_HP(i), 2)
        
        If Not checkfieldspecialcharacter(CoOwnedShare_HP(i)) Then
            msgbox_hprptfrm ("* Your Percentage of Share" & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidateSharePercent_HP = False
            Exit Function
        End If
        
        If Mid(CoOwnedYN_HP(i), 1, 1) = "Y" Then
            If CoOwnedShare_HP(i) = "" Or CoOwnedShare_HP(i) = 0 Then
            msgbox_hprptfrm ("* ""Sum of assessee percentage and co-owner(s) percentage must be equal to 100%  in schedule HP""")
            ValidateSharePercent_HP = False
            Exit Function
        End If
        End If
        
        If Mid(CoOwnedYN_HP(i), 1, 1) = "N" Then
            If CoOwnedShare_HP(i) <> 100 Then
                msgbox_hprptfrm ("* Your Percentage of Share" & i & " must be 100 if not Co-owned in schedule HP")
                ValidateSharePercent_HP = False
                Exit Function
            End If
        Else
        End If
    Next
End Function
Function ValidateCoOwnedYN_HP() As Boolean
    Dim i As Long
    
    ValidateCoOwnedYN_HP = True
    setTblinfo_hprptfrm
    ReDim CoOwnedYN_HP(end_hprptfrm)
    For i = 1 To end_hprptfrm
        CoOwnedYN_HP(i) = Sheet19.Range("HP.CoOwnedYN" & i).Value
        If isdropdownblank(CoOwnedYN_HP(i)) Then
            CoOwnedYN_HP(i) = ""
        End If
        If Not chkCompulsory(CoOwnedYN_HP(i)) Then
            msgbox_hprptfrm ("* ""Please select dropdown from Is property co-owned  in schedule HP""") & Chr(13)
            ValidateCoOwnedYN_HP = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(CoOwnedYN_HP(i)) Then
            msgbox_hprptfrm ("* IS Property Co-Owned" & i & "  characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidateCoOwnedYN_HP = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(CoOwnedYN_HP(i)) Then
            msgbox_hprptfrm ("* IS Property Co-Owned" & i & " characters < > & ' " & Chr(34) & " are not allowed in schedule HP")
            ValidateCoOwnedYN_HP = False
            Exit Function
        End If
    Next
End Function


Function ValidateCoownerRules_HP() As Boolean
    Dim i As Long
    ValidateCoownerRules_HP = True
    setTblinfo_hprptfrm
    For i = 1 To UBound(CoOwnedYN_HP)
        If isdropdownblank(CoOwnedYN_HP(i)) Then
            CoOwnedYN_HP(i) = ""
        End If

        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
          Dim CoOwnerShare As Variant
          CoOwnerShare = CoOwnedShare_HP(i)
          
        End If
  
        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
         Dim counter, name As Variant
         counter = 0
         For Each name In Sheet19.Range("HP.Co.Name" & i).Cells
         If name = "" Then counter = counter + 1
            If counter = Sheet19.Range("HP.Co.Name" & i).Cells.count Then
                msgbox_hprptfrm ("* ""There must be at least one Co-owner in schedule HP""")
                ValidateCoownerRules_HP = False
                Exit Function
                
             End If
        Next
                 
        Dim j As Long
        Dim coowener As Range
        Dim Pancoowener As Range
        Dim Aacoowener As Range
        Dim percoowener As Range
        For j = 1 To Sheet19.Range("HP.Co.Name" & i).count

            Set coowener = Sheet19.Range("HP.Co.Name" & i).Cells
            Set Pancoowener = Sheet19.Range("HP.Co.Pan" & i).Cells
            Set Aacoowener = Sheet19.Range("HP.Co.Aadhaar" & i).Cells
            Set percoowener = Sheet19.Range("HP.Co.Share" & i).Cells
            If coowener.item(j).Value = "" Then
                If Pancoowener.item(j).Value <> "" Or Aacoowener.item(j).Value <> "" Or percoowener.item(j).Value <> "" Then
                     msgbox_hprptfrm ("* ""Please enter the name of the other co-owner""")
                     
                    ValidateCoownerRules_HP = False
                Exit Function
                End If
            End If
            
            'Ankita_28/01/2026=======
            
'            If coowener.item(j).Value <> "" Then
'            If Pancoowener.item(j).Value = "" And Aacoowener.item(j).Value = "" Then
'                msgbox_hprptfrm ("* ""Any one of the field PAN/aadhaar is mandatory"" at Sr. No " & j & "")
'                ValidateCoownerRules_HP = False
'                Exit Function
'            End If
'            End If
 
        Next j
         End If
        Next
End Function

Function ValidatesharesRules_HP() As Boolean
     Dim i As Long
    ValidatesharesRules_HP = True
    setTblinfo_hprptfrm
    For i = 1 To UBound(CoOwnedYN_HP)
        If isdropdownblank(CoOwnedYN_HP(i)) Then
            CoOwnedYN_HP(i) = ""
        End If

        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
          Dim CoOwnerShare As Variant
          CoOwnerShare = CoOwnedShare_HP(i)
          
        End If
  
        If (UCase(CoOwnedYN_HP(i)) = "YES") Then
         Dim counter, name, name1 As Variant
         counter = 0
         For Each name In Sheet19.Range("HP.Co.Share" & i).Cells
          If name = "" Then counter = counter + 1
                
                    If counter = Sheet19.Range("HP.Co.Share" & i).Cells.count Then
                     msgbox_hprptfrm ("* ""Please enter Percentage Share of other Co-Owner in schedule HP """)
                        ValidatesharesRules_HP = False
                        Exit Function
                    End If
         Next
         
    Dim j As Long
    Dim coowener As Range
    Dim Pancoowener As Range
    Dim Aacoowener As Range
    Dim percoowener As Range
         For j = 1 To Sheet19.Range("HP.Co.Name" & i).count

         Set coowener = Sheet19.Range("HP.Co.Name" & i).Cells
         Set Pancoowener = Sheet19.Range("HP.Co.Pan" & i).Cells
         Set Aacoowener = Sheet19.Range("HP.Co.Aadhaar" & i).Cells
         Set percoowener = Sheet19.Range("HP.Co.Share" & i).Cells
         
         If percoowener.item(j).Value = "" Then
         If coowener.item(j).Value <> "" Or Pancoowener.item(j).Value <> "" Or Aacoowener.item(j).Value <> "" Then
         msgbox_hprptfrm ("* ""Please enter Percentage Share of other Co-Owner in schedule HP at S.No """ & j)
         ValidatesharesRules_HP = False
         Exit Function
         End If
         End If
         Next j
                 
        End If
    Next
End Function

Sub setTblinfo_hpco()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim scode As Variant
    
    ccount = 0
    
    mIntCells = Sheet19.Range("hp.co.name1").count
    Set rangecells = Sheet19.Range("hp.co.name1").Cells
    
    For mIntCtr = 1 To mIntCells
        scode = rangecells.item(mIntCtr).Value
        If isdropdownblank(scode) Then
            scode = ""
        End If
        
        If Not scode = "" Then
            ccount = ccount + 1
        End If
    Next
    end_hpco = ccount
    rngname_hpco = "HP.Co.Srno;HP.Co.Name;HP.Co.Pan;HP.Co.Share;"
End Sub
 
   
Sub Cmd_AddCo_Owners_Click()
AddPropertyCoOWners
End Sub
Sub AddPropertyCoOWners(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long
    Application.EnableEvents = False
    Sheet19.Activate
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet19.Range("PropertySectionCOunt").Value
    
        EfilingCommon.DefinedgridNameRange = "HP.Co.Srno||HP.Co.Name||HP.Co.Pan||HP.Co.Aadhaar||HP.Co.Share||Co_OwnnerBlock"
        ActiveCellRange = EfilingCommon.searchLastRow("HP.Co.Srno" & te)
        
        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
    
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Application.EnableEvents = False
    For i = 1 To Sheet19.Range("PropertySectionCOunt").Value
        newfrmsize = Sheet19.Range("NumRowsPropertyBlock").Value
        newfrmsize = newfrmsize + numberofrows
        Sheet19.Range("NumRowsPropertyBlock").Value = newfrmsize
    Next

endline:
'----------------Lock Password-------------------START---
   Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub

Sub Cmd_AddTenant_Click()
AddPropertyTenant
End Sub

Sub AddPropertyTenant(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long


    Application.EnableEvents = False
    Sheet19.Activate

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet19.Range("PropertySectionCOunt").Value


        EfilingCommon.DefinedgridNameRange = "HP.NameofTenant||HP.PANofTenant||HP.AadharofTenant||HP.TANofTenant"
        ActiveCellRange = EfilingCommon.searchLastRow("HP.NameofTenant" & te)

        If te > 1 Then
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(numberofrows, True, te)
        Else
            numberofrows = EfilingCommon.insertRowUnderSectionWithFormula(iRows, True, te)
        End If
    Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Application.EnableEvents = False
    For i = 1 To Sheet19.Range("PropertySectionCOunt").Value
        newfrmsize = Sheet19.Range("NumRowsPropertyBlock").Value
        newfrmsize = newfrmsize + numberofrows
        Sheet19.Range("NumRowsPropertyBlock").Value = newfrmsize
    Next
endline:

'----------------Lock Password-------------------START---
   Sheet19.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
  Application.EnableEvents = True
End Sub
'--end

'ANkita_23/01/2026--


Public Function ChkMaxDate_24b(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate_24b = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 AY 2024-25 Change
        'Ayush_25-26
'        If Year > 2024 Then
        If Year > CInt(Sheet5.Range("DOB_Year").Value) Then
            ChkMaxDate_24b = False
            Exit Function
        Else
'            If Year = 2024 Then
                If Year = CInt(Sheet5.Range("DOB_Year").Value) Then
                If month > 4 Then
                    ChkMaxDate_24b = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                           ChkMaxDate_24b = False
                            Exit Function
                        Else
                            If dat = 1 Then
                               ChkMaxDate_24b = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function


Function Validate24B_Table() As Boolean
    Dim i As Long
    Dim j As Long
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    Dim rangecells3 As Range
    Dim rangecells4 As Range
    Dim rangecells5 As Range
    Dim rangecells6 As Range
    Dim rangecells7 As Range
    
    Dim state_1 As Variant
    Dim sharepercent As Variant
    Dim cosharepercent As Double
    Dim end_hpco_24d As Long
    Validate24B_Table = True
    end_hprptfrm = Sheet19.Range("PropertySectionCOunt").Value
    
    For i = 1 To end_hprptfrm
        
        Set rangecells1 = Sheet19.Range("LoanfrmBankOrInstitute.24b" & i).Cells
        Set rangecells2 = Sheet19.Range("bankName.24b" & i).Cells
        Set rangecells3 = Sheet19.Range("loanAccNum.24b" & i).Cells
        Set rangecells4 = Sheet19.Range("loanDate.24b" & i).Cells
        Set rangecells5 = Sheet19.Range("loanAmt.24b" & i).Cells
        Set rangecells6 = Sheet19.Range("loanOutstanding.24b" & i).Cells
        Set rangecells7 = Sheet19.Range("Intrst.24b" & i).Cells
        
        setTblinfo_24bBankorInst (i)
        setTblinfo_24bankname (i)
        setTblinfo_24bLoan (i)
        setTblinfo_24bdate (i)
        setTblinfo_24bAmount (i)
        setTblinfo_24bloanout (i)
        setTblinfo_24bInt (i)
        
        end_hpco_24d = Application.WorksheetFunction.Max(0, end_hpco_24a, end_hpco_24b, end_hpco_24e, end_hpco_24f, end_hpco_24k, end_hpco_24s, end_hpco_24r)
        
        ReDim BankorInst_24b(end_hprptfrm, end_hpco_24d)
        ReDim BankName_24B(end_hprptfrm, end_hpco_24d)
        ReDim AccntNum_24b(end_hprptfrm, end_hpco_24d)
        ReDim LoanDate_24B(end_hprptfrm, end_hpco_24d)
        ReDim LoanAmt_24b(end_hprptfrm, end_hpco_24d)
        ReDim LoanOutStanding_24B(end_hprptfrm, end_hpco_24d)
        ReDim Intrst_24B(end_hprptfrm, end_hpco_24d)
        
         For j = 1 To end_hpco_24d
                BankorInst_24b(i, j) = rangecells1.item(j).Value
                BankName_24B(i, j) = rangecells2.item(j).Value
                AccntNum_24b(i, j) = rangecells3.item(j).Value
                LoanDate_24B(i, j) = rangecells4.item(j).text
                LoanAmt_24b(i, j) = rangecells5.item(j).text
                LoanOutStanding_24B(i, j) = rangecells6.item(j).text
                Intrst_24B(i, j) = rangecells7.item(j).text
                
                If (BankorInst_24b(i, j) = "(Select)") Or (BankorInst_24b(i, j) = "") Then
                    msgbox_hprptfrm ("*""Please select dropdown from ""Loan taken from"" in 24(b) schedule."" at Sr. No " & j & " in Table " & i & "") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                 If Not chkCompulsory(BankorInst_24b(i, j) <> "Select") Then
                    msgbox_hprptfrm ("*""Please select dropdown from ""Loan taken from"" in 24(b) schedule."" at Sr. No " & j & " in table " & i & "") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                If Not chkCompulsory(BankName_24B(i, j)) Then
                    msgbox_hprptfrm ("*""Please provide Name of the Bank/ Institution/Person from which the loan is taken in schedule 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                If Len(BankName_24B(i, j)) > 125 Then
                    msgbox_hprptfrm ("* Name of the Bank/ Institution/Person at Sr. No " & j & " in Sheet 24(b) should be less than or equal to 125 characters.")
                    Validate24B_Table = False
                    Exit Function
                End If
                
                If Not checkfieldSuperSpecialcharacter(BankName_24B(i, j)) Then
                    msgbox_hprptfrm ("* Name of the Bank/ Institution/Person in schedule 24b at Sl.no. " & j & " should not Contain <, >, characters.")
                    Validate24B_Table = False
                    Exit Function
                End If
                
                '--Acc number
                If Not chkCompulsory(AccntNum_24b(i, j)) Then
                    msgbox_hprptfrm ("* ""Please provide Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr. No " & j & " in table " & i & "")  'SIT-93563 updated by shrutika
                    Validate24B_Table = False
                    Exit Function
                End If
                
                If Len(AccntNum_24b(i, j)) > 20 Then
                    msgbox_hprptfrm ("* Loan Account number  at Sr. No " & j & " in schedule 24b less than 20 characters.") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                End If
                
                If UCase(BankorInst_24b(i, j)) = UCase("Bank ") Then
                    If Not ValidateBankAccountNumber_24b(AccntNum_24b(i, j)) Then
'                        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank or any reference number if loan is taken from other than Bank"" at Sr.No " & j & "") & Chr(13)
                    Validate24B_Table = False
                    Exit Function
                    End If
                End If
                
                If UCase(BankorInst_24b(i, j)) = UCase("Other than Bank") Then
                    If Not checkfieldspecialcharacter1(AccntNum_24b(i, j)) Then
                        msgbox_hprptfrm ("* ""Please provide correct Loan Account number of the Bank / Institution from which loan is taken in schedule 24(b)"" at Sr.No " & j & " in table " & i & "") & Chr(13)
                        Validate24B_Table = False
                        Exit Function
                    End If
                End If
             If Not chkCompulsory(LoanDate_24B(i, j)) Then
                 msgbox_hprptfrm ("* ""Please provide Date of sanction of Loan in schedule 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
             If Len(LoanDate_24B(i, j)) > 10 Then
              msgbox_hprptfrm ("* Date of sanction of Loan  at Sr. No " & j & " in schedule 24b less than 10 characters.") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
             If Not CheckDateddmmyyyy(LoanDate_24B(i, j)) Then
                msgbox_hprptfrm ("* ""Please enter date in valid format"" at Sr. No " & j & ".") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            'Ankita_24/01/2026=====
            
                Dim cutoff12 As Date
                cutoff12 = CDate(Sheet5.Range("DOB_1").Value)
            If Not ChkMaxDate_24b(Trim(LoanDate_24B(i, j)), Sheet5.Range("DOB_1").Value) Then
'                msgbox_hprptfrm ("* Date can not be after 31/03/2025 at Sr. No " & j & ".") & Chr(13)
                 msgbox_hprptfrm ("* Date can not be on or after " & Dformat1(cutoff12, "dd/mm/yyyy") & " at Sr. No " & i & ".") & Chr(13)
                Validate24B_Table = False
              Exit Function
            End If
            
            If Not chkCompulsory(LoanAmt_24b(i, j)) Then
                    msgbox_hprptfrm ("* ""Please provide Total amount of loan in schedule 24(b)"" at Sr. No " & j & " in table " & i & "") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            If Not IsNumeric(LoanAmt_24b(i, j)) Then
                msgbox_hprptfrm ("* Loan amount at Sr. No  " & j & "  in schedule 24b should be Numeric value") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            If LoanAmt_24b(i, j) > 99999999999999# Then
                msgbox_hprptfrm ("* Loan amount at Sr. No  " & j & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            If LoanAmt_24b(i, j) < 0 Or LoanAmt_24b(i, j) = 0 Then
               msgbox_hprptfrm ("* Total amount of loan should be more than 0 in schedule 24(b) at Sr. No  " & j & " in table " & i & "") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            If Not chkCompulsory(LoanOutStanding_24B(i, j)) Then
                msgbox_hprptfrm ("* ""Loan outstanding as on last date of financial year is mandatory in Table Section 24(b). Please enter 0 in case the entire loan is repaid during the year"" at Sr. No " & j & " in table " & i & "") & Chr(13)
                
                Validate24B_Table = False
                Exit Function
            End If
            If Not IsNumeric(LoanOutStanding_24B(i, j)) Then
                msgbox_hprptfrm ("* Loan outstanding at Sr. No  " & j & "  in schedule 24b should be Numeric value") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            If LoanOutStanding_24B(i, j) > 99999999999999# Then
                msgbox_hprptfrm ("* Loan outstanding at Sr. No  " & j & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            
            If LoanOutStanding_24B(i, j) < 0 Then
                 msgbox_hprptfrm ("* Loan outstanding as on last date of finacial year can't be less than 0 in schedule 24(b) at Sr. No  " & j & ". You may please enter as 0 if it become negative as result of excess payment.") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            If Not chkCompulsory(Intrst_24B(i, j)) Then
             msgbox_hprptfrm ("* ""Please provide Interest on Borrowed capital u/s 24(b)""") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            If Not IsNumeric(Intrst_24B(i, j)) Then
               msgbox_hprptfrm ("* Interest at Sr. No  " & j & "  in schedule 24b should be Numeric value") & Chr(13)
                
                Validate24B_Table = False
                Exit Function
            End If
            
            If Intrst_24B(i, j) > 99999999999999# Then
                msgbox_hprptfrm ("* Interest at Sr. No  " & j & "  in schedule 24b cannot exceed 14 digits") & Chr(13)
                Validate24B_Table = False
                Exit Function
            End If
            'Ankita_23/02/2026=============
            If Intrst_24B(i, j) < 0 Or Intrst_24B(i, j) = 0 Then
'               msgbox_hprptfrm ("* ""Please provide Interest on Borrowed capital u/s 24(b)""") & Chr(13)
                msgbox_hprptfrm ("* ""Interest u/s 24(b) should be more than 0 in schedule 24(b) at Sr. No  " & j & """")
                Validate24B_Table = False
                Exit Function
            End If
        Next
         If (Len(Sheet19.Range("TotAmt.24b" & i).Value) > 14) Then
            msgbox_hprptfrm ("*  Total of interest on borrowed capital u/s 24(b) cannot exceed 14 Digits.") & Chr(13)
            Validate24B_Table = False
            Exit Function
         End If
            
        
        
    Next
endlin1:
End Function
Function DefaultifLetOut_HP() As String
    DefaultifLetOut_HP = "S"
End Function

Function DefaultAnnualLetableValue_HP() As String
    DefaultAnnualLetableValue_HP = "0"
End Function

Function DefaultRentNotRealized_HP() As String
    DefaultRentNotRealized_HP = "0"
End Function

Function DefaultLocalTaxes_HP() As String
    DefaultLocalTaxes_HP = "0"
End Function

Function DefaultTotalUnrealizedAndTax_HP() As String
    DefaultTotalUnrealizedAndTax_HP = "0"
End Function

Function DefaultBalanceALV_HP() As String
    DefaultBalanceALV_HP = "0"
End Function

Function DefaultThirtyPercentOfBalance_HP() As String
    DefaultThirtyPercentOfBalance_HP = "0"
End Function

Function DefaultIntOnBorwCap_HP() As String
    DefaultIntOnBorwCap_HP = "0"
End Function

Function DefaultTotalDeduct_HP() As String
    DefaultTotalDeduct_HP = "0"
End Function

Function DefaultIncomeOfHP_HP() As String
    DefaultIncomeOfHP_HP = "0"
End Function


Function DefaultRentArearsSec25BAfter30pcDeduct_HP() As String
    DefaultRentArearsSec25BAfter30pcDeduct_HP = "0"
End Function

Function DefaultTotalIncomeChargeableUnHP_HP() As String
    DefaultTotalIncomeChargeableUnHP_HP = "0"
End Function

    

'Ankita_18/03/2026=======
'Function ValidateMandatorytable() As Boolean
'ValidateMandatorytable = True
'Dim i As Long
'Dim flag As Boolean
'
'Dim rangecells As Range
'
'
'
'flag = True
'For i = 1 To Sheet19.Range("HP.Co.Name2").Rows.count
'If Sheet19.Range("HP.AddrDetail2").Value = "" Or Sheet19.Range("HP.CityOrTownOrDistrict2").Value = "" Or (Sheet19.Range("HP.StateCode2").Value = "(Select)" Or Sheet19.Range("HP.StateCode2").Value = "") Or (Sheet19.Range("HP.CountryCode2").Value = "") Or (Sheet19.Range("HP.ifLetOut2").Value = "(Select)" Or Sheet19.Range("HP.ifLetOut2").Value = "") Then
'
'        If Sheet19.Range("HP.Co.Name2").item(i).Value <> "" Then
'            flag = False
'              msgbox_hprptfrm ("* Enter all mandatory fields of Block 1B in Schedule HP.")
'
'        End If
'
'
'        If flag = False Then
'            ValidateMandatorytable = False
'            Exit Function
'        End If
'
'    End If
'Next i
'End Function
'
'Function ValidateMandatorytable1() As Boolean
'ValidateMandatorytable1 = True
'Dim i As Long
'Dim flag As Boolean
'
'Dim rangecells As Range
'flag = True
'For i = 1 To Sheet19.Range("LoanfrmBankOrInstitute.24b2").Rows.count
'If Sheet19.Range("HP.AddrDetail2").Value = "" Or Sheet19.Range("HP.CityOrTownOrDistrict2").Value = "" Or (Sheet19.Range("HP.StateCode2").Value = "(Select)" Or Sheet19.Range("HP.StateCode2").Value = "") Or Sheet19.Range("HP.CountryCode2").Value = "" Or (Sheet19.Range("HP.ifLetOut2").Value = "(Select)" Or Sheet19.Range("HP.ifLetOut2").Value = "") Then
'
'        If Sheet19.Range("LoanfrmBankOrInstitute.24b2").item(i).Value <> "" And Sheet19.Range("bankName.24b2").item(i).Value <> "" And Sheet19.Range("loanAccNum.24b2").item(i).Value <> "" And Sheet19.Range("loanDate.24b2").item(i).Value <> "" And Sheet19.Range("loanAmt.24b2").item(i).Value <> "" And Sheet19.Range("loanOutstanding.24b2").item(i).Value <> "" And Sheet19.Range("Intrst.24b2").item(i).Value <> "" Then
'            flag = False
'              msgbox_hprptfrm ("* Enter all mandatory fields of Block 1B in Schedule HP.")
'
'        End If
'
'
'        If flag = False Then
'            ValidateMandatorytable1 = False
'            Exit Function
'        End If
'
'    End If
'Next i
'End Function
'
'
'
'
'
Function ValidateMandatorytable1() As Boolean
ValidateMandatorytable1 = True
Dim i As Long
Dim flag As Boolean

Dim rangecells As Range



flag = True
For i = 1 To Sheet19.Range("LoanfrmBankOrInstitute.24b2").Rows.count

    

If Sheet19.Range("HP.AddrDetail2").Value = "" Or Sheet19.Range("HP.CityOrTownOrDistrict2").Value = "" Or (Sheet19.Range("HP.StateCode2").Value = "(Select)" Or Sheet19.Range("HP.StateCode2").Value = "") Or Sheet19.Range("HP.CountryCode2").Value = "" Or (Sheet19.Range("HP.ifLetOut2").Value = "(Select)" Or Sheet19.Range("HP.ifLetOut2").Value = "") Then
        
        If Sheet19.Range("LoanfrmBankOrInstitute.24b2").item(i).Value <> "" And Sheet19.Range("bankName.24b2").item(i).Value <> "" And Sheet19.Range("loanAccNum.24b2").item(i).Value <> "" And Sheet19.Range("loanDate.24b2").item(i).Value <> "" And Sheet19.Range("loanAmt.24b2").item(i).Value <> "" And Sheet19.Range("loanOutstanding.24b2").item(i).Value <> "" And Sheet19.Range("Intrst.24b2").item(i).Value <> "" Then
            flag = False
              msgbox_hprptfrm ("* Enter all mandatory fields of Block 1B in Schedule HP.")

        End If
        

        If flag = False Then
            ValidateMandatorytable1 = False
            Exit Function
        End If
        
    End If
Next i
End Function
