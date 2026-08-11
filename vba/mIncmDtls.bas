Attribute VB_Name = "mIncmDtls"
Option Explicit
Public Others_NOI2, Others_NOI3, end_OthersNOI1, Others_Amt2, end_OthersAmt As Variant
Public Others_NOI4, Others_NOI5, end_OthersNOI2, Others_Amt3, end_OthersAmt1 As Variant
Public MobileCountryCode As Variant
'Ankita_12/01/2026=====
Public MobileCountryCode1 As Variant
Public MobileCountryCode2 As Variant
Public CountrycodeMobileNo2 As Variant
Public CountrycodeMobileNo3 As Variant
Public PhoneNo2 As Variant
Public PhoneNo3 As Variant

Public end_pran As Variant
Public pran_80CCC As Variant

Public DOF As Variant
Public DOF1 As Variant
Public SELECT80D As Variant
Public SELECT80DB As Variant
Public SELECT80DC As Variant
Public SELECT80DD As Variant
Public SELECT80DDS As Variant
Public SELECT80U As Variant

'Ankita_24/12/2025===========
Public end_80CCC As Variant
Public end_80CCCType As Variant
Public end_80CCCName As Variant
Public end_80CCCAmount As Variant
Public Type_80CCC As Variant
Public Name_80CCC As Variant
Public Amount_80CCC As Variant

Public end_80CCD_1 As Variant
Public end_80CCD_1_Type As Variant
Public end_80CCD_1_Name As Variant
Public end_80CCD_1_Amount As Variant
Public Type_1_80CCD As Variant
Public Name_1_80CCD As Variant
Public Amount_1_80CCD As Variant

Public end_80CCD_1b As Variant
Public end_80CCD_1b_Type As Variant
Public end_80CCD_1b_Name As Variant
Public end_80CCD_1b_Amount As Variant
Public Type_1b_80CCD As Variant
Public Name_1b_80CCD As Variant
Public Amount_1b_80CCD As Variant

Public NameRepAssessee As Variant
Public EmailRepAssessee As Variant
Public ContactRepAssessee As Variant
Public Representativeassesseeflg As Variant
'============================

Public ProvisoFlag As Variant
Public DepositAmountFlag As Variant
Public DepositAmount As Variant
Public AggrigateAmountFlag As Variant
Public AggrigateAmount As Variant
Public AggrigateAmountFlag1 As Variant
Public AggrigateAmount1 As Variant

Public Investment As Variant

Public stDcode As Variant
Public phoneNo As Variant
Public LastName As Variant
Public PAN As String

'Added by Shrutika(18/04/2025)NewDev
'Public PRAN As String
'Public IncDSection80CCD As String
'Public IncDSection80CCD1B As String
'Public IncDSection80CCD1 As String

'Malli------------------
Public PRAN As Variant
Public IncDSection80CCD As Variant
Public IncDSection80CCD1B As Variant
Public IncDSection80CCD1 As Variant
'---------------------------

Public IncDSection80GG As Variant  'Added by Ayush 23/05/2025
Public AckNum As Variant

'------------------------------
Public Flat As Variant
Public Flat1 As Variant
Public Area As Variant
Public Area1 As Variant
Public roadOrStreet As Variant
Public City As Variant
Public City1 As Variant
Public State As Variant
Public state1 As Variant
Public PinCode As Variant
Public PinCode1 As Variant
Public zipCode As Variant
Public zipCode1 As Variant

Public HASZIP As Variant
Public mobileNo As Variant

'Ankita_12/01/2026===========
Public mobileNo1 As Variant
Public mobileNo2 As Variant

Public Email As Variant
Public Email_1 As Variant
Public Email_2 As Variant

Public dob As String
Public ReturnFileSec As Variant
Public ReturnFurSec As Variant
Public ReturnType As Variant
'Public ResidentialStatus As String
Public status As String
Public residenceName As Variant
Public empcat As Variant
Public Gender As Variant
Public Portugese As Variant

Public firstName As Variant
Public middleName As Variant
Public Country As Variant
Public mobileNoSec As Variant

Public DesigOfficerWardorCircle As Variant
Public SpousePAN As Variant
Public NoticeDate As String
Public filingDate As String
Public VerDate As String
Public BankAccountNumber As Variant
Public agestatus  As String

Public Country1 As Variant
Public sCountry As Variant
Public sCountry1 As Variant
Public sNoticeNo As String
'Public sNoticeNoNew As String
Public sNoticeDate As String
Public sReceiptNo As String
Public sReceiptNo1 As String
Public sNotiDate As String
Public sResponse As String
Public ReturnTypeMsg As Boolean


Public IncomeFromSal_IncD As Variant
Public IncomeFromHP_IncD As Variant
Public IncomeFromOS_IncD As Variant
Public GrossTotIncome_IncD As Variant
Public Section80C_Usr_IncD As Variant
Public Section80CCC_Usr_IncD As Variant
Public Section80CCD_SE_Usr_IncD As Variant
Public Section80CCD_Usr_IncD As Variant
Public Section80CCD1B_SE_User_Incd As Variant
Public Section80D_Usr_IncD As Variant
Public Section80DD_Usr_IncD As Variant
Public Section80DDB_Usr_IncD As Variant
Public Section80E_Usr_IncD As Variant
Public Section80EE_Usr_IncD As Variant
Public Section80G_Usr_IncD As Variant
Public Section80GG_Usr_IncD As Variant
Public Section80GGA_Usr_IncD As Variant
Public Section80GGC_Usr_IncD As Variant
Public Section80U_Usr_IncD As Variant
Public Section80RRB_Usr_IncD As Variant
Public Section80QQB_Usr_IncD As Variant
Public Section80CCG_Usr_IncD As Variant
Public Section80TTA_Usr_IncD As Variant
Public Section80TTB_Usr_IncD As Variant
Public TotalChapVIADeductions_Usr_IncD As Variant

Public TotalIncome_IncD As Variant

Public Section80C_IncD As Variant
Public TypeOfHP_1 As Variant

Public TypeOfHP As Variant
Public IncomeFromHP As Variant
Public IncomeFromOS As Variant
Public GrossTotIncome As Variant
Public GrossTotIncomeIncLTCG112A As Variant
Public TotalChapVIADeductions As Variant
Public employercategory2 As Variant




Public salaries As Double
Public Section80C As Double
Public Section80CCC As Double
Public section80CCD As Double
Public section80CCD_SE As Double
Public Section80CCD1B_SE As Variant
Public section80CCG As Double
Public Section80D As Double
Public Section80DD As Double
Public Section80DDB As Double
Public Section80E As Double
Public Section80EE As Double
Public Section80G As Double
Public Section80GG As Double
Public section80RRB As Double
Public Section80GGA As Double
Public Section80GGC As Double
Public section80QQB As Double
Public Section80TTA As Double
Public Section80TTB As Double
Public Section80U As Double

Public section80C_Calc As Double
Public section80CCC_Calc As Double
Public section80CCD_Calc As Double
Public section80CCD_Calc_SE As Double
Public section80CCG_Calc As Double
Public section80D_Calc As Double
Public section80DD_Calc As Double
Public section80DDB_Calc As Double
Public section80E_Calc As Double
Public section80EE_Calc As Double
Public section80G_Calc As Double
Public section80GG_Calc As Double
Public section80RRB_Calc As Double
Public section80GGA_Calc As Double
Public section80GGC_Calc As Double
Public section80QQB_Calc As Double
Public section80TTA_Calc As Double
Public section80TTB_Calc As Double

Public section80U_Calc As Double

Dim TotalEligibleDonationsUs80G As Double

Public grossTotInc As Double
Dim gIncome As Double
Public totInc As Double
Public dedVIA As Double

Public TotalTDSSal As Variant
Public IncChrgSal As Variant
Public totalTCSSal As Variant
Public incChrgSalTCS As Variant
Public claimOutOfTotTDSOnAmtPaid As Variant
Public totTDSOnAmtPaid As Variant
Public amtClaimedBySpouse As Variant
Public taxDepDate As Variant
Public taxPAmt As Variant
Public BalTaxPayable As Double
'Ankita_04/02/2025
Public IncD_Sale_LTCG  As Variant
Public IncD_Cost_LTCG As Variant
Public IncD_CG_LTCG As Variant
Public msgError_LTCG As String

Dim taxPayable As Variant
Dim Rebate87A As Double
'Dim resStatus As String
Dim TaxPayableOnRebate As Double
Dim surchargeOnAboveCrore As Double
Dim eduCess As Double

Dim balTaxPay As Double
Dim AdvanceTax As Double
Dim TDS As Double
Dim TCS As Double
Dim SelfAssessmentTax As Double
Dim selfAssessmentTax234A As Double

Dim advanceTaxToDisplay As Double
Dim TDSToDisplay As Double
Dim TCSToDisplay As Double
Dim SATtoDisplay As Double

Public age As Long
Public bacage As Long
Dim currentdate As Variant

Public intrst234A As Double
Dim intrst234B As Double
Public intrst234C As Double
Public intrst234F As Double
Dim slab0 As Double
Dim slab1 As Double
Dim slab2 As Double
Dim slab3 As Double
Dim slab4 As Double

Public taxStatus As String   'variable to store tax status


'PAN Validation
Function ChkPAN() As Boolean
On Error Resume Next
    ChkPAN = True
    PAN = Sheet1.Range("sheet1.PAN")
    
    If Trim(PAN) = "" Or IsEmpty(PAN) Then
        ChkPAN = False
        Exit Function
    End If

    If Len(Trim(PAN)) > 10 Then
        fmsgbox ("* PAN in Sheet Income Details cannot exceed 10 characters ")
        ChkPAN = False
        Exit Function
        
    End If
 
End Function
'Added by Shrutika(18/04/2025)NewDev

'Commented by Ankita as per AY26-27 DE sheet v0.1

'Function ChkPRAN() As Boolean
'On Error Resume Next
'    ChkPRAN = True
''    PRAN = Sheet1.Range("sheet1.PRAN")
''    IncDSection80CCD = Sheet1.Range("IncD.section80CCD_SE")
''    IncDSection80CCD1B = Sheet1.Range("IncD.Section80CCD1B_SE")
''    IncDSection80CCD1 = Sheet1.Range("IncD.section80CCD")
'
'    PRAN = Sheet1.Range("sheet1.PRAN").Value
'    IncDSection80CCD = Sheet1.Range("IncD.section80CCD_SE").Value
'    IncDSection80CCD1B = Sheet1.Range("IncD.Section80CCD1B_SE").Value
'    'IncDSection80CCD1 = Sheet1.Range("IncD.section80CCD").Value
'
'    If PRAN = "" Then
'    If IncDSection80CCD > 0 Or IncDSection80CCD1B > 0 Or IncDSection80CCD1 > 0 Then
'    ChkPRAN = False
'    Exit Function
'    End If
'    End If
    
    
    
'    If Trim(PRAN) = "" Or IsEmpty(PRAN) Then
'    If Trim(IncDSection80CCD) > 0 Or Trim(IncDSection80CCD1B) > 0 Or Trim(IncDSection80CCD1) > 0 Then
'        ChkPRAN = False
'        Exit Function
'    End If
'    End If

'Ankita_29/12/2025=====
'    If Len(Trim(PRAN)) > 125 Then
'        fmsgbox ("* PRAN of the taxpayer in Sheet Income Details cannot exceed 125 characters ") & Chr(13)
'        ChkPRAN = False
'        Exit Function
'
'    End If
    
'     If Not checkfieldspecialcharacter(PRAN) Then
'     fmsgbox "* Invalid PPRAN of the taxpayer in Sheet Income Details. PRAN of the taxpayer format should be Alphanumeric in Sheet Income Details." & Chr(13)
'    ChkPRAN = True
'    End If
    
   
 
'End Function

'Added by Shrutika(18/04/2025)NewDev
Function ChkAckNum() As Boolean
On Error Resume Next
    ChkAckNum = True
    AckNum = Sheet1.Range("Sheet1.AckNum")
    
   Dim i As Long
    IncDSection80GG = Sheet1.Range("IncD.Section80GG")
    
    If Trim(AckNum) = "" Or IsEmpty(AckNum) Then
    If IncDSection80GG > 0 Then  'Ayush _23/05/2025
        ChkAckNum = False
        Exit Function
    End If
    End If
    
    If Len(Trim(AckNum)) > 15 Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + ("* Acknowledgement number of Form 10BA in Sheet Income Details cannot exceed 15 characters ")
        ChkAckNum = False
        Exit Function
        
    End If
    
          
          If Trim(AckNum) <> "" Then
        For i = 1 To Len(AckNum)
            If Not IsNumeric(Mid(AckNum, i, 1)) Then
                fmsgbox "* Acknowledgement number of Form 10BA should be numeric, Non negative, no decimal, upto 99,999,999,999,9999" & Chr(13)
                ChkAckNum = True
            Exit Function
            End If
        Next
    End If
      
  
 
End Function
        
'Name Validation
Function ChkName() As Boolean
On Error Resume Next
    
    ChkName = True
    LastName = Sheet1.Range("sheet1.SurNameOrOrgName")
    
    If Trim(LastName) = "" Or IsEmpty(LastName) Then
        ChkName = False
        Exit Function
    End If

    If Len(Trim(LastName)) > 125 Then
        fmsgbox ("* Name in Sheet Income Details  cannot exceed 125 characters ")
        ChkName = False
        Exit Function
    End If
    
    If Not checkfieldSuperSpecialcharactername(LastName) Then
   EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Last name should not Contain <, > " & Chr(34) & "& characters in Income Details" & Chr(13)
    End If
    
    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.FirstName").Value) Then
   EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* First name should not Contain <, > " & Chr(34) & "& characters in Income Details" & Chr(13)
    End If
    
    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.MiddleName").Value) Then
   EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Middle name name should not Contain <, > " & Chr(34) & "& characters in Income Details" & Chr(13)
    End If
    
End Function

'Flat/Door/Building Validation
Function ChkFlat() As Boolean
On Error Resume Next
    
    ChkFlat = True
    Flat = Sheet1.Range("sheet1.ResidenceNo")
    
    If Trim(Flat) = "" Or IsEmpty(Flat) Then
        ChkFlat = False
        Exit Function
    End If

    If Len(Trim(Flat)) > 50 Then
        fmsgbox ("* Flat/Door/Block No. under primary address in Sheet Income Details cannot exceed 50 characters ")
        ChkFlat = False
        Exit Function
    End If
End Function

'Ankita_14/01/2026_V0.2==========
Function ChkFlat1() As Boolean
On Error Resume Next
    ChkFlat1 = True
    Flat1 = Sheet1.Range("sheet1.ResidenceNo1")
    If Trim(Flat1) = "" Or IsEmpty(Flat1) Then
        ChkFlat1 = False
        Exit Function
    End If
    If Len(Trim(Flat1)) > 50 Then
        fmsgbox ("* Flat/Door/Block No. in Sheet Income Details cannot exceed 50 characters ")
        ChkFlat1 = False
        Exit Function
    End If
End Function

'================================
'Area / Locality Validation
Function ChkArea() As Boolean
On Error Resume Next
    ChkArea = True
    Area = Sheet1.Range("sheet1.LocalityOrArea")
    
    If Trim(Area) = "" Or IsEmpty(Area) Then
        ChkArea = False
        Exit Function
    End If

    If Len(Trim(Area)) > 50 Then
        fmsgbox ("* Area / Locality under primary address in Sheet Income Details cannot exceed 50 characters ")
        ChkArea = False
        Exit Function
    End If
End Function

'Ankita_14/01/2026_V0.2==========

Function ChkArea1() As Boolean
On Error Resume Next
    ChkArea1 = True
    Area1 = Sheet1.Range("sheet1.LocalityOrArea1")
    If Trim(Area1) = "" Or IsEmpty(Area1) Then
        ChkArea1 = False
        Exit Function
    End If
    If Len(Trim(Area1)) > 50 Then
        fmsgbox ("* Area / Locality under secondary address in Sheet Income Details cannot exceed 50 characters ")
        ChkArea1 = False
        Exit Function
    End If
End Function

'================================
'Town/City/District Validation
Function ChkCity() As Boolean
On Error Resume Next
    ChkCity = True
    City = Sheet1.Range("sheet1.CityOrTownOrDistrict")
    
    If Trim(City) = "" Or IsEmpty(City) Then
        ChkCity = False
        Exit Function
    End If

    If Len(Trim(City)) > 50 Then
        fmsgbox ("* Town/City/District under primary address in Sheet Income Details cannot exceed 50 characters ")
        ChkCity = False
        Exit Function
    End If
End Function

'Ankita_14/01/2026_V0.2==========
Function ChkCity1() As Boolean
On Error Resume Next
    ChkCity1 = True
    City1 = Sheet1.Range("sheet1.CityOrTownOrDistrict1")
    
    If Trim(City1) = "" Or IsEmpty(City1) Then
        ChkCity1 = False
        Exit Function
    End If

    If Len(Trim(City1)) > 50 Then
        fmsgbox ("* Town/City/District under secondary address in Sheet Income Details cannot exceed 50 characters ")
        ChkCity1 = False
        Exit Function
    End If
End Function


'================================
'State Validation
Function ChkState() As Boolean
On Error Resume Next
    ChkState = True
    State = Sheet1.Range("sheet1.StateCode1")
  
    If Trim(State) = "" Or Trim(State) = "(Select)" Then
        ChkState = False
        Exit Function
    End If
End Function

'Ankita_14/01/2026_V0.2==========
Function ChkState1() As Boolean
On Error Resume Next
    ChkState1 = True
    state1 = Sheet1.Range("sheet1.StateCode2")
    If Trim(state1) = "" Or Trim(state1) = "(Select)" Then
        ChkState1 = False
        Exit Function
    End If
End Function

'Pincode Validation
Function ChkPincode() As Boolean
On Error Resume Next
    ChkPincode = True
    Dim i As Long
    
    PinCode = Sheet1.Range("sheet1.PinCode")
    If sCountry <> "" And UCase(Mid(sCountry, 1, InStr(1, sCountry, "-") - 1)) = "91" Then
        If Trim(PinCode) = "" Or IsEmpty(PinCode) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Pin code under primary address is mandatory in Sheet Income Details" & Chr(13)
            ChkPincode = False
            Exit Function
        End If
        If Len(Trim(PinCode)) > 6 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Pin code under primary address in Sheet Income Details cannot exceed 6 characters " & Chr(13)
            ChkPincode = False
            Exit Function
        End If
        
         For i = 1 To Len(PinCode)
            If Not IsNumeric(Mid(PinCode, i, 1)) Then
                EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Pin code under primary address in Sheet Income Details  must contain only digits from 0 to 9" & Chr(13)
                ChkPincode = False
                Exit Function
            End If
        Next
    Else
        PinCode = ""
    End If
    
End Function

'Ankita_25/01/2026=========

'Pincode Validation
Function ChkPincode1() As Boolean
On Error Resume Next
    ChkPincode1 = True
    Dim i As Long
    
    PinCode1 = Sheet1.Range("sheet1.PinCode1")
    If sCountry <> "" And UCase(Mid(sCountry, 1, InStr(1, sCountry, "-") - 1)) = "91" Then
        If Trim(PinCode1) = "" Or IsEmpty(PinCode1) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Pin code under secondary address is mandatory in Sheet Income Details" & Chr(13)
            ChkPincode1 = False
            Exit Function
        End If
        If Len(Trim(PinCode1)) > 6 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Pin code under secondary address in Sheet Income Details cannot exceed 6 characters " & Chr(13)
            ChkPincode1 = False
            Exit Function
        End If
        
         For i = 1 To Len(PinCode1)
            If Not IsNumeric(Mid(PinCode, i, 1)) Then
                EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Pin code under secondary address in Sheet Income Details  must contain only digits from 0 to 9" & Chr(13)
                ChkPincode1 = False
                Exit Function
            End If
        Next
    Else
        PinCode1 = ""
    End If
    
End Function

'ZipCode Validation
Function ChkZipcode() As Boolean
On Error Resume Next
    ChkZipcode = True
    zipCode = Sheet1.Range("sheet1.ZipCode")

      If sCountry <> "" And UCase(Mid(sCountry, 1, InStr(1, sCountry, "-") - 1)) <> "91" Then
        If Trim(zipCode) = "" Or IsEmpty(zipCode) Then
            'EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "ZipCode in Sheet : Income Details  is Mandatory, if there is no ZipCode then select ""No ZIP Code""" & Chr(13)
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "ZipCode in Sheet : Zip Code under primary address is compulsory, if there is no ZipCode then select ""No ZIP Code""" & Chr(13)
            ChkZipcode = False
            Exit Function
        End If


        If Len(Trim(zipCode)) > 8 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "ZipCode in Sheet : Zip Code under primary address in Income Details cannot exceed 8 characters.Minimum 1 and up to 8 Characters." & Chr(13)
            ChkZipcode = False
            Exit Function
        End If

        If Not checkfieldspecialcharacter1(zipCode) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "ZipCode in Sheet : Zip Code under primary address in Income Details characters < > & ' " & Chr(34) & " are not allowed"" & chr(13)"
            ChkZipcode = False
            Exit Function
        End If
        
'        If Mid(ZipCode, 1, 1) = 0 Then
'            Sheet1.Range("sheet1.ZipCode").Value = "0" + ZipCode
'        End If
        
        
          
    Else
        zipCode = ""
    End If
End Function



'Ankita_25/01/2026=========

Function ChkZipcode1() As Boolean
On Error Resume Next
    ChkZipcode1 = True
    zipCode1 = Sheet1.Range("sheet1.ZipCode1")

      If sCountry1 <> "" And UCase(Mid(sCountry1, 1, InStr(1, sCountry1, "-") - 1)) <> "91" Then
        If Trim(zipCode1) = "" Or IsEmpty(zipCode1) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "ZipCode in Sheet : Zip Code under secondary address is compulsory, if there is no ZipCode then select ""No ZIP Code""" & Chr(13)
            ChkZipcode1 = False
            Exit Function
        End If


        If Len(Trim(zipCode1)) > 8 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "ZipCode in Sheet : Zip Code under secondary address in Income Details cannot exceed 8 characters.Minimum 1 and up to 8 Characters." & Chr(13)
            ChkZipcode1 = False
            Exit Function
        End If

        If Not checkfieldspecialcharacter1(zipCode1) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "ZipCode in Sheet : Zip Code under secondary address in Income Details characters < > & ' " & Chr(34) & " are not allowed"" & chr(13)"
            ChkZipcode1 = False
            Exit Function
        End If
    Else
        zipCode1 = ""
    End If
End Function

'Mobile No Validation

Function ChkMobileNo() As Boolean
On Error Resume Next
    Dim i As Long
    ChkMobileNo = True
    mobileNo = Sheet1.Range("sheet1.Mobileno")
    If Trim(mobileNo) = "" Or IsEmpty(mobileNo) Then
'     EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Primary Mobile No. is mandatory." & Chr(13)
      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Primary Mobile No. in Sheet: Personal Information is mandatory.""" & Chr(13)
        ChkMobileNo = False
        Exit Function
    End If
    
'    If Mid(Sheet1.Range("sheet1.Country").Value, 1, 2) = "91" Then   'Ankita_06/06/2025
'    If Len(mobileNo) <> 10 Then
'    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 10 digit Mobile Number in sheet Income Details" & Chr(13)
'    ChkMobileNo = False
'    Exit Function
'    End If
'    End If  'Ankita_06/06/2025

    If Mid(Sheet1.Range("sheet1.Country").Value, 1, 2) = "91" Then   'Ankita_06/06/2025
        If Len(mobileNo) <> 10 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 10 digit Primary Mobile No. in sheet Income Details" & Chr(13)
            ChkMobileNo = False
            Exit Function
        End If
    
    ElseIf Mid(Sheet1.Range("sheet1.Country").Value, 1, 2) <> "91" Then
        If Len(mobileNo) < 5 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 5 digit Primary Mobile No. in sheet Income Details" & Chr(13)
            ChkMobileNo = False
            Exit Function
        End If
    End If
   
    
    If Len(mobileNo) < 5 Or Len(mobileNo) > 10 Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 10 digit Primary Mobile No." & Chr(13)
    ChkMobileNo = False
    Exit Function
    End If
 
    If Len(Trim(mobileNo)) > 10 Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Primary Mobile No. in Sheet Income Details cannot exceed 10 digits " & Chr(34)
        ChkMobileNo = False
        Exit Function
    End If

 
    

'    If Sheet1.Range("sheet1.MobileCountryCode").Value = "91" Then

End Function
 
 'Ankita_13/01/2026_V0.2=============================
 Function ChkMobileNo2() As Boolean
On Error Resume Next
    Dim i As Long
    ChkMobileNo2 = True
    mobileNo2 = Sheet1.Range("sheet1.Mobileno1")
'    If Trim(mobileNo2) = "" Or IsEmpty(mobileNo2) Then
'     EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Secondary mobile number is mandatory." & Chr(13)
'        ChkMobileNo2 = False
'        Exit Function
'    End If
    If Mid(Sheet1.Range("sheet1.MobileCountryCode1").Value, 1, 2) = "91" Then
        If Len(mobileNo2) <> 10 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 10 digit Secondary Mobile Number in sheet Income Details" & Chr(13)
            ChkMobileNo2 = False
            Exit Function
        End If
    
    ElseIf Mid(Sheet1.Range("sheet1.MobileCountryCode1").Value, 1, 2) <> "91" Then
        If Len(mobileNo2) < 5 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 5 digit Secondary Mobile Number in sheet Income Details" & Chr(13)
            ChkMobileNo2 = False
            Exit Function
        End If
    End If
    If Len(mobileNo2) < 5 Or Len(mobileNo2) > 10 Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 10 digit Secondary Mobile Number." & Chr(13)
    ChkMobileNo2 = False
    Exit Function
    End If
 
    If Len(Trim(mobileNo2)) > 10 Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Secondary Mobile Number in Sheet Income Details cannot exceed 10 digits " & Chr(34)
        ChkMobileNo2 = False
        Exit Function
    End If
End Function

 ''==================================================
 'Ankita_12/01/2026_V0.2=============================
 
 Function ChkMobileNo1() As Boolean
 On Error Resume Next
    Dim i As Long
    ChkMobileNo1 = True
    mobileNo1 = Sheet1.Range("sheet1.ContactRepAssessee")
    If Trim(mobileNo1) = "" Or IsEmpty(mobileNo1) Then
'     EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Contact Number is mandatory." & Chr(13)
        ChkMobileNo1 = False
        Exit Function
    End If
    If Mid(Sheet1.Range("sheet1.Country").Value, 1, 2) = "91" Then
        If Len(mobileNo1) <> 10 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 10 digit Contact Number of the respresentative in sheet Income Details" & Chr(13)
            ChkMobileNo1 = False
            Exit Function
        End If
    
    ElseIf Mid(Sheet1.Range("sheet1.Country").Value, 1, 2) <> "91" Then
        If Len(mobileNo1) < 5 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 5 digit Contact Number of the respresentative in sheet Income Details" & Chr(13)
            ChkMobileNo1 = False
            Exit Function
        End If
    End If
   
    If Len(mobileNo1) < 5 Or Len(mobileNo1) > 10 Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 10 digit Contact Number of the respresentative in sheet Income Details ." & Chr(13)
        ChkMobileNo1 = False
        Exit Function
    End If
 
    If Len(Trim(mobileNo1)) > 10 Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Contact Number of the respresentative in Sheet Income Details cannot exceed 10 digits " & Chr(34)
        ChkMobileNo1 = False
        Exit Function
    End If
End Function

 '===================================================
Function ChkCountrycode() As Boolean
On Error Resume Next
ChkCountrycode = True
MobileCountryCode = Sheet1.Range("sheet1.MobileCountryCode").Value
    If mobileNo <> "" Then
    If Trim(MobileCountryCode) = "" Or IsEmpty(MobileCountryCode) Then
'           EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Primary Mobile Country Code is Mandatory" & Chr(13)
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*  Country Code in Primary Mobile No. in Sheet: Personal Information is mandatory." & Chr(13)
            ChkCountrycode = False
        Exit Function
    End If
    End If
    
    If Len(Trim(MobileCountryCode)) > 5 Then
       EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Primary Country code in Sheet Income Details cannot exceed 5 characters " & Chr(13)
        ChkCountrycode = False
        Exit Function
    End If
    
'    If MobileCountryCode = "91" Then
'If Len(mobileNo) <> 10 Then
'    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid 10 digit Mobile Number in sheet Income Details" & Chr(13)
'    ChkCountrycode = False
'    Exit Function
'End If
'
'Else
'If Len(mobileNo) < 5 Or Len(mobileNo) > 10 Then
'    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter a valid Mobile Number in sheet Income Details" & Chr(13)
'    ChkCountrycode = False
'    Exit Function
'End If
'End If
'
'If CountrycodeMobileNo1 = "" Then
'    msgError = msgError & "* Country code for Mobile number 1 is Mandatory in Sheet Income Details" & Chr(13)
'End If
'
If MobileCountryCode = "0" Or MobileCountryCode = "00" Or MobileCountryCode = "000" Or MobileCountryCode = "0000" Or MobileCountryCode = "00000" Then
     EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter valid Country code for Primary Mobile in Sheet Income Details " & Chr(13)
    ChkCountrycode = False
    Exit Function
End If
    
'    If Trim(MobileCountryCode) = "" Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Mobile No  in Sheet : Income Details  is mandatory" & Chr(13)
'        ChkMobileNo = False
'        Exit Function
'    End If
    
    Dim i As Variant
    If Trim(MobileCountryCode) <> "" Then
        For i = 1 To Len(MobileCountryCode)
            If Not IsNumeric(Mid(MobileCountryCode, i, 1)) Then
                EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Primary Country code in Sheet Income Details must contain only digits from 0 to 9" & Chr(13)
                ChkCountrycode = False
            Exit Function
            End If
        Next
    End If
End Function

'Ankita_13/01/2026_V0.2==========================

Function ChkCountrycode2() As Boolean
On Error Resume Next
ChkCountrycode2 = True
MobileCountryCode2 = Sheet1.Range("sheet1.MobileCountryCode1").Value
'    If mobileNo <> "" Then
'    If Trim(MobileCountryCode2) = "" Or IsEmpty(MobileCountryCode2) Then
'      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*  Country Code in Secondary mobile number in Sheet: Personal Information is mandatory." & Chr(13)
'        ChkCountrycode2 = False
'        Exit Function
'    End If
'    End If
    
    If Len(Trim(MobileCountryCode2)) > 5 Then
       EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Secondary Mobile Country code in Sheet Income Details cannot exceed 5 characters " & Chr(13)
        ChkCountrycode2 = False
        Exit Function
    End If
If MobileCountryCode2 = "0" Or MobileCountryCode2 = "00" Or MobileCountryCode2 = "000" Or MobileCountryCode2 = "0000" Or MobileCountryCode2 = "00000" Then
     EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter valid Country code for Secondary Mobile in Sheet Income Details " & Chr(13)
    ChkCountrycode2 = False
    Exit Function
End If
    
    Dim i As Variant
    If Trim(MobileCountryCode2) <> "" Then
        For i = 1 To Len(MobileCountryCode2)
            If Not IsNumeric(Mid(MobileCountryCode2, i, 1)) Then
                EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Secondary Country code in Sheet Income Details must contain only digits from 0 to 9" & Chr(13)
                ChkCountrycode2 = False
            Exit Function
            End If
        Next
    End If
End Function



'===============================================
'Ankita_12/01/2026_V0.2==========================
Function ChkCountrycode1() As Boolean
On Error Resume Next
ChkCountrycode1 = True
MobileCountryCode1 = Sheet1.Range("sheet1.CountryCodeRepAssessee").Value
    If mobileNo1 <> "" Then
    If Trim(MobileCountryCode1) = "" Or IsEmpty(MobileCountryCode1) Then
           EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*  ""Country code in contact No. of representative assessee in Sheet: Personal Information is mandatory." & Chr(13)
        ChkCountrycode1 = False
        Exit Function
    End If
    End If
    
    If Len(Trim(MobileCountryCode1)) > 5 Then
       EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Country Code in Contact No. of Representative assessee in Sheet Income Details cannot exceed 5 characters " & Chr(13)
        ChkCountrycode1 = False
        Exit Function
    End If
    If MobileCountryCode1 = "0" Or MobileCountryCode1 = "00" Or MobileCountryCode1 = "000" Or MobileCountryCode1 = "0000" Or MobileCountryCode1 = "00000" Then
         EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter valid Country Code in Contact No. of Representative assessee in Sheet Income Details " & Chr(13)
        ChkCountrycode1 = False
        Exit Function
    End If
    
    Dim i As Variant
    If Trim(MobileCountryCode1) <> "" Then
        For i = 1 To Len(MobileCountryCode1)
            If Not IsNumeric(Mid(MobileCountryCode1, i, 1)) Then
                EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Country Code in Contact No. of Representative assessee in Sheet Income Details must contain only digits from 0 to 9" & Chr(13)
                ChkCountrycode1 = False
            Exit Function
            End If
        Next
    End If
End Function

'================================================

'Email Address Validation
Function ChkEmail() As Boolean
On Error Resume Next
    ChkEmail = True
    Email = Sheet1.Range("sheet1.EmailAddress")
    
    If Trim(Email) = "" Or IsEmpty(Email) Then
        ChkEmail = False
        Exit Function
    End If
    
    If Len(Trim(Email)) > 125 Then
'       fmsgbox ("Email Address in Sheet Income Details cannot exceed 125 characters ")
        fmsgbox ("Primary Email Id of the taxpayer cannot exceed 125 characters in Sheet Income Details")
        ChkEmail = False
        Exit Function
    End If
End Function

'Ankita_13/01/2026_V0.2==============

'Email Address Validation
Function ChkEmail2() As Boolean
On Error Resume Next
    ChkEmail2 = True
    Email_1 = Sheet1.Range("sheet1.EmailAddress1")
    
    If Trim(Email_1) = "" Or IsEmpty(Email_1) Then
        ChkEmail2 = False
        Exit Function
    End If
    
    If Len(Trim(Email_1)) > 125 Then
        fmsgbox ("Seconadry Email ID in Sheet Income Details cannot exceed 125 characters ")
        ChkEmail2 = False
        Exit Function
    End If
End Function

'Ankita_12/01/2026_V0.2==========

Function ChkEmail1() As Boolean
On Error Resume Next
    ChkEmail1 = True
    Email_2 = Sheet1.Range("sheet1.EmailRepAssessee")

    If Trim(Email_2) = "" Or IsEmpty(Email_2) Then
        ChkEmail1 = False
        Exit Function
    End If

    If Len(Trim(Email_2)) > 125 Then
        fmsgbox ("Email Address of representative assessee in Sheet Income Details cannot exceed 125 characters ")
        ChkEmail1 = False
        Exit Function
    End If
End Function

'================================
'ReturnFileSec Validation
Function ChkReturnFileSec() As Boolean
On Error Resume Next

    ChkReturnFileSec = True
    ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1")
    ReturnFileSec = Mid(ReturnFileSec, 1, 2)
    
    If Trim(ReturnFileSec) = "" Or Trim(ReturnFileSec) = "(Select)" Or Trim(ReturnFileSec) = "(S" Then
        ChkReturnFileSec = False
        Exit Function
    End If
    
    
'    If IsNumeric(Mid(Sheet1.Range("sheet1.ReturnFileSec").Value, 1, 2)) Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Please select valid section from the dropdown in sheet : Income Details " & Chr(13)
'        ChkReturnFileSec = False
'        Exit Function
'    End If
    
End Function

'ReturnType Validation
'Function ChkReturnType() As Boolean
'On Error Resume Next
'    ChkReturnType = True
'    ReturnTypeMsg = True
'
'    ReturnType = Sheet1.Range("sheet1.ReturnType1")
'    ReturnType = Mid(ReturnType, 1, 1)
'
'    If Trim(ReturnType) = "" Or Trim(ReturnType) = "(Select)" Or Trim(ReturnType) = "(" Then
'        ChkReturnType = False
'        Exit Function
'    End If
'
'    If Trim(ReturnFileSec) = "18" Then
'        If Trim(ReturnType) <> "I" Then
'        ReturnTypeMsg = False
'        MsgBox "Section field in Part A-General information can be only " & " In response to notice under section 139(9)" & "for filing type " & " In response to a notice u/s 139(9)"
'        Exit Function
'        End If
'    End If
'
'    If Trim(ReturnFileSec) = "17" Then
'        If Trim(ReturnType) <> "R" Then
'        ReturnTypeMsg = False
'        MsgBox "Please select valid Filing type for the selected section."
'        Exit Function
'        End If
'    End If
'
'End Function

'ResidentialStatus Validation
'Function ChkResidentialStatus() As Boolean
'On Error Resume Next
'
'    ChkResidentialStatus = True
'    ResidentialStatus = Sheet1.Range("sheet1.ResidentialStatus1")
'    If Trim(ResidentialStatus) = "" Or Trim(ResidentialStatus) = "(Select)" Then
'        ChkResidentialStatus = False
'        Exit Function
'    End If
'End Function

'Status Validation
Function ChkStatus() As Boolean
On Error Resume Next
    ChkStatus = True
    status = Sheet1.Range("sheet1.Status")
    status = Mid(status, 1, 1)
    
    If Trim(status) = "" Or Trim(status) = "(Select)" Then
        ChkStatus = False
        Exit Function
    End If
End Function

'Country Validation
Function ChkCountry() As Boolean
On Error Resume Next
    ChkCountry = True
    sCountry = Sheet1.Range("sheet1.Country")
'
    If isdropdownblank(sCountry) Then
        sCountry = ""
    End If
'
    If Trim(sCountry) = "" Or Trim(sCountry) = "(Select)" Or Trim(sCountry) = "(" Then
        ChkCountry = False
        Exit Function
    End If
End Function

'Ankita_14/01/2026_V0.2==========
Function ChkCountry1() As Boolean
On Error Resume Next
    ChkCountry1 = True
    sCountry1 = Sheet1.Range("sheet1.Country1")
    If isdropdownblank(sCountry1) Then
        sCountry1 = ""
    End If
    If Trim(sCountry1) = "" Or Trim(sCountry1) = "(Select)" Or Trim(sCountry1) = "(" Then
        ChkCountry1 = False
        Exit Function
    End If
End Function


'================================
'Employer Category Validation
Function ChkEmpCategory() As Boolean
On Error Resume Next
    ChkEmpCategory = True
    empcat = Sheet1.Range("sheet1.EmployerCategory1")
    
    If Trim(empcat) = "" Or Trim(empcat) = "(Select)" Then
        ChkEmpCategory = False
        Exit Function
    End If
End Function


'Gender Validation
Function ChkGender() As Boolean
On Error Resume Next
    ChkGender = True
    Gender = Sheet1.Range("sheet1.Gender1")
 
    
    If Trim(Gender) = "" Or Trim(Gender) = "(Select)" Then
        ChkGender = False
        Exit Function
    End If
End Function

'Governed by Portugese Validation
'Function ChkGovernedbyPortugese() As Boolean
'On Error Resume Next
'    ChkGovernedbyPortugese = True
'    Portugese = Sheet1.Range("sheet1.PortugeseCC5A")
'
'    If Trim(Portugese) = "" Or Trim(Portugese) = "(Select)" Then
'        ChkGovernedbyPortugese = False
'        Exit Function
'    End If
'End Function
Function ChkDOB() As Boolean
On Error Resume Next
    ChkDOB = True
    dob = Sheet1.Range("sheet1.DOB")
    If Trim(dob) = "" Or IsEmpty(dob) Or dob = "00/00/0000" Then
        ChkDOB = False
        Exit Function
    End If
End Function
   
Function CheckPAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 1 is for individuals .So,4th character of PAN should be "P"

    CheckPAN = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPAN = False
            Exit Function
        End If
'        Enhancement
'        If Not ChkAlphabetP(Mid(PAN, 4, 1)) Then
'            CheckPAN = False
'            Exit Function
'        End If
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPAN = False
            Exit Function
        End If
       If Not checkfieldSuperSpecialcharacterDot(PAN) Then
       CheckPAN = False
        Exit Function
        End If
        
    End If
End Function

Function CheckDoneePAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 1 is for individuals .So,4th character of PAN should be "P"

    CheckDoneePAN = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckDoneePAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckDoneePAN = False
            Exit Function
        End If
    End If
End Function

Function ChkAlphabet(Char As String) As Boolean
    ChkAlphabet = True
    If ((asc(Char) < 65) Or (asc(Char) > 90)) Then
        ChkAlphabet = False
    End If
End Function

'Enhancement
'Function ChkAlphabetP(Char As String) As Boolean
'    ChkAlphabetP = True
'    If ((asc(Char) <> 80)) Then
'        ChkAlphabetP = False
'    End If
'End Function

Function ValidatePinCode() As Boolean
    Dim PinCode As String
    Dim i As Long
    
    ValidatePinCode = True
    PinCode = Sheet1.Range("sheet1.PinCode").Value
        
    If Sheet1.Range("sheet1.Country").Value = "91-INDIA" Then
    If Trim(PinCode) = "" Or IsEmpty(PinCode) Then
        ValidatePinCode = False
        Exit Function
    End If
    End If

     For i = 1 To Len(PinCode)
        If Not IsNumeric(Mid(PinCode, i, 1)) Then
             fmsgbox ("* Pin code under primary address in Sheet Income Details must contain only digits from 0 to 9")
            ValidatePinCode = False
            Exit Function
        End If
    Next
    If Len(PinCode) > 6 Then
             fmsgbox ("* Pin code under primary address in Sheet Income Details must contain only 6 digits")
            ValidatePinCode = False
            Exit Function
    End If
End Function

'Ankita_14/01/2026_V0.2==========
Function ValidatePinCode1() As Boolean
    Dim PinCode1 As String
    Dim i As Long
    
    ValidatePinCode1 = True
    PinCode1 = Sheet1.Range("sheet1.PinCode1").Value
        
    If Sheet1.Range("sheet1.Country1").Value = "91-INDIA" Then
    If Trim(PinCode1) = "" Or IsEmpty(PinCode1) Then
        ValidatePinCode1 = False
        Exit Function
    End If
    End If

     For i = 1 To Len(PinCode1)
        If Not IsNumeric(Mid(PinCode1, i, 1)) Then
             fmsgbox ("* Pin code under secondary address in Sheet Income Details must contain only digits from 0 to 9")
            ValidatePinCode1 = False
            Exit Function
        End If
    Next
    If Len(PinCode1) > 6 Then
             fmsgbox ("* Pin code under secondary address in Sheet Income Details must contain only 6 digits")
            ValidatePinCode1 = False
            Exit Function
    End If
End Function



'================================
Function ValidateMobileNumber() As Boolean
    Dim MobileNumber As String
    Dim i As Long
    
    ValidateMobileNumber = True
    MobileNumber = Sheet1.Range("sheet1.Mobileno").Value
  
    For i = 1 To Len(MobileNumber)
        If Not IsNumeric(Mid(MobileNumber, i, 1)) Then
             fmsgbox ("* ""Primary Mobile No. must contain only digits from 0 to 9 in Sheet Income Details""")
            ValidateMobileNumber = False
            Exit Function
        End If
        
       If (Mid(MobileNumber, 1, 1) = "0") Then
       fmsgbox ("* ""Primary Mobile No. cannot begin with '0' in sheet Income Details""")
            ValidateMobileNumber = False
            Exit Function
        End If
        
    Next
End Function

'Ankita_13/01/2026=========
Function ValidateMobileNumber2() As Boolean
    Dim MobileNumber2 As String
    Dim i As Long
    ValidateMobileNumber2 = True
    MobileNumber2 = Sheet1.Range("sheet1.MobileCountryCode1").Value
    For i = 1 To Len(MobileNumber1)
        If Not IsNumeric(Mid(MobileNumber2, i, 1)) Then
             fmsgbox ("* ""Mobile No. in contact No. of representative assessee must contain only digits from 0 to 9 in Sheet Income Details""")
            ValidateMobileNumber2 = False
            Exit Function
        End If
       If (Mid(MobileNumber2, 1, 1) = "0") Then
       fmsgbox ("* ""Mobile No. in contact No. of representative assessee cannot begin with '0' in sheet Income Details""")
            ValidateMobileNumber2 = False
            Exit Function
        End If
    Next
End Function


'=========================

'Ankita_12/01/2026==========
Function ValidateMobileNumber1() As Boolean
    Dim MobileNumber1 As String
    Dim i As Long
    
    ValidateMobileNumber1 = True
    MobileNumber1 = Sheet1.Range("sheet1.CodeRepAssessee").Value
  
    For i = 1 To Len(MobileNumber1)
        If Not IsNumeric(Mid(MobileNumber1, i, 1)) Then
             fmsgbox ("* ""Secondary Mobile Number must contain only digits from 0 to 9 in Sheet Income Details""")
            ValidateMobileNumber1 = False
            Exit Function
        End If
        
       If (Mid(MobileNumber1, 1, 1) = "0") Then
       fmsgbox ("* ""Secondary Mobile Number cannot begin with '0' in sheet Income Details""")
            ValidateMobileNumber1 = False
            Exit Function
        End If
        
    Next
End Function

Function ValidateMobileNumberSec() As Boolean
    Dim MobileNumber As String
    Dim i As Long
    
    ValidateMobileNumberSec = True
    MobileNumber = Sheet1.Range("sheet1.MobileNoSec").Value
  
    For i = 1 To Len(MobileNumber)
        If Not IsNumeric(Mid(MobileNumber, i, 1)) Then
             fmsgbox ("* ""Mobile Number 2 in Sheet Income Details must contain only digits from 0 to 9""")
            ValidateMobileNumberSec = False
            Exit Function
        End If
    Next
End Function



Function CheckDOB(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.

    CheckDOB = True
    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDate(dob) Then
            CheckDOB = False
           fmsgbox ("* Date of Birth in Sheet Income Details must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        
        'If Not ChkMaxDOBDate(dob, "31/03/2023") Then
'        If Not ChkMaxDOBDate(dob, "31/03/2024") Then
            'fmsgbox ("* Date of Birth in Sheet Income Details  should not be after 31/03/2023 for A.Y.2023-24.")
            'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
'            If Not ChkMaxDOBDate(dob, "31/03/2025") Then
'            fmsgbox ("* Date of Birth in Sheet Income Details should not be after 31/03/2025 for A.Y.2025-26.")
                       'Ankita_27/12/2025=============
    Dim cutoff As Date
    cutoff = CDate(Sheet5.Range("DOB_1").Value)
        If Not ChkMaxDOBDate23_24(dob, Sheet5.Range("DOB_1").Value) Then
                 fmsgbox ("* Date of birth in sheet income details should not be on or after " & Dformat1(cutoff, "dd/mm/yyyy"))
            CheckDOB = False
            Exit Function
        Else
            AssesseeDob = dob
        End If
    End If
 End Function
 Function CheckDateBefore(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.

    CheckDateBefore = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckDateBefore = False
           fmsgbox ("* Date of Verification in Sheet Tax Paid and Verification must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        
        'Changes year from 2023 to 2024
        'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
'        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2025") Then

'Ankita_28/01/2026==============
    Dim cutoff13 As Date
    cutoff13 = CDate(Sheet5.Range("DOB_1").Value)
                If EfilingCommon.checkFirstDateBefore(dob, Sheet5.Range("DOB_1").Value) Then
            fmsgbox ("* Date of Verification in Sheet Tax Paid and Verification should be on or after " & Dformat1(cutoff13, "dd/mm/yyyy"))
            CheckDateBefore = False
            Exit Function
        Else
        VerDate = dob
       End If
    End If
 End Function
Function CheckFilingDateBefore(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.


    CheckFilingDateBefore = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckFilingDateBefore = False
           fmsgbox ("* Date of Filing in Sheet Income Details  must be a valid dd/mm/yyyy format")
            Exit Function
        End If
         
        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2024") Then 'Year Changed from 2023 to 2024 by Ankita on 16/12/2024
          'AY changing from 2023-24 to 2024-25 Bindu
          'Year Changed by Ankita on 16/12/2024
            fmsgbox ("* Date of filing of original Return cannot be prior to 01/04/2024 for A.Y 2025-26 in Income Details")
            CheckFilingDateBefore = False
            Exit Function
            Else
            filingDate = dob
       End If
    End If
 End Function
 Function CheckFilingDateBefore1(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.


    CheckFilingDateBefore1 = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckFilingDateBefore1 = False
           fmsgbox ("* Date of Filing in Sheet Income Details  must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        
        If EfilingCommon.checkFirstDateBefore1(dob, "31/03/2024") Then  'Year Changed from 2023 to 2024 by Ankita on 16/12/2024
            'AY changing from 2023-24 to 2024-25 Bindu
            fmsgbox ("* Date of filing of original Return cannot be prior to  01/04/2024 for A.Y 2025-26 in Income Details")     'Year Changed from 2023 to 2024 by Ankita on 16/12/2024
            CheckFilingDateBefore1 = False
            Exit Function
            Else
            filingDate = dob
       End If
    End If
 End Function
 
  Function CheckNoticeDateBefore(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.

    CheckNoticeDateBefore = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckNoticeDateBefore = False
           fmsgbox ("* Date of Notice/Order in Sheet Income Details  must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        
        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2023") Then
            fmsgbox ("* Date of Notice/Order cannot be prior to 01/04/2023")
            CheckNoticeDateBefore = False
            Exit Function
             Else
            NoticeDate = dob
       End If
    End If
 End Function
 
Function ChkMinDOBDate(dob As Variant, minDefinedDOB As Variant) As Boolean
          

On Error Resume Next
     ChkMinDOBDate = True
     If Len(dob) > 0 Then
        If val(Mid(dob, 7, 4)) <= 2017 Then
            If val(Mid(dob, 4, 2)) <= 3 Then
                If val(Mid(dob, 1, 2)) <= 31 Then
                    ChkMinDOBDate = False
                    Exit Function
                End If
            End If
        End If
     End If
     End Function

'Chandru
Function ChkMinDOBDate_1(dob As Variant, minDefinedDOB As Variant) As Boolean
          

On Error Resume Next
     ChkMinDOBDate_1 = True
     If Len(dob) > 0 Then
        If val(Mid(dob, 7, 4)) <= 2002 Then
            If val(Mid(dob, 4, 2)) <= 3 Then
                If val(Mid(dob, 1, 2)) <= 31 Then
                    ChkMinDOBDate_1 = False
                    Exit Function
                End If
            End If
        End If
     End If
     End Function
 
 'Changin year from 2023 to 2024
Function FormatNCheckDate(ByRef dt As Variant, Optional Year As String = "2024") As Boolean
On Error Resume Next
    FormatNCheckDate = True
    If Len(dt) > 0 Then
        'Format the date in dd/mm/yyyy format
        If Mid(dt, 3, 1) <> "/" Then 'Checking between dd/mm
            If Mid(dt, 3, 1) <> "\" Or Mid(dt, 3, 1) <> "-" Or Mid(dt, 3, 1) <> "." Or Mid(dt, 3, 1) <> "," Then
                dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
            End If
        End If
        
        If Mid(dt, 6, 1) <> "/" Then 'Checking between mm/yyyy
            If Mid(dt, 6, 1) <> "\" Or Mid(dt, 6, 1) <> "-" Or Mid(dt, 6, 1) <> "." Or Mid(dt, 6, 1) <> "," Then
                dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
            End If
        End If
        
        'Checking Date if it is in correct format :Day(dd), Month(mm) & Year(yyyy)
        
        If val(Mid(dt, 1, 2)) < 0 Or val(Mid(dt, 1, 2)) > 31 Then FormatNCheckDate = False
        If val(Mid(dt, 4, 2)) < 0 Or val(Mid(dt, 4, 2)) > 12 Then FormatNCheckDate = False
        If Not IsDate(dt) Then FormatNCheckDate = False
    End If
End Function

Function ChkMaxDOBDate(dob As Variant, maxDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMaxDOBDate = True
     If Len(dob) > 0 Then
     'Changing year from 2023 to 2024
'        If val(Mid(dob, 7, 4)) >= 2025 Then   '17/12/2025  'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
       'Ankita_17/12/2025====================
        If val(Mid(dob, 7, 4)) >= Sheet5.Range("DOB_YEAR").Value Then
            If val(Mid(dob, 4, 2)) >= 4 Then
                If val(Mid(dob, 1, 2)) >= 1 Then
                    ChkMaxDOBDate = False
                    Exit Function
                End If
        Else
             'Changing year from 2023 to 2024
'              If val(Mid(dob, 4, 2)) <= 3 And val(Mid(dob, 7, 4)) <= 2025 Then           '17/12/2025         'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
          If val(Mid(dob, 4, 2)) <= 3 And val(Mid(dob, 7, 4)) <= Sheet5.Range("DOB_YEAR").Value Then
              Else
                ChkMaxDOBDate = False
              Exit Function
             End If
            End If
        Else
        End If
     End If
End Function

'Chandru
Function ChkMaxDOBDate_1(dob As Variant, maxDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMaxDOBDate_1 = True
     If Len(dob) > 0 Then
     'Changing year from 2023 to 2024
        If val(Mid(dob, 7, 4)) >= 2007 Then
            If val(Mid(dob, 4, 2)) >= 4 Then
                If val(Mid(dob, 1, 2)) >= 1 Then
                    ChkMaxDOBDate_1 = False
                    Exit Function
                End If
        Else
             'Changing year from 2023 to 2024
             
              If val(Mid(dob, 4, 2)) <= 3 And val(Mid(dob, 7, 4)) <= 2007 Then
              Else
                ChkMaxDOBDate_1 = False
                    Exit Function
             End If
            End If
        Else
        End If
     End If
End Function

Function CheckEmailAddress(emailAddress As String, ByRef vType As String) As Boolean
On Error Resume Next
    CheckEmailAddress = True
    If Len(emailAddress) > 0 Then
        If Not CheckSpecialCharacter(Mid(emailAddress, 1, 1)) Then
            vType = "1"
            CheckEmailAddress = False
            Exit Function
        End If
        
        If Len(emailAddress) > 125 Then
            vType = "2"
            CheckEmailAddress = False
            Exit Function
        End If
                
        If Not IsValidEmail(emailAddress) Then
            vType = "3"
            CheckEmailAddress = False
            Exit Function
        End If
        If Not CheckSpecialCharacter_findconsequtive(emailAddress) Then  'Added by Ankita 08/01/2025
            vType = "3"
            CheckEmailAddress = False
            Exit Function
        End If
    Else
            vType = "4"
            CheckEmailAddress = False
            Exit Function
    End If
End Function

'Ankita_13/02/2026==============
Function CheckEmailAddress_new(emailAddress As String, ByRef vType As String) As Boolean
On Error Resume Next
    CheckEmailAddress_new = True
    If Len(emailAddress) > 0 Then
        If Not CheckSpecialCharacter(Mid(emailAddress, 1, 1)) Then
            vType = "1"
            CheckEmailAddress_new = False
            Exit Function
        End If
        
        If Len(emailAddress) > 125 Then
            vType = "2"
            CheckEmailAddress_new = False
            Exit Function
        End If
                
        If Not IsValidEmail(emailAddress) Then
            vType = "3"
            CheckEmailAddress_new = False
            Exit Function
        End If
        If Not CheckSpecialCharacter_FindConsecutive_new(emailAddress) Then  'Added by Ankita 08/01/2025
            vType = "3"
            CheckEmailAddress_new = False
            Exit Function
        End If
    Else
            vType = "4"
            CheckEmailAddress_new = False
            Exit Function
    End If
End Function

'Ankita_23/02/2026========

Function CheckEmailAddress_new1(emailAddress As String, ByRef vType As String) As Boolean
On Error Resume Next
    CheckEmailAddress_new1 = True
    If Len(emailAddress) > 0 Then
        If Not CheckSpecialCharacter(Mid(emailAddress, 1, 1)) Then
            vType = "1"
            CheckEmailAddress_new1 = False
            Exit Function
        End If
        
        If Len(emailAddress) > 125 Then
            vType = "2"
            CheckEmailAddress_new1 = False
            Exit Function
        End If
                
        If Not IsValidEmail1(emailAddress) Then
            vType = "3"
            CheckEmailAddress_new1 = False
            Exit Function
        End If
        If Not CheckSpecialCharacter_FindConsecutive_new(emailAddress) Then  'Added by Ankita 08/01/2025
            vType = "3"
            CheckEmailAddress_new1 = False
            Exit Function
        End If
    Else
            vType = "4"
            CheckEmailAddress_new1 = False
            Exit Function
    End If
End Function

'Ankita_13/02/2026==============
Public Function CheckSpecialCharacter_FindConsecutive_new(ByVal emailid As String) As Boolean
    Dim i As Long
    Dim ch As String * 1
    Dim consecutive As Long
    If InStr(1, emailid, "_@", vbTextCompare) > 0 Then
        CheckSpecialCharacter_FindConsecutive_new = False
        Exit Function
    End If
    For i = 1 To Len(emailid)
        ch = Mid$(emailid, i, 1)
        If ch = "-" Or ch = "/" Or ch = "_" Or ch = "." Then
            consecutive = consecutive + 1
            If consecutive >= 2 Then
                CheckSpecialCharacter_FindConsecutive_new = False
                Exit Function
            End If
        Else
            consecutive = 0
        End If
    Next i
    CheckSpecialCharacter_FindConsecutive_new = True
End Function

'---------------------------
 
'===============================
'Added by Ankita 08/01/2025
Function CheckSpecialCharacter_findconsequtive(emailid As Variant) As Boolean

On Error Resume Next

    Dim specialCharArray As Variant

    Dim iCharCount, iSpecialChar As Long

    Dim M_a_L_L_I As Variant

    M_a_L_L_I = 0

    CheckSpecialCharacter_findconsequtive = True

    specialCharArray = Array("-", "/", "_", ".")

   Dim charcount As Boolean

   Dim d As Variant

   charcount = False

    For iCharCount = 1 To Len(emailid)

           If Mid(emailid, iCharCount, 1) = "-" Or Mid(emailid, iCharCount, 1) = "/" Or _
              Mid(emailid, iCharCount, 1) = "_" Or Mid(emailid, iCharCount, 1) = "." Then

                          M_a_L_L_I = M_a_L_L_I + 1

           Else:

                          M_a_L_L_I = 0

           End If

             If M_a_L_L_I >= 2 And d = iCharCount - 1 Then

                    CheckSpecialCharacter_findconsequtive = False

                    Exit Function

             End If

             d = iCharCount

        Next
 
End Function
 

Function CheckSpecialCharacter(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    
    CheckSpecialCharacter = True
    specialCharArray = Array("*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<")
    For iCharCount = 1 To Len(emailid)
        For iSpecialChar = 0 To UBound(specialCharArray)
        If Mid(emailid, iCharCount, 1) = specialCharArray(iSpecialChar) Then
            CheckSpecialCharacter = False
            Exit Function
        End If
        Next
    Next
End Function

Function IsValidEmail(strEmail)
On Error Resume Next
    Dim strFieldArray As Variant
    Dim strFieldItem As Variant
    Dim i As Long, c As String, blnIsItValid As Boolean
    blnIsItValid = True
     
    i = Len(strEmail) - Len(Application.Substitute(strEmail, "@", ""))
    If i <> 1 Then IsValidEmail = False: Exit Function
    ReDim strFieldArray(1 To 2)
    strFieldArray(1) = Left(strEmail, InStr(1, strEmail, "@", 1) - 1)
    strFieldArray(2) = Application.Substitute(Right(strEmail, Len(strEmail) - Len(strFieldArray(1))), "@", "")
    For Each strFieldItem In strFieldArray
        If Len(strFieldItem) <= 0 Then
            blnIsItValid = False
            IsValidEmail = blnIsItValid
            Exit Function
        End If
        For i = 1 To Len(strFieldItem)
            c = LCase(Mid(strFieldItem, i, 1))
            If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
                blnIsItValid = False
                IsValidEmail = blnIsItValid
                Exit Function
            End If
        Next i
        If Left(strFieldItem, 1) = "." Or Right(strFieldItem, 1) = "." Then
            blnIsItValid = False
            IsValidEmail = blnIsItValid
            Exit Function
        End If
    Next strFieldItem
    If InStr(strFieldArray(2), ".") <= 0 Then
        blnIsItValid = False
        IsValidEmail = blnIsItValid
        Exit Function
    End If
    
    If InStr(strEmail, "..") > 0 Then
        blnIsItValid = False
        IsValidEmail = blnIsItValid
        Exit Function
    End If
    IsValidEmail = blnIsItValid
End Function

'Ankita_23/02/2026======

Function IsValidEmail1(strEmail As String) As Boolean
    On Error Resume Next
    Dim localPart As String, domainPart As String
    Dim i As Long, c As String
    Dim allowedLocal As String, allowedDomain As String
    IsValidEmail1 = True
    
    i = Len(strEmail) - Len(Application.Substitute(strEmail, "@", ""))
    If i <> 1 Then IsValidEmail1 = False: Exit Function
    localPart = Left(strEmail, InStr(1, strEmail, "@", vbTextCompare) - 1)
    domainPart = Mid(strEmail, InStr(1, strEmail, "@", vbTextCompare) + 1)
    
    If Len(localPart) = 0 Or Len(domainPart) = 0 Then IsValidEmail1 = False: Exit Function
    
    If Left(localPart, 1) = "." Or Right(localPart, 1) = "." Then IsValidEmail1 = False: Exit Function
    If Left(domainPart, 1) = "." Or Right(domainPart, 1) = "." Then IsValidEmail1 = False: Exit Function
    
    If InStr(strEmail, "..") > 0 Then IsValidEmail1 = False: Exit Function
    
    If InStr(domainPart, ".") <= 0 Then IsValidEmail1 = False: Exit Function
    
    allowedLocal = "abcdefghijklmnopqrstuvwxyz_.-"   ' underscore allowed here
    allowedDomain = "abcdefghijklmnopqrstuvwxyz.-"   ' underscore NOT allowed
    
    For i = 1 To Len(localPart)
        c = LCase(Mid(localPart, i, 1))
        If InStr(allowedLocal, c) <= 0 And Not IsNumeric(c) Then
            IsValidEmail1 = False
            Exit Function
        End If
    Next i
    
    For i = 1 To Len(domainPart)
        c = LCase(Mid(domainPart, i, 1))
        If InStr(allowedDomain, c) <= 0 And Not IsNumeric(c) Then
            IsValidEmail1 = False
            Exit Function
        End If
    Next i
    Dim labels() As String, lbl As Variant
    labels = Split(domainPart, ".")
    For Each lbl In labels
        If Len(lbl) = 0 Then IsValidEmail1 = False: Exit Function
        If Left(lbl, 1) = "-" Or Right(lbl, 1) = "-" Then
            IsValidEmail1 = False
            Exit Function
        End If
    Next lbl
End Function

Function ValidateSTDcode() As Boolean
    ValidateSTDcode = True
    Dim stDcode As String
    Dim i As Long
    
    stDcode = Sheet1.Range("sheet1.STDcode").Value
    If Len(stDcode) > 5 Then
     fmsgbox ("* STDcode in Sheet Income Details should be at most 5 digits")
    ValidateSTDcode = False
    Exit Function
    End If
    
    If Trim(Mid(stDcode, 1, 1)) = 0 Then
        fmsgbox ("* STDCode in Sheet Income Details Do not prefix '0' before STD code")
        ValidateSTDcode = False
        Exit Function
    End If

    If Trim(stDcode) <> "" Then
        For i = 1 To Len(stDcode)
            If Not IsNumeric(Mid(stDcode, i, 1)) Then
                fmsgbox ("* STDcode in Sheet Income Details must contain only digits from 0 to 9")
                ValidateSTDcode = False
            Exit Function
            End If
        Next
    End If

End Function

Function ValidatePhoneNo() As Boolean
    Dim phoneNo, stDcode As String
    Dim i As Long
    
    ValidatePhoneNo = True
    phoneNo = Sheet1.Range("sheet1.PhoneNo").Value
    stDcode = Sheet1.Range("sheet1.STDcode").Value

    If Trim(phoneNo) <> "" Then
        For i = 1 To Len(phoneNo)
            If Not IsNumeric(Mid(phoneNo, i, 1)) Then
              fmsgbox ("* Phone Number in Sheet Income Details must contain only digits from 0 to 9")
              Sheet1.Range("sheet1.PhoneNo").Value = ""
            ValidatePhoneNo = False
            Exit Function
            End If
        Next
    End If

    If (Trim(phoneNo) <> "" Or Trim(stDcode) <> "") Then
        If (Len(phoneNo) + Len(stDcode)) <> 10 Then
            MsgPISheet = MsgPISheet + ("* Invalid Phone Number. STD Code + Landline Number should be 10 digits and cannot begin with '0' in Income Details.") & Chr(13)
            ValidatePhoneNo = False
            Exit Function
        End If
    End If
    
    If Mid((phoneNo), 1, 1) = 0 Then
            fmsgbox ("* Invalid Phone Number. STD Code + Landline Number should be 10 digits and cannot begin with '0' in Income Details.") & Chr(13)
            ValidatePhoneNo = False
            Exit Function
    End If

End Function

Function CheckState(State As String, ByRef vType As String) As Boolean
On Error Resume Next
'State is mandatory in tab: Personal Information
    
    CheckState = True
    If Trim(Range("sheet1.StateCode1").Value) = "" Or _
       IsEmpty(Range("sheet1.StateCode1").Value) Then
        vType = "1"
        CheckState = False
        Exit Function
    End If
    
EXITS:
    If Range("sheet1.StateCode1").Value = "(Select)" Then
    Range("sheet1.Country").Value = "(Select)"
        Range("sheet1.PinCode").Value = ""
    ElseIf Range("sheet1.StateCode1").Value = "99-FOREIGN" Then
        Range("sheet1.PinCode").Value = "999999"
        Range("sheet1.Country").Value = "(Select)"
      
        
    ElseIf Not Range("sheet1.StateCode1").Value = "99-FOREIGN" Or Range("sheet1.StateCode1").Value = "(Select)" Then
    Range("sheet1.Country").Value = "91-INDIA"
     Range("sheet1.PinCode").Value = ""
    End If
    
    If Range("sheet1.StateCode1").Value <> "99-FOREIGN" And Range("sheet1.Country").Value = "99-FOREIGN" Then
    fmsgbox "*Country cannot be other than India as you have selected an Indian state" & Chr(13)
    End If
    
End Function

'Ankita_14/01/2026_V0.2==========

Function CheckState1(State As String, ByRef vType As String) As Boolean
On Error Resume Next
    CheckState1 = True
    If Trim(Range("sheet1.StateCode2").Value) = "" Or _
       IsEmpty(Range("sheet1.StateCode2").Value) Then
        vType = "1"
        CheckState1 = False
        Exit Function
    End If
EXITS:
    If Range("sheet1.StateCode2").Value = "(Select)" Then
    Range("sheet1.Country1").Value = "(Select)"
        Range("sheet1.PinCode1").Value = ""
    ElseIf Range("sheet1.StateCode2").Value = "99-FOREIGN" Then
        Range("sheet1.PinCode1").Value = "999999"
        Range("sheet1.Country1").Value = "(Select)"
    ElseIf Not Range("sheet1.StateCode2").Value = "99-FOREIGN" Or Range("sheet1.StateCode2").Value = "(Select)" Then
    Range("sheet1.Country1").Value = "91-INDIA"
     Range("sheet1.PinCode1").Value = ""
    End If
    If Range("sheet1.StateCode2").Value <> "99-FOREIGN" And Range("sheet1.Country1").Value = "99-FOREIGN" Then
    fmsgbox "*Country cannot be other than India as you have selected an Indian state" & Chr(13)
    End If
    
End Function

'================================
Function ChkInvestment() As Boolean
On Error Resume Next
        ChkInvestment = False
End Function
Function ChkSeventhProvisoFlag() As Boolean
On Error Resume Next
    ChkSeventhProvisoFlag = True
    ProvisoFlag = Sheet1.Range("sheet1.SeventhProvisoFlag")
  
    If Trim(ProvisoFlag) = "" Or Trim(ProvisoFlag) = "(Select)" Then
        ChkSeventhProvisoFlag = False
        Exit Function
    End If
End Function
'Change-22.11.2022.102.16C
'Function ChkDepositAmountFlag() As Boolean
'On Error Resume Next
'    ChkDepositAmountFlag = True
'    DepositAmountFlag = Sheet1.Range("Sheet1.DepositAmountFlag")
'    DepositAmount = Sheet1.Range("Sheet1.DepositAmount")
'
'    If Sheet1.Range("sheet1.SeventhProvisoFlag").Value = "Yes" Then
'    If Trim(DepositAmountFlag) = "" Or Trim(DepositAmountFlag) = "(Select)" And Mid(ProvisoFlag, 1, 1) = "Y" Then
'        ChkDepositAmountFlag = False
'        Exit Function
'    End If
'
'
''    If Trim(DepositAmountFlag) = "Yes" And DepositAmount = "" Then
''    fmsgbox "you have selected ""Yes"" for ""Have you deposited amount or aggregate of amounts exceeding Rs. 1 Crore in one or more current account during the previous year?"", hence please enter the amount in Sheet Income details." & Chr(13)
''    ChkDepositAmountFlag = False
''        Exit Function
''    End If
'    End If
'
'End Function
'---end

Function ChkAggrigateAmountFlag() As Boolean
On Error Resume Next
    ChkAggrigateAmountFlag = True
    AggrigateAmountFlag = Sheet1.Range("Sheet1.AggrigateAmountFlag")
    AggrigateAmount = Sheet1.Range("Sheet1.AggrigateAmount")
    If Sheet1.Range("sheet1.SeventhProvisoFlag").Value = "Yes" Then
    If Trim(AggrigateAmountFlag) = "" Or Trim(AggrigateAmountFlag) = "(Select)" And Mid(ProvisoFlag, 1, 1) = "Y" Then
        ChkAggrigateAmountFlag = False
        Exit Function
    End If
    
    
'    If Trim(AggrigateAmountFlag) = "Yes" And AggrigateAmount = "" Then
'   fmsgbox "you have selected ""Yes"" for ""Have you incurred expenditure of an amount or aggregate of amount exceeding Rs. 2 lakhs for travel to a foreign country for yourself or for any other person?"", hence please enter the amount in Sheet Income details." & Chr(13)
'    ChkAggrigateAmountFlag = False
'        Exit Function
'    End If
    End If
    
End Function

Function ChkAggrigateAmountFlag1() As Boolean
On Error Resume Next
    ChkAggrigateAmountFlag1 = True
    AggrigateAmountFlag1 = Sheet1.Range("Sheet1.AggrigateAmountFlag1")
    AggrigateAmount1 = Sheet1.Range("Sheet1.AggrigateAmount1")
    If Sheet1.Range("sheet1.SeventhProvisoFlag").Value = "Yes" Then
    If Trim(AggrigateAmountFlag1) = "" Or Trim(AggrigateAmountFlag1) = "(Select)" And Mid(ProvisoFlag, 1, 1) = "Y" Then
        ChkAggrigateAmountFlag1 = False
        Exit Function
    End If
    
    
'    If Trim(AggrigateAmountFlag1) = "Yes" And AggrigateAmount1 = "" Then
' fmsgbox "you have selected ""Yes"" for ""Have you incurred expenditure of amount or aggregate of amount exceeding Rs. 1 lakh on consumption of electricity during the previous year? "", hence please enter the amount in Sheet Income details." & Chr(13)
'    ChkAggrigateAmountFlag1 = False
'        Exit Function
'    End If
    End If
End Function

Sub ValidateSchedulePI()
    
    If mIncmDtls.PIShtValidate Then
    End If

End Sub



Function PIShtValidate() As Boolean
    PIShtValidate = True
    Dim vbMessgaeCaption As String
   sPassword = EfilingCommon.getmsgstate
   Sheet1.Unprotect Password:=sPassword
    Sheet1.Activate
    subProcCaption = "Validating PI"
    noOfProcessSub = 10
   EfilingCommon.MsgPISheet = ""
    'vbMessgaeCaption = "ITR 1: AY: 2023-24"
    vbMessgaeCaption = "ITR 1: AY: 2026-27"             'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
    'Ankita_24/12/2025=============
    If Not Validate_80CCC Then PIShtValidate = False
If Sheet1.Range("IncD.Section80CCD1B_SE").Value > 0 Or Sheet1.Range("IncD.Section80CCD_SE").Value > 0 Then
    If Not Validate_Pran Then PIShtValidate = False
End If
'    If Not Validate_80CCD_1 Then PIShtValidate = False
'    If Not Validate_80CCD_1b Then PIShtValidate = False

    '==============================

    'Ankita_05/02/2026====
'If CountrycodeMobileNo2 <> "" Then
'If PhoneNo2 = "" Then
'    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Secondary Mobile No. is Mandatory in Sheet Income Details " & Chr(13)
'End If
'End If

    If Trim(Range("BacValue").Value) = "" Or IsEmpty(Range("BacValue").Value) Then
        'PAG_E3
        'EfilingCommon.MsgPISheet = "* Are you opting for new tax regime u/s 115BAC? is mandatory in Sheet : Income Details" & Chr(13)
        'PAG_C4 2024-25 BINDU
        EfilingCommon.MsgPISheet = "*Selection of either of the options for Do you wish to exercise the option u/s 115BAC(6) of Opting out of new tax regime ?" & Chr(13)
    End If
    
    If Trim(Range("IncD_q1div").Value) = "" Or IsEmpty(Range("IncD_q1div").Value) Then
        EfilingCommon.MsgPISheet = "* Upto 15/6 (i) is mandatory in Sheet : Income Details" & Chr(13)
    End If
    
    If Trim(Range("IncD_q2div").Value) = "" Or IsEmpty(Range("IncD_q2div").Value) Then
        EfilingCommon.MsgPISheet = "* From 16/6 to 15/9 (ii) is mandatory in Sheet : Income Details" & Chr(13)
    End If
    
    If Trim(Range("IncD_q3div").Value) = "" Or IsEmpty(Range("IncD_q3div").Value) Then
        EfilingCommon.MsgPISheet = "* From 16/9 to 15/12 (iii) is mandatory in Sheet : Income Details" & Chr(13)
    End If
    
   If Trim(Range("IncD_q4div").Value) = "" Or IsEmpty(Range("IncD_q4div").Value) Then
        EfilingCommon.MsgPISheet = "* From 16/12 to 15/3 (iv) is mandatory in Sheet : Income Details" & Chr(13)
    End If
    
    If Trim(Range("IncD_q5div").Value) = "" Or IsEmpty(Range("IncD_q5div").Value) Then
        EfilingCommon.MsgPISheet = "* From 16/3 to 31/3 (v) is mandatory in Sheet : Income Details" & Chr(13)
    End If
    
    'Malli----
     

 

age = calculateAge23_24(Sheet1.Range("sheet1.DOB").Value)
'---end
If age <= 60 Then

'Chandru
  If (Range("IncD.Div").Value <> Range("IncD_q1div").Value + Range("IncD_q2div").Value + Range("IncD_q3div").Value + Range("IncD_q4div").Value + Range("IncD_q5div").Value) = 0 Then
  
    If Range("IncD.Div").Value = "" Then
       Range("IncD.Div").Value = 0
    End If
    End If
'*****

If Range("IncD.Div").Value <> (Range("IncD_q1div").Value + Range("IncD_q2div").Value + Range("IncD_q3div").Value + Range("IncD_q4div").Value + Range("IncD_q5div").Value) Then
          EfilingCommon.MsgPISheet = "Quarterly breakup of dividend income shall match with total dividend income"
        End If

End If

 
    
    '---------
    
'    If Trim(Range("OSIncreliefus89A").Value) > Trim(Range("OSIncomeNotified89A").Value) Then
'        EfilingCommon.MsgPISheet = "* Relied u/s 89A cannot be claimed more than income offered in pension accrued in a pension fund maintained in a notified country u/s 89A." & Chr(13)
'        Sheet1.Range("OSIncreliefus89A").Value = 0
'    End If
    
    'Ankita_16/01/2026============
'    If Range("Increliefus89A").Value > Range("IncomeNotified89A").Value Then
'        EfilingCommon.MsgPISheet = "* Income claimed for relief from taxation u/s 89A  cannot be claimed more than income entered in income from retirement benefit account maintained in a notified country u/s 89A" & Chr(13)
'        Sheet1.Range("Increliefus89A").Value = 0
'    End If
    
'    If Sheet1.Range("IncD_qOS1").Value <> Sheet1.Range("OSIncomeNotified89ATot").Value Then
'    fmsgbox "* Total of this field should be equal to amount entered in 'Income from retirement benefit account maintained in a notified country u/s 89A' -  'Income claimed for relief from taxation u/s 89A'"
'    PIShtValidate = False
'    CloseMsg
'    End If
    
    If Not ChkName Then EfilingCommon.MsgPISheet = "* LastName is mandatory in Sheet : Income Details" & Chr(13)
    If Not ChkPAN Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* PAN is mandatory in Sheet Income Details" & Chr(13)
   
   'Added by Shrutika(18/04/2025)NewDev  'Commented by Ankita_29/12/2025=====
'    If Not ChkPRAN Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""PRAN of the Taxpayer is mandatory""" & Chr(13)

'Ankita_29/12/2025=========
    If Not ChkAckNum Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter Acknowledgement number of Form 10BA filed for AY 2026-27 in VIA deductions" & Chr(13)
    
    'Added by Ayush_23/05/2025
   'If Sheet1.Range("bacvalue").Value = 2 Then  'Ankita_13/06/2025
'    If Sheet1.Range("Sheet1.AckNum").Value <> "" Then
   If Len(Sheet1.Range("Sheet1.AckNum").Value) > 0 Then
    If Sheet5.Range("bacValue").Value = 2 Then
      If Sheet1.Range("IncD.Section80GG").Value = "" Or Sheet1.Range("IncD.Section80GG").Value = 0 Then
       EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* 80GG amount shall be more than 0 when acknowledgement no. of form 10BA is provided" & Chr(13)
      End If
    End If
   End If 'Ankita_13/06/2025
    '--------------------------


   'Added by Shrutika(06/05/2025)Newdev
 
If Sheet1.Range("Sheet1.HRA").Value > Application.WorksheetFunction.RoundUp(Sheet1.Range("IncD.Allowances").Value / 3, 0) Then
 
EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Exemption of HRA u/s 10(13A) claimed cannot be more than limit prescribed " & Chr(13)
 
End If
    
   '-------------------------
    If Not ChkFlat Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Flat/Door/Block No. under primary address is mandatory in Sheet Income Details" & Chr(13)
    'If Not ChkStatus Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Status is mandatory in Sheet : Income Details" & Chr(13)
   UpdateProgressBar
    If Not ChkArea Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Area / Locality under primary address is mandatory in Sheet Income Details" & Chr(13)
     If Not ChkDOB Then
         EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Date of Birth is mandatory in Sheet Income Details" & Chr(13)
    End If
       
   'If Not ChkGender Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Sex is mandatory in Sheet : Income Details" & Chr(13)
    
    
    If Not ChkCity Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Town/City/District under primary address is mandatory in Sheet Income Details" & Chr(13)
    If Not ChkState Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* State under primary address is mandatory in Sheet Income Details" & Chr(13)
   'Ankita_14/01/2026=====
   'Ankita_17/03/2026=====
    If Sheet1.Range("Secondary_Address").Value = "No" Then
'If Sheet1.Range("sheet1.ResidenceNo1").Value <> "" Or Sheet1.Range("sheet1.LocalityOrArea1").Value <> "" Or Sheet1.Range("sheet1.CityOrTownOrDistrict1").Value <> "" Or Sheet1.Range("sheet1.StateCode2").Value <> "(Select)" Or Sheet1.Range("sheet1.Country1").Value <> "(Select)" Or Sheet1.Range("sheet1.PinCode1").Value <> "" Then
        If Not ChkCity1 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Town/City/District under secondary address is mandatory in Sheet Income Details""" & Chr(13)
        If Not ChkState1 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""State under secondary address is mandatory in Sheet Income Details""" & Chr(13)
        If Not ChkCountry1 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Country under secondary address is mandatory in Part A General Information""" & Chr(13)
        If Not ChkFlat1 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Flat/Door/Block No. under secondary address is mandatory in Sheet Income Details""" & Chr(13)
        If Not ChkArea1 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Area / Locality under secondary address is mandatory in Sheet Income Details""" & Chr(13)
        If Not ValidatePinCode1 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Pin code under secondary address is mandatory in Sheet Income Details""" & Chr(13)
        If Not ChkZipcode1 Then PIShtValidate = False
End If
'    End If

    'If Not ChkCountry Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Country is mandatory in Sheet Income Details" & Chr(13)
    'AY_2024_25 Change 'Malli
    If Not ChkCountry Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Country under primary address is mandatory in Part A General Information" & Chr(13)
    
    UpdateProgressBar
     If Not ValidatePinCode Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Pin code is mandatory under primary address in Sheet Income Details" & Chr(13)

   ' If Not ChkPincode Then PIShtValidate = False
        'EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Pincode is mandatory in Sheet : Income Details " & Chr(13)
    'Else
    '   If Not (mIncmDtls.ValidatePinCode()) Then
                 'Sheet1.Range("sheet1.PinCode").Value = ""
    'End If
    'End If
    
    If Not ChkZipcode Then PIShtValidate = False
    
    'Ankita_17/03/2026=====
'    If Sheet1.Range("Secondary_Address").Value = "Yes" Then
'    If Not ChkZipcode1 Then PIShtValidate = False  'Ankita_28/01/2026
'    End If
    
    If Mid(Sheet1.Range("sheet1.ReturnFileSec"), 1, 7) = "139(8A)" Then
    Else
    If Not ChkSeventhProvisoFlag Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Selection of either of the options for 'Are you filing return of income under Seventh proviso to section 139(1) but otherwise not required to furnish return of income?' is mandatory" & Chr(13)
'Change-22.11.2022.102.16K
'    If Not ChkDepositAmountFlag Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option have you deposited amount or aggregate of amounts exceeding Rs. 1 Crore in one or more current account during the previous year in Sheet Income Details" & Chr(13)
'---end

    If Not ChkAggrigateAmountFlag Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option have you incurred expenditure of an amount or aggregate of amount exceeding Rs. 2 lakhs for travel to a foreign country for yourself or for any other person in Sheet Income Details" & Chr(13)
    
    If Not ChkAggrigateAmountFlag1 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option have you incurred expenditure of amount or aggregate of amount exceeding Rs. 1 lakh on consumption of electricity during the previous year in Sheet Income Details" & Chr(13)
    End If
     
     If Mid(Sheet1.Range("sheet1.SeventhProvisoFlag").Value, 1, 1) = "Y" Then
'Change-22.11.2022.102.16D
'     If Mid(Sheet1.Range("Sheet1.DepositAmountFlag").Value, 1, 1) = "N" And Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag").Value, 1, 1) = "N" And Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag1").Value, 1, 1) = "N" And Sheet1.Range("clauseiv7provisio139iFlg").Value = "No" Then
     If Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag").Value, 1, 1) = "N" And Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag1").Value, 1, 1) = "N" And Sheet1.Range("clauseiv7provisio139iFlg").Value = "No" Then
'---end

        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* You have selected ""Yes"" for ""Are you filing return of income under Seventh proviso to section 139(1) but otherwise not required to furnish return of income?"", hence please enter respective amount in Sheet Income details." & Chr(13)
     End If
'Change-22.11.2022.102.16E
'    If Mid(Sheet1.Range("Sheet1.DepositAmountFlag").Value, 1, 1) = "Y" Then
'
'        If Range("Sheet1.DepositAmount").Value = "" Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* You have selected ""Yes"" for ""Have you deposited amount or aggregate of amounts exceeding Rs. 1 Crore in one or more current account during the previous year?"", hence please enter the amount in Sheet Income details." & Chr(13)
'        Else
'            If Range("Sheet1.DepositAmount").Value < 10000000 Or Range("Sheet1.DepositAmount").Value = 10000000 Then
'                EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter valid amount of deposit in one or more current account during the previous year" & Chr(13)
'            End If
'
'        End If
'
'    End If
'---end

    
    If Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag").Value, 1, 1) = "Y" Then
    
     If Range("Sheet1.AggrigateAmount").Value = "" Then
       EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* You have selected ""Yes"" for ""Have you incurred expenditure of an amount or aggregate of amount exceeding Rs. 2 lakhs for travel to a foreign country for yourself or for any other person?"", hence please enter the amount in Sheet Income details." & Chr(13)
    Else
    If Range("Sheet1.AggrigateAmount").Value < 200000 Or Range("Sheet1.AggrigateAmount").Value = 200000 Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter valid  expenditure incurred for travel to a foreign country for yourself or for any other person" & Chr(13)
    End If
    End If
    
    End If
    
    If Mid(Sheet1.Range("Sheet1.AggrigateAmountFlag1").Value, 1, 1) = "Y" Then
    
    If Range("Sheet1.AggrigateAmount1").Value = "" Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* You have selected ""Yes"" for ""Have you incurred expenditure of amount or aggregate of amount exceeding Rs. 1 lakh on consumption of electricity during the previous year? "", hence please enter the amount in Sheet Income details." & Chr(13)
    Else
    If Range("Sheet1.AggrigateAmount1").Value < 100000 Or Range("Sheet1.AggrigateAmount1").Value = 100000 Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter valid amount of expenditure incurred on consumption of electricity during the previous year" & Chr(13)
    End If
    End If
    End If
    
    End If
    
    
    
      UpdateProgressBar
'    If Not ChkEmail Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Email Address is mandatory in Sheet Income Details" & Chr(13)
'Ankita_12/01/2026==================
        If Not ChkEmail Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Primary Email ID of the taxpayer is mandatory in Sheet Income Details""" & Chr(13)

    ' If Not ValidatePhoneNo Then PIShtValidate = False
'Ankita_12/01/2026==================
If Mid(Sheet1.Range("sheet1.RepAssessee").Value, 1, 1) = "Yes" Then
    If Not ChkEmail1 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Email Address of representative assessee is mandatory in Sheet Income Details.""" & Chr(13)
End If
'    If Not ChkEmail2 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Secondary Email ID is mandatory in Sheet Income Details.""" & Chr(13)
'===================================
'    If Not ChkMobileNo Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Mobile Number is mandatory in Sheet Income Details" & Chr(13)
'    If Not ChkCountrycode Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Country code is mandatory in Sheet Income Details" & Chr(13)
    
    'Added by aavula
     If Not ChkMobileNo Then PIShtValidate = False
     If Not ChkCountrycode Then PIShtValidate = False
    'Ankita_12/01/2026===========================
    If Not ChkMobileNo1 Then PIShtValidate = False
'   If Not ChkMobileNo2 Then PIShtValidate = False
    If Not ChkCountrycode1 Then PIShtValidate = False
    If Not ChkCountrycode2 Then PIShtValidate = False

'Ankita_05/02/2026=======
PhoneNo3 = Sheet1.Range("sheet1.Mobileno1").Value
CountrycodeMobileNo3 = Sheet1.Range("sheet1.MobileCountryCode1").Value
If PhoneNo3 <> "" Then
If CountrycodeMobileNo3 = "" Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*""Country Code in Secondary Mobile No. in Sheet: Personal Information is mandatory.""" & Chr(13)
End If
End If
If CountrycodeMobileNo3 <> "" Then
If PhoneNo3 = "" Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Secondary Mobile No. is Mandatory  in Sheet Income Details.""" & Chr(13)
End If
End If
 
     '===========================================
    If Not ChkEmpCategory Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Nature of Employment is mandatory in Sheet Income Details" & Chr(13)
    'If Not ChkGovernedbyPortugese Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Governed by Portuguese is mandatory in Sheet : Income Details" & Chr(13)
    UpdateProgressBar
   'If Not ChkReturnFileSec Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please Select ""Filed u/s"" or ""Filed in response to notice u/s  in Sheet Income Details" & Chr(13)
        'Ankita
    If Not ChkReturnFileSec Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Filing section is mandatory" & Chr(13)

  'Added by Shrutika(21/04/2025)NewDev
If (Sheet1.Range("SELECT80DDS").Value <> "(Select)" And Sheet1.Range("SELECT80DDS").Value <> "") And Sheet1.Range("IncD.Section80DDB") > 0 Then
If Sheet1.Range("Sheet1.Specified_Disease") = "(Select)" Or Sheet1.Range("Sheet1.Specified_Disease") = "" Then
EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Name of Specified disease is mandatory""" & Chr(13)
End If
End If
'----------------------------------------------------------


  'Ankita_05/02/2026
      'Ankita_30/03/2026=====

  If Sheet1.Range("Secondary_Address").Value = "(Select)" Or Sheet1.Range("Secondary_Address").Value = "" Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Is the secondary address same as primary address? is mandatory." & Chr(13)
  End If

    If Not ValidateReturnFileSection Then
    PIShtValidate = False
    CloseMsg
    End If
    UpdateProgressBar
     If Not ValidateAadharNumber Then
       'EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter the valid Aadhaar number/Aadhaar Enrolment Id in Sheet Income Details " & Chr(13)
       'Ankita
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please enter the Aadhaar number in valid format " & Chr(13)
Else
     End If
    
    If Not ValidatePropertyType Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Select the Type of House Property in Income Details." & Chr(13)
   
   
   'Ankita_21/01/2026
'   If (Range("IncD.TypeOfHP").Value = "Self Occupied") Then
'ElseIf (Range("IncD.TypeOfHP").Value = "Let Out") Or (Range("IncD.TypeOfHP").Value = "Deemed Let Out") Then
'       If (Mid(Range("IncD.TypeOfHP").Value, 1, 3) = "Dee") Or (Mid(Range("IncD.TypeOfHP").Value, 1, 3) = "Let") Then
'        If Range("IncD.GrossRentRecieved").Value = "" Or Range("IncD.GrossRentRecieved").Value = 0 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Gross rent received/ receivable/ letable value during the year should be greater than zero in Sheet Income Details" & Chr(13)
'        End If
'   End If
        
        
'        If (Range("IncD.GrossRentRecieved").Value = "" Or Range("IncD.GrossRentRecieved").Value) = 0 Then
'            If Range("IncD.TaxPaidLocalAuthorities").Value > 0 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Tax paid to local authorities can be claimed only if income from house property is declared in Income Details." & Chr(13)
'            End If
'        End If
        
'End If

     If Not ValidateSIncm Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* For a Self occupied House Property,Interest payable on borrowed capital value cannot exceed Rs. 2,00,000" & Chr(13)
     UpdateProgressBar
    If Not ValidateCountryState Then
  If Range("sheet1.StateCode1").Value = "(Select)" Then
     Else
     EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Country should be India in Sheet Income Details   " & Chr(13)
     End If
    Else
    End If
    
    UpdateProgressBar
'    If ValidatePANSPouseWithPorg Then
'    Else
'    If Range("sheet1.PANOFSPOUSE").Value <> "" Then
'
'    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Since PAN of spouse has been entered, please check the selection of drop down option in the field " & Chr(34) & "Whether governed by Portuguese Civil Code?" & Chr(34) & Chr(13)
'
'    Else
'
'    End If
   ' End If
    UpdateProgressBar
    
    
    
    If Len(Range("IncD.Allowances").Value) > 14 Then
      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Allowances value should not exceed 14 digits in Income Details" & Chr(13)
  End If
  
  If Len(Range("IncD.Perquisites").Value) > 14 Then
      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Perquisites value should not exceed 14 digits in Income Details" & Chr(13)
  End If
  
  If Len(Range("IncD.Profits").Value) > 14 Then
      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Profits salry should not exceed 14 digits in Income Details" & Chr(13)
  End If
  
  If Len(Range("IncD.Deduction16").Value) > 14 Then
      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Deduction u/s 16  should not exceed 14 digits in Income Details" & Chr(13)
  End If
  
  If Len(Range("IncD.Deduction16ia").Value) > 14 Then
      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Deduction u/s 16(ia)  should not exceed 14 digits in Income Details" & Chr(13)
  End If

  
  
  If Len(Range("IncD.TotalHeadSalaries").Value) > 14 Then
      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Total head salary should not exceed 14 digits in Income Details" & Chr(13)
  End If
  
  
'  If (Range("IncD.Deduction16").Value) > 7500 Then
'      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Deduction u/s 16 should not exceed 7,500." & Chr(13)
'  End If

'Added by Ankita on 06/12/2024
If Sheet5.Range("BacValue").Value = 2 Then
 If (Range("IncD.Deduction16ia").Value) > 50000 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Maximum Deduction u/s 16(ia) is  50,000/- only in Income Details" & Chr(13)
End If
End If

'Added by Ankita on 05/12/2024
If Sheet5.Range("BacValue").Value = 1 Then
  If (Range("IncD.Deduction16ia").Value) > 75000 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Maximum Deduction u/s 16(ia) is  75,000/- only in Income Details" & Chr(13)
  End If
End If

'=========================================

'If ((Range("IncD.Deduction16ia").Value) <= 40000 And (Range("IncD.Deduction16ia").Value > (Range("IncD.IncomeFromSal").Value + Range("IncD.Allowances").Value + Range("IncD.Perquisites").Value + Range("IncD.Profits").Value))) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Maximum deduction u/s 16(ia) is '40000' or sum of B1 [(i)+(ii)+(iii)+(iv)] in Part B whichever is lower." & Chr(13)
'End If
' If (Range("IncD.Deduction16ia").Value > (Range("IncD.IncomeFromSal").Value + Range("IncD.Allowances").Value + Range("IncD.Perquisites").Value + Range("IncD.Profits").Value)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Maximum deduction u/s 16(ia) is '40000' or sum of B1 [(i)+(ii)+(iii)+(iv)] in Part B whichever is lower." & Chr(13)
'End If




If empcat = "Central Government" Or empcat = "State Government" Or empcat = "Public Sector Undertaking" Then
        If (Range("IncD.Deduction16").Value) > 10000 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Deduction of Entertainment allowance u/s 16(ii) should not exceed 10,000 in Income Details." & Chr(13)
           ' Range("IncD.Deduction16").Value = ""
        End If
      End If
        
        
        If empcat = "Others" Then
        If (Range("IncD.Deduction16").Value) > 5000 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* For employee category " & """Others""" & " maximum amount that can be claimed for the Deduction of Entertainment allowance u/s 16(ii) is 5000 in Income Details" & Chr(13)
           ' Range("IncD.Deduction16").Value = ""
        End If
      End If
        
        If empcat = "Not Applicable (eg. Family pension etc)" Then
        If (Range("IncD.Deduction16").Value) > 0 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Deduction of Entertainment allowance u/s 16(ii) is not applicable for employer category" & " ""Not Applicable"" in Income Details." & Chr(13)
            'Range("IncD.Deduction16").Value = ""
        End If
      End If



  'Ankita_21/01/2026
'  If Len(Range("IncD.GrossRentRecieved").Value) > 14 Then
'      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Gross rent recieved should not exceed 14 digits in Income Details." & Chr(13)
'  End If
  
'  If Len(Range("IncD.TaxPaidLocalAuthorities").Value) > 14 Then
'      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Tax paid local authorities should not exceed 14 digits in Income Details." & Chr(13)
'  End If
  
'  If Len(Range("IncD.AnnualValue").Value) > 14 Then
'      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Annual value should not exceed 14 digits in Income Details." & Chr(13)
'  End If
  
'  If Len(Range("IncD.StandardDeduction").Value) > 14 Then
'      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Standard deduction  should not exceed 14 digits in Income Details." & Chr(13)
'  End If
  
'  If Len(Range("IncD.InterestBorrowedCapital").Value) > 14 Then
'      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Interest borrowed capital should not exceed 14 digits in Income Details." & Chr(13)
'  End If
  
  If Len(Range("IncD.IncomeHeadHouseProperty").Value) > 14 Then
      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Income Head House Property  should not exceed 14 digits in Income Details." & Chr(13)
  End If
  
'  If Len(Range("IncD.IncomeHeadHousePropertyINTER").Value) > 14 Then
'      EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Income Head House Property negative should not exceed 14 digits" & Chr(13)
'  End If
  
    'If ValidateAadharNumberWith Then
    'Else
    '    If Range("Sheet1.Aadhaar").Value <> "" Then
    '       EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Since Aadhaar Number has been entered, please check the selection of drop down option in the field " & Chr(34) & "Whether you have Aadhaar Number?" & Chr(34) & Chr(13)
    '   End If
    'End If
    
    UpdateProgressBar
      
      'Ankita 02/01/2025
      'Correctly updated by Bindu on 4th Feb 2025 as per DE V3
   'If (Sheet1.Range("IncD.TotalIncome").Value > 5000000) Then
   If ((Sheet1.Range("IncD.TotalIncome_New").Value - Sheet1.Range("IncD.CG_LTCG")) > 5000000) Then
   
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ITR 1 is for individuals being a resident other than not ordinarily resident having Income from Salaries, one house property, other sources (Interest etc.), agricultural income upto Rs.5 thousand and having total income upto Rs.50 lakh. Please file other ITR"
         EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""ITR 1 is for individuals being a resident (other than not ordinarily resident) having total income upto Rs.50 lakh, having Income from Salaries, one house property, interest income, Family pension income etc. and agricultural income upto Rs.5 thousand. Please file another ITR as applicable""" & Chr(13)

End If
'  Change 01, Satya, 24.01.2023
'   If (Sheet1.Range("IncD.TotalIncome").Value < 250000) Then
   If (Sheet1.Range("IncD.TotalIncome").Value < 250000) And Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
'End Change====
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* If total income of the assessee is less than or equal to 2.5 lakhs, then user is not allowed to file ITR-U" & Chr(13)
    End If
    
  'Ankita_03/06/2025
'    If Range("Less_allowance").Value > Range("IncD.IncomeFromSal").Value Then
'    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Total of allowances shall not exceed Gross salary in Schedule Income details" & Chr(13)
'    End If
  
  
  'Ankita_03/06/2025
      If Range("Less_allowance").Value > Range("TotAllowances_17abc").Value Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Total of allowances shall not exceed gross salary at sl. no. ia+ib+ic in Schedule Income details" & Chr(13)
    End If

  
     If Len(Range("sheet1.NoticeDate").Value) > 0 Then
        If Not FormatNCheckDate(Range("sheet1.NoticeDate").Value) Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter a date in Date of notice/order field in dd/mm/yyyy format" & Chr(13)
        End If
'Change-22.11.2022.102.23A
'        If EfilingCommon.checkFirstDateBefore(Range("sheet1.NoticeDate").Value, "31/03/2022") Then
'                EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*  Date of Notice/Order in Sheet Income Details cannot be prior to 01/04/2022 " & Chr(13)
'        End If
        Noticedate23_24
'---end
    End If
   

'    If (Sheet2.Range("TDS26QB.Sum").Value > 0) Then
'            If Sheet1.Range("IncD.TypeOfHP").Value = "Let Out" Then
'                If Not Sheet1.Range("IncD.GrossRentRecieved").Value > 0 Then
'                    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "If TDS3(16C) is filled, then please ensure that the 'Type of house property' should be 'Let out'  & 'Gross rent received /receivable /letable value' should be greater than zero" & Chr(13)
'
'                End If
'            Else
'                    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "If TDS3(16C) is filled, then please ensure that the 'Type of house property' should be 'Let out'  & 'Gross rent received /receivable /letable value' should be greater than zero" & Chr(13)
'
'            End If
'        End If
    
    
    If Len(Range("sheet1.OrigRetFiledDate").Value) > 0 Then
        If Not FormatNCheckDate(Range("sheet1.OrigRetFiledDate").Value) Then
           EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Date of Filing in Sheet Income Details must be a valid dd/mm/yyyy format" & Chr(13)
        End If
        
        'PAG_E10
        'If EfilingCommon.checkFirstDateBefore(Range("sheet1.OrigRetFiledDate").Value, "31/03/2023") Then
         '    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Date of Filing in Sheet Income Details  should be on or after 01/04/2023 for A.Y 2023-24" & Chr(13)
       'PAG_C10 2024-25 Bindu
       'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
       If EfilingCommon.checkFirstDateBefore(Range("sheet1.OrigRetFiledDate").Value, "31/03/2025") Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Date of Filing in Sheet Income Details should be on or after 01/04/2025 for A.Y 2025-26" & Chr(13)
       End If
    End If
    

    
    
    
'    If Not ValidateSelect80D Then PIShtValidate = False
'    If Not ValidateSelect80DB Then PIShtValidate = False
'    If Not ValidateSelect80DC Then PIShtValidate = False
'    If Not ValidateSystem80D Then PIShtValidate = False
    'If Not ValidateSelect80DD Then PIShtValidate = False
    If Not ValidateSELECT80DDS Then PIShtValidate = False
    'If Not ValidateSelect80U Then PIShtValidate = False
    
    Sheet1.Protect Password:=sPassword

'dpk1

If Sheet1.Range("clauseiv7provisio139iFlg").Value <> "No" And Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
    
    If Sheet1.Range("clauseiv7provisio139iFlg").Value = "Yes" Then
'Change-22.11.2022.102.19D
'If Sheet1.Range("clauseiv7provisio139iFlg_1").Value = "No" And Sheet1.Range("clauseiv7provisio139iFlg_2").Value = "No" And Sheet1.Range("clauseiv7provisio139iFlg_3").Value = "No" And Sheet1.Range("clauseiv7provisio139iFlg_4").Value = "No" Then
    If Sheet1.Range("clauseiv7provisio139iFlg_3").Value = "No" And Sheet1.Range("clauseiv7provisio139iFlg_4").Value = "No" Then
'---end
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* You have selected ""Yes"" for ""Are you required to file a return as per other conditions  prescribed under clause (iv) of seventh proviso to section 139(1)"", hence please enter respective amount in Sheet Income details." & Chr(13)
    End If
'Change-22.11.2022.102.19E
'    If Sheet1.Range("clauseiv7provisio139iFlg_1").Value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_1").Value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").Value = "Yes" Then
'    fmsgbox "*Selection of either of the options for 'total sales, turnover or gross receipts, as the case may be, of the person in the business exceeds sixty lakh rupees during the previous year; or ' is mandatory in Sheet Income Details" & Chr(13)
'    PIShtValidate = False
'    CloseMsg
'    End If
'
'    If (Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_1").Value = "") Then
'    fmsgbox "*Please enter amount of Total sales, turnover or gross receipts etc in Income Details"
'    PIShtValidate = False
'    CloseMsg
'    End If
'
'        If Sheet1.Range("clauseiv7provisio139iFlg_2").Value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_2").Value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").Value = "Yes" Then
'    fmsgbox "*Selection of either of the options for 'the total gross receipts of the person in profession exceeds ten lakh rupees during the previous year; or' is mandatory in Sheet Income Details" & Chr(13)
'    PIShtValidate = False
'    CloseMsg
'    End If
'
'    If (Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_2").Value = "") Then
'    fmsgbox "*Please enter amount of total gross receipts in Income Details"
'    PIShtValidate = False
'    CloseMsg
'    End If
'---end
'Ankita_30/05/2025
    If Sheet1.Range("clauseiv7provisio139iFlg_3").Value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_3").Value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").Value = "Yes" Then
'    fmsgbox "*Selection of either of the options for 'the aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is twenty-five thousand rupees or more; or' is mandatory in Sheet Income Details" & Chr(13)
    fmsgbox "* ""Please select atleast one of the conditions prescribed under clause (iv) of seventh proviso to section 139(1)""" & Chr(13)
    PIShtValidate = False
    CloseMsg
    End If


    If (Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_3").Value = "") Then
    fmsgbox "*Please enter amount of tax deducted at source and tax collected at source in Income Details"
    PIShtValidate = False
    CloseMsg
    End If
    
    If Sheet1.Range("clauseiv7provisio139iFlg_4").Value = "" Or Sheet1.Range("clauseiv7provisio139iFlg_4").Value = "(Select)" And Sheet1.Range("clauseiv7provisio139iFlg").Value = "Yes" Then
    fmsgbox "*Selection of either of the options for 'if his total deposits in a savings bank account is fifty lakh rupees or more, in the previous year.' is mandatory in Sheet Income Details" & Chr(13)
    PIShtValidate = False
    CloseMsg
    End If


    If (Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes") And (Sheet1.Range("clauseiv7provisio139iAmount_4").Value = "") Then
    fmsgbox "*Please enter amount of total deposits in Income Details"
    PIShtValidate = False
    CloseMsg
    End If
    
    Else
    
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* If the field is blank please select 'Yes' or 'No' to the question 'Are you required to file a return as per other conditions  prescribed under clause (iv) of seventh proviso to section 139(1)'" & Chr(13)
        
    End If
End If

If Mid(Sheet1.Range("clauseiv7provisio139iFlg").Value, 1, 1) = "Y" And Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
'Change-22.11.2022.102.19F
'    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_1").Value, 1, 1) = "Y" And Range("clauseiv7provisio139iAmount_1").Value <> "" Then
'    If Range("clauseiv7provisio139iAmount_1").Value < 6000000 Or Range("clauseiv7provisio139iAmount_1").Value = 6000000 Then
'    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select 'No' if total sales, turnover or gross receipts, as the case may be, of the person in the business is less than sixty lakh rupees during the previous year" & Chr(13)
'    End If
'    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_1").Value = "Yes" And Range("clauseiv7provisio139iAmount_1").Value = "" Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please select 'No' if total sales, turnover or gross receipts, as the case may be, of the person in the business is less than sixty lakh rupees during the previous year" & Chr(13)
'    End If
'
'    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_2").Value, 1, 1) = "Y" Then
'    If Range("clauseiv7provisio139iAmount_2").Value < 1000000 Or Range("clauseiv7provisio139iAmount_2").Value = 1000000 Then
'    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select 'No' if the total gross receipts of the person in profession is less than ten lakh rupees during the previous year" & Chr(13)
'    End If
'    End If
'
'     If Sheet1.Range("clauseiv7provisio139iFlg_2").Value = "Yes" And Range("clauseiv7provisio139iAmount_2").Value = "" Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please select 'No' if the total gross receipts of the person in profession is less than ten lakh rupees during the previous year" & Chr(13)
'    End If
'---end

    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_3").Value, 1, 1) = "Y" Then
    If Range("clauseiv7provisio139iAmount_3").Value < 25000 Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees " & Chr(13)
    End If
    End If
    
     If Sheet1.Range("clauseiv7provisio139iFlg_3").Value = "Yes" And Range("clauseiv7provisio139iAmount_3").Value = "" Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please select 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees " & Chr(13)
    End If
    
    If Mid(Sheet1.Range("clauseiv7provisio139iFlg_4").Value, 1, 1) = "Y" Then
    If Range("clauseiv7provisio139iAmount_4").Value < 5000000 Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select 'No' if total deposits in a savings bank account is less than fifty lakh rupees in the previous year" & Chr(13)
    End If
    End If
    
     If Sheet1.Range("clauseiv7provisio139iFlg_4").Value = "Yes" And Range("clauseiv7provisio139iAmount_4").Value = "" Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*Please select 'No' if total deposits in a savings bank account is less than fifty lakh rupees in the previous year" & Chr(13)
    End If
End If

'Ankita_12/01/2026_V0.2===================
'If Representativeassesseeflg = "" <> Representativeassesseeflg = "(Select)" Then
Representativeassesseeflg = Range("sheet1.RepAssessee").Value
If Representativeassesseeflg = "" Or Representativeassesseeflg = "(Select)" Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select whether return is filed by Representative assessee in Sheet Income details" & Chr(13)
End If

If (Sheet1.Range("sheet1.RepAssessee").Value = "" Or Sheet1.Range("sheet1.RepAssessee").Value = "(Select)") And Sheet3.Range("Ver.capacity").Value = "Representative" Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Details of representative assessee is mandatory as the return is being filed by representative in Sheet Tax Paid and Verification" & Chr(13)
End If

If Sheet1.Range("sheet1.RepAssessee").Value = "Yes" And Sheet1.Range("sheet1.NameRepAssessee").Value = "" Then
    If NameRepAssessee = "" Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter the Name of representative assessee in Sheet Income details" & Chr(13)
    End If
End If
If Sheet1.Range("sheet1.RepAssessee").Value = "Yes" And Sheet1.Range("sheet1.EmailRepAssessee").Value = "" Then
    If EmailRepAssessee = "" Then
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Email Address is mandatory in Sheet Income details" & Chr(13)
    End If
End If
If Sheet1.Range("sheet1.RepAssessee").Value = "Yes" And Sheet1.Range("sheet1.ContactRepAssessee").Value = "" Then
    If ContactRepAssessee = "" Then
'       EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Contact number is mandatory in Sheet Income details" & Chr(13)
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Contact number of representative assessee is mandatory in schedule Part A General""" & Chr(13)
 End If
End If

'=========================================

'dpk1
    
    'Updated from Range("IncD.GrossTotIncome") to Range("IncD.GrossTotIncome_New") as per DE V3 by Bindu on 4th Feb 2025
   ' If Len(Range("IncD.GrossTotIncome").Value) > 14 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Gross Total income in Sheet Income Details should not be greater than 14 digits" & Chr(13)
    If Len(Range("IncD.GrossTotIncome_New").Value) > 14 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Gross Total income in Sheet Income Details should not be greater than 14 digits" & Chr(13)
    
    If Len(Range("IncD.TotalChapVIADeductions_Input").Value) > 14 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Deductions in Sheet Income Details should not be greater than 14 digits" & Chr(13)
    If Len(Range("IncD.TotalIntrstPay").Value) > 14 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Total Interest in Sheet Income Details should not be greater than 14 digits" & Chr(13)
    If Len(Range("IncD.TotTaxPlusIntrstPay").Value) > 14 Then EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Total Tax and Interest Payable in Sheet Income Details should not be greater than 14 digits" & Chr(13)
    UpdateProgressBar
    ValidateOthersEI
    ValidateOthersEI1
    ValidateOthersEI2
    

    
    '=====================
    
    If EfilingCommon.MsgPISheet = "" Then
    
    Else
        'MsgBox EfilingCommon.MsgPISheet, vbOKOnly, "Error"
        fmsgbox (EfilingCommon.MsgPISheet)
         Sheet1.Activate
        PIShtValidate = False
        CloseMsg
    End If

End Function
Function ValidateAadharNumberWith() As Boolean
On Error Resume Next
    ValidateAadharNumberWith = True
    If Range("sheet1.AadhaarYN").Value = "No" Then
        ValidateAadharNumberWith = False
        Exit Function
    End If
End Function
Function ValidateTaxStatus() As Boolean
On Error Resume Next

   ValidateTaxStatus = True
   If Range("sheet1.TaxStatus").Value = "(Select)" Then
            ValidateTaxStatus = False
            Exit Function
        
    End If
End Function

Function ValidateCountryState() As Boolean
On Error Resume Next
  
   ValidateCountryState = True
End Function
Function ValidateAadharYN() As Boolean
On Error Resume Next

   ValidateAadharYN = True
    If Range("sheet1.AadhaarYN").Value = "(Select)" Then
        Exit Function
    End If
End Function
Function ValidateAadharNumber() As Boolean
On Error Resume Next

   ValidateAadharNumber = True
'   Dim AadharNumber, AadharEnrolNumber As Variant
   Dim AadharNumber As Variant

   AadharNumber = Sheet1.Range("Sheet1.Aadhaar").Value
'   Enhancement
'   AadharEnrolNumber = Sheet1.Range("Sheet1.AadhaarEnrol").Value
   
    If Trim(AadharNumber) <> "" Then
    
        If Not IsNumeric(AadharNumber) Then
            'errmsgID = "is invalid"
            ValidateAadharNumber = False
            Exit Function
        End If
        
        If Len(AadharNumber) <> 12 Then
            ValidateAadharNumber = False
            Exit Function
        End If
        
        If AadharNumber = "000000000000" Then
            'errmsgID = "is invalid"
            ValidateAadharNumber = False
            Exit Function
        End If
        
        If AadharNumber = "111111111111" Then
            'errmsgID = "is invalid"
            ValidateAadharNumber = False
            Exit Function
        End If
        
    End If

'    If Trim(AadharEnrolNumber) <> "" Then
'
'
'        If Not IsNumeric(AadharEnrolNumber) Then
'            'errmsgID = "is invalid"
'            ValidateAadharNumber = False
'            Exit Function
'        End If
'
'        If Len(AadharEnrolNumber) <> 28 Then
'            ValidateAadharNumber = False
'            Exit Function
'        End If
'        Enhancement
'        If AadharEnrolNumber = "0000000000000000000000000000" Then
'            'errmsgID = "is invalid"
'            ValidateAadharNumber = False
'            Exit Function
'        End If
        
'        If AadharEnrolNumber = "1111111111111111111111111111" Then
'            'errmsgID = "is invalid"
'            ValidateAadharNumber = False
'            Exit Function
'        End If
        
        'If Not ValidateAadharEnrol(AadharEnrolNumber) Then
        '    'errmsgID = "is invalid"
        '    ValidateAadharNumber = False
        '    Exit Function
        'End If
'    End If

End Function
Function ValidatePropertyType() As Boolean
On Error Resume Next

   ValidatePropertyType = True
   'If Not IIf(Range("IncD.IncomeFromHP").Value = "", 0, Range("IncD.IncomeFromHP").Value) = 0 Then
  'Ankita_21/01/2026
'      If (Range("IncD.GrossRentRecieved").Value <> "" Or Range("IncD.TaxPaidLocalAuthorities").Value <> "" Or Range("IncD.InterestBorrowedCapital").Value <> "") And (Range("IncD.GrossRentRecieved").Value <> 0 Or Range("IncD.TaxPaidLocalAuthorities").Value <> 0 Or Range("IncD.InterestBorrowedCapital").Value <> 0) Then
'      If Range("IncD.TypeOfHP").Value = "(Select)" Or Range("IncD.TypeOfHP").Value = "" Then
'            ValidatePropertyType = False
'            Exit Function
'      End If
'      End If
      
      
   ' Else
    'Range("IncD.IncomeFromHP").Value = 0

    'End If
End Function
Function ValidateReturnFileSection() As Boolean
On Error Resume Next
    ValidateReturnFileSection = True
 Dim vbMessgaeCaption As String
 
    If (Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = "18") Then
    If Not (ValidateReceiptNo And ValidateDOF) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Receipt Number for filed u/s 139(9) And Date of filing of Original Return for return filed u/s 139(9) is mandatory in Sheet Income Details" & Chr(13)
            Exit Function
    End If
    
    'Ankita_12/12/2024
     If Len(Sheet1.Range("sheet1.OrigRetFiledDate").Value) > 0 Then
          If Not CheckDateddmmyyyy(Sheet1.Range("sheet1.OrigRetFiledDate").Value) Then
          EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Date of filing of Original return must be a valid dd/mm/yyyy format""" & Chr(13)
          Exit Function
          End If
        End If
    '========================
        
    If (Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = "18") Then
    If Not (ValidateNoticeNo) Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Unique number / DIN  is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b) is mandatory in Sheet Income Details" & Chr(13)
    Exit Function
    End If

    If Not (ValidateNoticeDate) Then
    EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Date of Notice/Order for filed u/s 139(9)/142(1)/148/153C or 119(2)(b) is mandatory in Sheet Income Details" & Chr(13)
    Exit Function
    End If

    End If

    If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "18" Then
      If Not (ValidateDOF And ValidateReceiptNo And ValidateNoticeNo And ValidateNoticeDate) Then
          ValidateReturnFileSection = False
        Exit Function

    End If
   End If

End If



    If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "17" Then

        If Not (ValidateReceiptNo And ValidateDOF) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Receipt Number And Date of Filing Original Return is mandatory in Sheet Income Details" & Chr(13)
            Exit Function
        Else

         If (ValidateNoticeNo Or ValidateNoticeDate) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Unique number for filed u/s139(9)/142(1)/148/153C or 119(2)(b) Or Date of Notice/Order for return filed u/s 139(9)/142(1)/148/153C or 119(2)(b)  in Sheet Income Details are not mandatory" & Chr(13)
         End If
        End If
    End If

      If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "18" Then

        If Not (ValidateReceiptNo And ValidateDOF) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Receipt Number for filed u/s 139(9) And Date of filing of Original Return for return filed u/s 139(9) is mandatory in Sheet Income Details" & Chr(13)
            Exit Function
     End If
        End If

        If Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "13" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "14" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "15" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "16" Or Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "20" Then

        If Not (ValidateNoticeNo) Then
'        Change.28.02.2023.102.IDS.52
        'EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Unique number / DIN  is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148"" or Filed u/s is ""119(2)(b)-after condonation of delay is mandatory in Sheet Income Details" & Chr(13)
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Unique number / DIN  is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b) in Sheet Income Details" & Chr(13)
'        End Change
        Exit Function
        End If

        If Not ValidateNoticeDate Then
'        Change.28.02.2023.102.IDS.51
        'EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Date of notice or order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b)-after condonation of delay is mandatory in Sheet Income Details" & Chr(13)
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Date of notice or order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b) in Sheet Income Details" & Chr(13)
'        Changed by Ankita on 14/12/2024
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Date of notice or order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s is " & " ""139(9A)""" & Chr(13)
'        Changed by Ankita on 21/01/2025
         EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Date of notice or order is mandatory for ""Filed in response to notice u/s"" is  ""139(9)/142(1)/148/153C"" or Filed u/s is ""119(2)(b)""" & Chr(13)

'       End ChangeIDS
        Exit Function

        End If

        End If


End Function


Function ValidateStateAndCountry() As Boolean
On Error Resume Next
    ValidateStateAndCountry = True
 Dim vbMessgaeCaption As String
   End Function
   
   Function ValidateSTD() As Boolean
   On Error Resume Next
    ValidateSTD = True
 
   If Not Trim(Sheet1.Range("Sheet1.STDcode").Value) = "" Then
        If Trim(Sheet1.Range("sheet1.PhoneNo").Value) = "" Then
           ValidateSTD = False
            Exit Function
        End If
    End If
   End Function

Function ValidateReceiptNo() As Boolean
On Error Resume Next
    ValidateReceiptNo = True
    
    sReceiptNo = Sheet1.Range("sheet1.ReceiptNo")
    
    
    If Trim(sReceiptNo) = "" Or IsEmpty(sReceiptNo) Then
        ValidateReceiptNo = False
        Exit Function
    End If
End Function

Function ValidateDOF() As Boolean
On Error Resume Next
    ValidateDOF = True
    DOF = ""
    DOF = Sheet1.Range("sheet1.OrigRetFiledDate")
   If Trim(DOF) = "" Or Trim((DOF)) = "00/00/0000" Then
        ValidateDOF = False
        Exit Function
        ValidateDOF
    End If
End Function



Function ValidateNoticeNo() As Boolean
On Error Resume Next
    ValidateNoticeNo = True
    sNoticeNo = Sheet1.Range("sheet1.NoticeNo")
    If Trim(sNoticeNo) = "" Or IsEmpty(sNoticeNo) Then
        ValidateNoticeNo = False
        Exit Function
    End If
End Function

'Function ValidateNoticeNoNew() As Boolean
'On Error Resume Next
'    ValidateNoticeNoNew = True
'    sNoticeNoNew = Sheet1.Range("sheet1.NoticeNoNew")
'    If Trim(sNoticeNoNew) = "" Or IsEmpty(sNoticeNoNew) Then
'        ValidateNoticeNoNew = False
'        Exit Function
'    End If
'End Function



Function ValidateNoticeDate() As Boolean
On Error Resume Next
    ValidateNoticeDate = True
    sNoticeDate = Sheet1.Range("sheet1.NoticeDate")
    If Trim(sNoticeDate) = "" Or IsEmpty(sNoticeDate) Then
        ValidateNoticeDate = False
        Exit Function
    End If
End Function

Sub ValidateSIncmFrmHP()
On Error Resume Next
Dim type_HP As String
Dim incmHP As Double
'Ankita_21/01/2026
'type_HP = Sheet1.Range("IncD.TypeOfHP")
'incmHP = Sheet1.Range("IncD.IncomeFromHP")
incmHP = IIf(incmHP = "", 0, incmHP)

'If type_HP = "Self Occupied" Then
'
'         If (Range("IncD.InterestBorrowedCapital").Value) > 200000 Then
'            'MsgBox ("For a Self occupied House Property,Interest payable on borrowed capital value cannot exceed Rs. 2,00,000 in Income Details.")
'            fmsgbox ("* For a Self occupied House Property,Interest payable on borrowed capital value cannot exceed Rs. 2,00,000 in Income Details.")
'            Sheet1.Range("IncD.InterestBorrowedCapital").Value = 0
'            Sheet1.Range("IncD.InterestBorrowedCapital").Select
'         End If
'    End If
'ElseIf type_HP = "Let Out" Then
'    If incmHP < 0 Then
'         If Abs(incmHP) > 200000 Then
'            MsgBox "For a Let Out House Property, Loss cannot exceed Rs. 2,00,000"
'            Sheet1.Range("IncD.IncomeFromHP").Value = 0
'            Sheet1.Range("IncD.IncomeFromHP").Select
'         End If
'    End If
'End If

'ElseIf type_HP = "Let Out" Then
'    If incmHP < 0 Then
'         If Abs(incmHP) > 200000 Then
'            MsgBox "For a Let Out House Property, Loss cannot Setoff more than Rs. 2,00,000", vbExclamation
'
'         End If
'    End If
'End If

End Sub


Function ValidateSIncm()
On Error Resume Next

Dim type_HP As String
Dim incmHP As Double
ValidateSIncm = True
'Ankita_21/01/2026
'type_HP = Sheet1.Range("IncD.TypeOfHP")
'incmHP = Sheet1.Range("IncD.IncomeFromHP")
incmHP = IIf(incmHP = "", 0, incmHP)

'If type_HP = ("Self Occupied" Or "Let Out") Then
'    If incmHP < 0 Then
'         If Abs(incmHP) > 200000 Then
'            ValidateSIncm = False
'            Exit Function
'         End If
'    End If
'End If
'Interest Borrowed Capital are restriction by frontend formula.
'If type_HP = ("Self Occupied") Then
'  If Range("IncD.InterestBorrowedCapital").Value > 200000 Then
'        ValidateSIncm = False
'        Exit Function
'End If
'End If
End Function
Sub calcItr1()

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet1.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
subProcCaption = "Calculating Interest"
noOfProcessSub = 3
If Sheet5.Range("BacValue").Value = 1 Then

calcTaxPayableOnTINTR

calcTaxPayableOnTINTRQ1

calcTaxPayableOnTINTRQ2

calcTaxPayableOnTINTRQ3

calcTaxPayableOnTINTRQ4

calcTaxPayableOnTINTRQ5


Else

calcTaxPayableOnTI
calcTaxPayableOnTIQ1
calcTaxPayableOnTIQ2
calcTaxPayableOnTIQ3
calcTaxPayableOnTIQ4
calcTaxPayableOnTIQ5

End If
'mIncmDtls.calsurchargeOnAboveCrore
mIncmDtls.calcBalTaxPay

UpdateProgressBar

'----------------Lock Password-------------------START---
   Sheet1.Protect Password:=sPassword
'----------------Lock Password-------------------END-----


End Sub
Sub calcTaxPayableOnTIQ1()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q1Inc").Value
taxPayable = Sheet1.Range("IncD.Q1Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").Value)

'If (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 59 And age <= 79) Then
  If (age > 59 And age <= 79) Then
        If totInc <= 300000 Then
             taxPayable = 0
        ElseIf totInc >= 300001 And totInc <= 500000 Then
                tempTax = (totInc - 300000) * (0.05)
                taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                tempTax = (totInc - 500000) * (0.2)
                taxPayable = Round((tempTax + 10000), 0)
         ElseIf totInc >= 1000001 Then
                tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 110000), 0)
        End If
'ElseIf (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 79) Then
    ElseIf (age > 79) Then
        If totInc <= 500000 Then
            taxPayable = 0
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                  tempTax = (totInc - 500000) * (0.2)
                    taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 1000001 Then
                  tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 100000), 0)
        End If
           
ElseIf (totInc <= 250000) Then
        taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 1000000 Then
        tempTax = (totInc - 500000) * (0.2)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 1000001 Then
        tempTax = (totInc - 1000000) * (0.3)
        taxPayable = Round((tempTax + 112500), 0)
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q1Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate
End Sub


Sub calcTaxPayableOnTIQ2()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q2Inc").Value
taxPayable = Sheet1.Range("IncD.Q2Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").Value)

'If (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 59 And age <= 79) Then
  If (age > 59 And age <= 79) Then
        If totInc <= 300000 Then
             taxPayable = 0
        ElseIf totInc >= 300001 And totInc <= 500000 Then
                tempTax = (totInc - 300000) * (0.05)
                taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                tempTax = (totInc - 500000) * (0.2)
                taxPayable = Round((tempTax + 10000), 0)
         ElseIf totInc >= 1000001 Then
                tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 110000), 0)
        End If
'ElseIf (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 79) Then
    ElseIf (age > 79) Then
        If totInc <= 500000 Then
            taxPayable = 0
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                  tempTax = (totInc - 500000) * (0.2)
                    taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 1000001 Then
                  tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 100000), 0)
        End If
           
ElseIf (totInc <= 250000) Then
        taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 1000000 Then
        tempTax = (totInc - 500000) * (0.2)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 1000001 Then
        tempTax = (totInc - 1000000) * (0.3)
        taxPayable = Round((tempTax + 112500), 0)
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q2Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate
End Sub


Sub calcTaxPayableOnTIQ3()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q3Inc").Value
taxPayable = Sheet1.Range("IncD.Q3Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").Value)

'If (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 59 And age <= 79) Then
  If (age > 59 And age <= 79) Then
        If totInc <= 300000 Then
             taxPayable = 0
        ElseIf totInc >= 300001 And totInc <= 500000 Then
                tempTax = (totInc - 300000) * (0.05)
                taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                tempTax = (totInc - 500000) * (0.2)
                taxPayable = Round((tempTax + 10000), 0)
         ElseIf totInc >= 1000001 Then
                tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 110000), 0)
        End If
'ElseIf (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 79) Then
    ElseIf (age > 79) Then
        If totInc <= 500000 Then
            taxPayable = 0
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                  tempTax = (totInc - 500000) * (0.2)
                    taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 1000001 Then
                  tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 100000), 0)
        End If
           
ElseIf (totInc <= 250000) Then
        taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 1000000 Then
        tempTax = (totInc - 500000) * (0.2)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 1000001 Then
        tempTax = (totInc - 1000000) * (0.3)
        taxPayable = Round((tempTax + 112500), 0)
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q3Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate
End Sub


Sub calcTaxPayableOnTIQ4()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q4Inc").Value
taxPayable = Sheet1.Range("IncD.Q4Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").Value)

'If (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 59 And age <= 79) Then
  If (age > 59 And age <= 79) Then
        If totInc <= 300000 Then
             taxPayable = 0
        ElseIf totInc >= 300001 And totInc <= 500000 Then
                tempTax = (totInc - 300000) * (0.05)
                taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                tempTax = (totInc - 500000) * (0.2)
                taxPayable = Round((tempTax + 10000), 0)
         ElseIf totInc >= 1000001 Then
                tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 110000), 0)
        End If
'ElseIf (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 79) Then
    ElseIf (age > 79) Then
        If totInc <= 500000 Then
            taxPayable = 0
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                  tempTax = (totInc - 500000) * (0.2)
                    taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 1000001 Then
                  tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 100000), 0)
        End If
           
ElseIf (totInc <= 250000) Then
        taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 1000000 Then
        tempTax = (totInc - 500000) * (0.2)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 1000001 Then
        tempTax = (totInc - 1000000) * (0.3)
        taxPayable = Round((tempTax + 112500), 0)
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q4Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate
End Sub



Sub calcTaxPayableOnTIQ5()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q5Inc").Value
taxPayable = Sheet1.Range("IncD.Q5Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").Value)

'If (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 59 And age <= 79) Then
  If (age > 59 And age <= 79) Then
        If totInc <= 300000 Then
             taxPayable = 0
        ElseIf totInc >= 300001 And totInc <= 500000 Then
                tempTax = (totInc - 300000) * (0.05)
                taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                tempTax = (totInc - 500000) * (0.2)
                taxPayable = Round((tempTax + 10000), 0)
         ElseIf totInc >= 1000001 Then
                tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 110000), 0)
        End If
'ElseIf (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 79) Then

    ElseIf (age > 79) Then
        If totInc <= 500000 Then
            taxPayable = 0
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                  tempTax = (totInc - 500000) * (0.2)
                    taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 1000001 Then
                  tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 100000), 0)
        End If

ElseIf (totInc <= 250000) Then
        taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 1000000 Then
        tempTax = (totInc - 500000) * (0.2)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 1000001 Then
        tempTax = (totInc - 1000000) * (0.3)
        taxPayable = Round((tempTax + 112500), 0)
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q5Tax").Value = taxPayable - Sheet1.Range("IncD.Q4Tax").Value
Sheet1.Protect Password:=EfilingCommon.getmsgstate
End Sub

Sub calcTaxPayableOnTI()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.TotalIncome").Value
taxPayable = Sheet1.Range("IncD.TotalTaxPayable").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").Value)

'If (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 59 And age <= 79) Then
  If (age > 59 And age <= 79) Then
        If totInc <= 300000 Then
             taxPayable = 0
        ElseIf totInc >= 300001 And totInc <= 500000 Then
                tempTax = (totInc - 300000) * (0.05)
                taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                tempTax = (totInc - 500000) * (0.2)
                taxPayable = Round((tempTax + 10000), 0)
        ElseIf totInc >= 1000001 Then
                tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 110000), 0)
        End If
'ElseIf (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 79) Then
    ElseIf (age > 79) Then
        If totInc <= 500000 Then
            taxPayable = 0
        ElseIf totInc >= 500001 And totInc <= 1000000 Then
                  tempTax = (totInc - 500000) * (0.2)
                    taxPayable = Round(tempTax, 0)
        ElseIf totInc >= 1000001 Then
                  tempTax = (totInc - 1000000) * (0.3)
                taxPayable = Round((tempTax + 100000), 0)
        End If
           
ElseIf (totInc <= 250000) Then
        taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 1000000 Then
        tempTax = (totInc - 500000) * (0.2)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 1000001 Then
        tempTax = (totInc - 1000000) * (0.3)
        taxPayable = Round((tempTax + 112500), 0)
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.TotalTaxPayable").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate
End Sub
Sub calcTaxPayableOnTINTRQ1()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q1Inc").Value
taxPayable = Sheet1.Range("IncD.Q1Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

If (totInc <= 400000) Then  'Ankita_09/12/2025  changes from 3lpa to 4lpa
       taxPayable = 0
        
ElseIf totInc >= 400001 And totInc <= 800000 Then    'Ankita_09/12/2025 changes .... '6lpa to 7lpa Changed by Ankita on 16/12/2024
        tempTax = (totInc - 400000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 800001 And totInc <= 1200000 Then     'Ankita_09/12/2025 changes  .....'6lpa to 7lpa & 9lpa to 10lpa changed by Ankita on 16/12/2024
        tempTax = (totInc - 800000) * (0.1)
        taxPayable = Round((tempTax + 20000), 0)       'Ankita_09/12/2025 changes..........'15k to 20k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 1200001 And totInc <= 1600000 Then    'Ankita_09/12/2025 changes..........'Ankita_09/12/2025 changes  .....'9lpa to 10lpa changed by Ankita on 16/12/2024
        tempTax = (totInc - 1200000) * (0.15)
        taxPayable = Round((tempTax + 60000), 0)       'Ankita_09/12/2025 changes..........'45k to 50k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 1600001 And totInc <= 2000000 Then    'Ankita_09/12/2025 changes..........'Ankita_09/12/2025 changes
        tempTax = (totInc - 1600000) * (0.2)
        taxPayable = Round((tempTax + 120000), 0)      'Ankita_09/12/2025 changes.......... '90k to 80k changed by Ankita on 16/12/2024
               
ElseIf totInc >= 2000001 And totInc <= 2400000 Then    'Ankita_09/12/2025 changes..........'Ankita_09/12/2025 changes
        tempTax = (totInc - 2000000) * (0.25)
        taxPayable = Round((tempTax + 200000), 0)      'Ankita_09/12/2025 changes..........
        
ElseIf totInc >= 2400000 Then                          'Ankita_09/12/2025 changes..........'Ankita_09/12/2025 changes
        tempTax = (totInc - 2400000) * (0.3)
        taxPayable = Round((tempTax + 300000), 0)      'Ankita_09/12/2025 changes.......... '150k to 140k changed by Ankita on 16/12/2024

End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q1Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub

Sub calcTaxPayableOnTINTRQ1_OLD()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q1Inc").Value
taxPayable = Sheet1.Range("IncD.Q1Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

If (totInc <= 250000) Then
       taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 750000 Then
        tempTax = (totInc - 500000) * (0.1)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 750001 And totInc <= 1000000 Then
        tempTax = (totInc - 750000) * (0.15)
        taxPayable = Round((tempTax + 37500), 0)
ElseIf totInc >= 1000001 And totInc <= 1250000 Then
        tempTax = (totInc - 1000000) * (0.2)
        taxPayable = Round((tempTax + 75000), 0)
ElseIf totInc >= 1250001 And totInc <= 1500000 Then
        tempTax = (totInc - 1250000) * (0.25)
        taxPayable = Round((tempTax + 125000), 0)
ElseIf totInc >= 1500001 Then
        tempTax = (totInc - 1500000) * (0.3)
        taxPayable = Round((tempTax + 187500), 0)



End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q1Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub

Sub calcTaxPayableOnTINTRQ2()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q2Inc").Value
taxPayable = Sheet1.Range("IncD.Q2Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value


           
If (totInc <= 400000) Then
       taxPayable = 0
        
ElseIf totInc >= 400001 And totInc <= 800000 Then       'Ankita_09/12/2025 changes.........   '6lpa to 7lpa Changed by Ankita on 16/12/2024
        tempTax = (totInc - 400000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 800001 And totInc <= 1200000 Then     'Ankita_09/12/2025 changes .........  '6lpa to 7lpa & 9lpa to 10lpa changed by Ankita on 16/12/2024
        tempTax = (totInc - 800000) * (0.1)
        taxPayable = Round((tempTax + 20000), 0)        'Ankita_09/12/2025 changes......... '15k to 20k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 1200001 And totInc <= 1600000 Then      'Ankita_09/12/2025 changes.........'9lpa to 10lpa Changed by Ankita on 16/12/2024
        tempTax = (totInc - 1200000) * (0.15)
        taxPayable = Round((tempTax + 60000), 0)         'Ankita_09/12/2025 changes.........'45k to 50k changed by Ankita on 16/12/2024

ElseIf totInc >= 1600001 And totInc <= 2000000 Then
        tempTax = (totInc - 1600000) * (0.2)
        taxPayable = Round((tempTax + 120000), 0)         'Ankita_09/12/2025 changes.........'90k to 80k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 2000001 And totInc <= 2400000 Then       'uncommented by Ankita on 10/12/2025
        tempTax = (totInc - 2000000) * (0.25)
        taxPayable = Round((tempTax + 200000), 0)
        
ElseIf totInc >= 2400001 Then
        tempTax = (totInc - 2400000) * (0.3)
        taxPayable = Round((tempTax + 300000), 0)         'Ankita_09/12/2025 changes.........'150k to 140k changed by Ankita on 16/12/2024
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q2Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub
Sub calcTaxPayableOnTINTRQ2_OLD()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q2Inc").Value
taxPayable = Sheet1.Range("IncD.Q2Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value


           
If (totInc <= 250000) Then
       taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 750000 Then
        tempTax = (totInc - 500000) * (0.1)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 750001 And totInc <= 1000000 Then
        tempTax = (totInc - 750000) * (0.15)
        taxPayable = Round((tempTax + 37500), 0)
ElseIf totInc >= 1000001 And totInc <= 1250000 Then
        tempTax = (totInc - 1000000) * (0.2)
        taxPayable = Round((tempTax + 75000), 0)
ElseIf totInc >= 1250001 And totInc <= 1500000 Then
        tempTax = (totInc - 1250000) * (0.25)
        taxPayable = Round((tempTax + 125000), 0)
ElseIf totInc >= 1500001 Then
        tempTax = (totInc - 1500000) * (0.3)
        taxPayable = Round((tempTax + 187500), 0)



End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q2Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub
Sub calcTaxPayableOnTINTRQ3()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q3Inc").Value
taxPayable = Sheet1.Range("IncD.Q3Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

If (totInc <= 400000) Then
       taxPayable = 0
        
ElseIf totInc >= 400001 And totInc <= 800000 Then     'Ankita_09/12/2025 changes......... '6lpa to 7lpa Changed by Ankita on 16/12/2024
        tempTax = (totInc - 400000) * (0.05)
        taxPayable = Round(tempTax, 0)
         
ElseIf totInc >= 800001 And totInc <= 1200000 Then     'Ankita_09/12/2025 changes.........'6lpa to 7lpa & 9lpa to 10lpa changed by Ankita on 16/12/2024
        tempTax = (totInc - 800000) * (0.1)
        taxPayable = Round((tempTax + 20000), 0)       'Ankita_09/12/2025 changes.........'15k to 20k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 1200001 And totInc <= 1600000 Then    'Ankita_09/12/2025 changes.........'9lpa to 10lpa changed by Ankita on 16/12/2024
        tempTax = (totInc - 1200000) * (0.15)
        taxPayable = Round((tempTax + 60000), 0)       'Ankita_09/12/2025 changes.........'45k to 50k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 1600001 And totInc <= 2000000 Then    'Ankita_09/12/2025 changes.........
        tempTax = (totInc - 1600000) * (0.2)
        taxPayable = Round((tempTax + 120000), 0)       'Ankita_09/12/2025 changes.........'90k to 80k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 2000001 And totInc <= 2400000 Then    'Uncommted by 'Ankita_09/12/2025 changes
        tempTax = (totInc - 2000000) * (0.25)
        taxPayable = Round((tempTax + 200000), 0)
        
ElseIf totInc >= 2400001 Then
        tempTax = (totInc - 2400000) * (0.3)
        taxPayable = Round((tempTax + 300000), 0)       'Ankita_09/12/2025 changes.......'150k to 140k changed by Ankita on 16/12/2024

End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q3Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub
Sub calcTaxPayableOnTINTRQ3_OLD()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q3Inc").Value
taxPayable = Sheet1.Range("IncD.Q3Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value


           
If (totInc <= 250000) Then
       taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 750000 Then
        tempTax = (totInc - 500000) * (0.1)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 750001 And totInc <= 1000000 Then
        tempTax = (totInc - 750000) * (0.15)
        taxPayable = Round((tempTax + 37500), 0)
ElseIf totInc >= 1000001 And totInc <= 1250000 Then
        tempTax = (totInc - 1000000) * (0.2)
        taxPayable = Round((tempTax + 75000), 0)
ElseIf totInc >= 1250001 And totInc <= 1500000 Then
        tempTax = (totInc - 1250000) * (0.25)
        taxPayable = Round((tempTax + 125000), 0)
ElseIf totInc >= 1500001 Then
        tempTax = (totInc - 1500000) * (0.3)
        taxPayable = Round((tempTax + 187500), 0)



End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q3Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub
Sub calcTaxPayableOnTINTRQ4()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q4Inc").Value
taxPayable = Sheet1.Range("IncD.Q4Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value


           
If (totInc <= 400000) Then
       taxPayable = 0
        
ElseIf totInc >= 400001 And totInc <= 800000 Then     'Ankita_09/12/2025 changes......... '6lpa to 7lpa Changed by Ankita on 16/12/2024
        tempTax = (totInc - 400000) * (0.05)
        taxPayable = Round(tempTax, 0)
         
ElseIf totInc >= 800001 And totInc <= 1200000 Then     'Ankita_09/12/2025 changes.........'6lpa to 7lpa & 9lpa to 10lpa changed by Ankita on 16/12/2024
        tempTax = (totInc - 800000) * (0.1)
        taxPayable = Round((tempTax + 20000), 0)       'Ankita_09/12/2025 changes.........'15k to 20k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 1200001 And totInc <= 1600000 Then    'Ankita_09/12/2025 changes.........'9lpa to 10lpa changed by Ankita on 16/12/2024
        tempTax = (totInc - 1200000) * (0.15)
        taxPayable = Round((tempTax + 60000), 0)       'Ankita_09/12/2025 changes.........'45k to 50k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 1600001 And totInc <= 2000000 Then    'Ankita_09/12/2025 changes.........
        tempTax = (totInc - 1600000) * (0.2)
        taxPayable = Round((tempTax + 120000), 0)       'Ankita_09/12/2025 changes.........'90k to 80k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 2000001 And totInc <= 2400000 Then    'Uncommted by 'Ankita_09/12/2025 changes
        tempTax = (totInc - 2000000) * (0.25)
        taxPayable = Round((tempTax + 200000), 0)
        
ElseIf totInc >= 2400001 Then
        tempTax = (totInc - 2400000) * (0.3)
        taxPayable = Round((tempTax + 300000), 0)       'Ankita_09/12/2025 changes.......'150k to 140k changed by Ankita on 16/12/2024
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q4Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub
Sub calcTaxPayableOnTINTRQ4_OLD()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q4Inc").Value
taxPayable = Sheet1.Range("IncD.Q4Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value


           
If (totInc <= 250000) Then
       taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 750000 Then
        tempTax = (totInc - 500000) * (0.1)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 750001 And totInc <= 1000000 Then
        tempTax = (totInc - 750000) * (0.15)
        taxPayable = Round((tempTax + 37500), 0)
ElseIf totInc >= 1000001 And totInc <= 1250000 Then
        tempTax = (totInc - 1000000) * (0.2)
        taxPayable = Round((tempTax + 75000), 0)
ElseIf totInc >= 1250001 And totInc <= 1500000 Then
        tempTax = (totInc - 1250000) * (0.25)
        taxPayable = Round((tempTax + 125000), 0)
ElseIf totInc >= 1500001 Then
        tempTax = (totInc - 1500000) * (0.3)
        taxPayable = Round((tempTax + 187500), 0)



End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q4Tax").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub

Sub calcTaxPayableOnTINTRQ5()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q5Inc").Value
taxPayable = Sheet1.Range("IncD.Q5Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value


           
If (totInc <= 400000) Then
       taxPayable = 0
        
ElseIf totInc >= 400001 And totInc <= 800000 Then     'Ankita_09/12/2025 changes......... '6lpa to 7lpa Changed by Ankita on 16/12/2024
        tempTax = (totInc - 400000) * (0.05)
        taxPayable = Round(tempTax, 0)
         
ElseIf totInc >= 800001 And totInc <= 1200000 Then     'Ankita_09/12/2025 changes.........'6lpa to 7lpa & 9lpa to 10lpa changed by Ankita on 16/12/2024
        tempTax = (totInc - 800000) * (0.1)
        taxPayable = Round((tempTax + 20000), 0)       'Ankita_09/12/2025 changes.........'15k to 20k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 1200001 And totInc <= 1600000 Then    'Ankita_09/12/2025 changes.........'9lpa to 10lpa changed by Ankita on 16/12/2024
        tempTax = (totInc - 1200000) * (0.15)
        taxPayable = Round((tempTax + 60000), 0)       'Ankita_09/12/2025 changes.........'45k to 50k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 1600001 And totInc <= 2000000 Then    'Ankita_09/12/2025 changes.........
        tempTax = (totInc - 1600000) * (0.2)
        taxPayable = Round((tempTax + 120000), 0)       'Ankita_09/12/2025 changes.........'90k to 80k changed by Ankita on 16/12/2024
        
ElseIf totInc >= 2000001 And totInc <= 2400000 Then    'Uncommted by 'Ankita_09/12/2025 changes
        tempTax = (totInc - 2000000) * (0.25)
        taxPayable = Round((tempTax + 200000), 0)
        
ElseIf totInc >= 2400001 Then
        tempTax = (totInc - 2400000) * (0.3)
        taxPayable = Round((tempTax + 300000), 0)       'Ankita_09/12/2025 changes.......'150k to 140k changed by Ankita on 16/12/2024
End If


Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q5Tax").Value = taxPayable - Sheet1.Range("IncD.Q4Tax").Value
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub
Sub calcTaxPayableOnTINTRQ5_OLD()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.Q5Inc").Value
taxPayable = Sheet1.Range("IncD.Q5Tax").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value


           
If (totInc <= 250000) Then
       taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 750000 Then
        tempTax = (totInc - 500000) * (0.1)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 750001 And totInc <= 1000000 Then
        tempTax = (totInc - 750000) * (0.15)
        taxPayable = Round((tempTax + 37500), 0)
ElseIf totInc >= 1000001 And totInc <= 1250000 Then
        tempTax = (totInc - 1000000) * (0.2)
        taxPayable = Round((tempTax + 75000), 0)
ElseIf totInc >= 1250001 And totInc <= 1500000 Then
        tempTax = (totInc - 1250000) * (0.25)
        taxPayable = Round((tempTax + 125000), 0)
ElseIf totInc >= 1500001 Then
        tempTax = (totInc - 1500000) * (0.3)
        taxPayable = Round((tempTax + 187500), 0)



End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Q5Tax").Value = taxPayable - Sheet1.Range("IncD.Q4Tax").Value
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub


Sub calcTaxPayableOnTINTR()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.TotalIncome").Value
taxPayable = Sheet1.Range("IncD.TotalTaxPayable").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value


           
If (totInc <= 400000) Then
       taxPayable = 0
        
ElseIf totInc >= 400001 And totInc <= 800000 Then        'Changed by Ankita on 16/12/2024
        tempTax = (totInc - 400000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 800001 And totInc <= 1200000 Then
        tempTax = (totInc - 800000) * (0.1)
        taxPayable = Round((tempTax + 20000), 0)
        
ElseIf totInc >= 1200001 And totInc <= 1600000 Then
        tempTax = (totInc - 1200000) * (0.15)
        taxPayable = Round((tempTax + 60000), 0)
        
ElseIf totInc >= 1600001 And totInc <= 2000000 Then
        tempTax = (totInc - 1600000) * (0.2)
        taxPayable = Round((tempTax + 120000), 0)
        
ElseIf totInc >= 2000001 And totInc <= 2400000 Then
        tempTax = (totInc - 2000000) * (0.25)
        taxPayable = Round((tempTax + 200000), 0)
        
ElseIf totInc >= 2400000 Then
        tempTax = (totInc - 2400000) * (0.3)
        taxPayable = Round((tempTax + 300000), 0)
End If


Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.TotalTaxPayable").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub
Sub calcTaxPayableOnTINTR_OLD()
On Error Resume Next

Dim tempTax As Double

totInc = Sheet1.Range("IncD.TotalIncome").Value
taxPayable = Sheet1.Range("IncD.TotalTaxPayable").Value
'resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value


           
If (totInc <= 250000) Then
       taxPayable = 0
        
ElseIf totInc >= 250001 And totInc <= 500000 Then
        tempTax = (totInc - 250000) * (0.05)
        taxPayable = Round(tempTax, 0)
        
ElseIf totInc >= 500001 And totInc <= 750000 Then
        tempTax = (totInc - 500000) * (0.1)
        taxPayable = Round((tempTax + 12500), 0)
        
ElseIf totInc >= 750001 And totInc <= 1000000 Then
        tempTax = (totInc - 750000) * (0.15)
        taxPayable = Round((tempTax + 37500), 0)
ElseIf totInc >= 1000001 And totInc <= 1250000 Then
        tempTax = (totInc - 1000000) * (0.2)
        taxPayable = Round((tempTax + 75000), 0)
ElseIf totInc >= 1250001 And totInc <= 1500000 Then
        tempTax = (totInc - 1250000) * (0.25)
        taxPayable = Round((tempTax + 125000), 0)
ElseIf totInc >= 1500001 Then
        tempTax = (totInc - 1500000) * (0.3)
        taxPayable = Round((tempTax + 187500), 0)
End If

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.TotalTaxPayable").Value = taxPayable
Sheet1.Protect Password:=EfilingCommon.getmsgstate

End Sub

Sub calsurchargeOnAboveCroreAbc()
'
'On Error Resume Next
'
'
'Dim taxOnCutOffInc As Double
'Dim tempSurcharge As Double
'Dim extraInc As Double
'Dim marginalRelief As Double
'
'totInc = Sheet1.Range("IncD.IncD.TotalIncome").Value
'taxPayable = Sheet1.Range("IncD.TotalTaxPayable").Value
''resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value
'surchargeOnAboveCrore = Sheet1.Range("IncD.SurchargeOnTaxPayable").Value
'
'
'age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").Value)
'
'
'If ((resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 59) And age <= 79) Then
'    taxOnCutOffInc = ((10000000 - 1000000) * 0.3 + 120000)
'
'ElseIf ((resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") And (age > 79)) Then
'    taxOnCutOffInc = ((10000000 - 1000000) * 0.3 + 100000)
'Else
'    taxOnCutOffInc = ((10000000 - 1000000) * 0.3 + 125000)
'End If
'
'If (Round(totInc / 10, 0) * 10) > 10000000 Then
'    tempSurcharge = taxPayable * 0.15
'
'    'Check if eligible for marginal relief
'    extraInc = (Round(totInc / 10, 0) * 10) - 10000000
'
'    If ((taxPayable + tempSurcharge) > (taxOnCutOffInc + extraInc)) Then
'        marginalRelief = taxPayable + tempSurcharge - (taxOnCutOffInc + extraInc)
'        surchargeOnAboveCrore = Round(tempSurcharge - marginalRelief)
'
'    Else
'        surchargeOnAboveCrore = Round(tempSurcharge)
'    End If
'Else
'    surchargeOnAboveCrore = 0
'End If
' Sheet1.Range("IncD.SurchargeOnTaxPayable").Value = surchargeOnAboveCrore
'
'
End Sub


Sub calcBalTaxPay()
On Error Resume Next


Dim sec89 As Double
Dim sec89A As Double
Dim totTaxWithEduCess As Double

sec89 = Sheet1.Range("IncD.Section89").Value
sec89A = Sheet1.Range("IncD.Section89A").Value
TaxPayableOnRebate = Sheet1.Range("IncD.TaxPayableOnRebate").Value
'surchargeOnAboveCrore = Sheet1.Range("IncD.SurchargeOnTaxPayable").Value
eduCess = Sheet1.Range("IncD.EducationCess").Value

'totTaxWithEduCess = taxPayableOnRebate + surchargeOnAboveCrore + eduCess
totTaxWithEduCess = TaxPayableOnRebate + eduCess


balTaxPay = Round((totTaxWithEduCess - sec89 - sec89A), 0)


If balTaxPay < 0 Then
    balTaxPay = 0
End If

mIncmDtls.calcInterestPayable
mIncmDtls.calcTaxPayable15Minus17


End Sub

Sub calcInterestPayable()
On Error Resume Next

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
    Sheet3.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----



Dim intrst234Aprinciple As Double
'Dim MonthsAfterDueDate As Long
Dim origDate As Variant
Dim todayDate As Variant

AdvanceTax = 0
SelfAssessmentTax = 0
selfAssessmentTax234A = 0
TDS = 0
TCS = 0
advanceTaxToDisplay = Sheet3.Range("IncD.AdvanceTax").Value
TDSToDisplay = Sheet3.Range("IncD.TDS").Value
TCSToDisplay = Sheet3.Range("IncD.TCS").Value
SATtoDisplay = Sheet3.Range("IncD.SelfAssessmentTax").Value


BalTaxPayable = Sheet1.Range("IncD.NetTaxLiability").Value


 age = EfilingCommon.calculateAge(Sheet1.Range("sheet1.DOB").Value)
' resStatus = Sheet1.Range("sheet1.ResidentialStatus1").Value

mIncmDtls.calcTDSFrmSal
mIncmDtls.calcTDSOtherSal
mIncmDtls.calcTDS3OtherSal
mIncmDtls.calcTCSFrmSal


TDSToDisplay = TDS
TCSToDisplay = TCS
Sheet3.Range("IncD.TDS") = TDSToDisplay
Sheet3.Range("IncD.TCS") = TCSToDisplay
mIncmDtls.calcAdvSelfTax

Sheet3.Range("IncD.AdvanceTax").Value = advanceTaxToDisplay
Sheet3.Range("IncD.SelfAssessmentTax").Value = SATtoDisplay

If (BalTaxPayable - AdvanceTax - TDS - TCS - selfAssessmentTax234A < 0) Then
    intrst234Aprinciple = 0
Else
    intrst234Aprinciple = BalTaxPayable - AdvanceTax - TDS - TCS - selfAssessmentTax234A
        intrst234Aprinciple = Application.WorksheetFunction.Floor(intrst234Aprinciple, 100)
End If



currentdate = Sheet3.Range("Ver.Date").Value
todayDate = Sheet5.Range("TodayDate").Value

If EfilingCommon.checkFirstDateBefore(currentdate, todayDate) Then
    currentdate = todayDate
End If

    
bacage = age
If (Sheet5.Range("BacValue").Value) = 1 Then
age = 55
End If

If (age > 59) Then
'Ayush_DueDate_08/09/2025
    If (BalTaxPayable - AdvanceTax - TDS - TCS - Sheet2.Range("IT.Sat1").Value - Sheet2.Range("IT.Sat2").Value - Sheet2.Range("IT.Sat3").Value - Sheet2.Range("IT.Sat4").Value - Sheet2.Range("ExSAT").Value) > 100000 Then
    
'    Ayush_DueDate_08/09/2025
      MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(currentdate, Sheet5.Range("DueDate1").Value)

    Else
     
     MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(currentdate, Sheet5.Range("DueDate1").Value)

    End If
Else
    If (BalTaxPayable - AdvanceTax - TDS - TCS) > 100000 Then

'    Ayush_DueDate_08/09/2025
    MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(currentdate, Sheet5.Range("DueDate1").Value)


    Else

'    Ayush_DueDate_08/09/2025
    MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(currentdate, Sheet5.Range("DueDate1").Value)


    End If
End If
Dim UpdatedY As Boolean
    UpdatedY = False
    If (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
    UpdatedY = True
    End If

If (Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "17" Or UpdatedY Then
    origDate = Sheet1.Range("Sheet1.OrigRetFiledDate").Value
        
    If origDate <> "" Then
        
        
        If (age > 59) Then
'            If (BalTaxPayable - AdvanceTax - TDS - TCS - Sheet2.Range("IT.Sat1").Value - Sheet2.Range("IT.Sat2").Value - Sheet2.Range("IT.Sat3").Value - Sheet2.Range("IT.Sat4").Value) > 100000 Then
'Ayush_DueDate_08/09/2025
        If (BalTaxPayable - AdvanceTax - TDS - TCS - Sheet2.Range("IT.Sat1").Value - Sheet2.Range("IT.Sat2").Value - Sheet2.Range("IT.Sat3").Value - Sheet2.Range("IT.Sat4").Value - Sheet2.Range("ExSAT").Value) > 100000 Then

'    Ayush_DueDate_08/09/2025
            MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(origDate, Sheet5.Range("DueDate1").Value)
    

            Else

'    Ayush_DueDate_08/09/2025
            MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(origDate, Sheet5.Range("DueDate1").Value)

            End If
        Else
            If (BalTaxPayable - AdvanceTax - TDS - TCS) > 100000 Then

'    Ayush_DueDate_08/09/2025
            MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(origDate, Sheet5.Range("DueDate1").Value)

            Else

'    Ayush_DueDate_08/09/2025
            MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(origDate, Sheet5.Range("DueDate1").Value)

            End If
        End If
    Else
        MonthsAfterDueDate = 0
    End If
End If


If (Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2)) = "18") Then
'origDate = Sheet1.Range("sheet1.NoticeDate").Value
origDate = Sheet1.Range("Sheet1.OrigRetFiledDate").Value
    If origDate <> "" Then
        
        
        If (age > 59) Then

'Ayush_DueDate_08/09/2025
             If (BalTaxPayable - AdvanceTax - TDS - TCS - Sheet2.Range("IT.Sat1").Value - Sheet2.Range("IT.Sat2").Value - Sheet2.Range("IT.Sat3").Value - Sheet2.Range("IT.Sat4").Value - Sheet2.Range("ExSAT").Value) > 100000 Then

'    Ayush_DueDate_08/09/2025
            MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(origDate, Sheet5.Range("DueDate1").Value)
            Else

'    Ayush_DueDate_08/09/2025
            MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(origDate, Sheet5.Range("DueDate1").Value)
            End If
        Else
            If (BalTaxPayable - AdvanceTax - TDS - TCS) > 100000 Then

'    Ayush_DueDate_08/09/2025
            MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(origDate, Sheet5.Range("DueDate1").Value)
            Else

'    Ayush_DueDate_08/09/2025
            MonthsAfterDueDate = EfilingCommon.calcNoOfMonths(origDate, Sheet5.Range("DueDate1").Value)
            End If
        End If
    Else
        MonthsAfterDueDate = 0
    End If
End If
'End If


'234F Calculation
'Dim Exempt38 As Variant
Dim TempDueDate As Variant
Dim TempVerificationDate As Variant
Dim TempVerificationDate1 As Variant
Dim TempDateOfFiling As Variant
Dim TempDateOfFiling1 As Variant

  
  
If (Sheet5.Range("BacValue").Value) = 1 Then
age = 55
End If
    
  
If (age <= 59) Then
agestatus = "NC"
ElseIf ((age > 59) And (age <= 79)) Then
agestatus = "SC"
Else
agestatus = "SSC"
End If


'Ayush_Duedate_08/09/2025
TempDueDate = Sheet5.Range("DueDate1").Value


'Else

'End If



TempVerificationDate = Sheet3.Range("Ver.Date").Value
TempDueDate = Dformat(TempDueDate, "yyyy-mm-dd")
TempVerificationDate = Dformat(TempVerificationDate, "yyyy-mm-dd")

TempDateOfFiling = Range("sheet1.OrigRetFiledDate").Value
TempDateOfFiling = Dformat(TempDateOfFiling, "yyyy-mm-dd")

If Not TempDateOfFiling = "" Then
TempVerificationDate = TempDateOfFiling
Else
TempVerificationDate = TempVerificationDate
End If



If Sheet1.Range("sheet1.SeventhProvisoFlag").Value = "Yes" Then
NEW234F
Else
'INC_E37
'If agestatus = "NC" And (Range("IncD.GrossTotIncome").Value > 250000) Then
' Ankita_28/01/2026_AY26-27
'If agestatus = "NC" And (Range("IncD.GrossTotIncome").Value > 300000) And Sheet5.Range("BacValue").Value = 1 Then 'For 234F We need to consider GTI in which 112A should include as per DE V3 & BA - update dy Bindu on 4th Feb 2025
'If agestatus = "NC" And (Range("IncD.GrossTotIncome_New").Value > 300000) And Sheet5.Range("BacValue").Value = 1 Then
If agestatus = "NC" And (Range("IncD.GrossTotIncome_New").Value > 400000) And Sheet5.Range("BacValue").Value = 1 Then
  
  NEW234F
'ElseIf agestatus = "NC" And (Range("IncD.GrossTotIncome").Value > 250000) And Sheet5.Range("BacValue").Value = 2 Then ''For 234F We need to consider GTI in which 112A should include as per DE V3 & BA - update dy Bindu on 4th Feb 2025
ElseIf agestatus = "NC" And (Range("IncD.GrossTotIncome_New").Value > 250000) And Sheet5.Range("BacValue").Value = 2 Then
  
  NEW234F
'ElseIf agestatus = "SC" And (Range("IncD.GrossTotIncome").Value > 300000) Then ''For 234F We need to consider GTI in which 112A should include as per DE V3 & BA - update dy Bindu on 4th Feb 2025
 ElseIf agestatus = "SC" And (Range("IncD.GrossTotIncome_New").Value > 300000) Then
  NEW234F
'ElseIf agestatus = "SSC" And (Range("IncD.GrossTotIncome").Value > 500000) Then ''For 234F We need to consider GTI in which 112A should include as per DE V3 & BA - update dy Bindu on 4th Feb 2025
ElseIf agestatus = "SSC" And (Range("IncD.GrossTotIncome_New").Value > 500000) Then
  NEW234F
Else
  intrst234F = 0
End If
End If

'Sheet3.Range("EI.Sec10_38").Value = Sheet3.Range("EI.Sec10_38").Value




intrst234A = intrst234Aprinciple * (0.01) * MonthsAfterDueDate
intrst234B = 0
intrst234C = 0
slab0 = 0
slab1 = 0
slab2 = 0
slab3 = 0
slab4 = 0

SchTDS.setTableInfo_Grid3
Dim i As Long
Dim rangecells_I As Range
Set rangecells_I = Range("TaxP.DateDep").Cells

Dim rangecells_I2 As Range
Set rangecells_I2 = Range("TaxP.Amt").Cells

ReDim taxDepDate(ColCount2)
ReDim taxPAmt(ColCount2)

   Dim Temp_States, Temp_States_Flag As Variant
   Temp_States = Mid(Sheet1.Range("sheet1.StateCode1").Value, 1, 2)
   If Temp_States = "03" Or Temp_States = "04" Or Temp_States = "20" Or Temp_States = "21" Or Temp_States = "22" Or Temp_States = "23" Or Temp_States = "30" Then
   'Temp_States_Flag = True
   Temp_States_Flag = False
   Else
   Temp_States_Flag = False
   End If

 For i = 1 To ColCount2
    taxDepDate(i) = rangecells_I.item(i).Value
    taxPAmt(i) = rangecells_I2.item(i).Value

   
    If Temp_States_Flag = True Then
          If (EfilingCommon.checkFirstDateBefore("01/04/2019", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "15/06/2019")) Then
            slab0 = slab0 + taxPAmt(i)
        ElseIf (EfilingCommon.checkFirstDateBefore("16/06/2019", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "16/09/2019")) Then
            slab1 = slab1 + taxPAmt(i)
        ElseIf (EfilingCommon.checkFirstDateBefore("17/09/2019", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "31/12/2019")) Then
             slab2 = slab2 + taxPAmt(i)
        ElseIf (EfilingCommon.checkFirstDateBefore("01/01/2020", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "16/03/2020")) Then
             slab3 = slab3 + taxPAmt(i)
        End If
    Else
       'TAX DETAILS-C5 2024-25 Bindu Changed year from 2023 to 2024 & changed year from 2022 to 2023
        If (EfilingCommon.checkFirstDateBefore("01/04/2025", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "16/06/2025")) Then  'Year changed from 2023 to 2024 by Ankita on 27/12/2024
            slab0 = slab0 + taxPAmt(i)
        ElseIf (EfilingCommon.checkFirstDateBefore("17/06/2025", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "20/09/2025")) Then  'Year changed from 2023 to 2024 by Ankita on 27/12/2024
            slab1 = slab1 + taxPAmt(i)
        ElseIf (EfilingCommon.checkFirstDateBefore("21/09/2025", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "15/12/2025")) Then   'Year changed from 2023 to 2024 by Ankita on 27/12/2024
            slab2 = slab2 + taxPAmt(i)
        'ElseIf (EfilingCommon.checkFirstDateBefore("16/12/2023", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "15/03/2023")) Then 'Malli
        ElseIf (EfilingCommon.checkFirstDateBefore("16/12/2025", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "16/03/2026")) Then   'Year changed from 2023 to 2024 & 2024 to 2025 by Ankita on 27/12/2024
            slab3 = slab3 + taxPAmt(i)
        ElseIf (EfilingCommon.checkFirstDateBefore("17/03/2026", taxDepDate(i)) And EfilingCommon.checkFirstDateBefore(taxDepDate(i), "31/03/2026")) Then   'Year changed from 2024 to 2025 by Ankita on 27/12/2024
             slab4 = slab4 + taxPAmt(i)
        End If
    End If
Next
 mIncmDtls.calcIntrst234C
 
'   mIncmDtls.calcIntrst234B
    
    Application.EnableEvents = False
    ComputeInterest
    Application.EnableEvents = False
    Sheet1.Unprotect Password:=sPassword
    'Ayush_07/10
 '   Sheet1.Range("IncD.IntrstPayUs234A").Value = intrst234A
    Sheet1.Range("IncD.IntrstPayUs234B").Value = IntrstPayUs234B
    Sheet1.Range("IncD.IntrstPayUs234C").Value = intrst234C
    Sheet1.Range("IncD.IntrstPayUs234F").Value = intrst234F
       
    mIncmDtls.calcTaxPayable15Minus17
'----------------Lock Password-------------------START---
  Sheet3.Protect Password:=sPassword
  Sheet1.Protect Password:=sPassword
  Application.EnableEvents = True
'----------------Lock Password-------------------END-----

End Sub


Sub calcTDSFrmSal()
On Error Resume Next


SchTDS.setTblinfo_I2

Dim rangecells_I As Range
Set rangecells_I = Range("TDSal.TotalTDSSalary").Cells

Dim rangecells_I2 As Range
Set rangecells_I2 = Range("TDSal.IncChrgSalary").Cells

Dim i As Long
ReDim TotalTDSSal(end_I2)
ReDim IncChrgSal(end_I2)
 
 For i = 1 To end_I2
 
    TotalTDSSal(i) = rangecells_I.item(i).Value
    IncChrgSal(i) = rangecells_I2.item(i).Value
 
    If TotalTDSSal(i) > IncChrgSal(i) Then
            TotalTDSSal(i) = 0
        TDS = TDS + TotalTDSSal(i)
    Else
         TDS = TDS + TotalTDSSal(i)
    End If
Next

End Sub

Sub calcTCSFrmSal()
On Error Resume Next


SchTCS.setTblinfo_TCS

Dim rangecells_I As Range
Set rangecells_I = Range("TCS.TotalTCS").Cells

Dim rangecells_I2 As Range
Set rangecells_I2 = Range("TCS.AmtTCSClaimedThisYear").Cells

'Dim rangecells_I3 As Range
'Set rangecells_I3 = Range("TCS.AmtClaimedBySpouse").Cells




Dim i As Long
ReDim totalTCSSal(end_TCS)
ReDim incChrgSalTCS(end_TCS)
'ReDim incChrgSalTCSown(end_TCS)

 
 For i = 1 To end_TCS
 
    totalTCSSal(i) = rangecells_I.item(i).Value
    incChrgSalTCS(i) = rangecells_I2.item(i).Value
    'incChrgSalTCSown(i) = rangecells_I3.Item(i).Value
    
     
             If (incChrgSalTCS(i) > totalTCSSal(i)) Then
                                incChrgSalTCS(i) = 0
               ' amtClaimedBySpouse(i) = 0
                TCS = TCS + incChrgSalTCS(i)
            Else
                TCS = TCS + incChrgSalTCS(i)
            End If
    
Next

End Sub

Sub calcTDSOtherSal()
On Error Resume Next

'Dim portuVal As String

'portuVal = Sheet1.Range("Sheet1.PortugeseCC5A").Value


setTableInfo_Grid2

Dim rangecells_I As Range
Set rangecells_I = Range("TDSoth.TotTDSOnAmtPaid").Cells

Dim rangecells_I2 As Range
Set rangecells_I2 = Range("TDSoth.6income").Cells

'Dim rangecells_I3 As Range
'Set rangecells_I3 = Range("TDSoth.6TDS").Cells

ReDim claimOutOfTotTDSOnAmtPaid(ColCount1)
ReDim totTDSOnAmtPaid(ColCount1)
'ReDim amtClaimedBySpouse(ColCount1)
Dim i As Long

'If portuVal = "No" Then
'    For i = 1 To ColCount1_3
'
'    claimOutOfTotTDSOnAmtPaid(i) = rangecells_I.Item(i).Value
'    totTDSOnAmtPaid(i) = rangecells_I2.Item(i).Value
'
'        If claimOutOfTotTDSOnAmtPaid(i) > totTDSOnAmtPaid(i) Then
'                      claimOutOfTotTDSOnAmtPaid(i) = 0
'
'            TDS = TDS + claimOutOfTotTDSOnAmtPaid(i)
'        Else
'            TDS = TDS + claimOutOfTotTDSOnAmtPaid(i)
'        End If
'    Next

    For i = 1 To ColCount1
        claimOutOfTotTDSOnAmtPaid(i) = rangecells_I.item(i).Value
        totTDSOnAmtPaid(i) = rangecells_I2.item(i).Value
        'amtClaimedBySpouse(i) = rangecells_I3.Item(i).Value
        
             If (claimOutOfTotTDSOnAmtPaid(i) < totTDSOnAmtPaid(i)) Then
                                totTDSOnAmtPaid(i) = 0
                'amtClaimedBySpouse(i) = 0
                TDS = TDS + totTDSOnAmtPaid(i)
            Else
                TDS = TDS + totTDSOnAmtPaid(i)
            End If
    Next
'End If

End Sub

'NEW TDS3

Sub calcTDS3OtherSal()
On Error Resume Next

Dim portuVal As String

'portuVal = Sheet1.Range("Sheet1.PortugeseCC5A").Value


setTableInfo_Grid4

Dim rangecells1_I As Range
Set rangecells1_I = Range("TDS26QB.TotTDSOnAmtPaid").Cells

Dim rangecells1_I2 As Range
Set rangecells1_I2 = Range("TDS26QB.6income").Cells

'Dim rangecells1_I3 As Range
'Set rangecells1_I3 = Range("TDS26QB.6TDS").Cells

ReDim claimOutOfTotTDS3OnAmtPaid(ColCount3)
ReDim totTDS3OnAmtPaid(ColCount3)
'ReDim amtClaimedBySpouse3(ColCount3)
Dim i As Long

'If portuVal = "No" Then
'    For i = 1 To ColCount3_3
'
'    claimOutOfTotTDS3OnAmtPaid(i) = rangecells1_I.Item(i).Value
'    totTDS3OnAmtPaid(i) = rangecells1_I2.Item(i).Value
'
'        If claimOutOfTotTDS3OnAmtPaid(i) > totTDS3OnAmtPaid(i) Then
'                      claimOutOfTotTDS3OnAmtPaid(i) = 0
'
'            TDS = TDS + claimOutOfTotTDS3OnAmtPaid(i)
'        Else
'            TDS = TDS + claimOutOfTotTDS3OnAmtPaid(i)
'        End If
'    Next
'Else
    For i = 1 To ColCount3
        claimOutOfTotTDS3OnAmtPaid(i) = rangecells1_I.item(i).Value
        totTDS3OnAmtPaid(i) = rangecells1_I2.item(i).Value
        'amtClaimedBySpouse3(i) = rangecells1_I3.Item(i).Value
        
             If (claimOutOfTotTDS3OnAmtPaid(i) < totTDS3OnAmtPaid(i)) Then
                                totTDS3OnAmtPaid(i) = 0
                'amtClaimedBySpouse3(i) = 0
                TDS = TDS + totTDS3OnAmtPaid(i)
            Else
                TDS = TDS + totTDS3OnAmtPaid(i)
            End If
    Next
'End If

End Sub

Sub calcAdvSelfTax()
On Error Resume Next

SchTDS.setTableInfo1_Grid3


Dim rangecells_I As Range
Set rangecells_I = Range("TaxP.DateDep").Cells

Dim rangecells_I2 As Range
Set rangecells_I2 = Range("TaxP.Amt").Cells

Dim i As Long
ReDim taxDepDate(ColCount2_1)
ReDim taxPAmt(ColCount2_1)
 
 For i = 1 To ColCount2_1
  
    taxDepDate(i) = rangecells_I.item(i).Value
    taxPAmt(i) = rangecells_I2.item(i).Value
 'dpk101
 'TAX DETAILS-C5 - Changed from 2021 to 2022 & 2023 to 2024 Bindu 2024-25 AY
 'Year changed by Ankita on 06/12/2024
 
 'Year changed by Ankita on 29/01/2026
    If ((EfilingCommon.checkFirstDateBefore("01/04/2024", taxDepDate(i))) And _
     (EfilingCommon.checkFirstDateBefore(taxDepDate(i), "31/03/2026"))) Then
        AdvanceTax = AdvanceTax + taxPAmt(i)
    ElseIf (EfilingCommon.checkFirstDateBefore("01/04/2026", taxDepDate(i))) Then
          SelfAssessmentTax = SelfAssessmentTax + taxPAmt(i)
    End If
  'dpk101
2
    If (age > 59) Then
'        If (BalTaxPayable - AdvanceTax - TDS - TCS - Sheet2.Range("IT.Sat1").Value - Sheet2.Range("IT.Sat2").Value - Sheet2.Range("IT.Sat3").Value - Sheet2.Range("IT.Sat4").Value) > 100000 Then
'Ayush_DueDate_08/09/2025
        If (BalTaxPayable - AdvanceTax - TDS - TCS - Sheet2.Range("IT.Sat1").Value - Sheet2.Range("IT.Sat2").Value - Sheet2.Range("IT.Sat3").Value - Sheet2.Range("IT.Sat4").Value - Sheet2.Range("ExSAT").Value) > 100000 Then
           'TAX DETAILS-C5 Change 2023 to 2024 Bindu 2024-25 AY
           'Year Changed by Ankita on 06/12/2024
'           If ((EfilingCommon.checkFirstDateBefore("01/04/2025", taxDepDate(i))) And (EfilingCommon.checkFirstDateBefore(taxDepDate(i), "31/07/2025"))) Then
'           If ((EfilingCommon.checkFirstDateBefore("01/04/2025", taxDepDate(i))) And (EfilingCommon.checkFirstDateBefore(taxDepDate(i), "15/09/2025"))) Then  'Ankita_10/07/2025
'Ayush_DueDate_08/09/2025
            If ((EfilingCommon.checkFirstDateBefore("01/04/2026", taxDepDate(i))) And (EfilingCommon.checkFirstDateBefore(taxDepDate(i), Sheet5.Range("DueDate1").Value))) Then
                 selfAssessmentTax234A = selfAssessmentTax234A + taxPAmt(i)
           End If
        Else
        'TAX DETAILS-C5 Change 2023 to 2024 Bindu 2024-25 AY
            'Year Changed by Ankita on 06/12/2024
'           If ((EfilingCommon.checkFirstDateBefore("01/04/2025", taxDepDate(i))) And (EfilingCommon.checkFirstDateBefore(taxDepDate(i), "15/09/2025"))) Then  'Ankita_10/07/2025
'Ayush_DueDate_08/09/2025
            If ((EfilingCommon.checkFirstDateBefore("01/04/2026", taxDepDate(i))) And (EfilingCommon.checkFirstDateBefore(taxDepDate(i), Sheet5.Range("DueDate1").Value))) Then
                 selfAssessmentTax234A = selfAssessmentTax234A + taxPAmt(i)
           End If
        End If
    Else
        If (BalTaxPayable - AdvanceTax - TDS - TCS) > 100000 Then
       ' TAX DETAILS-C5 Change 2023 to 2024 Bindu 2024-25 AY
            'Year Changed by Ankita on 06/12/2024
'           If ((EfilingCommon.checkFirstDateBefore("01/04/2025", taxDepDate(i))) And (EfilingCommon.checkFirstDateBefore(taxDepDate(i), "15/09/2025"))) Then  'Ankita_10/07/2025
'Ayush_DueDate_08/09/2025
            If ((EfilingCommon.checkFirstDateBefore("01/04/2026", taxDepDate(i))) And (EfilingCommon.checkFirstDateBefore(taxDepDate(i), Sheet5.Range("DueDate1").Value))) Then
                 selfAssessmentTax234A = selfAssessmentTax234A + taxPAmt(i)
           End If
        Else
            ' TAX DETAILS-C5 Change 2023 to 2024 Bindu 2024-25 AY
            'Year Changed by Ankita on 06/12/2024
'           If ((EfilingCommon.checkFirstDateBefore("01/04/2025", taxDepDate(i))) And (EfilingCommon.checkFirstDateBefore(taxDepDate(i), "15/09/2025"))) Then  'Ankita_10/07/2025
'Ayush_DueDate_08/09/2025
            If ((EfilingCommon.checkFirstDateBefore("01/04/2026", taxDepDate(i))) And (EfilingCommon.checkFirstDateBefore(taxDepDate(i), Sheet5.Range("DueDate1").Value))) Then
                 selfAssessmentTax234A = selfAssessmentTax234A + taxPAmt(i)
           End If
        End If
    End If
  
Next
advanceTaxToDisplay = AdvanceTax
SATtoDisplay = SelfAssessmentTax

End Sub

Sub calcIntrst234C()
On Error Resume Next
 
Dim intrst234C0i As Double
Dim intrst234Ci As Double
Dim intrst234Cii As Double
Dim intrst234Ciii As Double
Dim intrst234Civ As Double
 
Dim tempintrst234C0i As Double
Dim tempintrst234Ci As Double
Dim tempintrst234Cii As Double
Dim tempintrst234Ciii As Double
Dim tempintrst234Civ As Double
Dim temp12PerQtr1, temp36PerQtr2 As Double
Dim sec89 As Double
Dim sec89A As Double
Dim baseTax As Double
 
    advanceTaxToDisplay = Sheet3.Range("IncD.AdvanceTax").Value
    TDSToDisplay = Sheet3.Range("IncD.TDS").Value
    TCSToDisplay = Sheet3.Range("IncD.TCS").Value
    SATtoDisplay = Sheet3.Range("IncD.SelfAssessmentTax").Value
    BalTaxPayable = Sheet1.Range("IncD.NetTaxLiability").Value
    baseTax = BalTaxPayable
    sec89 = Sheet1.Range("IncD.Section89").Value
    sec89A = Sheet1.Range("IncD.Section89A").Value
    Rebate87A = Sheet1.Range("IncD.Rebate87A").Value
    TDS = TDS + sec89 + sec89A
    BalTaxPayable = Sheet1.Range("IncD.Q1Tax").Value
    BalTaxPayable = BalTaxPayable - Rebate87A
    BalTaxPayable = BalTaxPayable * 1.04
'
'    BalTaxPayable = Sheet1.Range("IncD.NetTaxLiability").Value
    intrst234C0i = 0
    intrst234Ci = 0
    intrst234Cii = 0
    intrst234Ciii = 0
    intrst234Civ = 0
    temp12PerQtr1 = 0
    temp36PerQtr2 = 0
If ((BalTaxPayable - TDS - TCS) >= 0) Then
    temp12PerQtr1 = Application.WorksheetFunction.Floor(0.12 * (BalTaxPayable - TDS - TCS), 100)
'    MsgBox slab0
'    MsgBox slab1
'    MsgBox slab2
'    MsgBox slab3
'    MsgBox slab4
    If (slab0 < ((BalTaxPayable - TDS - TCS) * (0.15))) Then
        tempintrst234C0i = IIf(slab0 >= temp12PerQtr1, 0, ((BalTaxPayable - TDS - TCS) * (0.15)) - slab0)
        If tempintrst234C0i > 100 Then
            tempintrst234C0i = (Application.WorksheetFunction.RoundDown(tempintrst234C0i, -2))     'MRound(tempintrst234C0i, 100)
                    End If
            intrst234C0i = tempintrst234C0i * (0.01) * (3)

    End If
    BalTaxPayable = Sheet1.Range("IncD.Q2Tax").Value
    BalTaxPayable = BalTaxPayable - Rebate87A
    BalTaxPayable = BalTaxPayable * 1.04

   ' BalTaxPayable = Sheet1.Range("IncD.NetTaxLiability").Value
    temp36PerQtr2 = Application.WorksheetFunction.Floor(0.36 * (BalTaxPayable - TDS - TCS), 100)
    If (slab0 + slab1 < ((BalTaxPayable - TDS - TCS) * (0.45))) Then
        tempintrst234Ci = IIf((slab0 + slab1) >= temp36PerQtr2, 0, ((BalTaxPayable - TDS - TCS) * (0.45)) - slab0 - slab1)
        If tempintrst234Ci > 100 Then
            tempintrst234Ci = (Application.WorksheetFunction.RoundDown(tempintrst234Ci, -2))     'MRound(tempintrst234Ci, 100)
                    End If
            intrst234Ci = tempintrst234Ci * (0.01) * (3)
    End If
    BalTaxPayable = Sheet1.Range("IncD.Q3Tax").Value
    BalTaxPayable = BalTaxPayable - Rebate87A
    BalTaxPayable = BalTaxPayable * 1.04
    'BalTaxPayable = Sheet1.Range("IncD.NetTaxLiability").Value
    If (slab0 + slab1 + slab2) < ((BalTaxPayable - TDS - TCS) * (0.75)) Then
        tempintrst234Cii = ((BalTaxPayable - TDS - TCS) * (0.75)) - slab0 - slab1 - slab2
        If tempintrst234Cii > 100 Then
            tempintrst234Cii = Application.WorksheetFunction.RoundDown(tempintrst234Cii, -2)   'MRound(tempintrst234Cii, 100)
        End If
         intrst234Cii = tempintrst234Cii * (0.01) * 3
    End If
    BalTaxPayable = Sheet1.Range("IncD.Q4Tax").Value
    BalTaxPayable = BalTaxPayable - Rebate87A
    BalTaxPayable = BalTaxPayable * 1.04
'
    'BalTaxPayable = Sheet1.Range("IncD.NetTaxLiability").Value
    If ((slab0 + slab1 + slab2 + slab3) < (BalTaxPayable - TDS - TCS)) Then
            tempintrst234Ciii = (BalTaxPayable - TDS - TCS - slab0 - slab1 - slab2 - slab3)
        If tempintrst234Ciii > 100 Then
             tempintrst234Ciii = Application.WorksheetFunction.RoundDown(tempintrst234Ciii, -2)   'MRound(tempintrst234Ciii, 100)
        End If
        intrst234Ciii = tempintrst234Ciii * 0.01
    End If

    tempintrst234Civ = (slab0 + slab1 + slab2 + slab3) - BalTaxPayable - TDS - TCS
    If (tempintrst234Civ < 0) Then
            tempintrst234Civ = 0
            End If
    slab4 = slab4 + tempintrst234Civ
    BalTaxPayable = Sheet1.Range("IncD.Q5Tax").Value
    BalTaxPayable = BalTaxPayable - Rebate87A
    BalTaxPayable = BalTaxPayable * 1.04

 
    BalTaxPayable = BalTaxPayable - slab4
        If BalTaxPayable > 100 Then
             BalTaxPayable = Application.WorksheetFunction.RoundDown(BalTaxPayable, -2)   'MRound(BalTaxPayable, 100)
        End If
        intrst234Civ = BalTaxPayable * 0.01
    If (intrst234Civ < 0) Then
            intrst234Civ = 0
            End If
    Else
        intrst234C0i = 0
        intrst234Ci = 0
        intrst234Cii = 0
        intrst234Ciii = 0
        intrst234Civ = 0
End If
    intrst234C = intrst234C0i + intrst234Ci + intrst234Cii + intrst234Ciii + intrst234Civ
    If ((baseTax - TDS - TCS) < 10000) Then
    intrst234C = 0
    End If
'If (age > 59) And (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") Then
 
If (bacage > 59) Then
     intrst234C = 0
End If
'If (Sheet5.Range("BacValue").Value) = 1 And (age > 59) Then
'     intrst234C = 0
'End If
 
End Sub





Sub calcIntrst234C_Old()
On Error Resume Next



Dim intrst234C0i As Double
Dim intrst234Ci As Double
Dim intrst234Cii As Double
Dim intrst234Ciii As Double
Dim intrst234Civ As Double

Dim tempintrst234C0i As Double
Dim tempintrst234Ci As Double
Dim tempintrst234Cii As Double
Dim tempintrst234Ciii As Double
Dim tempintrst234Civ As Double
Dim temp12PerQtr1, temp36PerQtr2 As Double
Dim sec89 As Double
Dim sec89A As Double
Dim baseTax As Double

    advanceTaxToDisplay = Sheet3.Range("IncD.AdvanceTax").Value
    TDSToDisplay = Sheet3.Range("IncD.TDS").Value
    TCSToDisplay = Sheet3.Range("IncD.TCS").Value
    SATtoDisplay = Sheet3.Range("IncD.SelfAssessmentTax").Value
    BalTaxPayable = Sheet1.Range("IncD.NetTaxLiability").Value
    baseTax = BalTaxPayable
    
    sec89 = Sheet1.Range("IncD.Section89").Value
    sec89A = Sheet1.Range("IncD.Section89A").Value
'    Rebate87A = Sheet1.Range("IncD.Rebate87A").Value
    TDS = TDS + sec89 + sec89A
    BalTaxPayable = Sheet1.Range("IncD.Q1Tax").Value
    BalTaxPayable = BalTaxPayable * 1.04
    
    
    intrst234C0i = 0
    intrst234Ci = 0
    intrst234Cii = 0
    intrst234Ciii = 0
    intrst234Civ = 0
    
    temp12PerQtr1 = 0
    temp36PerQtr2 = 0
 
If ((BalTaxPayable - TDS - TCS) >= 0) Then
    
    temp12PerQtr1 = Application.WorksheetFunction.Floor(0.12 * (BalTaxPayable - TDS - TCS), 100)
'    MsgBox slab0
'    MsgBox slab1
'    MsgBox slab2
'    MsgBox slab3
'    MsgBox slab4
    
    If (slab0 < ((BalTaxPayable - TDS - TCS) * (0.15))) Then
    
        tempintrst234C0i = IIf(slab0 >= temp12PerQtr1, 0, ((BalTaxPayable - TDS - TCS) * (0.15)) - slab0)
        If tempintrst234C0i > 100 Then
            tempintrst234C0i = (Application.WorksheetFunction.RoundDown(tempintrst234C0i, -2))     'MRound(tempintrst234C0i, 100)
                    End If
            intrst234C0i = tempintrst234C0i * (0.01) * (3)
            
    
    
    End If
    BalTaxPayable = Sheet1.Range("IncD.Q2Tax").Value
    BalTaxPayable = BalTaxPayable * 1.04
    temp36PerQtr2 = Application.WorksheetFunction.Floor(0.36 * (BalTaxPayable - TDS - TCS), 100)
    
    If (slab0 + slab1 < ((BalTaxPayable - TDS - TCS) * (0.45))) Then
        
        tempintrst234Ci = IIf((slab0 + slab1) >= temp36PerQtr2, 0, ((BalTaxPayable - TDS - TCS) * (0.45)) - slab0 - slab1)
        If tempintrst234Ci > 100 Then
            tempintrst234Ci = (Application.WorksheetFunction.RoundDown(tempintrst234Ci, -2))     'MRound(tempintrst234Ci, 100)
                    End If
            intrst234Ci = tempintrst234Ci * (0.01) * (3)
    End If
    
    BalTaxPayable = Sheet1.Range("IncD.Q3Tax").Value
    BalTaxPayable = BalTaxPayable * 1.04
    
    If (slab0 + slab1 + slab2) < ((BalTaxPayable - TDS - TCS) * (0.75)) Then
    
        tempintrst234Cii = ((BalTaxPayable - TDS - TCS) * (0.75)) - slab0 - slab1 - slab2
        If tempintrst234Cii > 100 Then
            tempintrst234Cii = Application.WorksheetFunction.RoundDown(tempintrst234Cii, -2)   'MRound(tempintrst234Cii, 100)
        End If
         intrst234Cii = tempintrst234Cii * (0.01) * 3
    End If
    BalTaxPayable = Sheet1.Range("IncD.Q4Tax").Value
    BalTaxPayable = BalTaxPayable * 1.04
    
    
    If ((slab0 + slab1 + slab2 + slab3) < (BalTaxPayable - TDS - TCS)) Then
    
            tempintrst234Ciii = (BalTaxPayable - TDS - TCS - slab0 - slab1 - slab2 - slab3)
        If tempintrst234Ciii > 100 Then
             tempintrst234Ciii = Application.WorksheetFunction.RoundDown(tempintrst234Ciii, -2)   'MRound(tempintrst234Ciii, 100)
        End If
        intrst234Ciii = tempintrst234Ciii * 0.01
    End If
    
   
    tempintrst234Civ = (slab0 + slab1 + slab2 + slab3) - BalTaxPayable - TDS - TCS
    If (tempintrst234Civ < 0) Then
            tempintrst234Civ = 0
            End If
    slab4 = slab4 + tempintrst234Civ
    BalTaxPayable = Sheet1.Range("IncD.Q5Tax").Value
    BalTaxPayable = BalTaxPayable * 1.04
    BalTaxPayable = BalTaxPayable - slab4
        If BalTaxPayable > 100 Then
             BalTaxPayable = Application.WorksheetFunction.RoundDown(BalTaxPayable, -2)   'MRound(BalTaxPayable, 100)
        End If
        
        intrst234Civ = BalTaxPayable * 0.01
    If (intrst234Civ < 0) Then
            intrst234Civ = 0
            End If
     
Else
        intrst234C0i = 0
        intrst234Ci = 0
        intrst234Cii = 0
        intrst234Ciii = 0
        intrst234Civ = 0
End If
    
    intrst234C = intrst234C0i + intrst234Ci + intrst234Cii + intrst234Ciii + intrst234Civ
    If ((baseTax - TDS - TCS) < 10000) Then
    intrst234C = 0
    End If
           
'If (age > 59) And (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") Then

If (bacage > 59) Then
     intrst234C = 0
End If
'If (Sheet5.Range("BacValue").Value) = 1 And (age > 59) Then
'     intrst234C = 0
'End If

End Sub

Sub calcIntrst234B()
On Error Resume Next

Dim myCount As Variant
Dim intrst234Bi As Double
Dim earliestSelfAsspaidDate As Double
Dim noOfMonthsTillSelfasst As Double
Dim intrst234Bprinciple As Double

earliestSelfAsspaidDate = 0
noOfMonthsTillSelfasst = 0
Dim AdditionalTax As Variant
AdditionalTax = 0

If Sheet202.Range("U_Refund").Value <> "" Then
AdditionalTax = Sheet202.Range("U_Refund").Value
End If

If Sheet202.Range("U_TotRefund") <> "" Then
AdditionalTax = AdditionalTax + Sheet202.Range("U_TotRefund").Value
End If

If Mid(Range("sheet1.ReturnFileSec"), 1, 7) = "139(8A)" Then
BalTaxPayable = BalTaxPayable + AdditionalTax
End If
If (BalTaxPayable - TDS) >= 10000 Then
   If AdvanceTax < ((BalTaxPayable - TDS - TCS) * 0.9) Then
        intrst234Bprinciple = BalTaxPayable - AdvanceTax - TDS - TCS
        If intrst234Bprinciple > 100 Then
            intrst234Bprinciple = (Application.WorksheetFunction.RoundDown(intrst234Bprinciple, -2))   'MRound(intrst234Bprinciple, 100)
        End If
        
    'Interest 234B first part calc
        'Dim selfAsspaidDates() As Variant
        'Dim selfAsspaidAmts() As Variant
        Dim selfAsspaidDates As Variant
        Dim selfAsspaidAmts As Variant
        
        Dim x As Long
        Dim tempDate As Variant
        Dim tempAmt As Double
        Dim a As Long
            x = 0
            a = 0
            tempDate = 0
            tempAmt = 0
    
        SchTDS.setTableInfo_Grid3
f
        Dim rangecells_I As Range
        Set rangecells_I = Range("TaxP.DateDep").Cells
    
        Dim rangecells_I2 As Range
        Set rangecells_I2 = Range("TaxP.Amt").Cells
        Dim i As Long
 
        'to get all self assesment tax values
        ReDim taxDepDate(ColCount2)
        ReDim taxPAmt(ColCount2)
         
         For i = 1 To ColCount2
            taxDepDate(i) = rangecells_I.item(i).Value
            taxPAmt(i) = rangecells_I2.item(i).Value
         
            If (EfilingCommon.checkFirstDateBefore("01/04/2024", taxDepDate(i)) And checkFirstDateBefore(taxDepDate(i), currentdate)) Then
            If (taxPAmt(i) > 0) Then
                     a = a + 1
            End If
            End If
        Next
        
       'If a > 0 Then
        
        ReDim selfAsspaidDates(a)
        ReDim selfAsspaidAmts(a)
        
       'End If
       
        For i = 1 To ColCount2
         
            taxDepDate(i) = rangecells_I.item(i).Value
            taxPAmt(i) = rangecells_I2.item(i).Value
         
            If (EfilingCommon.checkFirstDateBefore("01/04/2024", taxDepDate(i)) And checkFirstDateBefore(taxDepDate(i), currentdate)) Then
               If (taxPAmt(i) > 0) Then
                    selfAsspaidDates(x) = taxDepDate(i)
                    selfAsspaidAmts(x) = taxPAmt(i)
                    x = x + 1
                End If
            End If
        Next
    
        'To sort all self assesment tax values according to date
        Dim j As Long
        Dim k As Long
        myCount = 0
        For i = 0 To UBound(selfAsspaidDates)
            If selfAsspaidDates(i) <> "" Then myCount = myCount + 1
        Next i
        'If Application.WorksheetFunction.CountA(selfAsspaidDates) > 1 Then
        If myCount > 1 Then  'UBound(selfAsspaidDates) >= 1 Then
            For j = 0 To (UBound(selfAsspaidDates))
                      For k = j + 1 To myCount - 1 'UBound(selfAsspaidDates) '(Application.WorksheetFunction.CountA(selfAsspaidDates) - 1)
                        If EfilingCommon.checkFirstDateBefore(selfAsspaidDates(j), selfAsspaidDates(k)) Then
                        Else
                            tempDate = selfAsspaidDates(j)
                            tempAmt = selfAsspaidAmts(j)

                            selfAsspaidDates(j) = selfAsspaidDates(k)
                            selfAsspaidAmts(j) = selfAsspaidAmts(k)

                            selfAsspaidDates(k) = tempDate
                            selfAsspaidAmts(k) = tempAmt
                        End If
                   Next
                Next
            
        Dim lastMonth As Long
        Dim lastIndex As Long
    
            lastMonth = 0
            lastIndex = -1
        For j = 0 To UBound(selfAsspaidDates) - 1
            If val(Mid(selfAsspaidDates(j), 4, 2)) = lastMonth Then
                selfAsspaidAmts(lastIndex) = selfAsspaidAmts(lastIndex) + selfAsspaidAmts(j)

            Else
                lastMonth = val(Mid(selfAsspaidDates(j), 4, 2))
                lastIndex = lastIndex + 1
                selfAsspaidAmts(lastIndex) = selfAsspaidAmts(j)
                selfAsspaidDates(lastIndex) = selfAsspaidDates(j)
            End If
        Next
 

           lastIndex = lastIndex + 1
          If lastIndex > 0 Then
                 ReDim Preserve selfAsspaidAmts(lastIndex - 1)
                ReDim Preserve selfAsspaidDates(lastIndex - 1)
         End If
        End If
        
        'If Application.WorksheetFunction.CountA(selfAsspaidDates) = 0 Then
        myCount = 0
        For i = 0 To UBound(selfAsspaidDates)
            If selfAsspaidDates(i) <> "" Then myCount = myCount + 1
        Next i
        If myCount = 0 Then  'UBound(selfAsspaidDates) = 0 Then
            noOfMonthsTillSelfasst = EfilingCommon.calcNoOfMonths(currentdate, "01/04/2024")
        Else
            noOfMonthsTillSelfasst = EfilingCommon.calcNoOfMonths(selfAsspaidDates(0), "01/04/2024")
        End If
            intrst234Bi = intrst234Bprinciple * (0.01) * noOfMonthsTillSelfasst
         
        
        'Interest 234B second part calc
        
        Dim intrst234Bprinciple2 As Double
        Dim selfAsspart As Double
        Dim noOfMonthsTillSelfasst2 As Long
        Dim intrst234Bii As Double
        Dim partialSelfAssPaid As Double
        Dim interestFrom As Variant
        Dim interestTill As Variant
             
         intrst234Bprinciple2 = 0
         selfAsspart = 0
         intrst234Bii = 0
         partialSelfAssPaid = 0
         k = 0
        
          Dim dateCount As Long
         'dateCount = Application.WorksheetFunction.CountA(selfAsspaidDates)
         dateCount = myCount 'UBound(selfAsspaidDates)
        
         If dateCount <> 0 Then
            For i = 0 To myCount - 1 'UBound(selfAsspaidDates) 'Application.WorksheetFunction.CountA(selfAsspaidDates) - 1
            
                partialSelfAssPaid = partialSelfAssPaid + selfAsspaidAmts(i)
                intrst234Bprinciple2 = Application.WorksheetFunction.Floor(BalTaxPayable - AdvanceTax - TDS, 100) + intrst234A + intrst234C + intrst234F + intrst234Bi + intrst234Bii - partialSelfAssPaid
                    
                    If (intrst234Bprinciple2 < 0) Then
                        intrst234Bprinciple2 = 0
                    End If
                    
                    If (intrst234Bprinciple2 > 100) Then
                        intrst234Bprinciple2 = (Application.WorksheetFunction.RoundDown(intrst234Bprinciple2, -2)) 'MRound(intrst234Bprinciple2, 100)
                    End If
                    
                    interestTill = currentdate
                    interestFrom = selfAsspaidDates(i)
                    
                   If i <> myCount - 1 Then 'UBound(selfAsspaidDates) Then '(Application.WorksheetFunction.CountA(selfAsspaidDates) - 1) Then
                        For k = i To myCount - 1   'UBound(selfAsspaidDates) '(Application.WorksheetFunction.CountA(selfAsspaidDates) - 1)
                            If selfAsspaidDates(k) <> selfAsspaidDates(k + 1) Then
                                interestTill = selfAsspaidDates(k + 1)
                                interestFrom = selfAsspaidDates(k)
                                k = (UBound(selfAsspaidDates) + 1)
                            End If
                        Next
                  End If
                noOfMonthsTillSelfasst2 = EfilingCommon.calcNoOfMonths(interestTill, interestFrom) - 1
                
                If intrst234Bprinciple2 < intrst234Bprinciple Then
                        intrst234Bii = intrst234Bii + (intrst234Bprinciple2 * (0.01) * noOfMonthsTillSelfasst2)
                Else
                        intrst234Bii = intrst234Bii + (intrst234Bprinciple * (0.01) * noOfMonthsTillSelfasst2)
                End If
            Next
        End If
    intrst234B = intrst234Bi + intrst234Bii
    
    End If
Else

    intrst234B = 0
End If
   
   ' If (age > 59) And (resStatus = "RES-Resident" Or resStatus = "NOR-Resident but not ordinarily Resident") Then
   
    If (bacage > 59) Then
            intrst234B = 0
    End If

End Sub

Sub NextIncD_Click()
Dim sourceSheet As Worksheet
    'Ankita  11/11/2024
    If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
       Sheet201.Activate
    Else:
'---------------------
'Ankita_08/05/2025
' Set sourceSheet = ThisWorkbook.Sheets("TDS")
'  Set sourceSheet = ThisWorkbook.Sheets("Schedule EA 10(13A)")
  Set sourceSheet = ThisWorkbook.Sheets("HP")

    sourceSheet.Activate
        End If 'Ankita 11/11/2024
End Sub

Sub calcTaxPayable15Minus17()
On Error Resume Next


 Dim taxPayable15M17 As Double
 Dim refund15M17 As Double
 Dim totTaxPaid As Double
 Dim totTaxIntrstPay As Double

totTaxPaid = Sheet3.Range("IncD.TotalTaxesPaid").Value
totTaxIntrstPay = Sheet1.Range("IncD.TotTaxPlusIntrstPay").Value

taxPayable15M17 = Sheet3.Range("IncD.balTaxPayable").Value
refund15M17 = Sheet3.Range("IncD.RefundDue").Value

 If totTaxPaid <= totTaxIntrstPay Then
     refund15M17 = 0
     taxPayable15M17 = totTaxIntrstPay - totTaxPaid
     taxPayable15M17 = Round(taxPayable15M17 / 10, 0) * 10
Else
        refund15M17 = totTaxPaid - totTaxIntrstPay
        refund15M17 = Round(refund15M17 / 10, 0) * 10
        taxPayable15M17 = 0
        taxStatus = "TR"
        'Sheet1.Range("sheet1.TaxStatus").Value = "Tax Refundable"
        
End If

 
If taxPayable15M17 > 0 Then
        taxStatus = "TP"
        'Sheet1.Range("sheet1.TaxStatus").Value = "Tax Payable"
        
ElseIf refund15M17 > 0 Then
        taxStatus = "TR"
        'Sheet1.Range("sheet1.TaxStatus").Value = "Tax Refundable"
Else
        taxStatus = "NT"
        'Sheet1.Range("sheet1.TaxStatus").Value = "Nil Tax Payable"
        
End If
 End Sub
Function checkfieldspecialcharacter2(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter2 = True
    Dim arr As Variant
    arr = Array("@", "-", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter2 = False
            Exit Function
        End If
        Next
    Next
End Function

Function ValidatePhoneNoWork() As Boolean
    Dim phoneNo, stDcode As String
    Dim i As Long
    
    ValidatePhoneNoWork = True
    phoneNo = Sheet1.Range("sheet1.PhoneNo").Value
    stDcode = Sheet1.Range("sheet1.STDcode").Value

    If Trim(phoneNo) <> "" Then
        For i = 1 To Len(phoneNo)
            If Not IsNumeric(Mid(phoneNo, i, 1)) Then
              fmsgbox ("* PhoneNo in Sheet PI  must contain only digits from 0 to 9 in Income Details")
            ValidatePhoneNoWork = False
            Exit Function
            End If
        Next
    End If

    If (Trim(phoneNo) <> "" Or Trim(stDcode) <> "") Then
        If (Len(phoneNo) + Len(stDcode)) <> 10 Then
          fmsgbox ("* Invalid Phone Number. STD Code + Landline Number should be 10 digits and cannot begin with '0' in Income Details.")
            ValidatePhoneNoWork = False
            Exit Function
        End If
        
        If Mid((phoneNo), 1, 1) = 0 Then
            fmsgbox ("* Invalid Phone Number. STD Code + Landline Number should be 10 digits and cannot begin with '0' in Income Details.")
            ValidatePhoneNoWork = False
            Exit Function
        End If

End If

End Function

'Function ValidateSelect80D() As Boolean
'ValidateSelect80D = True
'    Dim Usr80DVal As Variant
'    Dim dob_2 As Variant
'
'    SELECT80D = Sheet1.Range("SELECT80D").Value
'    SELECT80D = Mid(SELECT80D, 1, 1)
'    If isdropdownblank(SELECT80D) Then
'        SELECT80D = ""
'    End If
'
'    dob_2 = Sheet1.Range("sheet1.DOB").Value
'
'    Usr80DVal = Sheet1.Range("IncD.Section80D").Value
'
'    If SELECT80D <> "" Then
'        If IsEmpty(Usr80DVal) Or Usr80DVal = 0 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter the value for Part A Sec 80D deduction under Chapter VIA in Income Details" & Chr(13)
'            ValidateSelect80D = False
'            Exit Function
'        End If
'    End If
'
'    Usr80DVal = IIf(Trim(Usr80DVal) = "", 0, Usr80DVal)
'
'    If Usr80DVal > 0 Then
'        If isdropdownblank(SELECT80D) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option at dropdown of Part A Sec 80D (Health Insurance Premium) deduction under Chapter VIA in Income Details" & Chr(13)
'            ValidateSelect80D = False
'            Exit Function
'        End If
'    End If
'
'
'    If (SELECT80D = "7" And calculateAge(dob_2) <= 59) Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select a valid option from the dropdown of Part A Sec 80D (Health Insurance Premium) deduction under Chapter VIA in Income Details" & Chr(13)
'        ValidateSelect80D = False
'        Exit Function
'    End If
'
''    If (SELECT80D = "2" And calculateAge(dob_2) <= 59) Then
''        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Please select a valid option from the dropdown of Part A Sec 80D (Health Insurance Premium) deduction under Chapter VIA" & Chr(13)
''        ValidateSelect80D = False
''        Exit Function
''    End If
'
'End Function


'Function ValidateSelect80DB() As Boolean
'ValidateSelect80DB = True
'    Dim Usr80DBVal As Variant
'    Dim dob_2 As Variant
'
'    SELECT80DB = Sheet1.Range("SELECT80DB").Value
'    SELECT80DB = Mid(SELECT80DB, 1, 1)
'    If isdropdownblank(SELECT80DB) Then
'        SELECT80DB = ""
'    End If
'
'    dob_2 = Sheet1.Range("sheet1.DOB").Value
'
'    Usr80DBVal = Sheet1.Range("IncD.Section80DB").Value
'
'    If SELECT80DB <> "" Then
'        If IsEmpty(Usr80DBVal) Or Usr80DBVal = 0 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter the value for Part B Sec 80D (Medical Expenditure) deduction under Chapter VIA in Income Details" & Chr(13)
'            ValidateSelect80DB = False
'            Exit Function
'        End If
'    End If
'
'    Usr80DBVal = IIf(Trim(Usr80DBVal) = "", 0, Usr80DBVal)
'
'    If Usr80DBVal > 0 Then
'        If isdropdownblank(SELECT80DB) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option at dropdown of Part B Sec 80D (Medical Expenditure) Deduction under Chapter VIA in Income Details" & Chr(13)
'            ValidateSelect80DB = False
'            Exit Function
'        End If
'    End If
'
''    If ((SELECT80DB = "1" Or SELECT80DB = "2" Or SELECT80DB = "3") And calculateAge(dob_2) <= 59) Then
''        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Please select a valid option from the dropdown of Part A Sec.80D (Medical Expenditure) deduction under Chapter VIA." & Chr(13)
''        ValidateSelect80DB = False
''        Exit Function
''    End If
''
''    If ((SELECT80DB = "1") And calculateAge(dob_2) <= 59) Then
''        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Please select a valid option from the dropdown of Part A Sec.80D (Medical Expenditure) deduction under Chapter VIA." & Chr(13)
''        ValidateSelect80DB = False
''        Exit Function
''    End If
'
'
''    If (SELECT80DB = "3" And calculateAge(dob_2) < 80) Or (SELECT80DB = "1" And calculateAge(dob_2) < 80) Then
''        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Please select a valid option from the dropdown of Part B Sec 80D Deduction under Chapter VIA" & Chr(13)
''        ValidateSelect80DB = False
''        Exit Function
''    End If
'
'
'End Function
'
'
'Function ValidateSelect80DC() As Boolean
'ValidateSelect80DC = True
'    Dim Usr80DCVal As Variant
'    Dim dob_2 As Variant
'
'    SELECT80DC = Sheet1.Range("SELECT80DC").Value
'    SELECT80DC = Mid(SELECT80DC, 1, 1)
'    If isdropdownblank(SELECT80DC) Then
'        SELECT80DC = ""
'    End If
'
'    dob_2 = Sheet1.Range("sheet1.DOB").Value
'
'    Usr80DCVal = Sheet1.Range("IncD.Section80DC").Value
'
'    If SELECT80DC <> "" Then
'        If IsEmpty(Usr80DCVal) Or Usr80DCVal = 0 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter the value for Part C Sec 80D (Preventive Health check-up) deduction under Chapter VIA in Income Details" & Chr(13)
'            ValidateSelect80DC = False
'            Exit Function
'        End If
'    End If
'
'    Usr80DCVal = IIf(Trim(Usr80DCVal) = "", 0, Usr80DCVal)
'
'    If Usr80DCVal > 0 Then
'        If isdropdownblank(SELECT80DC) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option at dropdown of Part C Sec 80D (Preventive Health check-up) deduction under Chapter VIA in Income Details" & Chr(13)
'            ValidateSelect80DC = False
'            Exit Function
'        End If
'    End If
'
'End Function

'Function ValidateSystem80D() As Boolean
'ValidateSystem80D = True
'    Dim Usr80DVal As Variant
'    Dim Usr80DBVal As Variant
'    Dim Usr80DCVal As Variant
'    Dim SYSTEM80DB As Variant
'
'    Usr80DVal = Sheet1.Range("IncD.Section80D").Value
'    Usr80DBVal = Sheet1.Range("IncD.Section80DB").Value
'    Usr80DCVal = Sheet1.Range("IncD.Section80DC").Value
'    SYSTEM80DB = Sheet1.Range("IncD.Section80D_Calc").Value
'
'    Usr80DVal = IIf(Trim(Usr80DVal) = "", 0, Usr80DVal)
'    Usr80DBVal = IIf(Trim(Usr80DBVal) = "", 0, Usr80DBVal)
'    Usr80DCVal = IIf(Trim(Usr80DCVal) = "", 0, Usr80DCVal)
'    SYSTEM80DB = IIf(Trim(SYSTEM80DB) = "", 0, SYSTEM80DB)
'
'    If SYSTEM80DB > (Usr80DVal + Usr80DBVal + Usr80DCVal) Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Deduction u/s 80D should not be more than sum of amount claimed at 'Health Insurance, Medical Expenditure and Preventive Health Check Up' under Chapter VIA in Income Details" & Chr(13)
'        ValidateSystem80D = False
'        Exit Function
'    End If
'
'End Function

Function ValidateSelect80DD() As Boolean
ValidateSelect80DD = True
    Dim Usr80DDVal As Variant
    Dim dob_2 As Variant
    
    SELECT80DD = Sheet1.Range("SELECT80DD").Value
    SELECT80DD = Mid(SELECT80DD, 1, 1)
    If isdropdownblank(SELECT80DD) Then
        SELECT80DD = ""
    End If
    
    dob_2 = Sheet1.Range("sheet1.DOB").Value
    
    Usr80DDVal = Sheet1.Range("IncD.Section80DD").Value
    
    If SELECT80DD <> "" Then
        If IsEmpty(Usr80DDVal) Or Usr80DDVal = 0 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter the value for 80DD deduction under Chapter VIA in Income Details" & Chr(13)
            ValidateSelect80DD = False
            Exit Function
        End If
    End If
    
    Usr80DDVal = IIf(Trim(Usr80DDVal) = "", 0, Usr80DDVal)
        
    If Usr80DDVal > 0 Then
        If isdropdownblank(SELECT80DD) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option at dropdown of Sec.80DD under Chapter VIA in Income Details" & Chr(13)
            ValidateSelect80DD = False
            Exit Function
        End If
    End If

End Function

Function ValidateSELECT80DDS() As Boolean
ValidateSELECT80DDS = True
    Dim Usr80DDBVal As Variant
    Dim dob_2 As Variant
    Dim Sheet1_Specified_Disease As Variant  '25/05/2025
    
    Sheet1_Specified_Disease = Sheet1.Range("Sheet1.Specified_Disease").Value    '25/05/2025
      
    SELECT80DDS = Sheet1.Range("SELECT80DDS").Value
    SELECT80DDS = Mid(SELECT80DDS, 1, 1)
    If isdropdownblank(SELECT80DDS) Then
        SELECT80DDS = ""
    End If
    
    dob_2 = Sheet1.Range("sheet1.DOB").Value
    
    Usr80DDBVal = Sheet1.Range("IncD.Section80DDB").Value
    
    If SELECT80DDS <> "" Then
        If IsEmpty(Usr80DDBVal) Or Usr80DDBVal = 0 Then
            'Changed as per DE sheet v0.8 by Ankita
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter the value for 80DDB deduction under Chapter VIA in Income Details" & Chr(13)
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Please provide mandatory details under 80DDB deduction.""" & Chr(13)
             ValidateSELECT80DDS = False
             Exit Function
        End If
    End If
    
    

  
    Usr80DDBVal = IIf(Trim(Usr80DDBVal) = "", 0, Usr80DDBVal)
        
    If Usr80DDBVal > 0 Then
        If isdropdownblank(SELECT80DDS) Then
            'Changed as per DE sheet v0.8 by Ankita
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option at dropdown of Sec.80DDB under Chapter VIA in Income Details" & Chr(13)
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Please provide mandatory details under 80DDB deduction.""" & Chr(13)
            ValidateSELECT80DDS = False
            Exit Function
        End If
        
        '25/04/2025-------start
    ElseIf Sheet1_Specified_Disease <> "(Select)" And Sheet1_Specified_Disease <> "" Then
        If isdropdownblank(SELECT80DDS) Then
            'Changed as per DE sheet v0.8 by Ankita
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option at dropdown of Sec.80DDB under Chapter VIA in Income Details" & Chr(13)
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please provide mandatory details under 80DDB deduction.""" & Chr(13)
            ValidateSELECT80DDS = False
            Exit Function
        End If
    '------------END
    End If
    
    
'    If (SELECT80DDB = "1" And calculateAge(dob_2) > 59) Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Please select a valid option from the dropdown of Sec.80DDB under Chapter VIA" & Chr(13)
'        ValidateSelect80DDB = False
'        Exit Function
'    End If
'
'    If (SELECT80D = "2" And (calculateAge(dob_2) <= 59 Or calculateAge(dob_2) > 79)) Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Please select a valid option from the dropdown of Sec.80DDB under Chapter VIA" & Chr(13)
'        ValidateSelect80DDB = False
'        Exit Function
'    End If
'
'    If (SELECT80D = "3" And calculateAge(dob_2) < 80) Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Please select a valid option from the dropdown of Sec.80DDB under Chapter VIA" & Chr(13)
'        ValidateSelect80DDB = False
'        Exit Function
'    End If
End Function
Function ValidateSelect80U() As Boolean
ValidateSelect80U = True
    Dim Usr80UVal As Variant
    Dim dob_2 As Variant
    
    SELECT80U = Sheet1.Range("SELECT80U").Value
    SELECT80U = Mid(SELECT80U, 1, 1)
    If isdropdownblank(SELECT80U) Then
        SELECT80U = ""
    End If
    
    dob_2 = Sheet1.Range("sheet1.DOB").Value
    
    Usr80UVal = Sheet1.Range("IncD.Section80U").Value
    
    If SELECT80U <> "" Then
        If IsEmpty(Usr80UVal) Or Usr80UVal = 0 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter the value for 80U deduction under Chapter VIA in Income Details" & Chr(13)
            ValidateSelect80U = False
            Exit Function
        End If
    End If
    
    Usr80UVal = IIf(Trim(Usr80UVal) = "", 0, Usr80UVal)
        
    If Usr80UVal > 0 Then
        If isdropdownblank(SELECT80U) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option at dropdown of Sec.80U under Chapter VIA in Income Details" & Chr(13)
            ValidateSelect80U = False
            Exit Function
        End If
    End If
End Function


Sub NEW234F()


Dim Returnfiledstatus As String
Dim VerificationDate As Variant
Dim dateOfFiling As Variant
Dim VerificationDate1 As Variant
Dim VerificationDate2 As Variant
Dim VerificationDate3 As Variant
Dim DateOfFiling1 As Variant
Dim DateOfFiling2 As Variant
Dim DateOfFiling3 As Variant
Dim dueDate As Variant
Dim Altduedate As Variant
Dim UpdatedY As Boolean
    UpdatedY = False
    If (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
    UpdatedY = True
    End If

    Dim UpdatedN As Boolean
    UpdatedN = False
    If (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) <> "Y")) Then
    UpdatedN = True
    End If
Returnfiledstatus = Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))
VerificationDate = Range("Ver.Date").Value

dateOfFiling = Range("sheet1.OrigRetFiledDate").Value
'DateOfFiling2 = Range("sheet1.OrigRetFiledDate1").Value
    


'If Mid(Sheet1.Range("sheet1.StateCode1").Value, 1, 2) = "14" Or Mid(Sheet1.Range("sheet1.StateCode1").Value, 1, 2) = "37" Then
'Altduedate = "31/07/2025"               'Changed year from 2024 to 2025 by Ankita on 14/12/2023
' Altduedate = "15/09/2025"               'Changed year from 2024 to 2025 by Ankita on 10/07/2025
'Ayush_DueDate_08/09/2025
Altduedate = Sheet5.Range("DueDate1").Value

'2020 due date change
''Else
''Altduedate = "31/08/2020"
'End If


'Altduedate = "31/08/2020"

dueDate = Dformat(Altduedate, "yyyy-mm-dd")

VerificationDate1 = Dformat(VerificationDate, "yyyy-mm-dd")
'VerificationDate3 = Dformat(VerificationDate, "yyyy-mm-dd")

DateOfFiling1 = Dformat(dateOfFiling, "yyyy-mm-dd")
'DateOfFiling3 = Dformat(DateOfFiling2, "yyyy-mm-dd")

If Not ChkMinInclusiveDate(VerificationDate1, Range("DateOfProcessing").Value) Then
                VerificationDate1 = Range("DateOfProcessing").Value
End If



If Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or UpdatedY Then
VerificationDate1 = Dformat(dateOfFiling, "yyyy-mm-dd")
DateOfFiling1 = DateOfFiling1
Else
VerificationDate1 = VerificationDate1
End If

'If Returnfiledstatus = "18" Then
'VerificationDate1 = Dformat(DateOfFiling2, "yyyy-mm-dd")
'DateOfFiling1 = DateOfFiling3
'Else
'VerificationDate1 = VerificationDate1
'End If




If (Returnfiledstatus = "14" Or Returnfiledstatus = "15" Or Returnfiledstatus = "16" Or Returnfiledstatus = "20") Or (VerificationDate1 <= dueDate) Then
    intrst234F = 0
    'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome").Value <= 500000 Then ''For 234F We need to consider TotalIncome in which 112A should include as per DE V3 & BA - update dy Bindu on 4th Feb 2025
 ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome_New").Value <= 500000 Then
    intrst234F = 1000
    'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'Changed year from 2024 to 2025 by Ankita on 14/12/2023
'ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome").Value > 500000 And (VerificationDate1 > DueDate) And (VerificationDate1 <= "2024-12-31") Then
 'ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome").Value > 500000 And (VerificationDate1 > DueDate) And (VerificationDate1 <= "2025-12-31") Then ''For 234F We need to consider TotalIncome in which 112A should include as per DE V3 & BA - update dy Bindu on 4th Feb 2025
 ElseIf (Returnfiledstatus = "13" Or Returnfiledstatus = "11" Or Returnfiledstatus = "12" Or UpdatedN) And Range("IncD.TotalIncome_New").Value > 500000 And (VerificationDate1 > dueDate) And (VerificationDate1 <= "2025-12-31") Then
   
    intrst234F = 5000
    'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
    'Else
    'intrst234F = 10000
    'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
    'End If
ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or Returnfiledstatus = "19" Or UpdatedY) And (DateOfFiling1 <= dueDate) Then
    intrst234F = 0
    'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or Returnfiledstatus = "19" Or UpdatedY) And Range("IncD.TotalIncome").Value <= 500000 Then ''For 234F We need to consider TotalIncome in which 112A should include as per DE V3 & BA - update dy Bindu on 4th Feb 2025
ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or Returnfiledstatus = "19" Or UpdatedY) And Range("IncD.TotalIncome_New").Value <= 500000 Then
    intrst234F = 1000
    'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
'Changed year from 2024 to 2025 by Ankita on 14/12/2023
'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or Returnfiledstatus = "19" Or UpdatedY) And Range("IncD.TotalIncome").Value > 500000 And (DateOfFiling1 > DueDate) And (DateOfFiling1 <= "2024-12-31") Then
 'ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or Returnfiledstatus = "19" Or UpdatedY) And Range("IncD.TotalIncome").Value > 500000 And (DateOfFiling1 > DueDate) And (DateOfFiling1 <= "2025-12-31") Then ''For 234F We need to consider TotalIncome in which 112A should include as per DE V3 & BA - update dy Bindu on 4th Feb 2025
  ElseIf (Returnfiledstatus = "17" Or Returnfiledstatus = "18" Or Returnfiledstatus = "19" Or UpdatedY) And Range("IncD.TotalIncome_New").Value > 500000 And (DateOfFiling1 > dueDate) And (DateOfFiling1 <= "2025-12-31") Then
    intrst234F = 5000
    'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
Else
    'intrst234F = 10000
    intrst234F = 5000
    'Sheet9.Range("Sheet9.IntrstPayUs234F").value = intrstPayUs234F
End If
End Sub

Function checkfieldSuperSpecialcharacterDot(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldSuperSpecialcharacterDot = True
    Dim arr As Variant
    arr = Array(".")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldSuperSpecialcharacterDot = False
            Exit Function
        End If
        Next
    Next
End Function
Function ValidateOthersEI1() As Boolean
ValidateOthersEI1 = False

    If Len((Range("Others.NOI_2").item(1).Value) > 0) Then
        If Not ValidateNatureOfIncome1 Then ValidateOthersEI1 = False
        If Not ValidateAmount1 Then ValidateOthersEI1 = False
    End If

    setTblinfo_OthersNOI_2
    setTblinfo_OthersAmt2
    
     Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As String
    Dim cellRange1 As String
    Dim i As Long
    
    Set rangecells = Range("Others.NOI_2").Cells
    ReDim Others_NOI2(end_OthersAmt)
    
    For i = 1 To end_OthersAmt
   
    cellrange = GetMergedAddressCell(rangecells, i)
    
    Others_NOI2(i) = Sheet1.Range(cellrange).Value
       
        
       If (Others_NOI2(i) = "(Select)" Or Others_NOI2(i) = "") Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option from the drop down at B3 'Income from other Sources' at Sr. No  " & i & "  in Sheet Income Details" & Chr(13)
             ValidateOthersEI1 = False
             Exit Function
         End If

   Next
'
'    If ((end_OthersNOI1 <> end_OthersAmt)) Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Income From Other Sources : Enter all mandatory fields in Sheet Income Details" & Chr(13)
'    End If

End Function

Function ValidateNatureOfIncome1() As Boolean
ValidateNatureOfIncome1 = False

    setTblinfo_OthersNOI_2
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As String
    Dim cellRange1 As String
    Set rangecells = Range("Others.NOI_2").Cells
    Set rangecells1 = Range("Nature_Others_2").Cells
    Dim i As Long
    
    ReDim Others_NOI2(end_OthersNOI1)
    ReDim Others_NOI3(end_OthersNOI1)

    For i = 1 To end_OthersNOI1
    
        cellrange = GetMergedAddressCell(rangecells, i)
        cellRange1 = GetMergedAddressCell(rangecells1, i)
        
        Others_NOI2(i) = Sheet1.Range(cellrange).Value
        Others_NOI3(i) = Sheet1.Range(cellRange1).Value
    

         If (Others_NOI2(i) = "(Select)" Or Others_NOI2(i) = "") Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Nature of Income at Sr. No  " & i & "  in Sheet  Income Details  is mandatory" & Chr(13)
             ValidateNatureOfIncome1 = False
             Exit Function
         End If

         If (Others_NOI2(i) = "Any Other") Then
         If Others_NOI3(i) = "" Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter description in Nature of Income (Income from other Sources) at Sr. No  " & i & "  in Sheet Income Details " & Chr(13)
             ValidateNatureOfIncome1 = False
             Exit Function
         End If
         End If


         If Len(Others_NOI3(i)) > 125 Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Description at Sr. No  " & i & "  in Sheet Income Details cannot exceed 125 characters" & Chr(13)
             ValidateNatureOfIncome1 = False
             Exit Function
         End If
         
         
         
'          If checkfieldSuperSpecialcharactermax(Others_NOI3(i)) Then
'             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* In case any other is selected then in Description field permitted Special characters are (- , @ ;  / \ (  )  _ :) at Sr. No  " & i & "  in Sheet Income Details." & Chr(13)
'             ValidateNatureOfIncome1 = False
'             Exit Function
'         End If
    Next
End Function
Function ValidateAmount1() As Boolean
ValidateAmount1 = False

   setTblinfo_OthersNOI_2
    Dim cellrange As String
    Dim rangecells As Range
    Set rangecells = Range("Others.Amount_2").Cells
    Dim i As Long
    ReDim Others_Amt2(end_OthersNOI1)
    
    
    For i = 1 To end_OthersNOI1
    
        cellrange = GetMergedAddressCell(rangecells, i)
        Others_Amt2(i) = Sheet1.Range(cellrange).Value
        
      

        If Not chkCompulsory(Others_Amt2(i)) Then
        'Changed by Ankita 0n 30/12/2024
'           EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Income from other Sources - Please enter Amount at Sr. No  " & i & "  in Sheet Income Details is mandatory" & Chr(13)
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Income from other Sources - Please enter Amount at Sr. No  " & i & "  in Sheet Income Details""" & Chr(13)
            ValidateAmount1 = False
            Exit Function
        End If


        If Not IsNumeric(Others_Amt2(i)) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Income From Other Sources: Amount at Sr. No  " & i & "  in Sheet Income Details should be Numeric value" & Chr(13)
            ValidateAmount1 = False
            Exit Function
        End If

        If Others_Amt2(i) > 99999999999999# Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Income From Other Sources: Amount at Sr. No  " & i & "  in Sheet  Income Details cannot exceed 14 digits" & Chr(13)
            ValidateAmount1 = False
            Exit Function
        End If
    Next
End Function
Sub AddRows_Others2()
    Dim vRows As Long
    Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
    sourceSheet.Activate
    EfilingCommon.DefinedgridNameRange = "Others.NOI_2||Nature_Others_2||Others.Amount_2||Others.Amount_2_1||Others.Amount_2_2||Others.Amount_2_3||Others.Amount_2_4||Others.Amount_2_5"
    ActiveCellRange = EfilingCommon.searchLastRow("Others.NOI_2")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub setTblinfo_OthersNOI_2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Others.NOI_2").count
    Set rangecells = Range("Others.NOI_2").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "" Or rangecells.item(mIntCtr).Value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_OthersNOI1 = ccount
 End Sub
 Sub setTblinfo_OthersAmt2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Others.Amount_2").count
    Set rangecells = Range("Others.Amount_2").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_OthersAmt = ccount
End Sub

'Ankita_22/12/2025===========

Sub AddRows80CCC()
    Dim vRows As Long
    Dim sourceSheet As Worksheet
    
'    If Sheet1.Range("IncD.Section80CCC").Value > 0 Then
    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
    sourceSheet.Activate
    EfilingCommon.DefinedgridNameRange = "Sl_80CCC||Type_80CCC||Name_80CCC||Amount_80CCC"
    ActiveCellRange = EfilingCommon.searchLastRow("Type_80CCC")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula_80CCC_IncomeDetails
    
    Dim icnt, i As Variant
    icnt = Sheet1.Range("Sl_80CCC").Rows.count
    For i = 1 To icnt
        Application.EnableEvents = False
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("Sl_80CCC").Cells(i, 1).Value = i
        Application.EnableEvents = True
    Next
    
'    End If
End Sub


'Ankita_22/12/2025===========

'Sub AddRows80CCD1()
'    Dim vRows As Long
'    Dim sourceSheet As Worksheet
'
''    If Sheet1.Range("IncD.Section80CCD_SE").Value > 0 Then
'    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
'    sourceSheet.Activate
'    EfilingCommon.DefinedgridNameRange = "Sl_80CCD1||Type_80CCD1||Name_80CCD1||Amount_80CCD1"
'    ActiveCellRange = EfilingCommon.searchLastRow("Type_80CCD1")
'    vRows = EfilingCommon.insertRowUnderSectionWithFormula_80CCC_IncomeDetails
'
'    Dim icnt, i As Variant
'    icnt = Sheet1.Range("Sl_80CCD1").Rows.count
'    For i = 1 To icnt
'        Application.EnableEvents = False
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("Sl_80CCD1").Cells(i, 1).Value = i
'        Application.EnableEvents = True
'    Next
'
''    End If
'End Sub

'============================
'Ankita_22/12/2025===========

'Sub AddRows80CCD1b()
'    Dim vRows As Long
'    Dim sourceSheet As Worksheet
'
''    If Sheet1.Range("IncD.Section80CCD1B_SE").Value > 0 Then
'    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
'    sourceSheet.Activate
'    EfilingCommon.DefinedgridNameRange = "Sl_80CCD1b||Type_80CCD1b||Name_80CCD1b||Amount_80CCD1b"
'    ActiveCellRange = EfilingCommon.searchLastRow("Type_80CCD1b")
'    vRows = EfilingCommon.insertRowUnderSectionWithFormula_80CCC_IncomeDetails
'
'    Dim icnt, i As Variant
'    icnt = Sheet1.Range("Sl_80CCD1b").Rows.count
'    For i = 1 To icnt
'        Application.EnableEvents = False
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("Sl_80CCD1b").Cells(i, 1).Value = i
'        Application.EnableEvents = True
'    Next
'
''    End If
'End Sub

'Ankita_22/12/2025==============

Sub setTblinfo_80CCC_Type()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Type_80CCC").count
    Set rangecells = Range("Type_80CCC").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "") Then
                ccount = ccount + 1
            End If
    Next
 end_80CCCType = ccount
 End Sub
'======================
 Sub setTblinfo_80CCC_Name()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Name_80CCC").count
    Set rangecells = Range("Name_80CCC").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCCName = ccount
End Sub
'Ankita_22/12/2025===================
Sub setTblinfo_80CCC_Amount()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Amount_80CCC").count
    Set rangecells = Range("Amount_80CCC").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCCAmount = ccount
End Sub

'Ankita_25/12/2025===================
Sub setTblinfo_80CCD_1b_Type()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Type_80CCD1b").count
    Set rangecells = Range("Type_80CCD1b").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "") Then
                ccount = ccount + 1
            End If
    Next
 end_80CCD_1b_Type = ccount
 End Sub
'Ankita_25/12/2025===================
 Sub setTblinfo_80CCD_1b_Name()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Name_80CCD1b").count
    Set rangecells = Range("Name_80CCD1b").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCD_1b_Name = ccount
End Sub

'Ankita_25/12/2025===================

Sub setTblinfo_80CCD_1b_Amount()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Amount_80CCD1b").count
    Set rangecells = Range("Amount_80CCD1b").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCD_1b_Amount = ccount
End Sub
'Ankita_25/12/2025===================
'Sub setTblinfo_80CCD_1_Type()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Type_80CCD1").count
'    Set rangecells = Range("Type_80CCD1").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'            If Not (rangecells.item(mIntCtr).Value = "") Then
'                ccount = ccount + 1
'            End If
'    Next
' end_80CCD_1_Type = ccount
' End Sub
'Ankita_25/12/2025===================

 Sub setTblinfo_80CCD_1_Name()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Name_80CCD1").count
    Set rangecells = Range("Name_80CCD1").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCD_1_Name = ccount
End Sub

'Ankita_25/12/2025===================

Sub setTblinfo_80CCD_1_Amount()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Amount_80CCD1").count
    Set rangecells = Range("Amount_80CCD1").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80CCD_1_Amount = ccount
End Sub

'Ankita_22/12/2025===================

Function Validate_80CCC()
Validate_80CCC = True
setTblinfo_80CCC_Type
setTblinfo_80CCC_Name
setTblinfo_80CCC_Amount
end_80CCC = WorksheetFunction.Max(0, end_80CCCType, end_80CCCName, end_80CCCAmount)
If Not ValidateType_80CCC Then Validate_80CCC = False
If Not ValidateName_80CCC Then Validate_80CCC = False
If Not ValidateAmount_80CCC Then Validate_80CCC = False
End Function
'Ankita_22/12/2025===================

'Function Validate_80CCD_1()
'Validate_80CCD_1 = True
'setTblinfo_80CCD_1_Type
'setTblinfo_80CCD_1_Name
'setTblinfo_80CCD_1_Amount
'end_80CCD_1 = WorksheetFunction.Max(0, end_80CCD_1_Type, end_80CCD_1_Name, end_80CCD_1_Amount)
'If Not ValidateType_1_80CCD Then Validate_80CCD_1 = False
'If Not ValidateName_1_80CCD Then Validate_80CCD_1 = False
'If Not ValidateAmount_1_80CCD Then Validate_80CCD_1 = False
'End Function
'Ankita_22/12/2025===================

'Function Validate_80CCD_1b()
'Validate_80CCD_1b = True
'setTblinfo_80CCD_1b_Type
'setTblinfo_80CCD_1b_Name
'setTblinfo_80CCD_1b_Amount
'end_80CCD_1b = WorksheetFunction.Max(0, end_80CCD_1b_Type, end_80CCD_1b_Name, end_80CCD_1b_Amount)
'If Not ValidateType_1b_80CCD Then Validate_80CCD_1b = False
'If Not ValidateName_1b_80CCD Then Validate_80CCD_1b = False
'If Not ValidateAmount_1b_80CCD Then Validate_80CCD_1b = False
'End Function

'Ankita_22/12/2025===================

Function ValidateType_80CCC() As Boolean
    ValidateType_80CCC = True
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Type_80CCC").Cells
    ReDim Type_80CCC(end_80CCC)
    For i = 1 To end_80CCC
        Type_80CCC(i) = rangecells.item(i, 1).Value
        If Not chkCompulsory(Type_80CCC(i)) Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Type of Identifier is mandatory in schedule 80CCC at Sr. No " & i & """" & Chr(13)
            ValidateType_80CCC = False
            Exit Function
        End If
         If Len(Type_80CCC(i)) > 125 Then
          EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Type of Identifier in schedule 80CCC at Sr. No " & i & """ should be less than or equal to 125 characters." & Chr(13)
            ValidateType_80CCC = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(Type_80CCC(i)) Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Type of Identifier in schedule 80CCC at Sl.no. " & i & """ should not Contain special characters." & Chr(13)
            ValidateType_80CCC = False
            Exit Function
        End If
Next
End Function
'====================================

'Ankita_22/12/2025===================
Function ValidateName_80CCC() As Boolean
    ValidateName_80CCC = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Name_80CCC").Cells
    ReDim Name_80CCC(end_80CCC)
    For i = 1 To end_80CCC
        Name_80CCC(i) = rangecells.item(i, 1).Value
        If Not chkCompulsory(Name_80CCC(i)) Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Identifier No. is mandatory in schedule 80CCC at Sr. No " & i & """" & Chr(13)
            ValidateName_80CCC = False
            Exit Function
        End If
         If Len(Name_80CCC(i)) > 125 Then
          EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Identifier No. in schedule 80CCC at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
            ValidateName_80CCC = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Name_80CCC(i)) Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Identifier No. in schedule 80CCC at Sl.no. " & i & " should not Contain special characters." & Chr(13)
            ValidateName_80CCC = False
            Exit Function
        End If
         
Next
End Function
'Ankita_24/12/2025==============

Function ValidateAmount_80CCC() As Boolean
    ValidateAmount_80CCC = True
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("Amount_80CCC").Cells
    ReDim Amount_80CCC(end_80CCC)
    For i = 1 To end_80CCC
        Amount_80CCC(i) = rangecells.item(i, 1).Value
        If Not chkCompulsory(Amount_80CCC(i)) Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Amount is mandatory in schedule 80CCC at Sr. No " & i & """" & Chr(13)
            ValidateAmount_80CCC = False
            Exit Function
        End If
        If Not IsNumeric(Amount_80CCC(i)) Then
          EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Amount in schedule 80CCC at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
            ValidateAmount_80CCC = False
            Exit Function
        End If
        If Amount_80CCC(i) > 99999999999999# Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Amount in schedule 80CCC at Sl.no. " & i & " should not Contain special characters." & Chr(13)
            ValidateAmount_80CCC = False
            Exit Function
        End If
         
Next
End Function


'Ankita_16/04/2026================
'Commented as per v0.9
'Function ValidateType_1_80CCD() As Boolean
'    ValidateType_1_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet1.Range("Type_80CCD1").Cells
'    ReDim Type_1_80CCD(end_80CCD_1)
'    For i = 1 To end_80CCD_1
'        Type_1_80CCD(i) = rangecells.item(i, 1).Value
'        If Not chkCompulsory(Type_1_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Type of Identifier is mandatory in schedule 80CCD(1) at Sr. No " & i & """" & Chr(13)
'            ValidateType_1_80CCD = False
'            Exit Function
'        End If
'         If Len(Type_1_80CCD(i)) > 125 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Type of Identifier in schedule 80CCD(1) at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
'            ValidateType_1_80CCD = False
'            Exit Function
'        End If
'        If Not checkfieldspecialcharacter(Type_1_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Type of Identifier in schedule 80CCD(1) at Sl.no. " & i & " should not Contain special characters." & Chr(13)
'            ValidateType_1_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function

'===================================

'Function ValidateName_1_80CCD() As Boolean
'    ValidateName_1_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet1.Range("Name_80CCD1").Cells
'    ReDim Name_1_80CCD(end_80CCD_1)
'    For i = 1 To end_80CCD_1
'        Name_1_80CCD(i) = rangecells.item(i, 1).Value
'        If Not chkCompulsory(Name_1_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Name of Identifier is mandatory in schedule 80CCD(1) at Sr. No " & i & """" & Chr(13)
'            ValidateName_1_80CCD = False
'            Exit Function
'        End If
'         If Len(Name_1_80CCD(i)) > 125 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Name of Identifier in schedule 80CCD(1) length at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
'            ValidateName_1_80CCD = False
'            Exit Function
'        End If
'
'        If Not checkfieldspecialcharacter(Name_1_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Name of Identifier in schedule 80CCD(1) character at Sl.no. " & i & " should not Contain special characters." & Chr(13)
'            ValidateName_1_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function
'
'Function ValidateAmount_1_80CCD() As Boolean
'    ValidateAmount_1_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet1.Range("Amount_80CCD1").Cells
'    ReDim Amount_1_80CCD(end_80CCD_1)
'    For i = 1 To end_80CCD_1
'        Amount_1_80CCD(i) = rangecells.item(i, 1).Value
'        If Not chkCompulsory(Amount_1_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Amount is mandatory in schedule 80CCD(1) at Sr. No " & i & """" & Chr(13)
'            ValidateAmount_1_80CCD = False
'            Exit Function
'        End If
'
'
'        If Not IsNumeric(Amount_1_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Amount in schedule 80CCD(1) at Sr. No  " & i & "  in sheet Income Details should be Numeric value" & Chr(13)
'            ValidateAmount_1_80CCD = False
'            Exit Function
'        End If
'
'        If Amount_1_80CCD(i) > 99999999999999# Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Amount in schedule 80CCD(1) at Sr. No  " & i & "  in sheet Income Details cannot exceed 14 digits" & Chr(13)
'            ValidateAmount_1_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function
'===================================================

'Function ValidateType_1b_80CCD() As Boolean
'    ValidateType_1b_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet1.Range("Type_80CCD1b").Cells
'
'    ReDim Type_1b_80CCD(end_80CCD_1b)
'    For i = 1 To end_80CCD_1b
'        Type_1b_80CCD(i) = rangecells.item(i, 1).Value
'        If Not chkCompulsory(Type_1b_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Type of Identifier is mandatory in schedule 80CCD(1b) at Sr. No " & i & """" & Chr(13)
'            ValidateType_1b_80CCD = False
'            Exit Function
'        End If
'         If Len(Type_1b_80CCD(i)) > 125 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Type of Identifier in schedule 80CCD(1b) at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
'            ValidateType_1b_80CCD = False
'            Exit Function
'        End If
'
'        If Not checkfieldspecialcharacter(Type_1b_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Type of Identifier in schedule 80CCD(1b) at Sl.no. " & i & " should not Contain special characters." & Chr(13)
'            ValidateType_1b_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function


'Function ValidateName_1b_80CCD() As Boolean
'    ValidateName_1b_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet1.Range("Name_80CCD1b").Cells
'    ReDim Name_1b_80CCD(end_80CCD_1b)
'    For i = 1 To end_80CCD_1b
'        Name_1b_80CCD(i) = rangecells.item(i, 1).Value
'        If Not chkCompulsory(Name_1b_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Name of Identifier is mandatory in schedule 80CCD(1b) at Sr. No " & i & """" & Chr(13)
'            ValidateName_1b_80CCD = False
'            Exit Function
'        End If
'         If Len(Name_1b_80CCD(i)) > 125 Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Name of Identifier in schedule 80CCD(1b) length at Sr. No " & i & " should be less than or equal to 125 characters." & Chr(13)
'            ValidateName_1b_80CCD = False
'            Exit Function
'        End If
'        If Not checkfieldspecialcharacter(Name_1b_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Name of Identifier in schedule 80CCD(1b) character at Sl.no. " & i & " should not Contain special characters." & Chr(13)
'            ValidateName_1b_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function

'Function ValidateAmount_1b_80CCD() As Boolean
'    ValidateAmount_1b_80CCD = True
'
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet1.Range("Amount_80CCD1b").Cells
'    ReDim Amount_1b_80CCD(end_80CCD_1b)
'    For i = 1 To end_80CCD_1b
'        Amount_1b_80CCD(i) = rangecells.item(i, 1).Value
'        If Not chkCompulsory(Amount_1b_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""Amount is mandatory in schedule 80CCD(1b) at Sr. No " & i & """" & Chr(13)
'            ValidateAmount_1b_80CCD = False
'            Exit Function
'        End If
'        If Not IsNumeric(Amount_1b_80CCD(i)) Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Amount in schedule 80CCD(1b) at Sr. No  " & i & "  in sheet Income Details should be Numeric value" & Chr(13)
'            ValidateAmount_1b_80CCD = False
'            Exit Function
'        End If
'        If Amount_1b_80CCD(i) > 99999999999999# Then
'            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Amount in schedule 80CCD(1b) at Sr. No  " & i & "  in sheet Income Details cannot exceed 14 digits" & Chr(13)
'            ValidateAmount_1b_80CCD = False
'            Exit Function
'        End If
'
'Next
'End Function

'============================
'Ankita_16/04/2026========

Sub AddRows_Sec80CCC_pran()
    Dim vRows As Long
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
    sourceSheet.Activate
    EfilingCommon.DefinedgridNameRange = "pran_new||Pran_Sl"
    ActiveCellRange = EfilingCommon.searchLastRow("pran_new")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula_80CCC_IncomeDetails
End Sub

Sub AddRows_Others1()
    Dim vRows As Long
    Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
    sourceSheet.Activate
'    EfilingCommon.DefinedgridNameRange = "Others.NOI_1||Nature_Others_1||Others.Amount_1"
    EfilingCommon.DefinedgridNameRange = "Others.NOI_1||Others.Amount_1"
    ActiveCellRange = EfilingCommon.searchLastRow("Others.NOI_1")
    vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
Sub setTblinfo_OthersNOI_1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Others.NOI_1").count
    Set rangecells = Range("Others.NOI_1").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "" Or rangecells.item(mIntCtr).Value = "(Select)") Then
                ccount = ccount + 1
            End If
    Next
 end_OthersNOI2 = ccount
 End Sub
 Sub setTblinfo_OthersAmount_1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Others.Amount_1").count
    Set rangecells = Range("Others.Amount_1").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_OthersAmt1 = ccount
End Sub
Function ValidateOthersEI2() As Boolean
ValidateOthersEI2 = False

    If Len((Range("Others.NOI_1").item(1).Value) > 0) Then
        If Not ValidateNatureOfIncome2 Then ValidateOthersEI2 = False
        If Not ValidateAmount2 Then ValidateOthersEI2 = False
    End If

   setTblinfo_OthersNOI_1
   setTblinfo_OthersAmount_1
   
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim cellrange As String
    Dim cellRange1 As String
    Dim i As Long
    
    Set rangecells = Range("Others.NOI_1").Cells
 
    ReDim Others_NOI4(end_OthersAmt1)

    For i = 1 To end_OthersAmt1
   
    cellrange = GetMergedAddressCell(rangecells, i)
   
    Others_NOI4(i) = Sheet1.Range(cellrange).Value

        
       If (Others_NOI4(i) = "(Select)" Or Others_NOI4(i) = "") Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please select an option from the drop down of 1(ii) at Sr. No  " & i & "  in Sheet Income Details" & Chr(13)
             ValidateOthersEI2 = False
             Exit Function
         End If

   Next
'    If ((end_OthersNOI2 <> end_OthersAmt1)) Then
'        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "Income From Other Sources : Enter all mandatory fields in Sheet Income Details" & Chr(13)
'    End If

End Function
Function ValidateNatureOfIncome2() As Boolean
ValidateNatureOfIncome2 = False

    setTblinfo_OthersNOI_1
    Dim rangecells As Range
    Dim cellrange As String
     Dim cellRange1 As String
    Dim rangecells1 As Range
    Set rangecells = Range("Others.NOI_1").Cells
'     Set rangecells1 = Range("Nature_Others_1").Cells
    Dim i As Long
    
    ReDim Others_NOI4(end_OthersNOI2)
'    ReDim Others_NOI5(end_OthersNOI2)

    For i = 1 To end_OthersNOI2
    
         cellrange = GetMergedAddressCell(rangecells, i)
'        cellRange1 = GetMergedAddressCell(rangecells1, i)
        
         Others_NOI4(i) = Sheet1.Range(cellrange).Value
'        Others_NOI5(i) = Sheet1.Range(cellRange1).Value
        
         If (Others_NOI4(i) = "(Select)" Or Others_NOI4(i) = "") Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Nature of Income at Sr. No  " & i & "  in Sheet Income Details is mandatory" & Chr(13)
             ValidateNatureOfIncome2 = False
             Exit Function
         End If

        If (Others_NOI4(i) = "Sec 10(13A)-Allowance to meet expenditure incurred on house rent") Then
'        If (IsEmpty(Range("IncD.Section80GG_Calc").Value) = False) Then
        If Range("IncD.Section80GG_Calc").Value > 0 Then
              'MsgBox "* Deduction u/s 10(13A) & 80GG cannot be claimed for the same period"
              'Changed by Ankita on 30/12/2024
'              MsgBox "* Person receiving HRA cannot claim Deduction u/s 80GG"
               MsgBox "* ""Deduction u/s 10(13A) & 80GG cannot be claimed for the same period"""

             'EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Deduction u/s 10(13A) & 80GG cannot be claimed for the same period"
             'ValidateNatureOfIncome2 = False
             Exit Function
             End If
        End If
        
        'Ankita_10/03/2026=====
'         If (Others_NOI4(i) = "Any Other") Then
'         If Others_NOI5(i) = "" Then
'             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter description in Nature of Exempt Allowance at Sr. No  " & i & "  in Sheet Income Details " & Chr(13)
'             ValidateNatureOfIncome2 = False
'             Exit Function
'         End If
'         End If
'
'
'         If Len(Others_NOI5(i)) > 125 Then
'             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Description at Sr. No  " & i & "  in Sheet Income Details cannot exceed 125 characters" & Chr(13)
'             ValidateNatureOfIncome2 = False
'             Exit Function
'         End If
         
'         If checkfieldSuperSpecialcharactermax(Others_NOI5(i)) Then
'             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* In case any other is selected then in Description field permitted Special characters are (- , @ ;  / \ (  )  _ :) at Sr. No  " & i & "  in Sheet Income Details." & Chr(13)
'             ValidateNatureOfIncome2 = False
'             Exit Function
'         End If
    Next
End Function
Function ValidateAmount2() As Boolean
ValidateAmount2 = False

   setTblinfo_OthersNOI_1
    Dim cellrange As String
    Dim rangecells As Range
     Dim rangecells1 As Range
    Set rangecells = Range("Others.Amount_1").Cells
    Dim i As Long
    ReDim Others_Amt3(end_OthersNOI2)
    
    
    For i = 1 To end_OthersNOI2
    
        cellrange = GetMergedAddressCell(rangecells, i)
        Others_Amt3(i) = Sheet1.Range(cellrange).Value
        
        If Not chkCompulsory(Others_Amt3(i)) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter Amount against Nature of Exempt Allowance at Sr. No  " & i & "  in Sheet Income Details is mandatory" & Chr(13)
            ValidateAmount2 = False
            Exit Function
        End If


        If Not IsNumeric(Others_Amt3(i)) Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Income From Other Sources: Amount at Sr. No  " & i & "  in Sheet Income Details should be Numeric value" & Chr(13)
            ValidateAmount2 = False
            Exit Function
        End If

        If Others_Amt3(i) > 99999999999999# Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Income From Other Sources: Amount at Sr. No  " & i & "  in Sheet Income Details cannot exceed 14 digits" & Chr(13)
            ValidateAmount2 = False
            Exit Function
        End If
    Next
End Function
Function GetMergedAddressCell(rnge As Range, i As Long) As Variant
    Dim sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue As String
    Dim sTempCellValueInt As String

    sTempCellValue = Replace(rnge.AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
    sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
    sTempLastCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C6
    sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
    
    sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + i - 1)
    GetMergedAddressCell = sNewCellValue
End Function


Function CheckDoneePAN_80DD(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.
'ITR 1 is for individuals .So,4th character of PAN should be "P"

    CheckDoneePAN_80DD = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckDoneePAN_80DD = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckDoneePAN_80DD = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckDoneePAN_80DD = False
            Exit Function
        End If
        
        'Ankita_25-26
'        'Malli-----------
'        If Not UCase(Mid(PAN, 4, 1)) = UCase("P") Then
'            CheckDoneePAN_80DD = False
'            Exit Function
'        End If
        
        '-------------
        If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckDoneePAN_80DD = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckDoneePAN_80DD = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckDoneePAN_80DD = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckDoneePAN_80DD = False
            Exit Function
        End If
    End If
End Function

'Ankita_14/01/2026=====
    'Ankita_17/03/2026=====

'Sub LockUnlock_Secondary_Address()
'On Error Resume Next
'Application.EnableEvents = False
'    Sheet1.Unprotect Password:=getmsgstate
'     If Sheet1.Range("Secondary_Address").Value = "Yes" Then
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.ResidenceNo1").MergeArea.Locked = False
'            Sheet1.Range("sheet1.ResidenceNo1").MergeArea.Interior.Color = "&HCCFFCC"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.ResidenceName1").MergeArea.Locked = False
'            Sheet1.Range("sheet1.ResidenceName1").MergeArea.Interior.Color = "&HCCFFCC"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.Locked = False
'            Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.Interior.Color = "&HCCFFCC"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.Locked = False
'            Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.Interior.Color = "&HCCFFCC"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.Locked = False
'            Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.Interior.Color = "&HCCFFCC"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = False
'            Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HCCFFCC"
'            Sheet1.Range("sheet1.StateCode2").MergeArea.Value = "(Select)"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.Country1").MergeArea.Locked = False
'            Sheet1.Range("sheet1.Country1").MergeArea.Interior.Color = "&HCCFFCC"
'            Sheet1.Range("sheet1.Country1").MergeArea.Value = "91-INDIA"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.PinCode1").MergeArea.Locked = False
'            Sheet1.Range("sheet1.PinCode1").MergeArea.Interior.Color = "&HCCFFCC"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("HASZIP1").MergeArea.Locked = False
'            Sheet1.Range("HASZIP1").MergeArea.Interior.Color = "&HCCFFCC"
'
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = False
'            Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = "&HCCFFCC"
'
'    Else
'            Sheet1.Range("sheet1.ResidenceNo1").MergeArea.ClearContents
'            Sheet1.Range("sheet1.ResidenceNo1").MergeArea.Locked = True
'            Sheet1.Range("sheet1.ResidenceNo1").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("sheet1.ResidenceNo1").Value = ""
'
'            Sheet1.Range("sheet1.ResidenceName1").MergeArea.ClearContents
'            Sheet1.Range("sheet1.ResidenceName1").MergeArea.Locked = True
'            Sheet1.Range("sheet1.ResidenceName1").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("sheet1.ResidenceName1").Value = ""
'
'            Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.ClearContents
'            Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.Locked = True
'            Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("sheet1.RoadOrStreet1").Value = ""
'
'            Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.ClearContents
'            Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.Locked = True
'            Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("sheet1.LocalityOrArea1").Value = ""
'
'            Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.ClearContents
'            Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.Locked = True
'            Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("sheet1.CityOrTownOrDistrict1").Value = ""
'
'            Sheet1.Range("sheet1.StateCode2").MergeArea.ClearContents
'            Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = True
'            Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HD8D8D8"
'           Sheet1.Range("sheet1.StateCode2").Value = ""
'
'            Sheet1.Range("sheet1.Country1").MergeArea.ClearContents
'            Sheet1.Range("sheet1.Country1").MergeArea.Locked = True
'            Sheet1.Range("sheet1.Country1").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("sheet1.Country1").Value = ""
'
'            Sheet1.Range("sheet1.PinCode1").MergeArea.ClearContents
'            Sheet1.Range("sheet1.PinCode1").MergeArea.Locked = True
'            Sheet1.Range("sheet1.PinCode1").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("sheet1.PinCode1").Value = ""
'
'            Sheet1.Range("HASZIP1").MergeArea.ClearContents
'            Sheet1.Range("HASZIP1").MergeArea.Locked = True
'            Sheet1.Range("HASZIP1").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("HASZIP1").Value = ""
'
'            Sheet1.Range("sheet1.ZipCode1").MergeArea.ClearContents
'            Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = True
'            Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = "&HD8D8D8"
'            Sheet1.Range("sheet1.ZipCode1").Value = ""
'     End If
'
'    Application.EnableEvents = True
'    Sheet1.Protect Password:=getmsgstate
'End Sub

'======================


''AY_2025-26 Newly added by Ankita on 31st Jan 2025 as per V3 DE sheet

Function validateLTCG() As Boolean
On Error Resume Next
validateLTCG = True

IncD_Sale_LTCG = Sheet1.Range("IncD.Sale_LTCG").Value
IncD_Cost_LTCG = Sheet1.Range("IncD.Cost_LTCG").Value
IncD_CG_LTCG = Sheet1.Range("IncD.CG_LTCG").Value

msgError_LTCG = ""

If (IncD_Sale_LTCG > 0 And IncD_Sale_LTCG <> "") Or (IncD_Cost_LTCG > 0 And IncD_Cost_LTCG <> "") Then
            
'        If IncD_Sale_LTCG > 0 And (IncD_Cost_LTCG = "" Or IncD_Cost_LTCG = 0) Then
         If IncD_Sale_LTCG > 0 And (IncD_Cost_LTCG = "") Then

            validateLTCG = False
            msgError_LTCG = msgError_LTCG & "* ""Please fill Total cost of acquisition""" & Chr(13)
            Exit Function
        End If
        
'        If IncD_Cost_LTCG > 0 And (IncD_Sale_LTCG = "" Or IncD_Sale_LTCG = 0) Then
         If IncD_Cost_LTCG > 0 And (IncD_Sale_LTCG = "") Then
            validateLTCG = False
            msgError_LTCG = msgError_LTCG & "* ""Please fill Total sale consideration """ & Chr(13)
            Exit Function
        End If

End If

'IncD_Sale_LTCG
If IncD_Sale_LTCG <> "" And IncD_Sale_LTCG > 0 Then
    If Not IsNumeric(IncD_Sale_LTCG) Or IncD_Sale_LTCG > 99999999999999# Or IncD_Sale_LTCG < 0 Then
        validateLTCG = False
        msgError_LTCG = msgError_LTCG & "* ""Amount should be numeric, Non negative, no decimal, upto 99,999,999,999,999 at Total sale consideration""" & Chr(13)
        Sheet1.Range("IncD.Sale_LTCG").Value = ""
        'Exit Function
    End If
End If
            
'IncD_Cost_LTCG
If IncD_Cost_LTCG <> "" And IncD_Cost_LTCG > 0 Then
    If Not IsNumeric(IncD_Cost_LTCG) Or IncD_Cost_LTCG > 99999999999999# Or IncD_Cost_LTCG < 0 Then
        validateLTCG = False
        msgError_LTCG = msgError_LTCG & "* ""Amount should be numeric, Non negative, no decimal, upto 99,999,999,999,999 at Total cost of acquisition""" & Chr(13)
        Sheet1.Range("IncD.Cost_LTCG").Value = ""
        ' Exit Function
    End If
End If
            
 
If IncD_Sale_LTCG < IncD_Cost_LTCG Then
fmsgboxsmall """To avail the benefit of carry forward and set of loss, please use ITR -2"""
End If
            
'IncD_CG_LTCG
Dim IncD_Sale_Cost_LTCG As Variant
IncD_Sale_Cost_LTCG = IncD_Sale_LTCG - IncD_Cost_LTCG


If IncD_Sale_Cost_LTCG > 125000 Then
    validateLTCG = False
    msgError_LTCG = msgError_LTCG & "* ""In ITR 1, the maximum gains as per Section 112A can be INR 1,25,000/-. Please file ITR 2/3 if gains u/s 112A is more than INR 1,25,000/-""" & Chr(13)
    Sheet1.Range("IncD.Cost_LTCG").Value = ""
    Exit Function
End If
UpdateProgressBar

End Function

'Ankita_23/02/2026
Sub FitComments_DynamicHeight()
    Dim cmt As Comment
    Dim txt As String
    Dim charsPerLine As Double
    Dim lineCount As Long
    Dim lineHeight As Double
    Dim minLines As Long
    
    ' Tuning parameters
    charsPerLine = 35     ' approx characters per line (depends on width/font)
    lineHeight = 15       ' height of one text line
    minLines = 2          ' minimum number of lines
    
    For Each cmt In ActiveSheet.Comments
        With cmt.Shape
            .TextFrame.AutoSize = False
            .Width = 220        ' fixed width to enforce wrapping
            
            txt = cmt.text
            lineCount = Application.WorksheetFunction.RoundUp(Len(txt) / charsPerLine, 0)
            
            If lineCount < minLines Then lineCount = minLines
            
            .Height = lineCount * lineHeight
        End With
    Next cmt
End Sub

Sub AutoPopulateSecondaryAddress()

Application.EnableEvents = False
 
    If Sheet1.Range("Secondary_Address").Value = "Yes" Then
    
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceNo1").Value = Sheet1.Range("sheet1.ResidenceNo").Value
        Sheet1.Range("sheet1.ResidenceNo1").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.ResidenceNo1").MergeArea.Locked = True
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.LocalityOrArea1").Value = Sheet1.Range("sheet1.LocalityOrArea").Value
        Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.Locked = True
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceName1").Value = Sheet1.Range("sheet1.ResidenceName").Value
        Sheet1.Range("sheet1.ResidenceName1").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.ResidenceName1").MergeArea.Locked = True
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.RoadOrStreet1").Value = Sheet1.Range("sheet1.RoadOrStreet").Value
        Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.Locked = True
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.CityOrTownOrDistrict1").Value = Sheet1.Range("sheet1.CityOrTownOrDistrict").Value
        Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.Locked = True
        
        Sheet1.Unprotect Password:=getmsgstate
        If Sheet1.Range("sheet1.PinCode").Locked = True Then
            Sheet1.Range("sheet1.PinCode1").MergeArea.ClearContents
            Sheet1.Range("sheet1.PinCode1").MergeArea.Locked = True
            Sheet1.Range("sheet1.PinCode1").MergeArea.Interior.Color = "&HD8D8D8"
        
        Else
            Sheet1.Range("sheet1.PinCode1").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("sheet1.PinCode1").Value = Sheet1.Range("sheet1.PinCode").Value
            Sheet1.Range("sheet1.PinCode1").MergeArea.Locked = True
        End If
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.StateCode2").Value = Sheet1.Range("sheet1.StateCode1").Value
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = True
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.Country1").Value = Sheet1.Range("sheet1.Country").Value
        Sheet1.Range("sheet1.Country1").MergeArea.Interior.Color = RGB(217, 217, 217)
        Sheet1.Range("sheet1.Country1").MergeArea.Locked = True
        Sheet1.Unprotect Password:=getmsgstate
        
        If Sheet1.Range("HASZIP").Locked = True Then
            Sheet1.Range("HASZIP1").MergeArea.ClearContents
            Sheet1.Range("HASZIP1").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("HASZIP1").MergeArea.Locked = True
        Else
            If Sheet1.Range("HASZIP").Value = "YES" Then
            Sheet1.Range("sheet1.ZipCode1").Value = "XXXXXX"
            Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = True
            Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = "&HD8D8D8"
            Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("sheet1.ZipCode1").Value = Sheet1.Range("sheet1.ZipCode").Value
            Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = True
        End If
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("HASZIP1").MergeArea.Interior.Color = RGB(217, 217, 217)
            Sheet1.Range("HASZIP1").Value = Sheet1.Range("HASZIP").Value
            Sheet1.Range("HASZIP1").MergeArea.Locked = True
        End If
        
        Sheet1.Protect Password:=getmsgstate
Else
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceNo1").MergeArea.ClearContents
        Sheet1.Range("sheet1.ResidenceNo1").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("sheet1.ResidenceNo1").MergeArea.Locked = False
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.ClearContents
        Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("sheet1.LocalityOrArea1").MergeArea.Locked = False
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ResidenceName1").MergeArea.ClearContents
        Sheet1.Range("sheet1.ResidenceName1").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("sheet1.ResidenceName1").MergeArea.Locked = False
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.ClearContents
        Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("sheet1.RoadOrStreet1").MergeArea.Locked = False
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.ClearContents
        Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("sheet1.CityOrTownOrDistrict1").MergeArea.Locked = False
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.PinCode1").MergeArea.ClearContents
        Sheet1.Range("sheet1.PinCode1").MergeArea.Locked = False
        Sheet1.Range("sheet1.PinCode1").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.StateCode2").Value = "(Select)"
        Sheet1.Range("sheet1.StateCode2").MergeArea.Locked = False
        Sheet1.Range("sheet1.StateCode2").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.Country1").Value = "(Select)"
        Sheet1.Range("sheet1.Country1").MergeArea.Locked = False
        Sheet1.Range("sheet1.Country1").MergeArea.Interior.Color = "&HCCFFCC"
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("HASZIP1").MergeArea.ClearContents
        Sheet1.Range("HASZIP1").MergeArea.Locked = False
        Sheet1.Range("HASZIP1").MergeArea.Interior.Color = "&HD8D8D8"
        
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range("sheet1.ZipCode1").MergeArea.ClearContents
        Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = False
        Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = "&HD8D8D8"
        
    End If

Application.EnableEvents = True
 
End Sub
 
'Ankita_16/04/2026===
 Function Validate_Pran()
Validate_Pran = True
setTblinfo_pran
If Not Validate_pran_new Then Validate_Pran = False
End Function
'Ankita_16/04/2026===
Sub setTblinfo_pran()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("pran_new").count
    Set rangecells = Range("pran_new").Cells
    Dim countrycd As Variant
    For mIntCtr = 1 To mIntCells
            If Not (rangecells.item(mIntCtr).Value = "") Then
                ccount = ccount + 1
            End If
    Next
 end_pran = ccount
 End Sub


'Ankita_16/04/2026===
Function Validate_pran_new() As Boolean
    Validate_pran_new = True

    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet1.Range("pran_new").Cells
    ReDim pran_80CCC(end_pran)
    If end_pran = 0 Then
            EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* ""PRAN no. shall be mandatory if deduction u/s 80CCD(1) or 80CCD(1B) is claimed""" & Chr(13)
            Validate_pran_new = False
            Exit Function
    End If
    For i = 1 To end_pran
        pran_80CCC(i) = rangecells.item(i, 1).Value
         If Len(pran_80CCC(i)) > 12 Then
          EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter a valid 12 digit PRAN at Sr. No " & i & "" & Chr(13)
            Validate_pran_new = False
            Exit Function
        End If
        
        If Not IsNumeric(pran_80CCC(i)) Then
             EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "* Please enter a valid 12 digit PRAN at Sl.no. " & i & "" & Chr(13)
            Validate_pran_new = False
            Exit Function
        End If
         
Next
End Function


