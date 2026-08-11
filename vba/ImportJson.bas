Attribute VB_Name = "ImportJson"
'30/11/2021
Option Explicit

Dim end_TaxP
Private Type json_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-JSON will use String for numbers longer than 15 characters that contain only digits
    ' to override set `JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True`
    UseDoubleForLargeNumbers As Boolean

    ' The JSON standard requires object keys to be quoted (" or '), use this option to allow unquoted keys
    AllowUnquotedKeys As Boolean

    ' The solidus (/) is not required to be escaped, use this option to escape them as \/ in ConvertToJson
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

Sub ImportJson()
Dim filepath, jsonText As String
Dim msgq As Variant
Dim answer As Integer
Unload UserForm3
'fmsgbox "Import personal/tax details from downloaded Pre-filled JSON or Import from already generated JSON of the current assessment year."
msgq = "Import JSON" + Chr(13) + "functionality should be used only in a blank utility to import previously generated JSON. In case you are using blank utility click Yes, otherwise click No."
answer = MsgBox(msgq, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
If answer = vbYes Then
 imported = 1
Else
imported = 0
Exit Sub

End If



With Application.FileDialog(msoFileDialogFilePicker)
    'Makes sure the user can select only one file
    .AllowMultiSelect = False
    .Title = "Please select a Json file."
    'Filter to just the following types of files to narrow down selection options
    .Filters.add "Json File", "*.json", 1
    'Show the dialog box
    If .Show = True Then
        filepath = .SelectedItems.item(1)
    End If
    On Error Resume Next
End With


Open filepath For Input As #1
jsonText = Input$(LOF(1), 1)
Close #1



Dim jsonObject As Object

Set jsonObject = ParseJson(jsonText)

If Not jsonObject.exists("ITR") Then
imported = 0
fmsgbox ("Please select a valid Import JSON.")
Exit Sub
End If



ImportPersonalInfo (jsonText)
ImportFilingStatus (jsonText)
ImportITR1_IncomeDeductions (jsonText)
'Konda Added AY2025-26 on 04-02-2025
ImportLTCG112A (jsonText)
ImportTaxComputation (jsonText)
ImportTDSonOthThanSals (jsonText)
ImportTDSonSalaries (jsonText)
ImportScheduleTDS3Dtls (jsonText)
ImportTaxPayments (jsonText)
ImportScheduleTCS (jsonText)
ImportTaxPaid (jsonText)
ImportRefund (jsonText)
ImportVerification (jsonText)
ImportTaxReturnPreparer (jsonText)
If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
    ImportPartA_139_8A (jsonText)
    ImportPartB_ATI (jsonText)
End If
If (Sheet5.Range("BacValue").Value) = 2 Then
    Sheet4.Visible = True
    Sheet12.Visible = True
    Sheet9.Visible = True
    ImportSchedule80D (jsonText)
    ImportSchedule80GGA (jsonText)
    ImportSchedule80G100NoAppr (jsonText)
    ImportSchedule80G50NoAppr (jsonText)
    ImportSchedule80G100Appr (jsonText)
    ImportSchedule80G50Appr (jsonText)
    
    '80GGC_C1 2024-25 Bindu
    ImportSchedule80GGC (jsonText)
    
    '80U_DD_C1 2024_25 Malli
    ImportSchedule80DD (jsonText)
    ImportSchedule80U (jsonText)
    
     'Konda_AY_2025_26  09/04/2025
    ImportSchedule_80C (jsonText)
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
    'ImportSchedule_80CCC (jsonText)
    ImportSchedule_80E (jsonText)
    ImportSchedule_80EE (jsonText)
    ImportSchedule_80EEA (jsonText)
    ImportSchedule_80EEB (jsonText)
    'ImportSchedule_Int24B (jsonText)
Else
    Sheet4.Visible = False
    Sheet12.Visible = False
    Sheet9.Visible = False
End If

'ImportScheduleHP (jsonObject)   'Add by Konda on 22-01-2026
'Commented by Konda on 22-01-2026
'Malli---25/05/2025
'If Sheet5.Range("BacValue").Value = 1 And Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "S" And Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" Then
'     ImportSchedule_Int24B (jsonText)
'ElseIf Sheet5.Range("BacValue").Value = 2 And Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" Then
'      ImportSchedule_Int24B (jsonText)
'End If
'================================

'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
If Sheet5.Range("BacValue").Value = 2 And Mid(Range("sheet1.EmployerCategory1").Value, 1, 1) <> "N" Then
    ImportScheduleEA10_13A (jsonText)

End If
'----------------------------

'22/04/2026
Application.EnableEvents = False

Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Sheet1.Range("IncD.Section80C").Formula = "=IF(BacValue=1,0,TotAmount.80C)"
Sheet1.Range("IncD.Section80DValue").Formula = "=IF(BacValue=1,0,Eligible_Amount_80D)"
Sheet1.Range("IncD.Section80GGC").Formula = "=IF(BacValue=1,0,Total_Donation_80GGC)"
Sheet1.Range("IncD.Section80GGC_Calc").Formula = "=IF(BacValue=1,0,Total_Donation_Eligible_80GGC)"
Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
Application.EnableEvents = True

'----------
 


End Sub
Function ImportPersonalInfo(jsonText As String)
 On Error Resume Next
Dim jsonObject As Object
Dim jsonDictionary As Object
'E nhancement
'Dim PAN, firstName, middleName, LastName, residenceNo, residenceName, roadOrStreet, localityOrArea, cityOrTownOrDistrict, StateCode, CountryCode, PinCode, zipCode, countryCodeMobile, mobileNo, emailAddress, dob, iEmpCat, sEmpCat, aadhaarCardNo, AadhaarEnrolmentIds As Variant
Dim PAN, firstName, middleName, LastName, residenceNo, residenceName, roadOrStreet, localityOrArea, cityOrTownOrDistrict, StateCode, CountryCode, PinCode, zipCode, countryCodeMobile, mobileNo, emailAddress, dob, iEmpCat, sEmpCat, aadhaarCardNo As Variant
Dim sCountry, iCountry As Variant
Dim sState, iState As Variant

'Konda updated on 20-01-2026
'Dim WantUpdateAdd1, residenceNo1, residenceName1, roadOrStreet1, localityOrArea1, cityOrTownOrDistrict1, StateCode1, CountryCode1, PinCode1, zipCode1, countryCodeMobile3, mobileNo3, emailAddress3 As Variant
Dim SecondaryAdd1, residenceNo1, residenceName1, roadOrStreet1, localityOrArea1, cityOrTownOrDistrict1, StateCode1, CountryCode1, PinCode1, zipCode1, countryCodeMobile3, mobileNo3, emailAddress3 As Variant

Dim sCountry1, iCountry1 As Variant
Dim sState1, iState1 As Variant
'==========================
Dim YYYY, MM, DD, strDate As String

Set jsonObject = ParseJson(jsonText)

firstName = jsonObject("ITR")("ITR1")("PersonalInfo")("AssesseeName")("FirstName")
middleName = jsonObject("ITR")("ITR1")("PersonalInfo")("AssesseeName")("MiddleName")
LastName = jsonObject("ITR")("ITR1")("PersonalInfo")("AssesseeName")("SurNameOrOrgName")
PAN = jsonObject("ITR")("ITR1")("PersonalInfo")("PAN")
residenceNo = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("ResidenceNo")
residenceName = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("ResidenceName")
roadOrStreet = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("RoadOrStreet")
localityOrArea = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("LocalityOrArea")
cityOrTownOrDistrict = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("CityOrTownOrDistrict")
iState = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("StateCode")
sState = Findtext(CStr(iState), "StateList")
iCountry = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("CountryCode")
sCountry = Findtext(CStr(iCountry), "CountList")
countryCodeMobile = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("CountryCodeMobile")
mobileNo = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("MobileNo")
emailAddress = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("EmailAddress")
PinCode = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("PinCode")
zipCode = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("ZipCode")
'Konda updated on 20-01-2026===========
'WantUpdateAdd1 = jsonObject("ITR")("ITR1")("PersonalInfo")("WantUpdateAdd")
'V0.8
SecondaryAdd1 = jsonObject("ITR")("ITR1")("PersonalInfo")("SecondaryAdd")
residenceNo1 = jsonObject("ITR")("ITR1")("PersonalInfo")("AlternateAddress")("ResidenceNo")
residenceName1 = jsonObject("ITR")("ITR1")("PersonalInfo")("AlternateAddress")("ResidenceName")
roadOrStreet1 = jsonObject("ITR")("ITR1")("PersonalInfo")("AlternateAddress")("RoadOrStreet")
localityOrArea1 = jsonObject("ITR")("ITR1")("PersonalInfo")("AlternateAddress")("LocalityOrArea")
cityOrTownOrDistrict1 = jsonObject("ITR")("ITR1")("PersonalInfo")("AlternateAddress")("CityOrTownOrDistrict")
iState1 = jsonObject("ITR")("ITR1")("PersonalInfo")("AlternateAddress")("StateCode")
sState1 = Findtext(CStr(iState1), "StateList")
iCountry1 = jsonObject("ITR")("ITR1")("PersonalInfo")("AlternateAddress")("CountryCode")
sCountry1 = Findtext(CStr(iCountry1), "CountList")
countryCodeMobile3 = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("CountryCodeMobileNoSec")
mobileNo3 = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("MobileNoSec")
emailAddress3 = jsonObject("ITR")("ITR1")("PersonalInfo")("Address")("EmailAddressSec")
PinCode1 = jsonObject("ITR")("ITR1")("PersonalInfo")("AlternateAddress")("PinCode")
zipCode1 = jsonObject("ITR")("ITR1")("PersonalInfo")("AlternateAddress")("ZipCode")
'====================================
dob = jsonObject("ITR")("ITR1")("PersonalInfo")("DOB")
    YYYY = Mid(dob, 1, 4)
    MM = Mid(dob, 6, 2)
    DD = Mid(dob, 9, 2)
    strDate = DD & "/" & MM & "/" & YYYY
iEmpCat = jsonObject("ITR")("ITR1")("PersonalInfo")("EmployerCategory")
aadhaarCardNo = jsonObject("ITR")("ITR1")("PersonalInfo")("AadhaarCardNo")
'Enhancement
'AadhaarEnrolmentIds = Trim(jsonObject("ITR")("ITR1")("PersonalInfo")("AadhaarEnrolmentId"))

If iEmpCat = "CGOV" Then
    sEmpCat = "Central Government"
ElseIf iEmpCat = "SGOV" Then
    sEmpCat = "State Government"
'Konda updated on 27-02-2026--Schema-V0.4
'ElseIf iEmpCat = "SCJ" Then   'Konda_12/02/2026
'    sEmpCat = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958"
ElseIf iEmpCat = "PSU" Then
    sEmpCat = "Public Sector Undertaking"
ElseIf iEmpCat = "OTH" Then
    sEmpCat = "Others"
ElseIf iEmpCat = "PE" Then
    sEmpCat = "Pensioners - Central Government"
ElseIf iEmpCat = "PESG" Then
    sEmpCat = "Pensioners - State Government"
ElseIf iEmpCat = "PEPS" Then
    sEmpCat = "Pensioners - Public sector undertaking "
ElseIf iEmpCat = "PEO" Then
    sEmpCat = "Pensioners - Others"
ElseIf iEmpCat = "NA" Then
    sEmpCat = "Not Applicable (eg. Family pension etc)"
End If
    Sheet1.Unprotect Password:=getmsgstate
    If firstName <> "" Then
        Sheet1.Range("sheet1.FirstName").Value = firstName
    End If
    Sheet1.Unprotect Password:=getmsgstate
    If middleName <> "" Then
        Sheet1.Range("sheet1.MiddleName").Value = middleName
    End If
    Sheet1.Unprotect Password:=getmsgstate
    If LastName <> "" Then
        Sheet1.Range("sheet1.SurNameOrOrgName").Value = LastName
    End If
    Sheet1.Unprotect Password:=getmsgstate
    If PAN <> "" Then
        Sheet1.Range("sheet1.PAN").Value = PAN
    End If
    Sheet1.Protect Password:=getmsgstate
    If aadhaarCardNo <> "" Then
        Sheet1.Range("Sheet1.Aadhaar").Value = aadhaarCardNo
    End If
'    Enhancement
'    If CStr(AadhaarEnrolmentIds) <> "" Then
'        Sheet1.Range("Sheet1.AadhaarEnrol").Value = CStr(AadhaarEnrolmentIds)
'    End If
    Sheet1.Unprotect Password:=getmsgstate
    
    If strDate <> "" Then
        Sheet1.Range("sheet1.DOB").Value = strDate
    End If
    
    If sEmpCat <> "" Then
        Sheet1.Range("sheet1.EmployerCategory1").Value = sEmpCat
    End If
    
    If emailAddress <> "" Then
        Sheet1.Range("sheet1.EmailAddress").Value = emailAddress
    End If
    
    If emailAddress3 <> "" Then
        Sheet1.Range("sheet1.EmailAddress1").Value = emailAddress3
    End If
    
    Sheet1.Protect Password:=getmsgstate
    If residenceNo <> "" Then
        Sheet1.Range("sheet1.ResidenceNo").Value = residenceNo
    End If
    If residenceName <> "" Then
        Sheet1.Range("sheet1.ResidenceName").Value = residenceName
    End If
    If roadOrStreet <> "" Then
        Sheet1.Range("sheet1.RoadOrStreet").Value = roadOrStreet
    End If
    If localityOrArea <> "" Then
        Sheet1.Range("sheet1.LocalityOrArea").Value = localityOrArea
    End If
    If cityOrTownOrDistrict <> "" Then
        Sheet1.Range("sheet1.CityOrTownOrDistrict").Value = cityOrTownOrDistrict
    End If
    
    If sState <> "" Then
        Sheet1.Range("sheet1.StateCode1").Value = sState
    End If
    If sCountry <> "" Then
        Sheet1.Range("sheet1.Country").Value = sCountry
    End If
    If PinCode <> "" Then
        Sheet1.Range("sheet1.PinCode").Value = PinCode
    End If
    If iState = "99" Then
        If zipCode <> "" Then
            Sheet1.Range("HASZIP").Value = "No"
        Else
            Sheet1.Range("HASZIP").Value = "Yes"
        End If
    End If
    
    If zipCode <> "" Then
        Sheet1.Range("sheet1.ZipCode").Value = zipCode
    End If
'Konda updated on 17-03-2026--V0.6
''Konda updated on 20-01-2026===========================
'    Dim Add_flag
'    If WantUpdateAdd1 <> "" Then
'        If WantUpdateAdd1 = "Y" Then
'            Add_flag = "Yes"
'        Else
'        Add_flag = "No"
'        End If
'        Sheet1.Range("Secondary_Address").Value = Add_flag
'    End If
'If Sheet1.Range("Secondary_Address").Value = "Yes" Then
'===========================================
'+==
Dim Add_flag
If SecondaryAdd1 <> "" Then
        If SecondaryAdd1 = "Y" Then
            Add_flag = "Yes"
        Else
        Add_flag = "No"
        End If
        Sheet1.Range("Secondary_Address").Value = Add_flag
    End If
'If Sheet1.Range("Secondary_Address").Value = "Yes" Then
'V0.8
 If Sheet1.Range("Secondary_Address").Value = "No" Then
'+=

    If residenceNo1 <> "" Then
        Sheet1.Range("sheet1.ResidenceNo1").Value = residenceNo1
    End If
    If residenceName1 <> "" Then
        Sheet1.Range("sheet1.ResidenceName1").Value = residenceName1
    End If
    If roadOrStreet1 <> "" Then
        Sheet1.Range("sheet1.RoadOrStreet1").Value = roadOrStreet1
    End If
    If localityOrArea1 <> "" Then
        Sheet1.Range("sheet1.LocalityOrArea1").Value = localityOrArea1
    End If
    If cityOrTownOrDistrict1 <> "" Then
        Sheet1.Range("sheet1.CityOrTownOrDistrict1").Value = cityOrTownOrDistrict1
    End If
    
    If sState1 <> "" Then
        Sheet1.Range("sheet1.StateCode2").Value = sState1
    End If
    If sCountry1 <> "" Then
        Sheet1.Range("sheet1.Country1").Value = sCountry1
    End If
    If PinCode1 <> "" Then
        Sheet1.Range("sheet1.PinCode1").Value = PinCode1
    End If
    If iState1 = "99" Then
        If zipCode1 <> "" Then
            Sheet1.Range("HASZIP1").Value = "No"
        Else
            Sheet1.Range("HASZIP1").Value = "Yes"
        End If
    End If
    
    If zipCode1 <> "" Then
        Sheet1.Range("sheet1.ZipCode1").Value = zipCode1
    End If
End If     'Konda updated on 17-03-2026--V0.6 V0.8
'=========================
    If countryCodeMobile <> "" Then
        Sheet1.Range("sheet1.MobileCountryCode").Value = countryCodeMobile
    End If
    If mobileNo <> "" Then
        Sheet1.Range("sheet1.Mobileno").Value = mobileNo
    End If
'Konda updated on 20-01-2026=============
    
    If countryCodeMobile3 <> "" Then
        Sheet1.Range("sheet1.MobileCountryCode1").Value = countryCodeMobile3
    End If
    If mobileNo3 <> "" Then
        Sheet1.Range("sheet1.Mobileno1").Value = mobileNo3
    End If
    
'==================================
End Function
Function ImportFilingStatus(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim ReturnFileSec, SeventhProvisio139, NewTaxRegime, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, OptOutNewTaxRegime, ItrFilingDueDate As Variant
Dim sReturnFile, iReturnFile As Variant
Dim iProvisoFlag, sProvisoFlag As Variant
Dim iDepositAmountFlag, sDepositAmountFlag As Variant
Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
Dim DateofOriginalfile, DateofOriginalfile1, NoticeDateussec, Filingtype, ItrFilingDueDate_1 As Variant
Dim YYYY, MM, DD, strDate As String
Dim init As Variant
Dim node As Object


Set jsonObject = ParseJson(jsonText)

iReturnFile = jsonObject("ITR")("ITR1")("FilingStatus")("ReturnFileSec")
    If iReturnFile = "11" Then
       Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date"
    ElseIf iReturnFile = "12" Then
       Range("sheet1.ReturnFileSec").Value = "139(4)-Belated"
    ElseIf iReturnFile = "13" Then
    Sheet1.Shapes("RadioButton2").OLEFormat.Object.Value = 1
        RadioButton2_Click
       Range("sheet1.ReturnFileSec").Value = "142(1)"
    ElseIf iReturnFile = "14" Then
     Sheet1.Shapes("RadioButton2").OLEFormat.Object.Value = 1
        RadioButton2_Click
       Range("sheet1.ReturnFileSec").Value = "148"
    ElseIf iReturnFile = "15" Then
     Sheet1.Shapes("RadioButton2").OLEFormat.Object.Value = 1
        RadioButton2_Click
       Range("sheet1.ReturnFileSec").Value = "153A"
    ElseIf iReturnFile = "16" Then
        Sheet1.Shapes("RadioButton2").OLEFormat.Object.Value = 1
        RadioButton2_Click
        Range("sheet1.ReturnFileSec").Value = "153C"
    ElseIf iReturnFile = "17" Then
       Range("sheet1.ReturnFileSec").Value = "139(5)-Revised"
    ElseIf iReturnFile = "18" Then
        Sheet1.Shapes("RadioButton2").OLEFormat.Object.Value = 1
        RadioButton2_Click
         Range("sheet1.ReturnFileSec").Value = "139(9)"
    ElseIf iReturnFile = "20" Then
       'Range("sheet1.ReturnFileSec").Value = "139(9A) - After condonation of delay u/s 119(2)(b)"
       Range("sheet1.ReturnFileSec").Value = "119(2)(b)- After condonation of delay"
    ElseIf iReturnFile = "21" Then
       Range("sheet1.ReturnFileSec").Value = "139(8A)"
End If

iProvisoFlag = jsonObject("ITR")("ITR1")("FilingStatus")("SeventhProvisio139")
    If iProvisoFlag = "Y" Then
    sProvisoFlag = "Yes"
    ElseIf iProvisoFlag = "N" Then
    sProvisoFlag = "No"
End If


'Sheet1.Range("sheet1.SeventhProvisoFlag").Value = sProvisoFlag

'Enhancement_06/11/24
 If Sheet1.Range("Sheet1.SeventhProvisoFlag").Locked = False Then
    Sheet1.Range("Sheet1.SeventhProvisoFlag").Value = sProvisoFlag
    End If
   
'===============================
   


'PAG_C4
'If jsonObject("ITR")("ITR1")("FilingStatus")("NewTaxRegime") = "Y" Then
'    ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 1
'   Sheet5.Unprotect Password:=sPassword
'   Sheet5.Range("BacValue").Value = 1
'   ThisWorkbook.Unprotect Password:=sPassword
'   Sheet4.Unprotect Password:=sPassword
'  Sheet4.Visible = xlSheetHidden
'  Sheet4.Protect Password:=sPassword
'  Sheet12.Unprotect Password:=sPassword
'  Sheet12.Visible = xlSheetHidden
'  Sheet12.Protect Password:=sPassword
'  Sheet9.Unprotect Password:=sPassword
'  Sheet9.Visible = xlSheetHidden
'  Sheet9.Protect Password:=sPassword
'  sPassword = EfilingCommon.getmsgstate
'  Sheet5.Unprotect Password:=sPassword
'  Sheet5.Range("BacValue").Value = 1
'  Sheet5.Protect Password:=sPassword
'  Sheet1.Unprotect Password:=sPassword
'  resetBacYes
'  Sheet1.Protect Password:=sPassword
'
'  Sheet1.Activate
'  ThisWorkbook.Protect Password:=sPassword
'ElseIf jsonObject("ITR")("ITR1")("FilingStatus")("NewTaxRegime") = "N" Then
' Sheet5.Unprotect Password:=sPassword
'   Sheet5.Range("BacValue").Value = 2
'    ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 1
'    ThisWorkbook.Unprotect Password:=sPassword
'Sheet4.Unprotect Password:=sPassword
'  Sheet4.Visible = xlSheetVisible
'  Sheet4.Protect Password:=sPassword
'  Sheet12.Unprotect Password:=sPassword
'  Sheet12.Visible = xlSheetVisible
'  Sheet12.Protect Password:=sPassword
'  Sheet9.Unprotect Password:=sPassword
'  Sheet9.Visible = xlSheetVisible
'  Sheet9.Protect Password:=sPassword
'  sPassword = EfilingCommon.getmsgstate
'  Sheet5.Unprotect Password:=sPassword
'  Sheet5.Range("BacValue").Value = 2
'   Sheet5.Protect Password:=sPassword
'  Sheet1.Unprotect Password:=sPassword
'  resetBacNo
'  Sheet1.Protect Password:=sPassword
'
'  Sheet1.Activate
'  ThisWorkbook.Protect Password:=sPassword
'Else
'    ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 0
'    ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 0
'End If


'PAG_C4 2024-25 BIndu



If jsonObject("ITR")("ITR1")("FilingStatus")("OptOutNewTaxRegime") = "Y" Then
'   ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 1
   Sheet1.Activate
   ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 2  'Ankita_11/02/2026
   Sheet5.Unprotect Password:=sPassword
   Sheet5.Range("BacValue").Value = 2
   Sheet1.Activate
   Sheet1.Protect Password:=sPassword
   ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 1
   ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 0
   ThisWorkbook.Unprotect Password:=sPassword
   Sheet4.Unprotect Password:=sPassword
   Sheet4.Visible = xlSheetVisible
   Sheet4.Protect Password:=sPassword
  Sheet12.Unprotect Password:=sPassword
  Sheet12.Visible = xlSheetVisible
  Sheet12.Protect Password:=sPassword
  Sheet9.Unprotect Password:=sPassword
  Sheet9.Visible = xlSheetVisible
  Sheet9.Protect Password:=sPassword
  sPassword = EfilingCommon.getmsgstate
  Sheet5.Unprotect Password:=sPassword
  Sheet5.Range("BacValue").Value = 2
  Sheet5.Protect Password:=sPassword
  Sheet1.Unprotect Password:=sPassword
  
  '80GGC_C1 2024-25 Bindu
  Sheet13.Unprotect Password:=sPassword
  Sheet13.Visible = xlSheetVisible
  Sheet13.Protect Password:=sPassword
  
  '80U-DD_C1 2024-25 Malli
  Sheet14.Unprotect Password:=sPassword
  Sheet14.Visible = xlSheetVisible
  Sheet14.Protect Password:=sPassword
  
'Commented by Konda on 05-03-2026--SIT-113407
  'Malli-----------23/04/2025
  'Schedule 24(B)
'  If Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" Then
' 'Malli------23/04/2025
'  Sheet16.Unprotect Password:=sPassword
'  Sheet16.Visible = xlSheetVisible
'  Sheet16.Protect Password:=sPassword
'
'  Else
'  Sheet16.Unprotect Password:=sPassword
'  Sheet16.Visible = xlSheetHidden
'  Sheet16.Protect Password:=sPassword
'  End If
'----------------------------------
'=====================================
  
  
  
  '---------------------------------------
  '80c_80CCC
  
  Sheet15.Unprotect Password:=sPassword
  Sheet15.Visible = xlSheetVisible
  Sheet15.Protect Password:=sPassword
  
  '80E_80EE_80EEA_80EEB
  
  Sheet17.Unprotect Password:=sPassword
  Sheet17.Visible = xlSheetVisible
  Sheet17.Protect Password:=sPassword
  
  '---------------------------
  
  resetBacNo ' Changes
  Sheet1.Protect Password:=sPassword
 
  Sheet1.Activate
  ThisWorkbook.Protect Password:=sPassword
ElseIf jsonObject("ITR")("ITR1")("FilingStatus")("OptOutNewTaxRegime") = "N" Then
 Sheet5.Unprotect Password:=sPassword
    Sheet5.Range("BacValue").Value = 1
    ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 1
    Sheet1.Activate
   Sheet1.Protect Password:=sPassword
   ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 0
   ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 1
ThisWorkbook.Unprotect Password:=sPassword
Sheet4.Unprotect Password:=sPassword
  Sheet4.Visible = xlSheetHidden
  Sheet4.Protect Password:=sPassword
  Sheet12.Unprotect Password:=sPassword
  Sheet12.Visible = xlSheetHidden
  Sheet12.Protect Password:=sPassword
  Sheet9.Unprotect Password:=sPassword
  Sheet9.Visible = xlSheetHidden
  Sheet9.Protect Password:=sPassword
  sPassword = EfilingCommon.getmsgstate
  Sheet5.Unprotect Password:=sPassword
  Sheet5.Range("BacValue").Value = 1
   Sheet5.Protect Password:=sPassword
  Sheet1.Unprotect Password:=sPassword
  
  '80GGC_C1 2024-25 Bindu
  Sheet13.Unprotect Password:=sPassword
  Sheet13.Visible = xlSheetHidden
  Sheet13.Protect Password:=sPassword
  
  '80U-DD_C1 2024-25 Malli
  Sheet14.Unprotect Password:=sPassword
  Sheet14.Visible = xlSheetHidden
  Sheet14.Protect Password:=sPassword
  
  'Commented by Konda on 05-03-2026--SIT-113407
  'Malli---------23/04/2025
  'Schedule 24(B)
'  If Mid(Range("IncD.TypeOfHP").Value, 1, 1) = "S" Or Mid(Range("IncD.TypeOfHP").Value, 1, 1) = "(" Then
'  Sheet16.Unprotect Password:=sPassword
'      Sheet16.Visible = xlSheetHidden
'  Sheet16.Protect Password:=sPassword
'Else
'   ThisWorkbook.Unprotect Password:=getmsgstate
'     Sheet16.Visible = xlSheetVisible
'   ThisWorkbook.Protect Password:=getmsgstate
'End If
'-------------------------------------
'======================================
  
  '80c_80CCC
  
  Sheet15.Unprotect Password:=sPassword
  Sheet15.Visible = xlSheetHidden
  Sheet15.Protect Password:=sPassword
  
  '80E_80EE_80EEA_80EEB
  
  Sheet17.Unprotect Password:=sPassword
  Sheet17.Visible = xlSheetHidden
  Sheet17.Protect Password:=sPassword
  
  '-----------------------------------
  
  
  resetBacYes ' Changes
  Sheet1.Protect Password:=sPassword
 
  Sheet1.Activate
  ThisWorkbook.Protect Password:=sPassword
Else
    ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 0
    ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 0
End If

iDepositAmountFlag = jsonObject("ITR")("ITR1")("FilingStatus")("DepAmtAggAmtExcd1CrPrYrFlg")
    If iDepositAmountFlag = "Y" Then
    sDepositAmountFlag = "Yes"
    ElseIf iDepositAmountFlag = "N" Then
    sDepositAmountFlag = "No"
End If
'Change-22.11.2022.102.16I
'Sheet1.Range("Sheet1.DepositAmountFlag").Value = sDepositAmountFlag
'
'If sDepositAmountFlag = "Yes" Then
'    Sheet1.Range("Sheet1.DepositAmount").Value = jsonObject("ITR")("ITR1")("FilingStatus")("AmtSeventhProvisio139i")
'End If
'---end


'    'Enhancement 06/11/2024  'Please check ankita  08/11/2024
'    If DepositAmount <> "" And Sheet1.Range("Sheet1.AggrigateAmount").Locked = False Then
'        Sheet1.Range("Sheet1.AggrigateAmount").Value = AggrigateAmount
'    End If
'    '------------

iAggrigateAmountFlag = jsonObject("ITR")("ITR1")("FilingStatus")("IncrExpAggAmt2LkTrvFrgnCntryFlg")
If iAggrigateAmountFlag = "Y" Then
sAggrigateAmountFlag = "Yes"
ElseIf iAggrigateAmountFlag = "N" Then
sAggrigateAmountFlag = "No"
End If


'Enhancement_06/11/24
'Sheet1.Range("Sheet1.AggrigateAmountFlag").Value = sAggrigateAmountFlag
 If Sheet1.Range("Sheet1.AggrigateAmountFlag").Locked = False Then
    Sheet1.Range("Sheet1.AggrigateAmountFlag").Value = sAggrigateAmountFlag
    End If


'Ankita 08/11/2024
If Sheet1.Range("Sheet1.AggrigateAmount").Locked = False Then
If sAggrigateAmountFlag = "Yes" Then
    Sheet1.Range("Sheet1.AggrigateAmount").Value = jsonObject("ITR")("ITR1")("FilingStatus")("AmtSeventhProvisio139ii")
End If
End If

iAggrigateAmountFlag1 = jsonObject("ITR")("ITR1")("FilingStatus")("IncrExpAggAmt1LkElctrctyPrYrFlg")
If iAggrigateAmountFlag1 = "Y" Then
    sAggrigateAmountFlag1 = "Yes"
ElseIf iAggrigateAmountFlag1 = "N" Then
    sAggrigateAmountFlag1 = "No"
End If


'Sheet1.Range("Sheet1.AggrigateAmountFlag1").Value = sAggrigateAmountFlag1
'Changed by Riyaz on 31/12/2024 for SIT-85221
        If Sheet1.Range("Sheet1.AggrigateAmountFlag1").Locked = False Then
'        Sheet1.Range("Sheet1.AggrigateAmountFlag1").Value = AggrigateAmountFlag1
        Sheet1.Range("Sheet1.AggrigateAmountFlag1").Value = sAggrigateAmountFlag1
        End If

'Enhancement_06/11/2024
If Sheet1.Range("Sheet1.AggrigateAmount1").Locked = False Then
If sAggrigateAmountFlag1 = "Yes" Then
    Sheet1.Range("Sheet1.AggrigateAmount1").Value = jsonObject("ITR")("ITR1")("FilingStatus")("AmtSeventhProvisio139iii")
End If
End If
'======

ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1")
ReturnFileSec = Mid(ReturnFileSec, 1, 2)

If Sheet1.Range("Sheet1.ReceiptNo").Locked = False Then
    Sheet1.Range("Sheet1.ReceiptNo").Value = jsonObject("ITR")("ITR1")("FilingStatus")("ReceiptNo")
End If
If Sheet1.Range("Sheet1.OrigRetFiledDate").Locked = False Then
     DateofOriginalfile = jsonObject("ITR")("ITR1")("FilingStatus")("OrigRetFiledDate")
        YYYY = Mid(DateofOriginalfile, 1, 4)
        MM = Mid(DateofOriginalfile, 6, 2)
        DD = Mid(DateofOriginalfile, 9, 2)
        strDate = DD & "/" & MM & "/" & YYYY
    Sheet1.Range("Sheet1.OrigRetFiledDate").Value = strDate
End If
If Sheet1.Range("sheet1.NoticeNo").Locked = False Then
    Sheet1.Range("sheet1.NoticeNo").Value = jsonObject("ITR")("ITR1")("FilingStatus")("NoticeNo")
End If
If Sheet1.Range("sheet1.NoticeDate").Locked = False Then
     NoticeDateussec = jsonObject("ITR")("ITR1")("FilingStatus")("NoticeDateUnderSec")
            YYYY = Mid(NoticeDateussec, 1, 4)
            MM = Mid(NoticeDateussec, 6, 2)
            DD = Mid(NoticeDateussec, 9, 2)
            strDate = DD & "/" & MM & "/" & YYYY
        Sheet1.Range("sheet1.NoticeDate").Value = strDate
End If

'PAG_C9 2024-25 Bindu
ItrFilingDueDate_1 = jsonObject("ITR")("ITR1")("FilingStatus")("ItrFilingDueDate")
YYYY = Mid(ItrFilingDueDate_1, 1, 4)
MM = Mid(ItrFilingDueDate_1, 6, 2)
DD = Mid(ItrFilingDueDate_1, 9, 2)
strDate = DD & "/" & MM & "/" & YYYY
Sheet1.Range("Due_DateITR1").Value = strDate


    ' import 139 clause iv seventh proviso
    Set init = jsonObject("ITR")("ITR1")("FilingStatus")
 If init.exists("clauseiv7provisio139i") Then
        
    If init("clauseiv7provisio139i") = "Y" Then
        
        'Enhancement_06/11/2024
        If Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
           Sheet1.Range("clauseiv7provisio139iFlg") = "Yes"
        End If

         If Sheet1.Range("clauseiv7provisio139iFlg_3").Locked = False Then
        Sheet1.Range("clauseiv7provisio139iFlg_3") = "No"
        End If
        
            If Sheet1.Range("clauseiv7provisio139iFlg_4").Locked = False Then
            Sheet1.Range("clauseiv7provisio139iFlg_4") = "No"
            End If

        Application.EnableEvents = True
'        Sheet1.Range("clauseiv7provisio139iFlg") = "Yes"
'Change-22.11.2022.102.19H
'        Sheet1.Range("clauseiv7provisio139iFlg_1") = "No"
'        Sheet1.Range("clauseiv7provisio139iFlg_2") = "No"
'--end
'........Ankita_06/11/2024
'        Sheet1.Range("clauseiv7provisio139iFlg_3") = "No"
'        Sheet1.Range("clauseiv7provisio139iFlg_4") = "No"
            
            For Each node In init("clauseiv7provisio139iDtls")
'Change-22.11.2022.102.19I
'                If Node("clauseiv7provisio139iNature") = "1" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_1") = Node("clauseiv7provisio139iAmount")
'                ElseIf Node("clauseiv7provisio139iNature") = "2" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_2") = Node("clauseiv7provisio139iAmount")
'---end

'                If Node("clauseiv7provisio139iNature") = "3" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_3") = Node("clauseiv7provisio139iAmount")
'                ElseIf Node("clauseiv7provisio139iNature") = "4" Then
'                    Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes"
'                    Sheet1.Range("clauseiv7provisio139iAmount_4") = Node("clauseiv7provisio139iAmount")
'                End If

               'Enhancement_06/11/2024
                If node("clauseiv7provisio139iNature") = "1" Then
                   If Sheet1.Range("clauseiv7provisio139iFlg_3").Locked = False Then
                    Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes"
                    End If
                    If Sheet1.Range("clauseiv7provisio139iAmount_3").Locked = False Then
                    Sheet1.Range("clauseiv7provisio139iAmount_3") = node("clauseiv7provisio139iAmount")
                    End If
                ElseIf node("clauseiv7provisio139iNature") = "2" Then
                    
                    If Sheet1.Range("clauseiv7provisio139iFlg_4").Locked = False Then
                    Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes"
                    End If
                   If Sheet1.Range("clauseiv7provisio139iAmount_4").Locked = False Then
                    Sheet1.Range("clauseiv7provisio139iAmount_4") = node("clauseiv7provisio139iAmount")
                End If
                End If
            Next node
            
        Else
'                   Sheet1.Range("clauseiv7provisio139iFlg") = "No"

        'Enhancement_06/11/2024
            If Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
            Sheet1.Range("clauseiv7provisio139iFlg") = "No"
            End If
        End If
        
        
    End If

 Set init = jsonObject("ITR")("ITR1")("FilingStatus")
 If init.exists("AsseseeRepFlg") Then
        
    If init("AsseseeRepFlg") = "Y" Then
        
     If Sheet1.Range("sheet1.RepAssessee").Locked = False Then
       Sheet1.Range("sheet1.RepAssessee") = "Yes"
    End If
        If Sheet1.Range("sheet1.RepAssessee").Value <> "" Then
    
             If Sheet1.Range("sheet1.NameRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.NameRepAssessee") = jsonObject("ITR")("ITR1")("FilingStatus")("AssesseeRep")("RepName")
            End If
            
            If Sheet1.Range("sheet1.EmailRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.EmailRepAssessee") = jsonObject("ITR")("ITR1")("FilingStatus")("AssesseeRep")("RepEmailID")
            End If
            
            If Sheet1.Range("sheet1.CountryCodeRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.CountryCodeRepAssessee") = jsonObject("ITR")("ITR1")("FilingStatus")("AssesseeRep")("CountryCodeRepMobileNo")
            End If
            
            If Sheet1.Range("sheet1.ContactRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.ContactRepAssessee") = jsonObject("ITR")("ITR1")("FilingStatus")("AssesseeRep")("RepMobileNo")
            End If
        End If
    Else
    If Sheet1.Range("sheet1.RepAssessee").Locked = False Then
       Sheet1.Range("sheet1.RepAssessee") = "No"
    End If

End If
End If


End Function
Function ImportPartB_ATI(jsonText As String)

On Error Resume Next
Dim jsonObject As Object
Set jsonObject = ParseJson(jsonText)
Set jsonObject = jsonObject("ITR")("ITR1")("PartB-ATI")

If Sheet202.Range("U_Salaries").Locked = False Then
    Sheet202.Range("U_Salaries").Value = jsonObject("HeadOfInc")("Salaries")
End If
If Sheet202.Range("U_IncomeFromHP").Locked = False Then
    Sheet202.Range("U_IncomeFromHP").Value = jsonObject("HeadOfInc")("IncomeFromHP")
End If
'If Sheet202.Range("U_IncomeFromBP").Locked = False Then
'    Sheet202.Range("U_IncomeFromBP").Value = jsonObject("HeadOfInc")("IncomeFromBP")
'End If
'If Sheet202.Range("U_IncomeFromCG").Locked = False Then
'    Sheet202.Range("U_IncomeFromCG").Value = jsonObject("HeadOfInc")("IncomeFromCG")
'End If
If Sheet202.Range("U_IncomeFromOS").Locked = False Then
    Sheet202.Range("U_IncomeFromOS").Value = jsonObject("HeadOfInc")("IncomeFromOS")
End If




If Sheet202.Range("U_LatestTotInc").Locked = False Then
    Sheet202.Range("U_LatestTotInc").Value = jsonObject("LatestTotInc")
End If

If Sheet202.Range("U_LastAmtPayable").Locked = False Then
    Sheet202.Range("U_LastAmtPayable").Value = jsonObject("LastAmtPayable")
End If

If Sheet202.Range("U_Refund").Locked = False Then
    Sheet202.Range("U_Refund").Value = jsonObject("Refund")
End If

If Sheet202.Range("U_TotRefund").Locked = False Then
    Sheet202.Range("U_TotRefund").Value = jsonObject("TotRefund")
End If
If Sheet202.Range("U_RegAssessementTAX").Locked = False Then
    Sheet202.Range("U_RegAssessementTAX").Value = jsonObject("RegAssessementTAX")
End If



Dim Nodelist, node, itemp, vRows

If jsonObject.exists("ScheduleIT1") Then
    If jsonObject("ScheduleIT1").exists("TaxPayment1") Then
        Set Nodelist = jsonObject("ScheduleIT1")("TaxPayment1")("ITTaxPayments")
        itemp = 0
        If Nodelist.count > Sheet202.Range("U_BSRCode1").Rows.count Then
            Sheet202.Activate
            EfilingCommon.DefinedgridNameRange = ("U_slno1||U_BSRCode1||U_DateDep1||U_SrlNoOfChaln1||U_Amt1||U_DateOfAmount1")
            ActiveCellRange = EfilingCommon.searchLastRow("U_BSRCode1")
            vRows = EfilingCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet202.Range("U_BSRCode1").Rows.count)
        End If
        
        
        For Each node In Nodelist
            itemp = itemp + 1
            
            Sheet202.Range("U_BSRCode1").Rows(itemp).Cells(1).Value = node("BSRCode")
            
            Sheet202.Range("U_DateDep1").Rows(itemp).Cells(1).Value = Mid(node("DateDep"), 9, 2) & "/" & Mid(node("DateDep"), 6, 2) & "/" & Mid(node("DateDep"), 1, 4)
            
            Sheet202.Range("U_SrlNoOfChaln1").Rows(itemp).Cells(1).Value = node("SrlNoOfChaln")
            
            Sheet202.Range("U_Amt1").Rows(itemp).Cells(1).Value = node("Amt")
        
        Next node
    
    End If
    


End If
Nodelist = New Collection


If jsonObject.exists("ScheduleIT2") Then
    If jsonObject("ScheduleIT2").exists("TaxPayment2") Then
        Set Nodelist = jsonObject("ScheduleIT2")("TaxPayment2")("ITTaxPayments")
        itemp = 0
        If Nodelist.count > Sheet202.Range("U_BSRCode2").Rows.count Then
            Sheet202.Activate
            EfilingCommon.DefinedgridNameRange = ("U_slno2||U_BSRCode2||U_DateDep2||U_SrlNoOfChaln2||U_Amt2")
            ActiveCellRange = EfilingCommon.searchLastRow("U_BSRCode2")
            vRows = EfilingCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet202.Range("U_BSRCode2").Rows.count)
        End If
        
        
        For Each node In Nodelist
            itemp = itemp + 1
            
            Sheet202.Range("U_BSRCode2").Rows(itemp).Cells(1).Value = node("BSRCode")
            
            Sheet202.Range("U_DateDep2").Rows(itemp).Cells(1).Value = Mid(node("DateDep"), 9, 2) & "/" & Mid(node("DateDep"), 6, 2) & "/" & Mid(node("DateDep"), 1, 4)
            
            Sheet202.Range("U_SrlNoOfChaln2").Rows(itemp).Cells(1).Value = node("SrlNoOfChaln")
            
            Sheet202.Range("U_Amt2").Rows(itemp).Cells(1).Value = node("Amt")
        
        Next node
    
    End If
    


End If



If Sheet202.Range("U_ReleifUS89").Locked = False Then
    Sheet202.Range("U_ReleifUS89").Value = jsonObject("ReleifUS89")
End If




End Function
Function ImportPartA_139_8A(jsonText As String)

On Error Resume Next
Dim jsonObject As Object
Set jsonObject = ParseJson(jsonText)
Set jsonObject = jsonObject("ITR")("ITR1")
If Sheet201.Range("U_AadhaarCardNo").Locked = False Then
    Sheet201.Range("U_AadhaarCardNo").Value = jsonObject("PartA_139_8A")("AadhaarCardNo")
End If

'
'If Sheet201.Range("U_AadhaarEnrolmentId").Locked = False Then
'    Sheet201.Range("U_AadhaarEnrolmentId").Value = jsonObject("PartA_139_8A")("AadhaarEnrolmentId")
'End If


If Sheet201.Range("U_PreviouslyFiledForThisAY").Locked = False Then
    If jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY") = "Y" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY").Value = "Yes"
        
    ElseIf jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY") = "N" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY").Value = "No"
    End If
End If


If Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Locked = False Then
    If jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY_139_8A") = "1" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Value = "139(1)"
        
    ElseIf jsonObject("PartA_139_8A")("PreviouslyFiledForThisAY_139_8A") = "2" Then
        Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Value = "Other"
    End If
End If


If Sheet201.Range("U_ITRForm").Locked = False Then
    Sheet201.Range("U_ITRForm").Value = jsonObject("PartA_139_8A")("Applicable_139_8A")("ITRForm")
End If

If Sheet201.Range("U_AcknowledgementNo").Locked = False Then
    Sheet201.Range("U_AcknowledgementNo").Value = jsonObject("PartA_139_8A")("Applicable_139_8A")("AcknowledgementNo")
End If

If Sheet201.Range("U_OrigRetFiledDate").Locked = False Then
    Sheet201.Range("U_OrigRetFiledDate").Value = Mid(jsonObject("PartA_139_8A")("Applicable_139_8A")("OrigRetFiledDate"), 9, 2) & "/" & Mid(jsonObject("PartA_139_8A")("Applicable_139_8A")("OrigRetFiledDate"), 6, 2) & "/" & Mid(jsonObject("PartA_139_8A")("Applicable_139_8A")("OrigRetFiledDate"), 1, 4)
End If

If Sheet201.Range("U_LaidOutIn_139_8A").Locked = False Then
    If jsonObject("PartA_139_8A")("LaidOutIn_139_8A") = "Y" Then
        Sheet201.Range("U_LaidOutIn_139_8A").Value = "Yes"
        
    ElseIf jsonObject("PartA_139_8A")("LaidOutIn_139_8A") = "N" Then
        Sheet201.Range("U_LaidOutIn_139_8A").Value = "No"
    End If
End If


Dim Nodelist, node, itemp, vRows

If jsonObject("PartA_139_8A").exists("UpdatingInc") Then

    Set Nodelist = jsonObject("PartA_139_8A")("UpdatingInc")("ReasonsForUpdatingIncDtls")
    If Nodelist.count > 0 Then
'        Dim U_ReasonsForUpdatingIncomeCol
'
'        U_ReasonsForUpdatingIncomeCol = Sheet201.Range("U_ReasonsForUpdatingIncome").Column
        

        itemp = 0
        
        If Nodelist.count > Sheet201.Range("U_ReasonsForUpdatingIncome").Rows.count Then
            Sheet201.Activate
            EfilingCommon.DefinedgridNameRange = ("U_ReasonsForUpdatingIncome")
            ActiveCellRange = EfilingCommon.searchLastRow("U_ReasonsForUpdatingIncome")
            vRows = EfilingCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet201.Range("U_ReasonsForUpdatingIncome").Rows.count)
        End If
        
        For Each node In Nodelist
        itemp = itemp + 1
        Select Case node("ReasonsForUpdatingIncome")
        
        Case "1"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).Value = "Return previously not filed"
        Case "2"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).Value = "Income not reported correctly"
        Case "3"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).Value = "Wrong heads of income chosen"
        Case "4"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).Value = "Reduction of carried forward loss"
        Case "5"
             Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).Value = "Reduction of unabsorbed depreciation"
        Case "6"
              Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).Value = "Reduction of tax credit u/s 115JB/115JC"
        Case "7"
              Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).Value = "Wrong rate of tax"
        Case "OTH"
               Sheet201.Range("U_ReasonsForUpdatingIncome").Rows(itemp).Cells(1).Value = "Others "
        End Select
        
        
        Next node
    
    End If
    
End If

If Sheet201.Range("U_UpdatedReturnDuringPeriod").Locked = False Then
    If jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "1" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").Value = "Up to 12 months from the end of Relevant Assessment Year"
        
    ElseIf jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "2" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").Value = "Between 12 to 24 Months from the end of Relevant Assessment  Year"
          
    ElseIf jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "3" Then  'Ankita_04/11/2025
        Sheet201.Range("U_UpdatedReturnDuringPeriod").Value = "Between 24 to 36 Months from the end of Relevant Assessment Year"

    ElseIf jsonObject("PartA_139_8A")("UpdatedReturnDuringPeriod") = "4" Then
        Sheet201.Range("U_UpdatedReturnDuringPeriod").Value = "Between 36 to 48 Months from the end of Relevant Assessment Year"

    End If
End If



If Sheet201.Range("U_UnabsorbedDepreciation").Locked = False Then
    If jsonObject("PartA_139_8A")("RetrntoRedCarriedFL")("UnabsorbedDepreciation") = "Y" Then
        Sheet201.Range("U_UnabsorbedDepreciation").Value = "Yes"
        
    ElseIf jsonObject("PartA_139_8A")("RetrntoRedCarriedFL")("UnabsorbedDepreciation") = "N" Then
        Sheet201.Range("U_UnabsorbedDepreciation").Value = "No"
    End If
End If

Set Nodelist = New Collection


If jsonObject("PartA_139_8A").exists("RetrntoRedCarriedFL") Then
If jsonObject("PartA_139_8A")("RetrntoRedCarriedFL").exists("UDYear") Then
    Set Nodelist = jsonObject("PartA_139_8A")("RetrntoRedCarriedFL")("UDYear")("UnabsorbedDepreciationYearDtls")
    If Nodelist.count > 0 Then
'        Dim U_ReasonsForUpdatingIncomeCol
'
'        U_ReasonsForUpdatingIncomeCol = Sheet201.Range("U_ReasonsForUpdatingIncome").Column
        
        
        itemp = 0
        
        If Nodelist.count > Sheet201.Range("U_UnabsorbedDepreciationYear").Rows.count Then
            Sheet201.Activate
            EfilingCommon.DefinedgridNameRange = ("U_UnabsorbedDepreciationYear")
            ActiveCellRange = EfilingCommon.searchLastRow("U_UnabsorbedDepreciationYear")
            vRows = EfilingCommon.insertRowUnderSectionWithFormula(Nodelist.count - Sheet201.Range("U_UnabsorbedDepreciationYear").Rows.count)
        End If
        
        For Each node In Nodelist
        itemp = itemp + 1
            
            Dim YearUD As Variant
            YearUD = node("UnabsorbedDepreciationYear")
            'Ankita_UR
'            If YearUD <> "" Then
'                If YearUD = "2024" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).Value = "2024-25"
'                ElseIf YearUD = "2025" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).Value = "2025-26"
'                End If
'            End If
            
            
            If YearUD <> "" Then
                If YearUD = "2026" Then
                     Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).Value = "2026-27"
'                    ElseIf YearUD = "2025" Then
'                    Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).Value = "2025-26"
                     ElseIf YearUD = "2027" Then
                     Sheet201.Range("U_UnabsorbedDepreciationYear").Rows(itemp).Cells(1).Value = "2027-28"
                End If
            End If

            If Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Locked = False Then
'                If node("ReturnFiledEffectFlg") = "Y" Then
                 If node("RevisedReturnFile") = "Y" Then
                    Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Value = "Yes"

'                ElseIf node("ReturnFiledEffectFlg") = "N" Then
                 ElseIf node("RevisedReturnFile") = "N" Then
                    Sheet201.Range("U_RevisedReturnFile").Rows(itemp).Cells(1).Value = "No"
                End If
             End If
             
             
            'Ankita_04/11/2025
            If Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Locked = False Then
                If node("UpdatedReturnFile") = "Y" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Value = "Yes"

                ElseIf node("UpdatedReturnFile") = "N" Then
                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Value = "No"
                End If
             End If
             '------------------------------

'              If Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Locked = False Then
'                If node("ReturnType") = "1" Then
'                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Value = "Original Return 139(1)/139(4)"
'                ElseIf node("ReturnType") = "2" Then
'                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Value = "Revised return"
'                ElseIf node("ReturnType") = "3" Then
'                    Sheet201.Range("U_UpdatedReturnFile").Rows(itemp).Cells(1).Value = "Updated return"
'                End If
'             End If
        
        Next node
    
    End If
End If
End If




End Function
Function ImportITR1_IncomeDeductions(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary, node, Node2, Nodelist As Object
Set jsonObject = ParseJson(jsonText)
Dim NatureColNo, DescriptionColNo, AmtColNo, TotalExRow, TotalDiffRow, TotalXMLRow, rowcount, cnt, RecTDS1 As Long
Dim SalNatureDesc, DeductionUs57iia, IncomeFromSal, TotalIncomeOfHP, IncomeOthSrc, GrossTotIncome, Section80C, Section80CCC, Section80CCDEmployeeOrSE, Section80CCD1B, Section80CCDEmployer, Section80D, Section80DD, Section80DDB, AnyOthSec80CCH As Variant
Dim InterestPayable, Section80E, Section80EE, Section80G, Section80GG, Section80GGA, Section80GGC, Section80U, Section80TTA, Section80TTB, Section80DDUsrType, Section80DDBUsrType, Section80EEA, Section80EEB, Section80UUsrType As Variant
Dim ArrearsUnrealizedRentRcvd, TotalIncome, GrossSalary, Salary, PerquisitesValue, IncomeNotified89A, IncomeNotifiedOther89A, ProfitsInSalary, DeductionUs16, DeductionUs16ia, EntertainmentAlw16ii, ProfessionalTaxUs16iii, TypeOfHP, GrossRentReceived, TaxPaidlocalAuth As Variant
Dim test As Variant
Dim init As Variant

GrossSalary = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("GrossSalary")
Salary = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("Salary")
PerquisitesValue = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("PerquisitesValue")
ProfitsInSalary = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("ProfitsInSalary")
'Konda-20-01-2026===========================
'IncomeNotified89A = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("IncomeNotified89A")
'IncomeNotifiedOther89A = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("IncomeNotifiedOther89A")
'====================================
DeductionUs16 = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("DeductionUs16")
DeductionUs16ia = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("DeductionUs16ia")
EntertainmentAlw16ii = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("EntertainmentAlw16ii")
ProfessionalTaxUs16iii = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("ProfessionalTaxUs16iii")

'Konda-20-01-2026======================
'TypeOfHP = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("TypeOfHP")
'GrossRentReceived = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("GrossRentReceived")
'TaxPaidlocalAuth = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("TaxPaidlocalAuth")
'InterestPayable = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("InterestPayable")
'ArrearsUnrealizedRentRcvd = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("ArrearsUnrealizedRentRcvd")
'====================================
DeductionUs57iia = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("DeductionUs57iia")
IncomeOthSrc = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("IncomeOthSrc")
Section80C = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80C")
Section80CCC = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCC")
Section80CCDEmployeeOrSE = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCDEmployeeOrSE")
Section80CCD1B = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCD1B")
Section80CCDEmployer = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCDEmployer")
Section80D = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80D")
'Section80DD = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80DD")
Section80DDB = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80DDB")
'Section80E = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80E")
'Section80EE = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80EE")
Section80G = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80G")
Section80GG = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80GG")
Section80GGA = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80GGA")
Section80GGC = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80GGC")
'Section80U = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80U")
Section80TTA = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80TTA")
Section80TTB = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80TTB")
'Section80DDUsrType = Findtext(jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80DDUsrType"), "Selection80DD")
Section80DDBUsrType = Findtext(jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80DDBUsrType"), "Selection80DDB")
'Section80EEA = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80EEA")
'Section80EEB = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80EEB")
'Section80UUsrType = Findtext(jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80UUsrType"), "Selection80U")
AnyOthSec80CCH = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("AnyOthSec80CCH")
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.IncomeFromSal").Value = GrossSalary
    Sheet1.Range("IncD.Allowances").Value = Salary
    Sheet1.Range("IncD.Perquisites").Value = PerquisitesValue
    Sheet1.Range("IncD.Profits").Value = ProfitsInSalary
    Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")
'Konda-20-01-2026====================================
'    If init.exists("IncomeNotified89AType") Then
'        For Each node In init("IncomeNotified89AType")
'            If node("NOT89ACountrycode") = "US" Then
'                Sheet1.Range("IncomeNotified89A_AmountUS").Value = node("NOT89AAmount")
'            ElseIf node("NOT89ACountrycode") = "UK" Then
'                Sheet1.Range("IncomeNotified89A_AmountUK").Value = node("NOT89AAmount")
'            ElseIf node("NOT89ACountrycode") = "CA" Then
'                Sheet1.Range("IncomeNotified89A_AmountCan").Value = node("NOT89AAmount")
'            End If
'        Next node
'    End If
    
    
'    If Sheet1.Range("IncomeNotified89A").Locked = False Then
'    Sheet1.Range("IncomeNotified89A").Value = IncomeNotified89A
'    End If
'    If Sheet1.Range("IncomeNotifiedOther89A").Locked = False Then
'    Sheet1.Range("IncomeNotifiedOther89A").Value = IncomeNotifiedOther89A
'    End If
'=====================================================
    
   ' Sheet1.Unprotect Password:=getmsgstate
'Konda-20-01-2026====================================
'    Sheet1.Range("IncD.GrossRentRecieved").Value = GrossRentReceived
    'Sheet1.Unprotect Password:=getmsgstate
'Konda-20-01-2026====================================
'    Sheet1.Range("IncD.TaxPaidLocalAuthorities").Value = TaxPaidlocalAuth
'Konda updated V0.6.2 0n 23-04-2025
    'Sheet1.Range("IncD.InterestBorrowedCapital").Value = InterestPayable
'Konda-20-01-2026====================================
'    Sheet1.Range("IncD.Arrears").Value = ArrearsUnrealizedRentRcvd
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.IncomeFromOS").Value = IncomeOthSrc

'Konda updated V0.6.2 0n 21-04-2025
'    Sheet1.Range("IncD.Section80C").Value = Section80C
'Konda_07/05/2025_Uncommented as per DESheet_v0.7
    Sheet1.Range("IncD.Section80CCC").Value = Section80CCC
'--------------------------------------------------07/05/2025

'New Schema updated by Konda as AY-2026-27 on 26-12-2025

'SIT-109565-Updated by Konda on 04-02-2026
'If Sheet1.Range("IncD.Section80CCC").Value <> "" And Sheet1.Range("IncD.Section80CCC").Value > 0 Then
If Section80CCC <> "" And Section80CCC > 0 Then
Dim Nodelist80CCC As Object
Set Nodelist80CCC = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")
If Nodelist80CCC.exists("PensionContribution80CCC") Then

Dim Type_8CCC, Name_80CCC, Amount_80CCC

Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")
    
    If init.exists("PensionContribution80CCC") Then
    
        Set Nodelist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("PensionContribution80CCC")

        
        Type_8CCC = Sheet1.Range("Type_80CCC").Column
        Name_80CCC = Sheet1.Range("Name_80CCC").Column
        Amount_80CCC = Sheet1.Range("Amount_80CCC").Column
        
        TotalExRow = Range("Type_80CCC").Rows.count
        
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
    
        If (TotalXMLRow > 0) Then
           
            Sheet1.Range("Type_80CCC").ClearContents
         
            Sheet1.Range("Name_80CCC").ClearContents
         
            Sheet1.Range("Amount_80CCC").ClearContents
        End If
        
        
            If (TotalDiffRow > 0) Then

                    Dim P80CCC_idi As Long
                    P80CCC_idi = 0
                    For P80CCC_idi = 1 To TotalDiffRow
                    Sheet1.Activate
                    AddDiffRows_80CC (1)
                    Next
                End If

    
       rowcount = getRowNo(Sheet1.Range("Type_80CCC").name)
       rowcount = rowcount - 1
       cnt = 0
    
        For Each node In Nodelist
    
            rowcount = rowcount + 1
                 
               
            If Sheet1.Cells(rowcount, Type_8CCC).Locked = False Then
            
                    'Malli updated as per V0.9 17/04/2026
                    
                    Dim TypeofIdentifier_80ccc, TypeofIdentifier_80cccomport As Variant
                      TypeofIdentifier_80ccc = node("TypeofIdentifier")
                    
                       If TypeofIdentifier_80ccc <> "" Then
                                    If TypeofIdentifier_80ccc = "PRAN" Then
                                       TypeofIdentifier_80cccomport = "PRAN"
                                    ElseIf TypeofIdentifier_80ccc = "OTHPRAN" Then
                                       TypeofIdentifier_80cccomport = "Other than PRAN"
                                    Else
                                       TypeofIdentifier_80cccomport = "(Select)"
                                    End If
                    
                            'Sheet1.Cells(rowcount, Type_8CCC).Value = node("TypeofIdentifier")
                             Sheet1.Cells(rowcount, Type_8CCC).Value = TypeofIdentifier_80cccomport
                        
                        End If
                
            End If
            
            If Sheet1.Cells(rowcount, Name_80CCC).Locked = False Then
                Sheet1.Cells(rowcount, Name_80CCC).Value = node("NameofIdentifier")
            End If
            
                If Sheet1.Cells(rowcount, Amount_80CCC).Locked = False Then
                Sheet1.Cells(rowcount, Amount_80CCC).Value = node("Amount")
            End If
               

            cnt = cnt + 1
            
        Next node
        RecTDS1 = cnt
        End If
    End If
End If
'==========================================================

    Sheet1.Range("IncD.Section80CCD_SE").Value = Section80CCDEmployeeOrSE
   


''Malli commented as per V0.9.1 version changes
''New Schema updated by Konda as AY-2026-27 on 26-12-2025
'
''SIT-109570-Updated by Konda on 04-02-2026
''If Sheet1.Range("IncD.Section80CCD_SE").Value <> "" And Sheet1.Range("IncD.Section80CCD_SE").Value > 0 Then
'
'If Section80CCDEmployeeOrSE <> "" And Section80CCDEmployeeOrSE > 0 Then
'Dim Nodelist80CCD1 As Object
'Set Nodelist80CCD1 = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")
'If Nodelist80CCD1.exists("PensionContribution80CCD1") Then
'
'Dim Type_8CCD1, Name_80CCD1, Amount_80CCD1
'
'Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")
'
'    If init.exists("PensionContribution80CCD1") Then
'
'        Set Nodelist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("PensionContribution80CCD1")
'
'
'        Type_8CCD1 = Sheet1.Range("Type_80CCD1").Column
'        Name_80CCD1 = Sheet1.Range("Name_80CCD1").Column
'        Amount_80CCD1 = Sheet1.Range("Amount_80CCD1").Column
'
'        TotalExRow = Range("Type_80CCD1").Rows.count
'
'        TotalXMLRow = Nodelist.count
'        TotalDiffRow = TotalXMLRow - TotalExRow
'
'        If (TotalXMLRow > 0) Then
'
'            Sheet1.Range("Type_80CCD1").ClearContents
'
'            Sheet1.Range("Name_80CCD1").ClearContents
'
'            Sheet1.Range("Amount_80CCD1").ClearContents
'        End If
'
'
'            If (TotalDiffRow > 0) Then
'
'                    Dim P80CCD1_idi As Long
'                    P80CCD1_idi = 0
'                    For P80CCD1_idi = 1 To TotalDiffRow
'                    Sheet1.Activate
'                    AddDiffRows_80CCD1 (1)
'                    Next
'                End If
'
'
'       rowcount = getRowNo(Sheet1.Range("Type_80CCD1").name)
'       rowcount = rowcount - 1
'       cnt = 0
'
'        For Each node In Nodelist
'
'            rowcount = rowcount + 1
'
'
'               If Sheet1.Cells(rowcount, Type_8CCD1).Locked = False Then
'                Sheet1.Cells(rowcount, Type_8CCD1).Value = node("TypeofIdentifier")
'            End If
'
'            If Sheet1.Cells(rowcount, Name_80CCD1).Locked = False Then
'                Sheet1.Cells(rowcount, Name_80CCD1).Value = node("NameofIdentifier")
'            End If
'
'                If Sheet1.Cells(rowcount, Amount_80CCD1).Locked = False Then
'                Sheet1.Cells(rowcount, Amount_80CCD1).Value = node("Amount")
'            End If
'
'
'            cnt = cnt + 1
'
'        Next node
'        RecTDS1 = cnt
'        End If
'    End If
'End If
'
'
''===============================================
    
    Sheet1.Range("IncD.Section80CCD1B_SE").Value = Section80CCD1B
    
'New Schema updated by Konda as AY-2026-27 on 26-12-2025

'SIT-109573-Updated by Konda on 04-02-2026
'If Sheet1.Range("IncD.Section80CCD1B_SE").Value <> "" And Sheet1.Range("IncD.Section80CCD1B_SE").Value > 0 Then

'Malli commented as per V0.9.1 version changes
'If Section80CCD1B <> "" And Section80CCD1B > 0 Then
'Dim Nodelist80CCD1b As Object
'Set Nodelist80CCD1b = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")
'If Nodelist80CCD1b.exists("PensionContribution80CCD1B") Then
'
'Dim Type_8CCD1b, Name_80CCD1b, Amount_80CCD1b
'
'Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")
'
'    If init.exists("PensionContribution80CCD1B") Then
'
'        Set Nodelist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("PensionContribution80CCD1B")
'
'
'        Type_8CCD1b = Sheet1.Range("Type_80CCD1b").Column
'        Name_80CCD1b = Sheet1.Range("Name_80CCD1b").Column
'        Amount_80CCD1b = Sheet1.Range("Amount_80CCD1b").Column
'
'        TotalExRow = Range("Type_80CCD1b").Rows.count
'
'        TotalXMLRow = Nodelist.count
'        TotalDiffRow = TotalXMLRow - TotalExRow
'
'        If (TotalXMLRow > 0) Then
'
'            Sheet1.Range("Type_80CCD1b").ClearContents
'
'            Sheet1.Range("Name_80CCD1b").ClearContents
'
'            Sheet1.Range("Amount_80CCD1b").ClearContents
'        End If
'
'            If (TotalDiffRow > 0) Then
'
'                    Dim P80CCD1b_idi As Long
'                    P80CCD1b_idi = 0
'                    For P80CCD1b_idi = 1 To TotalDiffRow
'                    Sheet1.Activate
'                    AddDiffRows_80CCD1b (1)
'                    Next
'                End If
'
'       rowcount = getRowNo(Sheet1.Range("Type_80CCD1b").name)
'       rowcount = rowcount - 1
'       cnt = 0
'
'        For Each node In Nodelist
'
'            rowcount = rowcount + 1
'
'                If Sheet1.Cells(rowcount, Type_8CCD1b).Locked = False Then
'                Sheet1.Cells(rowcount, Type_8CCD1b).Value = node("TypeofIdentifier")
'                End If
'
'                If Sheet1.Cells(rowcount, Name_80CCD1b).Locked = False Then
'                Sheet1.Cells(rowcount, Name_80CCD1b).Value = node("NameofIdentifier")
'                End If
'
'                If Sheet1.Cells(rowcount, Amount_80CCD1b).Locked = False Then
'                Sheet1.Cells(rowcount, Amount_80CCD1b).Value = node("Amount")
'                End If
'
'            cnt = cnt + 1
'
'        Next node
'        RecTDS1 = cnt
'        End If
'    End If
'End If
'=======================================

    Sheet1.Range("IncD.Section80CCD").Value = Section80CCDEmployer

'As per AY-2026-27 New Schema PRANNum commented by Konda on 26-12-2025
''Konda updated V0.6.2 0n 21-04-2025
'    Dim PranNum_IncmDedutn
'    PranNum_IncmDedutn = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("PRANNum")
'    If Sheet1.Range("sheet1.PRAN").Locked = False Then
'        Sheet1.Range("sheet1.PRAN").Value = PranNum_IncmDedutn
'    End If

Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")
If init.exists("PRANDtls") Then
Dim PRANNumlist As Object
Dim PRANNumlist_TypeColNo

Set PRANNumlist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("PRANDtls")
   PRANNumlist_TypeColNo = Sheet1.Range("pran_new").Cells(1, 1).Column

            TotalExRow = Sheet1.Range("Pran_Sl").Rows.count
            TotalXMLRow = PRANNumlist.count
            TotalDiffRow = TotalXMLRow - TotalExRow

           If (TotalXMLRow > 0) Then
                    If Sheet1.Range("pran_new").Locked = False Then
                        Sheet1.Range("pran_new").ClearContents
                    End If
           End If
           
           If (TotalDiffRow > 0) Then
                Dim PRANNum_diff As Long
                PRANNum_diff = 0
                
                For PRANNum_diff = 1 To TotalDiffRow
                Sheet1.Activate
                AddDiffRows_PRANNum (1)
                Next
           
           End If
           
           rowcount = getRowNo(Sheet1.Range("Pran_Sl").name)
           rowcount = rowcount - 1
           cnt = 0
           
           For Each node In PRANNumlist
           
             rowcount = rowcount + 1
             
             Dim PRANNum_Import
             PRANNum_Import = node("PRANNum")
             If PRANNum_Import <> "" Then
                If Sheet1.Cells(rowcount, PRANNumlist_TypeColNo).Locked = False Then
                   Sheet1.Cells(rowcount, PRANNumlist_TypeColNo).Value = PRANNum_Import
                End If
             End If
           Next node
End If



' '--------------------------------

'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Section80DValue").Value = Section80D
     
    'Sheet1.Range("SELECT80DD").Value = Section80DDUsrType
    'Sheet1.Range("IncD.Section80DD").Value = Section80DD
    Sheet1.Range("SELECT80DDS").Value = Section80DDBUsrType
    
 Dim NameOfSpecDisease80DDB_G, NameOfSpecDisease80DDB_G_enum As Variant
    
    NameOfSpecDisease80DDB_G = Sheet1.Range("Sheet1.Specified_Disease").Value
    NameOfSpecDisease80DDB_G_enum = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("NameOfSpecDisease80DDB")
                                    
    If NameOfSpecDisease80DDB_G_enum <> "" Then
        

            If NameOfSpecDisease80DDB_G_enum = "a" Then
            NameOfSpecDisease80DDB_G = "(a) Dementia"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "b" Then
            NameOfSpecDisease80DDB_G = "(b) Dystonia Musculorum Deformans"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "c" Then
            NameOfSpecDisease80DDB_G = "(c) Motor Neuron Disease"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "d" Then
            NameOfSpecDisease80DDB_G = "(d) Ataxia"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "e" Then
            NameOfSpecDisease80DDB_G = "(e) Chorea"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "f" Then
            NameOfSpecDisease80DDB_G = "(f) Hemiballismus"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "g" Then
            NameOfSpecDisease80DDB_G = "(g) Aphasia"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "h" Then
            NameOfSpecDisease80DDB_G = "(h) Parkinsons Disease"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "i" Then
            NameOfSpecDisease80DDB_G = "(i) Malignant Cancers"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "j" Then
            NameOfSpecDisease80DDB_G = "(j) Full Blown Acquired Immuno-Deficiency Syndrome (AIDS)"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "k" Then
            NameOfSpecDisease80DDB_G = "(k) Chronic Renal failure"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "l" Then
            NameOfSpecDisease80DDB_G = "(l) Hematological disorders"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "m" Then
            NameOfSpecDisease80DDB_G = "(m) Hemophilia"
            
            ElseIf NameOfSpecDisease80DDB_G_enum = "n" Then
            NameOfSpecDisease80DDB_G = "(n) Thalassaemia"
            Else
            NameOfSpecDisease80DDB_G = ""
            
            End If
    
    If NameOfSpecDisease80DDB_G_enum <> "" Then
    'Sheet1.Range("Sheet1.Specified_Disease").Value = (Trim(NameOfSpecDisease80DDB_G))
     Sheet1.Range("Sheet1.Specified_Disease").Value = NameOfSpecDisease80DDB_G
    End If
    
    End If
    '--------------------------------------

    
    
    
    Sheet1.Range("IncD.Section80DDB").Value = Section80DDB

'Konda updated V0.6.2 0n 21-04-2025
'    Sheet1.Range("IncD.Section80E").Value = Section80E
'    Sheet1.Range("IncD.Section80EE").Value = Section80EE
'   ' Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Section80EEA").Value = Section80EEA
'    Sheet1.Range("IncD.Section80EEB").Value = Section80EEB
'-------------------------------------------------
     Sheet1.Unprotect Password:=getmsgstate   '22/05/2025
'    Sheet1.Range("IncD.Section80G").Value = Section80G
    Sheet1.Range("IncD.Section80GG").Value = Section80GG

'Konda updated V0.6.2 0n 21-04-2025
    Dim AckNum_F10BA
    AckNum_F10BA = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Form10BAAckNum")
    If Sheet1.Range("Sheet1.AckNum").Locked = False Then
    Sheet1.Range("Sheet1.AckNum").Value = AckNum_F10BA
    End If

 '--------------------------------------

    
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Section80GGA").Value = Section80GGA
    Sheet1.Range("IncD.Section80TTA").Value = Section80TTA
   ' Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80TTB").Value = Section80TTB
   ' Sheet1.Range("IncD.Section80GGC").Value = Section80GGC
    
    'Sheet1.Range("SELECT80U").Value = Section80UUsrType
    
    'Change-06.04.2023.103.ID.04
    'INC_E30
'    If Sheet1.Range("IncD.AnyOther").Locked = False Then
'        Sheet1.Range("IncD.AnyOther").Value = "80CCH-Contribution to Agnipath Scheme"
'        Sheet1.Range("IncD.AnyOtherDeductions").Value = AnyOthSec80CCH
'    End If
'
    
   'INC_C30 2024-25 Bindu
    'If Sheet1.Range("IncD.AnyOther").Locked = False Then
     '   Sheet1.Range("IncD.AnyOther").Value = "80CCH-Contribution to Agnipath Scheme"
        Sheet1.Range("IncD.AnyOtherDeductions").Value = AnyOthSec80CCH
    'End If
    
    
    'End Change
    
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.AnnualValue").Value = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("AnnualValue")
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.StandardDeduction").Value = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("StandardDeduction")

   
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.TotalChapVIADeductions_Input").Value = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("TotalChapVIADeductions")
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.TotalChapVIADeductions").Value = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("DeductUndChapVIA")("TotalChapVIADeductions")
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.TotalIncome").Value = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("TotalIncome")

'Konda updated on 27-02-2026 SIT-110467
'    Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")
'    If init.exists("AllwncExemptUs10") Then

    Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("AllwncExemptUs10")
    If init.exists("AllwncExemptUs10Dtls") Then
    
        Set Nodelist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("AllwncExemptUs10")("AllwncExemptUs10Dtls")
    
'        If init.exists("AllwncExemptUs10Dtls") Then 'Konda updated on 01-09-2025 SIT-95655
        NatureColNo = Sheet1.Range("Others.NOI_1").Column
        DescriptionColNo = Sheet1.Range("Nature_Others_1").Column
        AmtColNo = Sheet1.Range("Others.Amount_1").Column
        
        TotalExRow = Range("Others.NOI_1").Rows.count
        
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
    
        If (TotalXMLRow > 0) Then
            'Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Others.NOI_1").ClearContents
          '  Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Nature_Others_1").ClearContents
          '  Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Others.Amount_1").ClearContents
        End If
        
        If (TotalDiffRow > 0) Then
        'Konda add-------------------
            If UCase(node("SalNatureDesc")) = "10(13A)" Then
                TotalDiffRow = TotalDiffRow - 1
                AddDiffRows_Exempt2 (TotalDiffRow)
            Else
             AddDiffRows_Exempt2 (TotalDiffRow)
            End If
'         AddDiffRows_Exempt2 (TotalDiffRow)
        End If
    
       rowcount = getRowNo(Sheet1.Range("Others.NOI_1").name)
       rowcount = rowcount - 1
       cnt = 0
    
        For Each node In Nodelist
       ' Sheet1.Unprotect Password:=getmsgstate
       
        If UCase(node("SalNatureDesc")) <> "10(13A)" Then 'Newly added by Bindu as per DE V4
            rowcount = rowcount + 1
                 
                If UCase(node("SalNatureDesc")) = "10(10B)(I)" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette"
                ElseIf UCase(node("SalNatureDesc")) = "10(10B)(II)" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government"
                ElseIf UCase(node("SalNatureDesc")) = "10(14)(I)(115BAC)" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Section 10(14)(i) - Allowances referred in sub-clauses (a) to (c) of sub-rule (1) in Rule 2BB"
                ElseIf UCase(node("SalNatureDesc")) = "10(14)(II)(115BAC)" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Section 10(14)(ii) -  Transport allowance granted to certain physically handicapped assessee"
 '       Konda updated on 10-03-2026--V0.5
'                ElseIf UCase(node("SalNatureDesc")) = "OTH" Then
'                    Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
                ElseIf UCase(node("SalNatureDesc")) = "10(17)" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(17)-Allowance MP/MLA/MLC " 'Ankita
'======================================
                ElseIf UCase(node("SalNatureDesc")) = "EIC" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Exempt income received by a judge covered under the payment of salaries to Supreme Court/High Court judges Act /Rules"
                Else
                    test = Findtext("Sec " & (node("SalNatureDesc")), "PART_Nature_1")
                    Sheet1.Range("J" & rowcount).Value = test
                End If

'       Konda updated on 10-03-2026--V0.5
'               If Not node("SalOthNatOfInc") = "" Then
'                Sheet1.Range("Z" & rowcount).Value = node("SalOthNatOfInc")
'               End If
'===================================
'                Sheet1.Range("AO" & rowcount).Value = node("SalOthAmount")
                If UCase(node("SalOthAmount")) <> "" Then
                Sheet1.Cells(rowcount, AmtColNo).Value = node("SalOthAmount")
                End If

            cnt = cnt + 1
            End If ''Bindu
        Next node
        RecTDS1 = cnt
'        End If 'Konda updated on 01-09-2025 SIT-95655
    End If
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Deductions_16").Value = DeductionUs16
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Deduction16ia").Value = DeductionUs16ia
  '  Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Deduction16").Value = EntertainmentAlw16ii
 '   Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Deduction16ic").Value = ProfessionalTaxUs16iii
 
'Konda-20-01-2026========================
'    If TypeOfHP <> "" Then
'        If TypeOfHP = "L" Then
'            Sheet1.Range("IncD.TypeOfHP").Value = "Let Out"
'        ElseIf TypeOfHP = "S" Then
'            Sheet1.Range("IncD.TypeOfHP").Value = "Self Occupied"
'        ElseIf TypeOfHP = "D" Then
'            Sheet1.Range("IncD.TypeOfHP").Value = "Deemed Let Out"
'        End If
'    End If
'===================================
     Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("OthersInc")
    If init.exists("OthersIncDtlsOthSrc") Then
        Set Nodelist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("OthersInc")("OthersIncDtlsOthSrc")
    
        NatureColNo = Sheet1.Range("Others.NOI_2").Column
        DescriptionColNo = Sheet1.Range("Nature_Others_2").Column
        AmtColNo = Sheet1.Range("Others.Amount_2").Column
        Dim i
        i = 0
        For Each node In Nodelist
            i = i + 1
                If UCase(node("OthSrcNatureDesc")) = "DIV" Then
                    'Newly added by Bindu
                    
 'Konda updated on 13-04-2026
'                    Sheet1.Range("IncD_q1div").Value = node("DividendInc")("DateRange")("Upto15Of6")
'                    Sheet1.Range("IncD.Div").Value = node("OthSrcOthAmount")
'
'                    Sheet1.Range("IncD_q4div").Value = node("DividendInc")("DateRange")("Up16Of12To15Of3")
'                    Sheet1.Range("IncD_q5div").Value = node("DividendInc")("DateRange")("Up16Of3To31Of3")
'                    Sheet1.Range("IncD_q3div").Value = node("DividendInc")("DateRange")("Up16Of9To15Of12")
'
'                    Sheet1.Range("IncD_q2div").Value = node("DividendInc")("DateRange")("Upto15Of9")
                    
        
                    Sheet1.Range("IncD.Div").Value = node("OthSrcOthAmount")
                    
                    Sheet1.Range("IncD_q1div").Value = node("DividendInc")("DateRange")("Upto15Of6")
                    Sheet1.Range("IncD_q2div").Value = node("DividendInc")("DateRange")("Upto15Of9")
                    Sheet1.Range("IncD_q3div").Value = node("DividendInc")("DateRange")("Up16Of9To15Of12")
                    Sheet1.Range("IncD_q4div").Value = node("DividendInc")("DateRange")("Up16Of12To15Of3")
                    Sheet1.Range("IncD_q5div").Value = node("DividendInc")("DateRange")("Up16Of3To31Of3")
    '=======================on 13-04-2026

                    
                    
                    Nodelist.Remove i
                    i = i - 1
                
'Konda-20-01-2026===================================
'                ElseIf UCase(node("OthSrcNatureDesc")) = "OTHNOT89A" Then
'                    Sheet1.Range("OSIncomeNotifiedOther89A").Value = node("OthSrcOthAmount")
'                    Nodelist.Remove i
'                    i = i - 1
'
'                ElseIf UCase(node("OthSrcNatureDesc")) = "NOT89A" Then
'                    Sheet1.Range("IncD_q4OS1").Value = node("NOT89AInc")("DateRange")("Up16Of12To15Of3")
'                    Sheet1.Range("IncD_q5OS1").Value = node("NOT89AInc")("DateRange")("Up16Of3To31Of3")
'                    Sheet1.Range("IncD_q3OS1").Value = node("NOT89AInc")("DateRange")("Up16Of9To15Of12")
'                    Sheet1.Range("IncD_q1OS1").Value = node("NOT89AInc")("DateRange")("Upto15Of6")
'                    Sheet1.Range("IncD_q2OS1").Value = node("NOT89AInc")("DateRange")("Upto15Of9")

'                    If node.exists("NOT89A") Then
'                        For Each Node2 In node("NOT89A")
'                            If Node2("NOT89ACountrycode") = "US" Then
'                                Sheet1.Range("OSIncomeNotified89A_AmountUS").Value = Node2("NOT89AAmount")
'                            ElseIf Node2("NOT89ACountrycode") = "UK" Then
'                                Sheet1.Range("OSIncomeNotified89A_AmountUK").Value = Node2("NOT89AAmount")
'                            ElseIf Node2("NOT89ACountrycode") = "CA" Then
'                                Sheet1.Range("OSIncomeNotified89A_AmountCan").Value = Node2("NOT89AAmount")
'                            End If
'                        Next Node2
'                    End If
'======================================================
                    Nodelist.Remove i
                    i = i - 1
                End If
        Next node
        
        TotalExRow = Range("Others.NOI_2").Rows.count
        
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
    
        If (TotalXMLRow > 0) Then
         '   Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Others.NOI_2").ClearContents
           ' Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Nature_Others_2").ClearContents
           ' Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Others.Amount_2").ClearContents
            
        End If
    
        If (TotalDiffRow > 0) Then
         AddDiffRows_Exempt1 (TotalDiffRow)
        End If
 
        rowcount = getRowNo(Sheet1.Range("Others.NOI_2").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each node In Nodelist
            rowcount = rowcount + 1
        
                If UCase(node("OthSrcNatureDesc")) = "SAV" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Savings Bank Account"
                End If
                
                If UCase(node("OthSrcNatureDesc")) = "IFD" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Deposit (Bank/Post Office/Cooperative Society)"
                End If
                
                If UCase(node("OthSrcNatureDesc")) = "TAX" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Income Tax Refund"
                End If
                
                If UCase(node("OthSrcNatureDesc")) = "FAP" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Family pension"
                End If
                
                If UCase(node("OthSrcNatureDesc")) = "10(11)(IP)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(11)"
                End If
                
                If UCase(node("OthSrcNatureDesc")) = "10(11)(IIP)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(11)"
                End If
                
                If UCase(node("OthSrcNatureDesc")) = "10(12)(IP)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(12)"
                End If
                
                If UCase(node("OthSrcNatureDesc")) = "10(12)(IIP)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(12)"
                End If
                
                If UCase(node("OthSrcNatureDesc")) = "OTH" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
                End If
              '  Sheet1.Unprotect Password:=getmsgstate
               If UCase(node("OthSrcNatureDesc")) <> "DIV" And UCase(node("OthSrcNatureDesc")) <> "NOT89A" And UCase(node("OthSrcNatureDesc")) <> "OTHNOT89A" Then
                Sheet1.Cells(rowcount, DescriptionColNo).Value = node("OthSrcOthNatOfInc")
                Sheet1.Cells(rowcount, AmtColNo).Value = node("OthSrcOthAmount")
              End If
            cnt = cnt + 1
        Next node
        RecTDS1 = cnt
    End If
        
            
            
            
     Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")
    If init.exists("ExemptIncAgriOthUs10") Then
        Set Nodelist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("ExemptIncAgriOthUs10")("ExemptIncAgriOthUs10Dtls")
    
        NatureColNo = Sheet1.Range("Others.NOI").Column
        DescriptionColNo = Sheet1.Range("Nature_Others").Column
        AmtColNo = Sheet1.Range("Others.Amount").Column
        
        'Malli_AY_2026_27 17/06/2026
        Dim DescEIColNo As Variant
        DescEIColNo = Sheet1.Range("Sheet5.DescEI").Column
        '---------------------------
        
        TotalExRow = Range("Others.NOI").Rows.count
        
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        If (TotalXMLRow > 0) Then
            'Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Others.NOI").ClearContents
           ' Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Nature_Others").ClearContents
          '  Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("Others.Amount").ClearContents
            
            Sheet1.Range("Sheet5.DescEI").ClearContents
        End If
    
       If (TotalDiffRow > 0) Then
           AddDiffRows_Exempt (TotalDiffRow)
       End If
    
       rowcount = getRowNo(Sheet1.Range("Others.NOI").name)
       rowcount = rowcount - 1
       cnt = 0
        
        For Each node In Nodelist
            rowcount = rowcount + 1
                
 'Commented by Konda on 10-03-2026---V0.5
'                If UCase(node("NatureDesc")) = "OTH" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(34)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(34) (Exempted Dividend Income)"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(26AAA)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(26AAA) Any income as referred to in section 10(26AAA)"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(26)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(26) Any income as referred to in section 10(26)"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(19)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(19) Armed Forces Family pension in case of death during operational duty"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(10D)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10D) Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(1)"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(11)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(11) Statutory Provident Fund received"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(12)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(12) Recognized Provident Fund received"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(13)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(13) Approved superannuation fund received"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(16)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(16) Scholarships granted to meet the cost of education"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(17)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(17) Allowance MP/MLA/MLC"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(17A)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(17A) Award instituted by Government"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(18)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(18) Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(10BC)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10BC) Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster"
'                End If
'
'                If UCase(node("NatureDesc")) = "DMDP" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Defense Medical Disability Pension"
'                End If
'
'                If UCase(node("NatureDesc")) = "CG1L" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Any exempt income including LTCG on which tax is not payable"
'                End If
'
'                If UCase(node("NatureDesc")) = "AGRI" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Agriculture Income (less than equal to Rs.5000)"
'                End If
'
'                If UCase(node("NatureDesc")) = "10(12C)" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(12C) Any payment from the Agniveer Corpus Fund to a person enrolled under the Agnipath Scheme, or to his nominee"
'                End If
'
''                If UCase(Node("NatureDesc")) = "LTCG" Then
''                Sheet1.Cells(rowcount, NatureColNo).Value = "LTCG u/s 112A not exceeding  Rs. 1 Lakh"
''                End If
'
'               ' Sheet1.Unprotect Password:=getmsgstate
'                Sheet1.Cells(rowcount, DescriptionColNo).Value = node("OthNatOfInc")

'Konda updated on 10-03-2026----V0.5
        If Sheet1.Cells(rowcount, NatureColNo).Locked = False Then
        
            If UCase(node("Category")) = "AGRI" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Agricultural & related incomes"
            
            ElseIf UCase(node("Category")) = "GOVC" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Compensation/other sums received by government or other approved entities"
            
            ElseIf UCase(node("Category")) = "ISI" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Income from specified Investments"
            
            ElseIf UCase(node("Category")) = "SSRA" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Specified sums received by armed forces personnel"
            
            ElseIf UCase(node("Category")) = "SRSC" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sums received by Senior Citizens/Minors "
            
            ElseIf UCase(node("Category")) = "SRST" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sums received by specified Category of Taxpayers"
            
            ElseIf UCase(node("Category")) = "SRPC" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana"
            
            ElseIf UCase(node("Category")) = "OTH" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Other Incomes"
            
            End If
           
    End If
        
    If Sheet1.Cells(rowcount, DescriptionColNo).Locked = False Then
        
        If UCase(node("SubCategory")) = "10(1)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(1)-Agricultural income(Less than or equal to 5000)"
        
        ElseIf UCase(node("SubCategory")) = "10(30)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(30)-subsidy received from or through the Tea Board"
        
        ElseIf UCase(node("SubCategory")) = "10(31)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(31)-Rubber/Coffee/Tea development accounts/funds"
        
        ElseIf UCase(node("SubCategory")) = "10(10BB)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(10BB)-payments made under the Bhopal Gas Leak Disaster"
        'Ankita_23/03/2026============
        ElseIf UCase(node("SubCategory")) = "10(10BC)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(10BC)-Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster"
        
        ElseIf UCase(node("SubCategory")) = "10(17A)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(17A)-Award instituted by Government"
        
        ElseIf UCase(node("SubCategory")) = "10(12AB)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(12AB)-any sum received as lump sum amount as per clause (vi) of paragraph 2 of the notification number FX-1/3/2024-PR"
        
        ElseIf UCase(node("SubCategory")) = "10(15)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(15)-Interest on specified securities/investments"
          'Ankita_removed as per V0.6
'        ElseIf UCase(node("SubCategory")) = "10(23EA)" Then
'        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(23EA)-Contributions received from recognised stock exchanges"
        
        ElseIf UCase(node("SubCategory")) = "10(23FBB)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(23FBB)-income referred to in section 115UB, accruing or arising to, or received by, a unit holder of an investment fund"
        
        ElseIf UCase(node("SubCategory")) = "10(23FD)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(23FD)Unit holder income from Business Trust (certain parts)"
        
        ElseIf UCase(node("SubCategory")) = "10(35)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(35)-Income from specified Mutual Funds "
        
        ElseIf UCase(node("SubCategory")) = "10(35A)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(35A)-distributed income referred to in section 115TA received from a securitisation trust"
        
        ElseIf UCase(node("SubCategory")) = "10(12C)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(12C)-Agniveer Corpus Fund income"
        
        ElseIf UCase(node("SubCategory")) = "10(18)" Then
'        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(18)-Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award"
         Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(18)-Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award"
        
        ElseIf UCase(node("SubCategory")) = "10(19)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(19)-Armed Forces Family pension in case of death during operational duty"
        
        ElseIf UCase(node("SubCategory")) = "10(23AA)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(23AA)-Sum received by any person on behalf of any Fund established by the armed forces"
        
        ElseIf UCase(node("SubCategory")) = "DMD" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "Defense Medical Disability Pension"
        
        ElseIf UCase(node("SubCategory")) = "10(32)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(32)-Minor child’s income—small exemption"
        
        ElseIf UCase(node("SubCategory")) = "10(43)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(43)-Reverse mortgage—payments to senior citizens"
        
        ElseIf UCase(node("SubCategory")) = "10(19A)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(19A)-Annual value of one palace in occupation of ex-ruler"
        
        ElseIf UCase(node("SubCategory")) = "10(26)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(26)-Any income as referred to in section 10(26)"
        
        ElseIf UCase(node("SubCategory")) = "10(26AAA)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(26AAA)-Any income as referred to in section 10(26AAA)"
        
        ElseIf UCase(node("SubCategory")) = "10(10D)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(10D)-Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(10D)"
        
        ElseIf UCase(node("SubCategory")) = "10(11)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(11)-Statutory Provident Fund received"
        
        ElseIf UCase(node("SubCategory")) = "10(11A)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(11A)-Sum received from an account opened under the Sukanya Samriddhi Yojana"
        
        ElseIf UCase(node("SubCategory")) = "10(12)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(12)-Recognized Provident Fund received"
        
        ElseIf UCase(node("SubCategory")) = "10(12A)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(12A)-Any payment from the National Pension System Trust to an assessee"
        
        ElseIf UCase(node("SubCategory")) = "10(12AA)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(12AA)-any payment from the National Pension System Trust "
        
        ElseIf UCase(node("SubCategory")) = "10(12B)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(12B)-Any payment from the National Pension System Trust to an Central Govt. Employee"
        
        ElseIf UCase(node("SubCategory")) = "10(12BA)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(12BA)-partial withdrawal made from the National Pension System"
        
        ElseIf UCase(node("SubCategory")) = "10(13)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(13)-Approved superannuation fund received"
        
        ElseIf UCase(node("SubCategory")) = "10(25)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(25)-Sum received by trustees on behalf of approved superannuation, gratuity, or pension funds"
        
'        ElseIf UCase(node("SubCategory")) = "10(25A)" Then
'        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(25A)-any income under Employees' State Insurance Fund"
        
        ElseIf UCase(node("SubCategory")) = "10(44)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(44)-Income received by any person for, or on behalf of, the New Pension System Trust"
        
        ElseIf UCase(node("SubCategory")) = "10(2)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(2)-Member’s share from HUF"
        
        ElseIf UCase(node("SubCategory")) = "10(16)" Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "10(16)-Scholarships for education"
        
        'Malli_AY_2026_27  17/06/2026
        ElseIf UCase(node("SubCategory")) = UCase("Incmexmptcircular") Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "Income exempt as per CBDT Circular"
        
        ElseIf UCase(node("SubCategory")) = UCase("Incmexmptnotification") Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "Income exempt as per CBDT Notification"
        
        ElseIf UCase(node("SubCategory")) = UCase("Receiptnotincme") Then
        Sheet1.Cells(rowcount, DescriptionColNo).Value = "Receipts not in the nature of income"
        
        '----------------------------
        
        End If
    End If
'================================

                'Malli_AY_2026_27
                If Sheet1.Cells(rowcount, DescEIColNo).Locked = False Then
                   Sheet1.Cells(rowcount, DescEIColNo).Value = node("Description")
                End If
                
                '----------------
                Sheet1.Cells(rowcount, AmtColNo).Value = node("OthAmount")
    
            cnt = cnt + 1
        Next node
        RecTDS1 = cnt
    End If
    
    'Sheet1.Range("IncD.Section80U").Value = Section80U
'Konda-20-01-2026=============================
'    Sheet1.Range("Increliefus89A").Value = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("Increliefus89A")

'    Sheet1.Range("OSIncreliefus89A").Value = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("Increliefus89AOS")
'==============================
    'Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.LessDeduction57").Value = DeductionUs57iia
    
    ImportScheduleHP (jsonObject)   'Add by Konda on 22-01-2026
End Function
'New Schema updated by Konda as AY-2026-27 on 26-12-2025
Sub AddDiffRows_80CC(DiffRows As Long)
    setDiffTblinfo_80CCC
    Sheet1.Activate
    searchLastRow1 ("Type_80CCC")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'New Schema updated by Konda as AY-2026-27 on 26-12-2025
Sub setDiffTblinfo_80CCC()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim end_TaxP
    
    ccount = 0
    mIntCells = Sheet1.Range("Type_80CCC").item(0.1).count
    Set rangecells = Sheet1.Range("Type_80CCC").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    DefinedgridNameRange = "Sl_80CCC||Type_80CCC||Name_80CCC||Amount_80CCC"         'Updated by Ankita on 04/05/2026
 End Sub
 'New Schema updated by Konda as AY-2026-27 on 26-12-2025
 Sub AddDiffRows_80CCD1(DiffRows As Long)
    setDiffTblinfo_80CCD1
    Sheet1.Activate
    searchLastRow1 ("Type_80CCD1")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'New Schema updated by Konda as AY-2026-27 on 26-12-2025
Sub setDiffTblinfo_80CCD1()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim end_TaxP
    
    ccount = 0
    mIntCells = Sheet1.Range("Type_80CCD1").item(0.1).count
    Set rangecells = Sheet1.Range("Type_80CCD1").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    DefinedgridNameRange = "Type_80CCD1||Name_80CCD1||Amount_80CCD1"
 End Sub
 'New Schema updated by Konda as AY-2026-27 on 26-12-2025
 Sub AddDiffRows_80CCD1b(DiffRows As Long)
    setDiffTblinfo_80CCD1b
    Sheet1.Activate
    searchLastRow1 ("Type_80CCD1b")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'New Schema updated by Konda as AY-2026-27 on 26-12-2025
Sub setDiffTblinfo_80CCD1b()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim end_TaxP
    
    ccount = 0
    mIntCells = Sheet1.Range("Type_80CCD1b").item(0.1).count
    Set rangecells = Sheet1.Range("Type_80CCD1b").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    DefinedgridNameRange = "Type_80CCD1b||Name_80CCD1b||Amount_80CCD1b"
 End Sub
 'New Schema updated by Konda as AY-2026-27 on 26-12-2025
 Function searchLastRow1(ByVal gridRangeName As String) As String
On Error Resume Next
Dim searchLastRow
 strCurrActiveCellRange = Replace(ActiveSheet.Range(gridRangeName).AddressLocal, "$", "")
 strNewActiveCellRange = Mid(strCurrActiveCellRange, InStr(1, strCurrActiveCellRange, ":") + 1, Len(strCurrActiveCellRange))
 ActiveSheet.Range(strNewActiveCellRange).Select
 searchLastRow = strCurrActiveCellRange
End Function

'Konda added AY 2025-26 on 04/02/2025
Function ImportLTCG112A(jsonText As String)

On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist  As Object

Dim text

Dim TotalSales, TotalCost 'LongCap

Set jsonObject = ParseJson(jsonText)

TotalSales = jsonObject("ITR")("ITR1")("LTCG112A")("TotSaleCnsdrn")
TotalCost = jsonObject("ITR")("ITR1")("LTCG112A")("TotCstAcqisn")
'LongCap = jsonObject("ITR")("ITR1")("LTCG112A")("LongCap112A")

    If TotalSales <> "" And Sheet1.Range("IncD.Sale_LTCG").Locked = False Then
            Sheet1.Range("IncD.Sale_LTCG").Value = TotalSales
    End If
    
    If TotalCost <> "" And Sheet1.Range("IncD.Cost_LTCG").Locked = False Then
            Sheet1.Range("IncD.Cost_LTCG").Value = TotalCost
    End If
    
'    If LongCap <> "" And Sheet1.Range("IncD.CG_LTCG").Locked = False Then
'            Sheet1.Range("IncD.CG_LTCG").Value = LongCap
'    End If

End Function
Function ImportTaxComputation(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set jsonObject = ParseJson(jsonText)

'Sheet1.Unprotect Password:=getmsgstate
Sheet1.Range("IncD.TotalTaxPayable").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("TotalTaxPayable")
'Sheet1.Unprotect Password:=getmsgstate
'Sheet1.Range("IncD.Rebate87A").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("Rebate87A")
'Sheet1.Unprotect Password:=getmsgstate
'Sheet1.Range("IncD.TaxPayableOnRebate").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("TaxPayableOnRebate")
'Sheet1.Unprotect Password:=getmsgstate
'Sheet1.Range("IncD.EducationCess").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("EducationCess")
'Sheet1.Unprotect Password:=getmsgstate
'Sheet1.Range("IncD.GrossTaxLiability").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("GrossTaxLiability")
'Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False
Sheet1.Range("IncD.Section89").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("Section89")
'Sheet1.Range("IncD.Section89A").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("Section89A")
Application.EnableEvents = True
'Sheet1.Unprotect Password:=getmsgstate
'Sheet1.Range("IncD.NetTaxLiability").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("NetTaxLiability")
'Sheet1.Unprotect Password:=getmsgstate
'Sheet1.Range("IncD.TotalIntrstPay").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("TotalIntrstPay")
'Sheet1.Unprotect Password:=getmsgstate
Sheet1.Range("IncD.IntrstPayUs234A").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("IntrstPay")("IntrstPayUs234A")
'Sheet1.Unprotect Password:=getmsgstate
Sheet1.Range("IncD.IntrstPayUs234B").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("IntrstPay")("IntrstPayUs234B")
'Sheet1.Unprotect Password:=getmsgstate
Sheet1.Range("IncD.IntrstPayUs234C").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("IntrstPay")("IntrstPayUs234C")
'Sheet1.Unprotect Password:=getmsgstate
Sheet1.Range("IncD.IntrstPayUs234F").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("IntrstPay")("LateFilingFee234F")

'Sheet1.Unprotect Password:=getmsgstate
'Sheet1.Range("IncD.TotTaxPlusIntrstPay").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("TotTaxPlusIntrstPay")

End Function
Sub teet()
ImportTDSonOthThanSals ("ABC")
End Sub
Function ImportTDSonOthThanSals(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim TANColNo, DEDNameColNo, DEDAmountDeducted, UTNColNo, SectionTDSColNo, FYColNo, TaxColNo, ClaimColNo, BroughtFwdTDSAmt As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, cnt, rowcount As Long

Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR1")("TDSonOthThanSals")("TDSonOthThanSal")
    
    TANColNo = Sheet2.Range("TDSoth.TAN").Column
    DEDNameColNo = Sheet2.Range("TDSoth.EmployerOrDeductorOrCollecterName").Column
    DEDAmountDeducted = Sheet2.Range("TDSoth.AmountDeducted").Column
'    SectionTDSColNo = Sheet2.Range("TDsOthr.SectionTDS").Column   'Konda added 0n 05-02-2025
    SectionTDSColNo = Sheet2.Range("TDsOthr.SectionTDS").Column   'Konda added 0n 22-04-2025
    FYColNo = Sheet2.Range("TDSoth.DeductedYear").Column
    BroughtFwdTDSAmt = Sheet2.Range("TDSoth.TotTDSOnAmtPaid").Column
    TaxColNo = Sheet2.Range("TDSoth.6income").Column
    
    TotalExRow = Range("TDSoth.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    
    Application.EnableEvents = False
    
        Sheet2.Range("TDSoth.TAN").ClearContents
        Sheet2.Range("TDSoth.EmployerOrDeductorOrCollecterName").ClearContents
 '       Sheet2.Range("TDsOthr.SectionTDS").ClearContents    'Konda added 0n 05-02-2025
        Sheet2.Range("TDsOthr.SectionTDS").ClearContents    'Konda added 0n 22-04-2025
        Sheet2.Range("TDSoth.AmountDeducted").ClearContents
        Sheet2.Range("TDSoth.DeductedYear").ClearContents
        Sheet2.Range("TDSoth.TotTDSOnAmtPaid").ClearContents
        Sheet2.Range("TDSoth.6income").ClearContents
        
        Application.EnableEvents = True
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDSoth (TotalDiffRow)
    End If
      
    rowcount = getRowNo(Sheet2.Range("TDSoth.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
  If TotalXMLRow > 0 Then
    For Each node In Nodelist
        rowcount = rowcount + 1
            Sheet2.Cells(rowcount, TANColNo).Value = node("EmployerOrDeductorOrCollectDetl")("TAN")
            Sheet2.Cells(rowcount, DEDNameColNo).Value = node("EmployerOrDeductorOrCollectDetl")("EmployerOrDeductorOrCollecterName")
            Sheet2.Cells(rowcount, DEDAmountDeducted).Value = node("AmtForTaxDeduct")
            
 'Konda AY 2025-26 updated V0.4 on 12-02-2025
'Konda AY 2025-26 added on 05-02-2025
'            Dim SecTDSDeducted
'        If Sheet2.Cells(rowcount, SectionTDSColNo).Locked = False Then
'            Dim SectionTDS As Variant
'            SectionTDS = node("SecTDSDeducted")
''            If SectionTDS = "92A" Then
''            SectionTDS = "192-Salary-Payment to Government employees other than Indian Government employees"
''            ElseIf SectionTDS = "92B" Then
''            SectionTDS = "192-Salary-Payment to employees other than Government employees"
''            ElseIf SectionTDS = "92C" Then
''            SectionTDS = "192-Salary-Payment to Indian Government employees"
'            If SectionTDS = "192A" Then
'            SectionTDS = "192A-TDS on PF withdrawal"
'            ElseIf SectionTDS = "193" Then
'            SectionTDS = "193 -Interest on Securities"
'            ElseIf SectionTDS = "194" Then
'            SectionTDS = "194 -Dividends"
'            ElseIf SectionTDS = "94A" Then
'            SectionTDS = "194A-Interest other than 'Interest on securities'"
'            ElseIf SectionTDS = "4DA" Then
'            SectionTDS = "194DA-Payment in respect of life insurance policy"
'            ElseIf SectionTDS = "4EE" Then
'            SectionTDS = "194EE-Payments in respect of deposits under National Savings"
'            ElseIf SectionTDS = "4IB" Then
'            SectionTDS = "194IB-Payment of rent by certain individuals or Hindu undivided"
'            ElseIf SectionTDS = "94K" Then
'            SectionTDS = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
'            ElseIf SectionTDS = "4BA1" Then
'            SectionTDS = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
'            ElseIf SectionTDS = "4BA2" Then
'            SectionTDS = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
'            ElseIf SectionTDS = "LBB" Then
'            SectionTDS = "194LBB-Income in respect of units of investment fund"
'            ElseIf SectionTDS = "94P" Then
'            SectionTDS = "194P-Deduction of tax in case of specified senior citizen"
'            End If
'            Sheet2.Cells(rowcount, SectionTDSColNo).Value = SectionTDS
'        End If
'    '--end--
 'End AY 2025-26 updated V0.4 on 12-02-2025
 '================================================================================================================================
    'Konda AY 2025-26 updated V0.4 on 12-02-2025

        Dim TDSSection_TDS2
        If Sheet2.Cells(rowcount, SectionTDSColNo).Locked = False Then
            Dim SectionTDS As Variant
            SectionTDS = node("TDSSection")

            If SectionTDS = "92A" Then
            TDSSection_TDS2 = "192-Salary-Payment to Government employees other than Indian Government employees"
            ElseIf SectionTDS = "92B" Then
            TDSSection_TDS2 = "192-Salary-Payment to employees other than Government employees"
            ElseIf SectionTDS = "92C" Then
            TDSSection_TDS2 = "192-Salary-Payment to Indian Government employees"
            ElseIf SectionTDS = "192A" Then
            TDSSection_TDS2 = "192A-TDS on PF withdrawal"
            ElseIf SectionTDS = "193" Then
            TDSSection_TDS2 = "193-Interest on Securities"
            ElseIf SectionTDS = "194" Then
            TDSSection_TDS2 = "194-Dividends"
            ElseIf SectionTDS = "94A" Then
            TDSSection_TDS2 = "194A-Interest other than 'Interest on securities'"
            ElseIf SectionTDS = "94B" Then
            TDSSection_TDS2 = "194B-Winning from lottery or crossword puzzle"
            ElseIf SectionTDS = "94BA" Then
            TDSSection_TDS2 = "194BA-Winnings from online games"
            ElseIf SectionTDS = "4BB" Then
            TDSSection_TDS2 = "194BB-Winning from horse race"
            ElseIf SectionTDS = "94C" Then
            TDSSection_TDS2 = "194C-Payments to contractors and sub-contractors"
            ElseIf SectionTDS = "94D" Then
            TDSSection_TDS2 = "194D-Insurance commission"
            ElseIf SectionTDS = "4DA" Then
            TDSSection_TDS2 = "194DA-Payment in respect of life insurance policy"
            ElseIf SectionTDS = "94E" Then
            TDSSection_TDS2 = "194E-Payments to non-resident sportsmen or sports associations"
            ElseIf SectionTDS = "4EE" Then
            TDSSection_TDS2 = "194EE-Payments in respect of deposits under National Savings"
            ElseIf SectionTDS = "4F" Then
            TDSSection_TDS2 = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
            ElseIf SectionTDS = "4G" Then
            TDSSection_TDS2 = "194G-Commission, price, etc. on sale of lottery tickets"
            ElseIf SectionTDS = "4H" Then
            TDSSection_TDS2 = "194H-Commission or brokerage"
            ElseIf SectionTDS = "4-IA" Then
            TDSSection_TDS2 = "194I(a)-Rent on hiring of plant and machinery"
            ElseIf SectionTDS = "4-IB" Then
            TDSSection_TDS2 = "194I(b)-Rent on other than plant and machinery"
            ElseIf SectionTDS = "4IA" Then
            TDSSection_TDS2 = "194IA-TDS on Sale of immovable property"
            ElseIf SectionTDS = "4IB" Then
            TDSSection_TDS2 = "194IB-Payment of rent by certain individuals or Hindu undivided"
            ElseIf SectionTDS = "4IC" Then
            TDSSection_TDS2 = "194IC-Payment under specified agreement"
            ElseIf SectionTDS = "94J-A" Then
            TDSSection_TDS2 = "194J(a)-Fees for technical services"
            ElseIf SectionTDS = "94J-B" Then
            TDSSection_TDS2 = "194J(b)-Fees for professional  services or royalty etc"
            ElseIf SectionTDS = "94K" Then
            TDSSection_TDS2 = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
            ElseIf SectionTDS = "4LA" Then
            TDSSection_TDS2 = "194LA-Payment of compensation on acquisition of certain immovable"
            ElseIf SectionTDS = "4LB" Then
            TDSSection_TDS2 = "194LB-Income by way of Interest from Infrastructure Debt fund"
            ElseIf SectionTDS = "4LC1" Then
            TDSSection_TDS2 = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
            ElseIf SectionTDS = "4LC2" Then
            TDSSection_TDS2 = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
            ElseIf SectionTDS = "4LC3" Then
            TDSSection_TDS2 = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
            ElseIf SectionTDS = "4BA1" Then
            TDSSection_TDS2 = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
            ElseIf SectionTDS = "4BA2" Then
            TDSSection_TDS2 = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
            ElseIf SectionTDS = "LBA1" Then
            TDSSection_TDS2 = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
            ElseIf SectionTDS = "LBA2" Then
            TDSSection_TDS2 = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
            ElseIf SectionTDS = "LBA3" Then
            TDSSection_TDS2 = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
            ElseIf SectionTDS = "LBB" Then
            TDSSection_TDS2 = "194LBB-Income in respect of units of investment fund"
            ElseIf SectionTDS = "94R" Then
            TDSSection_TDS2 = "194R-Benefits or perquisites of business or profession"
            ElseIf SectionTDS = "94S" Then
            TDSSection_TDS2 = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
            ElseIf SectionTDS = "94B-P" Then
            TDSSection_TDS2 = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
            ElseIf SectionTDS = "94R-P" Then
            TDSSection_TDS2 = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
            ElseIf SectionTDS = "94S-P" Then
            TDSSection_TDS2 = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
            ElseIf SectionTDS = "LBC" Then
            TDSSection_TDS2 = "194LBC-Income in respect of investment in securitization trust"
            ElseIf SectionTDS = "4LD" Then
            TDSSection_TDS2 = "194LD-TDS on interest on bonds / government securities"
            ElseIf SectionTDS = "94M" Then
            TDSSection_TDS2 = "194M-Payment of certain sums by certain individuals or HUF"
            ElseIf SectionTDS = "94N" Then
            TDSSection_TDS2 = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
            ElseIf SectionTDS = "94N-F" Then
            TDSSection_TDS2 = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
            ElseIf SectionTDS = "94N-C" Then
            TDSSection_TDS2 = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
            ElseIf SectionTDS = "94N-FT" Then
            TDSSection_TDS2 = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
            ElseIf SectionTDS = "94O" Then
            TDSSection_TDS2 = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
            ElseIf SectionTDS = "94P" Then
            TDSSection_TDS2 = "194P-Deduction of tax in case of specified senior citizen"
            ElseIf SectionTDS = "94Q" Then
            TDSSection_TDS2 = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
            ElseIf SectionTDS = "195" Then
            TDSSection_TDS2 = "195-Other sums payable to a non-resident"
            ElseIf SectionTDS = "96A" Then
            TDSSection_TDS2 = "196A-Income in respect of units of non-residents"
            ElseIf SectionTDS = "96B" Then
            TDSSection_TDS2 = "196B-Payments in respect of units to an offshore fund"
            ElseIf SectionTDS = "96C" Then
            TDSSection_TDS2 = "196C-Income from foreign currency bonds or shares of Indian"
            ElseIf SectionTDS = "96D" Then
            TDSSection_TDS2 = "196D-Income of foreign institutional investors from securities"
            ElseIf SectionTDS = "96DA" Then
            TDSSection_TDS2 = "196D(1A)-Income of specified fund from securities"
            ElseIf SectionTDS = "94BA-P" Then
            TDSSection_TDS2 = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
            Else
            TDSSection_TDS2 = ""
            End If
            If TDSSection_TDS2 <> "" Then
            Sheet2.Cells(rowcount, SectionTDSColNo).Value = TDSSection_TDS2
            TDSSection_TDS2 = ""
            End If
            End If
            
            
        '--------------------------------------------------
            
            
            'TAX DETAILS-E2
            'Sheet2.Cells(rowcount, FYColNo).Value = Node("DeductedYr")
            'TAX DETAILS-C2 2024-25 Bindu
            Sheet2.Cells(rowcount, FYColNo).Value = Trim("2023-24")
            
            'Newly added by Chetan C M
            If Sheet2.Cells(rowcount, FYColNo).Locked = False Then
        
                Sheet2.Cells(rowcount, FYColNo).Value = node("DeductedYr")
    
                Dim DeductedYr_1 As Variant
                DeductedYr_1 = UCase(node("DeductedYr"))
                
                'New Schema updated by Konda as AY-2026-27 on 26-12-2025
                If DeductedYr_1 = "2025" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2025-26"
         
                ElseIf DeductedYr_1 = "2024" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2024-25"
      '--------------------------
                
                ElseIf DeductedYr_1 = "2023" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2023-24"
                
                ElseIf DeductedYr_1 = "2022" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2022-23"
                
                ElseIf DeductedYr_1 = "2021" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2021-22"
                
                ElseIf DeductedYr_1 = "2020" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2020-21"
                
                ElseIf DeductedYr_1 = "2019" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2019-20"
                
                ElseIf DeductedYr_1 = "2018" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2018-19"
                
                ElseIf DeductedYr_1 = "2017" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2017-18"
                
                ElseIf DeductedYr_1 = "2016" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2016-17"
                
                ElseIf DeductedYr_1 = "2015" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2015-16"
                
                ElseIf DeductedYr_1 = "2014" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2014-15"
                
                ElseIf DeductedYr_1 = "2013" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2013-14"
                
                ElseIf DeductedYr_1 = "2012" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2012-13"
                
                ElseIf DeductedYr_1 = "2011" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2011-12"
                
                ElseIf DeductedYr_1 = "2010" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2010-11"
                
                ElseIf DeductedYr_1 = "2009" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2009-10"
                
                ElseIf DeductedYr_1 = "2008" Then
                Sheet2.Cells(rowcount, FYColNo).Value = "2008-09"
                Else
                 Sheet2.Cells(rowcount, FYColNo).Value = "(Select)"
                End If
            End If
            'addition end
            
            Sheet2.Cells(rowcount, BroughtFwdTDSAmt).Value = node("TotTDSOnAmtPaid")
            Sheet2.Cells(rowcount, TaxColNo).Value = node("ClaimOutOfTotTDSOnAmtPaid")
        
        cnt = cnt + 1
    Next node
 End If
    RecTDS1 = cnt
    
    
   End Function
Function ImportTDSonSalaries(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim Capacity As Variant
Dim TANNoEmployer, TDSNameOfEmployer, TDSIncomeCharge, TDSTotalTax As Long
Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, RecTDS1, rowcount As Long

Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR1")("TDSonSalaries")("TDSonSalary")

    TANNoEmployer = Range("TDSal.TAN").Column
    TDSNameOfEmployer = Range("TDSal.EmployerOrDeductorOrCollecterName").Column
    TDSIncomeCharge = Range("TDSal.IncChrgSalary").Column
    TDSTotalTax = Range("TDSal.TotalTDSSalary").Column
    
    TotalExRow = Range("TDSal.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
       Application.EnableEvents = False
       
        Range("TDSal.TAN").ClearContents
        Range("TDSal.EmployerOrDeductorOrCollecterName").ClearContents
        Range("TDSal.IncChrgSalary").ClearContents
        Range("TDSal.TotalTDSSalary").ClearContents
        
        Application.EnableEvents = False
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDS1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet2.Range("TDSal.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        UpdateProgressBar
            Sheet2.Cells(rowcount, TANNoEmployer).Value = node("EmployerOrDeductorOrCollectDetl")("TAN")
            Sheet2.Cells(rowcount, TDSNameOfEmployer).Value = node("EmployerOrDeductorOrCollectDetl")("EmployerOrDeductorOrCollecterName")
            Sheet2.Cells(rowcount, TDSIncomeCharge).Value = node("IncChrgSal")
            Sheet2.Cells(rowcount, TDSTotalTax).Value = node("TotalTDSSal")
            
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Function
Function ImportScheduleTDS3Dtls(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim Node1, nodeList1 As Object
Dim TANColNo1, TDSAADhaar, DEDNameColNo1, DEDAmountDeducted1, UTNColNo1, SectionTDSColNo1, FYColNo1, TaxColNo1, ClaimColNo1, BroughtFwdTDSAmt1 As Variant
Dim TotalExRow1, TotalXMLRow1, TotalDiffRow1, cnt, RecTDS11, rowcount As Long

Set jsonObject = ParseJson(jsonText)

    Set nodeList1 = jsonObject("ITR")("ITR1")("ScheduleTDS3Dtls")("TDS3Details")
    
    TANColNo1 = Sheet2.Range("TDS26QB.PAN").Column
    TDSAADhaar = Sheet2.Range("TDS26QB.Aadhar_Number").Column
    DEDNameColNo1 = Sheet2.Range("TDS26QB.EmployerOrDeductorName").Column
    DEDAmountDeducted1 = Sheet2.Range("TDS26QB.AmountDeducted").Column
    FYColNo1 = Sheet2.Range("TDS26QB.DeductedYear").Column
    'SectionTDSColNo1 = Sheet2.Range("TDsOthr.SectionTDS_ii").Column   'Konda added 0n 05-02-2025
    SectionTDSColNo1 = Sheet2.Range("TDsOthr2.SectionTDSDeducted").Column   'Konda added 0n 22-04-2025
    BroughtFwdTDSAmt1 = Sheet2.Range("TDS26QB.TotTDSOnAmtPaid").Column
    TaxColNo1 = Sheet2.Range("TDS26QB.6income").Column
    
    TotalExRow1 = Range("TDS26QB.PAN").Rows.count
    
    TotalXMLRow1 = nodeList1.count
    TotalDiffRow1 = TotalXMLRow1 - TotalExRow1
    
    If (TotalXMLRow1 > 0) Then
    
    Application.EnableEvents = False
    
        Sheet2.Range("TDS26QB.PAN").ClearContents
        Sheet2.Range("TDS26QB.Aadhar_Number").ClearContents
        Sheet2.Range("TDS26QB.EmployerOrDeductorName").ClearContents
 '       Sheet2.Range("TDsOthr.SectionTDS_ii").ClearContents     'Konda added 0n 05-02-2025
        Sheet2.Range("TDsOthr2.SectionTDSDeducted").ClearContents   'Konda added on 22-04-2025
        Sheet2.Range("TDS26QB.AmountDeducted").ClearContents
        Sheet2.Range("TDS26QB.DeductedYear").ClearContents
        Sheet2.Range("TDS26QB.TotTDSOnAmtPaid").ClearContents
        Sheet2.Range("TDS26QB.6income").ClearContents
        
        Application.EnableEvents = True
        
    End If
    
    If (TotalDiffRow1 > 0) Then
     AddDiffRows_TDSoth1 (TotalDiffRow1)
    End If
    
    
    rowcount = getRowNo(Sheet2.Range("TDS26QB.PAN").name)
    rowcount = rowcount - 1
    cnt = 0
  If TotalXMLRow1 > 0 Then
    For Each Node1 In nodeList1
        rowcount = rowcount + 1
            Sheet2.Cells(rowcount, TANColNo1).Value = Node1("PANofTenant")
            Sheet2.Cells(rowcount, TDSAADhaar).Value = Node1("AadhaarofTenant")
            Sheet2.Cells(rowcount, DEDNameColNo1).Value = Node1("NameOfTenant")
            
'================================================================================================================================
    'Konda AY 2025-26 updated V0.6.2 on 22-04-2025

        Dim TDSSection_TDS3
        If Sheet2.Cells(rowcount, SectionTDSColNo1).Locked = False Then
            Dim SectionTDS As Variant
            SectionTDS = Node1("TDSSection")

            If SectionTDS = "92A" Then
            TDSSection_TDS3 = "192-Salary-Payment to Government employees other than Indian Government employees"
            ElseIf SectionTDS = "92B" Then
            TDSSection_TDS3 = "192-Salary-Payment to employees other than Government employees"
            ElseIf SectionTDS = "92C" Then
            TDSSection_TDS3 = "192-Salary-Payment to Indian Government employees"
            ElseIf SectionTDS = "192A" Then
            TDSSection_TDS3 = "192A-TDS on PF withdrawal"
            ElseIf SectionTDS = "193" Then
            TDSSection_TDS3 = "193-Interest on Securities"
            ElseIf SectionTDS = "194" Then
            TDSSection_TDS3 = "194-Dividends"
            ElseIf SectionTDS = "94A" Then
            TDSSection_TDS3 = "194A-Interest other than 'Interest on securities'"
            ElseIf SectionTDS = "94B" Then
            TDSSection_TDS3 = "194B-Winning from lottery or crossword puzzle"
            ElseIf SectionTDS = "94BA" Then
            TDSSection_TDS3 = "194BA-Winnings from online games"
            ElseIf SectionTDS = "4BB" Then
            TDSSection_TDS3 = "194BB-Winning from horse race"
            ElseIf SectionTDS = "94C" Then
            TDSSection_TDS3 = "194C-Payments to contractors and sub-contractors"
            ElseIf SectionTDS = "94D" Then
            TDSSection_TDS3 = "194D-Insurance commission"
            ElseIf SectionTDS = "4DA" Then
            TDSSection_TDS3 = "194DA-Payment in respect of life insurance policy"
            ElseIf SectionTDS = "94E" Then
            TDSSection_TDS3 = "194E-Payments to non-resident sportsmen or sports associations"
            ElseIf SectionTDS = "4EE" Then
            TDSSection_TDS3 = "194EE-Payments in respect of deposits under National Savings"
            ElseIf SectionTDS = "4F" Then
            TDSSection_TDS3 = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
            ElseIf SectionTDS = "4G" Then
            TDSSection_TDS3 = "194G-Commission, price, etc. on sale of lottery tickets"
            ElseIf SectionTDS = "4H" Then
            TDSSection_TDS3 = "194H-Commission or brokerage"
            ElseIf SectionTDS = "4-IA" Then
            TDSSection_TDS3 = "194I(a)-Rent on hiring of plant and machinery"
            ElseIf SectionTDS = "4-IB" Then
            TDSSection_TDS3 = "194I(b)-Rent on other than plant and machinery"
            ElseIf SectionTDS = "4IA" Then
            TDSSection_TDS3 = "194IA-TDS on Sale of immovable property"
            ElseIf SectionTDS = "4IB" Then
            TDSSection_TDS3 = "194IB-Payment of rent by certain individuals or Hindu undivided"
            ElseIf SectionTDS = "4IC" Then
            TDSSection_TDS3 = "194IC-Payment under specified agreement"
            ElseIf SectionTDS = "94J-A" Then
            TDSSection_TDS3 = "194J(a)-Fees for technical services"
            ElseIf SectionTDS = "94J-B" Then
            TDSSection_TDS3 = "194J(b)-Fees for professional  services or royalty etc"
            ElseIf SectionTDS = "94K" Then
            TDSSection_TDS3 = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
            ElseIf SectionTDS = "4LA" Then
            TDSSection_TDS3 = "194LA-Payment of compensation on acquisition of certain immovable"
            ElseIf SectionTDS = "4LB" Then
            TDSSection_TDS3 = "194LB-Income by way of Interest from Infrastructure Debt fund"
            ElseIf SectionTDS = "4LC1" Then
            TDSSection_TDS3 = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
            ElseIf SectionTDS = "4LC2" Then
            TDSSection_TDS3 = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
            ElseIf SectionTDS = "4LC3" Then
            TDSSection_TDS3 = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
            ElseIf SectionTDS = "4BA1" Then
            TDSSection_TDS3 = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
            ElseIf SectionTDS = "4BA2" Then
            TDSSection_TDS3 = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
            ElseIf SectionTDS = "LBA1" Then
            TDSSection_TDS3 = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
            ElseIf SectionTDS = "LBA2" Then
            TDSSection_TDS3 = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
            ElseIf SectionTDS = "LBA3" Then
            TDSSection_TDS3 = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
            ElseIf SectionTDS = "LBB" Then
            TDSSection_TDS3 = "194LBB-Income in respect of units of investment fund"
            ElseIf SectionTDS = "94R" Then
            TDSSection_TDS3 = "194R-Benefits or perquisites of business or profession"
            ElseIf SectionTDS = "94S" Then
            TDSSection_TDS3 = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
            ElseIf SectionTDS = "94B-P" Then
            TDSSection_TDS3 = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
            ElseIf SectionTDS = "94R-P" Then
            TDSSection_TDS3 = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
            ElseIf SectionTDS = "94S-P" Then
            TDSSection_TDS3 = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
            ElseIf SectionTDS = "LBC" Then
            TDSSection_TDS3 = "194LBC-Income in respect of investment in securitization trust"
            ElseIf SectionTDS = "4LD" Then
            TDSSection_TDS3 = "194LD-TDS on interest on bonds / government securities"
            ElseIf SectionTDS = "94M" Then
            TDSSection_TDS3 = "194M-Payment of certain sums by certain individuals or HUF"
            ElseIf SectionTDS = "94N" Then
            TDSSection_TDS3 = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
            ElseIf SectionTDS = "94N-F" Then
            TDSSection_TDS3 = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
            ElseIf SectionTDS = "94N-C" Then
            TDSSection_TDS3 = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
            ElseIf SectionTDS = "94N-FT" Then
            TDSSection_TDS3 = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
            ElseIf SectionTDS = "94O" Then
            TDSSection_TDS3 = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
            ElseIf SectionTDS = "94P" Then
            TDSSection_TDS3 = "194P-Deduction of tax in case of specified senior citizen"
            ElseIf SectionTDS = "94Q" Then
            TDSSection_TDS3 = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
            ElseIf SectionTDS = "195" Then
            TDSSection_TDS3 = "195-Other sums payable to a non-resident"
            ElseIf SectionTDS = "96A" Then
            TDSSection_TDS3 = "196A-Income in respect of units of non-residents"
            ElseIf SectionTDS = "96B" Then
            TDSSection_TDS3 = "196B-Payments in respect of units to an offshore fund"
            ElseIf SectionTDS = "96C" Then
            TDSSection_TDS3 = "196C-Income from foreign currency bonds or shares of Indian"
            ElseIf SectionTDS = "96D" Then
            TDSSection_TDS3 = "196D-Income of foreign institutional investors from securities"
            ElseIf SectionTDS = "96DA" Then
            TDSSection_TDS3 = "196D(1A)-Income of specified fund from securities"
            ElseIf SectionTDS = "94BA-P" Then
            TDSSection_TDS3 = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
            Else
            TDSSection_TDS3 = ""
            End If
            If TDSSection_TDS3 <> "" Then
            Sheet2.Cells(rowcount, SectionTDSColNo1).Value = TDSSection_TDS3
            TDSSection_TDS3 = ""
            End If
            End If
    '-------------------------------------------------------------------------
            Sheet2.Cells(rowcount, DEDAmountDeducted1).Value = Node1("GrsRcptToTaxDeduct")
            
'Konda AY 2025-26 Commeted as V0.4 on 12-02-2025
'  'Konda AY 2025-26 added on 05-02-2025
'            Dim SecTDSDeducted
'        If Sheet2.Cells(rowcount, SectionTDSColNo1).Locked = False Then
'            Dim SectionTDS1 As Variant
'            SectionTDS1 = Node1("SecTDSDeducted")
''            If SectionTDS1 = "92A" Then
''            SectionTDS1 = "192-Salary-Payment to Government employees other than Indian Government employees"
''            ElseIf SectionTDS1 = "92B" Then
''            SectionTDS1 = "192-Salary-Payment to employees other than Government employees"
''            ElseIf SectionTDS1 = "92C" Then
''            SectionTDS1 = "192-Salary-Payment to Indian Government employees"
'            If SectionTDS1 = "192A" Then
'            SectionTDS1 = "192A-TDS on PF withdrawal"
'            ElseIf SectionTDS1 = "193" Then
'            SectionTDS1 = "193 -Interest on Securities"
'            ElseIf SectionTDS1 = "194" Then
'            SectionTDS1 = "194 -Dividends"
'            ElseIf SectionTDS1 = "94A" Then
'            SectionTDS1 = "194A-Interest other than 'Interest on securities'"
'            ElseIf SectionTDS1 = "4DA" Then
'            SectionTDS1 = "194DA-Payment in respect of life insurance policy"
'            ElseIf SectionTDS1 = "4EE" Then
'            SectionTDS1 = "194EE-Payments in respect of deposits under National Savings"
'            ElseIf SectionTDS1 = "4IB" Then
'            SectionTDS1 = "194IB-Payment of rent by certain individuals or Hindu undivided"
'            ElseIf SectionTDS1 = "94K" Then
'            SectionTDS1 = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
'            ElseIf SectionTDS1 = "4BA1" Then
'            SectionTDS1 = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
'            ElseIf SectionTDS1 = "4BA2" Then
'            SectionTDS1 = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
'            ElseIf SectionTDS1 = "LBB" Then
'            SectionTDS1 = "194LBB-Income in respect of units of investment fund"
'            ElseIf SectionTDS1 = "94P" Then
'            SectionTDS1 = "194P-Deduction of tax in case of specified senior citizen"
'            End If
'            Sheet2.Cells(rowcount, SectionTDSColNo1).Value = SectionTDS1
'        End If
'    '--end--
'End AY 2025-26 Commeted as V0.4 on 12-02-2025
            'TAX DETAILS-E3
            'Sheet2.Cells(rowcount, FYColNo1).Value = Node1("DeductedYr")
            
            'TAX DETAILS-C3 2024-25 Bindu
            Sheet2.Cells(rowcount, FYColNo1).Value = Trim("2023-24")
            
            'Newly added by Chetan C M
            If Sheet2.Cells(rowcount, FYColNo1).Locked = False Then
        
                Sheet2.Cells(rowcount, FYColNo1).Value = Node1("DeductedYr")
    
                Dim DeductedYr_2 As Variant
                DeductedYr_2 = UCase(Node1("DeductedYr"))
                
                ' New Schema updated by Konda as AY-2026-27 on 26-12-2025
                If DeductedYr_2 = "2025" Then
                Sheet2.Cells(rowcount, FYColNo1).Value = "2025-26"
         
                ElseIf DeductedYr_2 = "2024" Then
                Sheet2.Cells(rowcount, FYColNo1).Value = "2024-25"
        '------------------
                
                ElseIf DeductedYr_2 = "2023" Then
                Sheet2.Cells(rowcount, FYColNo1).Value = "2023-24"
                
                ElseIf DeductedYr_2 = "2022" Then
                Sheet2.Cells(rowcount, FYColNo1).Value = "2022-23"
                
                ElseIf DeductedYr_2 = "2021" Then
                Sheet2.Cells(rowcount, FYColNo1).Value = "2021-22"
                
                ElseIf DeductedYr_2 = "2020" Then
                Sheet2.Cells(rowcount, FYColNo1).Value = "2020-21"
                
                ElseIf DeductedYr_2 = "2019" Then
                Sheet2.Cells(rowcount, FYColNo1).Value = "2019-20"
                
                ElseIf DeductedYr_2 = "2018" Then
                Sheet2.Cells(rowcount, FYColNo1).Value = "2018-19"
                
                ElseIf DeductedYr_2 = "2017" Then
                Sheet2.Cells(rowcount, FYColNo1).Value = "2017-18"
                
                Else
                 Sheet2.Cells(rowcount, FYColNo1).Value = "(Select)"
                End If
            End If
            'addition end
            
            Sheet2.Cells(rowcount, BroughtFwdTDSAmt1).Value = Node1("TDSDeducted")
            Sheet2.Cells(rowcount, TaxColNo1).Value = Node1("TDSClaimed")
        cnt = cnt + 1
    Next Node1
 End If
    RecTDS11 = cnt
  
End Function
Function ImportTaxPayments(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim BSRCodeColNo, DateDepColNo, SrlNoChallanColNo, AmtColNo As Long
Dim TotalXMLRow, RecTDS1, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim strDate As String
Dim YYYY, MM, DD As String

Set jsonObject = ParseJson(jsonText)

    Set Nodelist = jsonObject("ITR")("ITR1")("TaxPayments")("TaxPayment")

    BSRCodeColNo = Sheet2.Range("TaxP.BSRCode").Column
    DateDepColNo = Sheet2.Range("TaxP.DateDep").Column
    SrlNoChallanColNo = Sheet2.Range("TaxP.SrlNoOfChaln").Column
    AmtColNo = Sheet2.Range("TaxP.Amt").Column
    
    TotalExRow = Range("TaxP.BSRCode").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
    Application.EnableEvents = False
    
        Sheet2.Range("TaxP.BSRCode").ClearContents
        Sheet2.Range("TaxP.DateDep").ClearContents
        Sheet2.Range("TaxP.SrlNoOfChaln").ClearContents
        Sheet2.Range("TaxP.Amt").ClearContents
        
    Application.EnableEvents = True
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_IT (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet2.Range("TaxP.BSRCode").name)
    rowcount = rowcount - 1
    cnt = 0
    
    If TotalXMLRow > 0 Then
        For Each node In Nodelist
            rowcount = rowcount + 1
                Sheet2.Cells(rowcount, BSRCodeColNo).Value = node("BSRCode")
                strDate = node("DateDep")
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
          
                Sheet2.Cells(rowcount, DateDepColNo).Value = strDate
                Sheet2.Cells(rowcount, SrlNoChallanColNo).Value = node("SrlNoOfChaln")
                Sheet2.Cells(rowcount, AmtColNo).Value = node("Amt")
            cnt = cnt + 1
        Next node
         RecTDS1 = cnt
    End If
   

End Function
Function ImportScheduleTCS(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim TANColNo, DEDNameColNo, TaxColNo, ClaimColNo, AmntClaimedBySpouseTCS, AmtTaxCollected, CollectedYr As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTCS, rowcount, cnt As Long

Set jsonObject = ParseJson(jsonText)

    Set Nodelist = jsonObject("ITR")("ITR1")("ScheduleTCS")("TCS")
    
    TANColNo = Sheet11.Range("TCS.TAN").Column
    DEDNameColNo = Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").Column
    AmtTaxCollected = Sheet11.Range("TCS.AmountCollected").Column
    CollectedYr = Sheet11.Range("TCS.CollectionYear").Column
    TaxColNo = Sheet11.Range("TCS.TotalTCS").Column
    ClaimColNo = Sheet11.Range("TCS.AmtTCSClaimedThisYear").Column
    
    TotalExRow = Range("TCS.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet11.Range("TCS.TAN").ClearContents
    Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").ClearContents
    Sheet11.Range("TCS.AmountCollected").ClearContents
    Sheet11.Range("TCS.CollectionYear").ClearContents
    Sheet11.Range("TCS.TotalTCS").ClearContents
    Sheet11.Range("TCS.AmtTCSClaimedThisYear").ClearContents
    
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_TCS (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet11.Range("TCS.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
If TotalXMLRow > 0 Then
    For Each node In Nodelist
        rowcount = rowcount + 1
            Sheet11.Cells(rowcount, TANColNo).Value = node("EmployerOrDeductorOrCollectDetl")("TAN")
            Sheet11.Cells(rowcount, DEDNameColNo).Value = node("EmployerOrDeductorOrCollectDetl")("EmployerOrDeductorOrCollecterName")
            Sheet11.Cells(rowcount, AmtTaxCollected).Value = node("AmtTaxCollected")
            
            'TAX DETAILS-E4
            'Sheet11.Cells(rowcount, CollectedYr).Value = Node("CollectedYr")
            'TAX DETAILS-C4 2024-25 Bindu
            Sheet11.Cells(rowcount, CollectedYr).Value = Trim("2023-24")
            
            'Newly added by Chetan C M
            If Sheet11.Cells(rowcount, CollectedYr).Locked = False Then
        
            Sheet11.Cells(rowcount, CollectedYr).Value = node("CollectedYr")
            
                Dim CollectedYr_1 As Variant
                CollectedYr_1 = UCase(node("CollectedYr"))
                
                'New Schema updated by Konda as AY-2026-27 on 26-12-2025
                If CollectedYr_1 = "2025" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2025-26"
      
                ElseIf CollectedYr_1 = "2024" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2024-25"
     '----------------------
                ElseIf CollectedYr_1 = "2023" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2023-24"
                
                ElseIf CollectedYr_1 = "2022" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2022-23"
                
                ElseIf CollectedYr_1 = "2021" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2021-22"
                
                ElseIf CollectedYr_1 = "2020" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2020-21"
                
                ElseIf CollectedYr_1 = "2019" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2019-20"
                
                ElseIf CollectedYr_1 = "2018" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2018-19"
                
                ElseIf CollectedYr_1 = "2017" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2017-18"
                
                ElseIf CollectedYr_1 = "2016" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2016-17"
                
                ElseIf CollectedYr_1 = "2015" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2015-16"
                
                ElseIf CollectedYr_1 = "2014" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2014-15"
                
                ElseIf CollectedYr_1 = "2013" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2013-14"
                
                ElseIf CollectedYr_1 = "2012" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2012-13"
                
                ElseIf CollectedYr_1 = "2011" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2011-12"
                
                ElseIf CollectedYr_1 = "2010" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2010-11"
                
                ElseIf CollectedYr_1 = "2009" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2009-10"
                
                ElseIf CollectedYr_1 = "2008" Then
                Sheet11.Cells(rowcount, CollectedYr).Value = "2008-09"
                Else
                 Sheet11.Cells(rowcount, CollectedYr).Value = "(Select)"
                
                End If
            End If
            
        'addition end
            
            Sheet11.Cells(rowcount, TaxColNo).Value = node("TotalTCS")
            Sheet11.Cells(rowcount, ClaimColNo).Value = node("AmtTCSClaimedThisYear")
        cnt = cnt + 1
    Next node
End If
    RecTCS = cnt
    
End Function
Function ImportTaxPaid(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim item As Object

Set jsonObject = ParseJson(jsonText)

'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.AdvanceTax").Value) = "" Then
        Sheet3.Range("IncD.AdvanceTax").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("AdvanceTax")
    End If
'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.TDS").Value) = "" Then
        Sheet3.Range("IncD.TDS").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("TDS")
    End If
'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.TCS").Value) = "" Then
        Sheet3.Range("IncD.TCS").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("TCS")
    End If
'Sheet1.Unprotect Password:=getmsgstate
    If Trim(Sheet3.Range("IncD.SelfAssessmentTax").Value) = "" Then
        Sheet3.Range("IncD.SelfAssessmentTax").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("SelfAssessmentTax")
    End If
'Sheet1.Unprotect Password:=getmsgstate
'    If Trim(Sheet3.Range("IncD.TotalTaxesPaid").Value) = "" Then
'        Sheet3.Range("IncD.TotalTaxesPaid").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("TotalTaxesPaid")
'    End If
'Sheet1.Unprotect Password:=getmsgstate
'    If Trim(Sheet3.Range("IncD.BalTaxPayable").Value) = "" Then
'        Sheet3.Range("IncD.BalTaxPayable").Value = jsonObject("ITR")("ITR1")("TaxPaid")("BalTaxPayable")
'    End If
End Function

Function ImportRefund(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim IFSC, BankName, ACCNO, CheckBox As Variant

'INC_C43
Dim BankAccType As Variant
Dim BankAccType_Value As String

Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long

Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR1")("Refund")("BankAccountDtls")("AddtnlBankDetails")

    IFSC = Range("SchBA.IFSC").Column
    BankName = Range("SchBA.BankName").Column
    ACCNO = Range("SchBA.AcntNo").Column
    'Konda----------------------Uncomented--AY_2025-26
    CheckBox = Range("tempxml").Column
    'End------------------
    'INC_C43
    BankAccType = Range("SchBA.AcntType").Column
    
    TotalExRow = Range("SchBA.IFSC").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Range("SchBA.IFSC").ClearContents
        Range("SchBA.BankName").ClearContents
        Range("SchBA.AcntNo").ClearContents
   'Konda----------------------Uncomented--AY_2025-26
        Range("SchBA.CheckBox").ClearContents
        Range("tempxml").ClearContents
  'End------------------
        Range("BankAccType").ClearContents
        
    End If
    
    If (TotalDiffRow > 0) Then
    Dim idi As Long
    idi = 0
    For idi = 1 To TotalDiffRow
     AddDiffRows_BANK (1)
     Next
    End If
    
    rowcount = getRowNo(Sheet3.Range("SchBA.IFSC").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
            Sheet3.Cells(rowcount, IFSC).Value = node("IFSCCode")
            Sheet3.Cells(rowcount, BankName).Value = node("BankName")
            Sheet3.Cells(rowcount, ACCNO).Value = node("BankAccountNo")
            
'
'            'INC_C43 2024-25 Bindu
'            BankAccType_Value = Node("AccountType")
'
'            If BankAccType_Value = "SAV" Then
'               BankAccType_Value = "Savings Account"
'            ElseIf BankAccType_Value = "CUR" Then
'                BankAccType_Value = "Current Account"
'            ElseIf BankAccType_Value = "CC" Then
'                BankAccType_Value = "Cash Credit Account"
'            ElseIf BankAccType_Value = "OD" Then
'                BankAccType_Value = "Over draft account"
'            ElseIf BankAccType_Value = "NR" Then
'                BankAccType_Value = "Non Resident Account"
'             ElseIf BankAccType_Value = "OTH" Then
'                BankAccType_Value = "Other"
'            End If
'            Sheet3.Cells(rowcount, BankAccType).Value = BankAccType_Value
'
            'INC_C43 2024-25  Malli
            BankAccType_Value = node("AccountType")
            
            If BankAccType_Value = "SB" Then
               BankAccType_Value = "Savings Account"
            ElseIf BankAccType_Value = "CA" Then
                BankAccType_Value = "Current Account"
            ElseIf BankAccType_Value = "CC" Then
                BankAccType_Value = "Cash Credit Account"
            ElseIf BankAccType_Value = "OD" Then
                BankAccType_Value = "Over draft account"
            ElseIf BankAccType_Value = "NRO" Then
                BankAccType_Value = "Non Resident Account"
             ElseIf BankAccType_Value = "OTH" Then
                BankAccType_Value = "Other"
            End If
            Sheet3.Cells(rowcount, BankAccType).Value = BankAccType_Value
'Malli-----------------------------------
'AY_2024_25 Change
'Konda------------------Uncomented--AY_2025-26
            Sheet3.Cells(rowcount, CheckBox).Value = node("UseForRefund")
            
            If CheckBox = True Then
            CheckBox = "true"
            ElseIf CheckBox = False Then
            CheckBox = "false"
            ElseIf CheckBox = "" Then
            CheckBox = "false"
            End If

        LinkCheckBoxes
'End-------------------
'Malli-----------------------
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet3.Range("IncD.RefundDue").Value = jsonObject("ITR")("ITR1")("Refund")("RefundDue")

End Function
Function ImportVerification(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim Capacity As Variant

Set jsonObject = ParseJson(jsonText)

Capacity = jsonObject("ITR")("ITR1")("Verification")("Capacity")
If Capacity = "S" Then
Capacity = "Self"
ElseIf Capacity = "R" Then
Capacity = "Representative"
Else
Capacity = ""
End If
Sheet3.Range("Ver.capacity").Value = Capacity

Sheet3.Range("Ver.Place").Value = jsonObject("ITR")("ITR1")("Verification")("Place")
Sheet3.Range("Ver.AssesseeVerName").Value = jsonObject("ITR")("ITR1")("Verification")("Declaration")("AssesseeVerName")
Sheet3.Range("Ver.FatherName").Value = jsonObject("ITR")("ITR1")("Verification")("Declaration")("FatherName")
Sheet3.Range("Ver.PAN").Value = jsonObject("ITR")("ITR1")("Verification")("Declaration")("AssesseeVerPAN")

End Function
Function ImportTaxReturnPreparer(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object

Set jsonObject = ParseJson(jsonText)

    Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value = jsonObject("ITR")("ITR1")("TaxReturnPreparer")("IdentificationNoOfTRP")
    Sheet3.Range("Sheet2.NameOfTRP").Value = jsonObject("ITR")("ITR1")("TaxReturnPreparer")("NameOfTRP")
    Sheet3.Range("Sheet2.ReImbFrmGov").Value = jsonObject("ITR")("ITR1")("TaxReturnPreparer")("ReImbFrmGov")

End Function
'Function ImportSchedule80D(jsonText As String)
' On Error Resume Next
'Dim jsonObject, jsonDictionary As Object
'Dim node, Nodelist As Object
'
'Set jsonObject = ParseJson(jsonText)
'Set node = jsonObject("ITR")("ITR1")("Schedule80D")("Sec80DSelfFamSrCtznHealth")
'
'            If Trim(node("SeniorCitizenFlag")) <> "" Then
'                If node("SeniorCitizenFlag") = "Y" Then
'                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Yes"
'                ElseIf node("SeniorCitizenFlag") = "N" Then
'                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "No"
'                ElseIf node("SeniorCitizenFlag") = "S" Then
'                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Not claiming for Self/ Family"
'                End If
'            End If
'
'            If Trim(node("HealthInsPremSlfFam")) <> "" And Sheet9.Range("Health_Insurance_80D").Locked = False Then
'                Sheet9.Range("Health_Insurance_80D").Value = node("HealthInsPremSlfFam")
'            End If
'            If Trim(node("PrevHlthChckUpSlfFam")) <> "" And Sheet9.Range("Preventive_Health_80D").locekd = False Then
'                Sheet9.Range("Preventive_Health_80D").Value = node("PrevHlthChckUpSlfFam")
'            End If
'            If Trim(node("HlthInsPremSlfFamSrCtzn")) <> "" And Sheet9.Range("Health_InsuranceSC_80D").Locked = False Then
'                Sheet9.Range("Health_InsuranceSC_80D").Value = node("HlthInsPremSlfFamSrCtzn")
'            End If
'            If Trim(node("PrevHlthChckUpSlfFamSrCtzn")) <> "" And Sheet9.Range("Preventive_Health_SC_80D").Locked = False Then
'                Sheet9.Range("Preventive_Health_SC_80D").Value = node("PrevHlthChckUpSlfFamSrCtzn")
'            End If
'            If Trim(node("MedicalExpSlfFamSrCtzn")) <> "" And Sheet9.Range("Medical_Expenditure_SC_80D").Locked = False Then
'                Sheet9.Range("Medical_Expenditure_SC_80D").Value = node("MedicalExpSlfFamSrCtzn")
'            End If
'
'            If Trim(node("ParentsSeniorCitizenFlag")) <> "" Then
'                If node("ParentsSeniorCitizenFlag") = "Y" Then
'                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Yes"
'                ElseIf node("ParentsSeniorCitizenFlag") = "N" Then
'                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "No"
'                ElseIf node("ParentsSeniorCitizenFlag") = "P" Then
'                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Not claiming for Parents"
'                End If
'            End If
'
'            If Trim(node("HlthInsPremParents")) <> "" And Sheet9.Range("Health_Insurance2_80D").Locked = False Then
'                Sheet9.Range("Health_Insurance2_80D").Value = node("HlthInsPremParents")
'            End If
'            If Trim(node("PrevHlthChckUpParents")) <> "" And Sheet9.Range("Preventive_Health2_80D") = False Then
'                Sheet9.Range("Preventive_Health2_80D").Value = node("PrevHlthChckUpParents")
'            End If
'            If Trim(node("HlthInsPremParentsSrCtzn")) <> "" And Sheet9.Range("Health_Insurance3_80D").Locked = False Then
'                Sheet9.Range("Health_Insurance3_80D").Value = node("HlthInsPremParentsSrCtzn")
'            End If
'            If Trim(node("PrevHlthChckUpParentsSrCtzn")) <> "" And Sheet9.Range("Preventive_Health3_80D").Locked = False Then
'                Sheet9.Range("Preventive_Health3_80D").Value = node("PrevHlthChckUpParentsSrCtzn")
'            End If
'            If Trim(node("MedicalExpParentsSrCtzn")) <> "" And Sheet9.Range("Medical_Expenditure2_80D").Locked = False Then
'                Sheet9.Range("Medical_Expenditure2_80D").Value = node("MedicalExpParentsSrCtzn")
'            End If
'
'End Function
Function ImportSchedule80D(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object

Set jsonObject = ParseJson(jsonText)
Set node = jsonObject("ITR")("ITR1")("Schedule80D")("Sec80DSelfFamSrCtznHealth")
    
            If Trim(node("SeniorCitizenFlag")) <> "" Then
                If node("SeniorCitizenFlag") = "Y" Then
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Yes"
                ElseIf node("SeniorCitizenFlag") = "N" Then
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "No"
                ElseIf node("SeniorCitizenFlag") = "S" Then
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Not claiming for Self/ Family"
                End If
            End If
                                            
               'Malli--------17/04/2025
               'This field autopopulate AY_2025-26
'            If Trim(node("HealthInsPremSlfFam")) <> "" And Sheet9.Range("Health_Insurance_80D").Locked = False Then
'                Sheet9.Range("Health_Insurance_80D").Value = node("HealthInsPremSlfFam")
'            End If
            '_________________________________
             
            Dim TotalExRow, TotalXMLRow, TotalDiffRow, rowcount, cnt As Long
                Dim node_80DbSFHI, Nodelist_80DbSFHI
                'Set Nodelist_80Db = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("InsDtls80Db")("Sch80DInsDtls")
                Set Nodelist_80DbSFHI = node("Sec80DSelfFamHIDtls")("Sch80DInsDtls")
                
                Dim InsurerName_80DbSFHI, PolicyNo_80DbSFHI, ReceiptNo_80DbSFHI, HealthInsAmt_80DbSFHI, RecInt80DbSFHI
                
                InsurerName_80DbSFHI = Sheet9.Range("NameInsurerA1.80D").Column
                PolicyNo_80DbSFHI = Sheet9.Range("PolicyNumA1.80D").Column
'                ReceiptNo_80DbSFHI = Sheet9.Range("ReceiptNumA1.80D").Column
                HealthInsAmt_80DbSFHI = Sheet9.Range("AmtA1.80D").Column
                
                 TotalExRow = Sheet9.Range("NameInsurerA1.80D").Rows.count
                 
                  TotalXMLRow = Nodelist_80DbSFHI.count
                  
                  TotalDiffRow = TotalXMLRow - TotalExRow
                  
                    If (TotalXMLRow > 0) Then
                    
                    If Sheet9.Range("NameInsurerA1.80D").Locked = False Then
                       Sheet9.Range("NameInsurerA1.80D").ClearContents
                    End If
                    
                    If Sheet9.Range("PolicyNumA1.80D").Locked = False Then
                       Sheet9.Range("PolicyNumA1.80D").ClearContents
                    End If
                    'Ankita_06/05/2025_Commented as per DESheet_v0.7

'                    If Sheet9.Range("ReceiptNumA1.80D").Locked = False Then
'                       Sheet9.Range("ReceiptNumA1.80D").ClearContents
'                    End If
                    
                    If Sheet9.Range("AmtA1.80D").Locked = False Then
                       Sheet9.Range("AmtA1.80D").ClearContents
                    End If
               End If
               
                    If (TotalDiffRow > 0) Then
                         AddDiffRows_Sec80DSelfFamHIDtls (TotalDiffRow)
                     End If
                     
                    rowcount = getRowNo(Sheet9.Range("NameInsurerA1.80D").name)
                    rowcount = rowcount - 1
                    cnt = 0
                     
                     For Each node_80DbSFHI In Nodelist_80DbSFHI
                         rowcount = rowcount + 1
                         
                         If Sheet9.Cells(rowcount, InsurerName_80DbSFHI).Locked = False Then
                             Sheet9.Cells(rowcount, InsurerName_80DbSFHI).Value = node_80DbSFHI("InsurerName")
                         End If
                         
                         If Sheet9.Cells(rowcount, PolicyNo_80DbSFHI).Locked = False Then
                             Sheet9.Cells(rowcount, PolicyNo_80DbSFHI).Value = node_80DbSFHI("PolicyNo")
                         End If
                         
'                         'Konda---------08/05/2025
'                         If Sheet9.Cells(rowcount, ReceiptNo_80DbSFHI).Locked = False Then
'                             Sheet9.Cells(rowcount, ReceiptNo_80DbSFHI).Value = node_80DbSFHI("ReceiptNo")
'                         End If
'                         '--------------------------
                         
                         If Sheet9.Cells(rowcount, HealthInsAmt_80DbSFHI).Locked = False Then
                             Sheet9.Cells(rowcount, HealthInsAmt_80DbSFHI).Value = node_80DbSFHI("HealthInsAmt")
                         End If
                         
                         
                        cnt = cnt + 1
                     Next node_80DbSFHI
                     
                 RecInt80DbSFHI = cnt
                     
                
                
                '-------------------------
            
            '---------------------------------
            If Trim(node("PrevHlthChckUpSlfFam")) <> "" And Sheet9.Range("Preventive_Health_80D").locekd = False Then
                Sheet9.Range("Preventive_Health_80D").Value = node("PrevHlthChckUpSlfFam")
            End If
            '-------------------
            'This field autopopulate AY_2025-26
'            If Trim(node("HlthInsPremSlfFamSrCtzn")) <> "" And Sheet9.Range("Health_InsuranceSC_80D").Locked = False Then
'                Sheet9.Range("Health_InsuranceSC_80D").Value = node("HlthInsPremSlfFamSrCtzn")
'            End If
            '__________________________________
           ' Dim TotalExRow, TotalXMLRow, TotalDiffRow, rowcount, cnt As Long
                Dim node_80DbSFsrCHI, Nodelist_80DbSFsrCHI
                'Set Nodelist_80Db = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("InsDtls80Db")("Sch80DInsDtls")
                Set Nodelist_80DbSFsrCHI = node("Sec80DSelfFamSrCtznHIDtls")("Sch80DInsDtls")
                
                Dim InsurerName_80DbSFsrCHI, PolicyNo_80DbSFsrCHI, ReceiptNo_80DbSFsrCHI, HealthInsAmt_80DbSFsrCHI, RecInt80DbSFsrCHI
                
                InsurerName_80DbSFsrCHI = Sheet9.Range("NameInsurerB1.80D").Column
                PolicyNo_80DbSFsrCHI = Sheet9.Range("PolicyNumB1.80D").Column
'                ReceiptNo_80DbSFsrCHI = Sheet9.Range("ReceiptNumB1.80D").Column
                HealthInsAmt_80DbSFsrCHI = Sheet9.Range("AmtB1.80D").Column
                
                 TotalExRow = Sheet9.Range("NameInsurerB1.80D").Rows.count
                 
                  TotalXMLRow = Nodelist_80DbSFsrCHI.count
                  
                  TotalDiffRow = TotalXMLRow - TotalExRow
                  
                    If (TotalXMLRow > 0) Then
                    
                    If Sheet9.Range("NameInsurerB1.80D").Locked = False Then
                       Sheet9.Range("NameInsurerB1.80D").ClearContents
                    End If
                    
                    If Sheet9.Range("PolicyNumB1.80D").Locked = False Then
                       Sheet9.Range("PolicyNumB1.80D").ClearContents
                    End If
                    
'                    If Sheet9.Range("ReceiptNumB1.80D").Locked = False Then
'                       Sheet9.Range("ReceiptNumB1.80D").ClearContents
'                    End If
                    
                    If Sheet9.Range("AmtB1.80D").Locked = False Then
                       Sheet9.Range("AmtB1.80D").ClearContents
                    End If
               End If
               
                    If (TotalDiffRow > 0) Then
                         AddDiffRows_Sec80DSelfFamSrCtznHIDtls (TotalDiffRow)
                     End If
                     
                    rowcount = getRowNo(Sheet9.Range("NameInsurerB1.80D").name)
                    rowcount = rowcount - 1
                    cnt = 0
                     
                     For Each node_80DbSFsrCHI In Nodelist_80DbSFsrCHI
                         rowcount = rowcount + 1
                         
                         If Sheet9.Cells(rowcount, InsurerName_80DbSFsrCHI).Locked = False Then
                             Sheet9.Cells(rowcount, InsurerName_80DbSFsrCHI).Value = node_80DbSFsrCHI("InsurerName")
                         End If
                         
                         If Sheet9.Cells(rowcount, PolicyNo_80DbSFsrCHI).Locked = False Then
                             Sheet9.Cells(rowcount, PolicyNo_80DbSFsrCHI).Value = node_80DbSFsrCHI("PolicyNo")
                         End If
                         
'                         'Konda--------08/05/2025
'                         If Sheet9.Cells(rowcount, ReceiptNo_80DbSFsrCHI).Locked = False Then
'                             Sheet9.Cells(rowcount, ReceiptNo_80DbSFsrCHI).Value = node_80DbSFsrCHI("ReceiptNo")
'                         End If
'                         '----------
                         
                         If Sheet9.Cells(rowcount, HealthInsAmt_80DbSFsrCHI).Locked = False Then
                             Sheet9.Cells(rowcount, HealthInsAmt_80DbSFsrCHI).Value = node_80DbSFsrCHI("HealthInsAmt")
                         End If
                         
                         
                        cnt = cnt + 1
                     Next node_80DbSFsrCHI
                     
                 RecInt80DbSFsrCHI = cnt
            
            '----------------------------------
            If Trim(node("PrevHlthChckUpSlfFamSrCtzn")) <> "" And Sheet9.Range("Preventive_Health_SC_80D").Locked = False Then
                Sheet9.Range("Preventive_Health_SC_80D").Value = node("PrevHlthChckUpSlfFamSrCtzn")
            End If
            If Trim(node("MedicalExpSlfFamSrCtzn")) <> "" And Sheet9.Range("Medical_Expenditure_SC_80D").Locked = False Then
                Sheet9.Range("Medical_Expenditure_SC_80D").Value = node("MedicalExpSlfFamSrCtzn")
            End If
            
            If Trim(node("ParentsSeniorCitizenFlag")) <> "" Then
                If node("ParentsSeniorCitizenFlag") = "Y" Then
                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Yes"
                ElseIf node("ParentsSeniorCitizenFlag") = "N" Then
                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "No"
                ElseIf node("ParentsSeniorCitizenFlag") = "P" Then
                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Not claiming for Parents"
                End If
            End If
            '-----------------------------
            'This field autopopulate AY_2025-26
'            If Trim(node("HlthInsPremParents")) <> "" And Sheet9.Range("Health_Insurance2_80D").Locked = False Then
'                Sheet9.Range("Health_Insurance2_80D").Value = node("HlthInsPremParents")
'            End If
            '_________________________________
            
            Dim node_80DbPHI, Nodelist_80DbPHI
                'Set Nodelist_80Db = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("InsDtls80Db")("Sch80DInsDtls")
                Set Nodelist_80DbPHI = node("Sec80DParentsHIDtls")("Sch80DInsDtls")
                
                Dim InsurerName_80DbPHI, PolicyNo_80DbPHI, ReceiptNo_80DbPHI, HealthInsAmt_80DbPHI, RecInt80DbPHI
                
                InsurerName_80DbPHI = Sheet9.Range("NameInsurerA2.80D").Column
                PolicyNo_80DbPHI = Sheet9.Range("PolicyNumA2.80D").Column
'                ReceiptNo_80DbPHI = Sheet9.Range("ReceiptNumA2.80D").Column
                HealthInsAmt_80DbPHI = Sheet9.Range("AmtA2.80D").Column
                
                 TotalExRow = Sheet9.Range("NameInsurerA2.80D").Rows.count
                 
                  TotalXMLRow = Nodelist_80DbPHI.count
                  
                  TotalDiffRow = TotalXMLRow - TotalExRow
                  
                    If (TotalXMLRow > 0) Then
                    
                    If Sheet9.Range("NameInsurerA2.80D").Locked = False Then
                       Sheet9.Range("NameInsurerA2.80D").ClearContents
                    End If
                    
                    If Sheet9.Range("PolicyNumA2.80D").Locked = False Then
                       Sheet9.Range("PolicyNumA2.80D").ClearContents
                    End If
                    
'                    If Sheet9.Range("ReceiptNumA2.80D").Locked = False Then
'                       Sheet9.Range("ReceiptNumA2.80D").ClearContents
'                    End If
                    
                    If Sheet9.Range("AmtA2.80D").Locked = False Then
                       Sheet9.Range("AmtA2.80D").ClearContents
                    End If
               End If
               
                    If (TotalDiffRow > 0) Then
                         AddDiffRows_Sec80DParentsHIDtls (TotalDiffRow)
                     End If
                     
                    rowcount = getRowNo(Sheet9.Range("NameInsurerA2.80D").name)
                    rowcount = rowcount - 1
                    cnt = 0
                     
                     For Each node_80DbPHI In Nodelist_80DbPHI
                         rowcount = rowcount + 1
                         
                         If Sheet9.Cells(rowcount, InsurerName_80DbPHI).Locked = False Then
                             Sheet9.Cells(rowcount, InsurerName_80DbPHI).Value = node_80DbPHI("InsurerName")
                         End If
                         
                         If Sheet9.Cells(rowcount, PolicyNo_80DbPHI).Locked = False Then
                             Sheet9.Cells(rowcount, PolicyNo_80DbPHI).Value = node_80DbPHI("PolicyNo")
                         End If
                         
'                         'Konda--------08/05/2025
'                         If Sheet9.Cells(rowcount, ReceiptNo_80DbPHI).Locked = False Then
'                             Sheet9.Cells(rowcount, ReceiptNo_80DbPHI).Value = node_80DbPHI("ReceiptNo")
'                         End If
'                         '------------
                         
                         If Sheet9.Cells(rowcount, HealthInsAmt_80DbPHI).Locked = False Then
                             Sheet9.Cells(rowcount, HealthInsAmt_80DbPHI).Value = node_80DbPHI("HealthInsAmt")
                         End If
                         
                         
                        cnt = cnt + 1
                     Next node_80DbPHI
                     
                 RecInt80DbPHI = cnt
            
            '----------------------------------
            '---------------------------------
            If Trim(node("PrevHlthChckUpParents")) <> "" And Sheet9.Range("Preventive_Health2_80D") = False Then
                Sheet9.Range("Preventive_Health2_80D").Value = node("PrevHlthChckUpParents")
            End If
            '-------------------------------------
            'This field autopopulate AY_2025-26
'            If Trim(node("HlthInsPremParentsSrCtzn")) <> "" And Sheet9.Range("Health_Insurance3_80D").Locked = False Then
'                Sheet9.Range("Health_Insurance3_80D").Value = node("HlthInsPremParentsSrCtzn")
'            End If
            '_______________________________________
             Dim node_80DbPsrCHI, Nodelist_80DbPsrCHI
                'Set Nodelist_80Db = jsonObject("ITR")("ITR4")("Schedule80D")("Sec80DSelfFamSrCtznHealth")("InsDtls80Db")("Sch80DInsDtls")
                Set Nodelist_80DbPsrCHI = node("Sec80DParentsSrCtznHIDtls")("Sch80DInsDtls")
                
                Dim InsurerName_80DbPsrCHI, PolicyNo_80DbPsrCHI, ReceiptNo_80DbPsrCHI, HealthInsAmt_80DbPsrCHI, RecInt80DbPsrCHI
                
                InsurerName_80DbPsrCHI = Sheet9.Range("NameInsurerB2.80D").Column
                PolicyNo_80DbPsrCHI = Sheet9.Range("PolicyNumB2.80D").Column
'                ReceiptNo_80DbPsrCHI = Sheet9.Range("ReceiptNumB2.80D").Column
                HealthInsAmt_80DbPsrCHI = Sheet9.Range("AmtB2.80D").Column
                
                 TotalExRow = Sheet9.Range("NameInsurerB2.80D").Rows.count
                 
                  TotalXMLRow = Nodelist_80DbPsrCHI.count
                  
                  TotalDiffRow = TotalXMLRow - TotalExRow
                  
                    If (TotalXMLRow > 0) Then
                    
                    If Sheet9.Range("NameInsurerB2.80D").Locked = False Then
                       Sheet9.Range("NameInsurerB2.80D").ClearContents
                    End If
                    
                    If Sheet9.Range("PolicyNumB2.80D").Locked = False Then
                       Sheet9.Range("PolicyNumB2.80D").ClearContents
                    End If
                    
'                    If Sheet9.Range("ReceiptNumB2.80D").Locked = False Then
'                       Sheet9.Range("ReceiptNumB2.80D").ClearContents
'                    End If
                    
                    If Sheet9.Range("AmtB2.80D").Locked = False Then
                       Sheet9.Range("AmtB2.80D").ClearContents
                    End If
               End If
               
                    If (TotalDiffRow > 0) Then
                         AddDiffRows_Sec80DParentsSrCtznHIDtls (TotalDiffRow)
                     End If
                     
                    rowcount = getRowNo(Sheet9.Range("NameInsurerB2.80D").name)
                    rowcount = rowcount - 1
                    cnt = 0
                     
                     For Each node_80DbPsrCHI In Nodelist_80DbPsrCHI
                         rowcount = rowcount + 1
                         
                         If Sheet9.Cells(rowcount, InsurerName_80DbPsrCHI).Locked = False Then
                             Sheet9.Cells(rowcount, InsurerName_80DbPsrCHI).Value = node_80DbPsrCHI("InsurerName")
                         End If
                         
                         If Sheet9.Cells(rowcount, PolicyNo_80DbPsrCHI).Locked = False Then
                             Sheet9.Cells(rowcount, PolicyNo_80DbPsrCHI).Value = node_80DbPsrCHI("PolicyNo")
                         End If
                         
'                         'Konda--------08/05/2025
'                         If Sheet9.Cells(rowcount, ReceiptNo_80DbPsrCHI).Locked = False Then
'                             Sheet9.Cells(rowcount, ReceiptNo_80DbPsrCHI).Value = node_80DbPsrCHI("ReceiptNo")
'                         End If
'                         '------------------
                         
                         If Sheet9.Cells(rowcount, HealthInsAmt_80DbPsrCHI).Locked = False Then
                             Sheet9.Cells(rowcount, HealthInsAmt_80DbPsrCHI).Value = node_80DbPsrCHI("HealthInsAmt")
                         End If
                         
                         
                        cnt = cnt + 1
                     Next node_80DbPsrCHI
                     
                 RecInt80DbPsrCHI = cnt
            '---------------------------------------
            If Trim(node("PrevHlthChckUpParentsSrCtzn")) <> "" And Sheet9.Range("Preventive_Health3_80D").Locked = False Then
                Sheet9.Range("Preventive_Health3_80D").Value = node("PrevHlthChckUpParentsSrCtzn")
            End If
            If Trim(node("MedicalExpParentsSrCtzn")) <> "" And Sheet9.Range("Medical_Expenditure2_80D").Locked = False Then
                Sheet9.Range("Medical_Expenditure2_80D").Value = node("MedicalExpParentsSrCtzn")
            End If

End Function

Function ImportSchedule80G100NoAppr(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long

'Konda updated 10-03-2026--V0.5
Dim TransactionRefNum_80GA, IFSCCode_80GA
'==============

Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80G")("Don100Percent")("DoneeWithPan")
    
    NameColNo = Sheet4.Range("Per10080G.DoneeName").Column
    AddressColNo = Sheet4.Range("Per10080G.AddrDetail").Column
    CityColNo = Sheet4.Range("Per10080G.CityOrTownOrDistrict").Column
    StateCodeColNo = Sheet4.Range("Per10080G.StateCode").Column
    PincodeColNo = Sheet4.Range("Per10080G.PinCode").Column
    PanofDoneeColNo = Sheet4.Range("Per10080G.DoneePAN").Column
    AmountColNo = Sheet4.Range("Per10080G.DonationAmt").Column
    DonationColNo = Sheet4.Range("Per10080G.DonationAmtOther").Column
'Konda updated 10-03-2026--V0.5
    TransactionRefNum_80GA = Sheet4.Range("Per10080G.Traref").Column
    IFSCCode_80GA = Sheet4.Range("Per10080G.IFSC").Column
'============================================

    TotalExRow = Range("Per10080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet4.Range("Per10080G.DoneeName").ClearContents
        Sheet4.Range("Per10080G.AddrDetail").ClearContents
        Sheet4.Range("Per10080G.CityOrTownOrDistrict").ClearContents
        Sheet4.Range("Per10080G.StateCode").ClearContents
        Sheet4.Range("Per10080G.PinCode").ClearContents
        Sheet4.Range("Per10080G.DoneePAN").ClearContents
        Sheet4.Range("Per10080G.DonationAmt").ClearContents
        Sheet4.Range("Per10080G.DonationAmtOther").ClearContents
'Konda updated 10-03-2026--V0.5
        Sheet4.Range("Per10080G.Traref").ClearContents
        Sheet4.Range("Per10080G.IFSC").ClearContents
    '====================
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_A (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet4.Range("Per10080G.DoneeName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
            Sheet4.Cells(rowcount, NameColNo).Value = node("DoneeWithPanName")
            Sheet4.Cells(rowcount, AddressColNo).Value = node("AddressDetail")("AddrDetail")
            Sheet4.Cells(rowcount, CityColNo).Value = node("AddressDetail")("CityOrTownOrDistrict")
            Dim iState As Variant
            iState = node("AddressDetail")("StateCode")
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            
            
            If iState = "99" Then
            iState = ""
            End If
            
            Sheet4.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet4.Cells(rowcount, PincodeColNo).Value = node("AddressDetail")("PinCode")
            Sheet4.Cells(rowcount, PanofDoneeColNo).Value = node("DoneePAN")
            Sheet4.Cells(rowcount, AmountColNo).Value = node("DonationAmtCash")
            Sheet4.Cells(rowcount, DonationColNo).Value = node("DonationAmtOtherMode")
    'Konda updated on 10-03-2026--v0.5
            Sheet4.Cells(rowcount, TransactionRefNum_80GA).Value = node("TransactionRefNum")
            Sheet4.Cells(rowcount, IFSCCode_80GA).Value = node("IFSCCode")
    '============================
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt

End Function
Function ImportSchedule80G50NoAppr(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, cnt, rowcount As Long

'Konda updated 10-03-2026--V0.5
Dim TransactionRefNum_80GB, IFSCCode_80GB
'==============


Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80G")("Don50PercentNoApprReqd")("DoneeWithPan")
    
    NameColNo = Sheet4.Range("PerNO5080G.DoneeName").Column
    AddressColNo = Sheet4.Range("PerNO5080G.AddrDetail").Column
    CityColNo = Sheet4.Range("PerNO5080G.CityOrTownOrDistrict").Column
    StateCodeColNo = Sheet4.Range("PerNO5080G.StateCode").Column
    PincodeColNo = Sheet4.Range("PerNO5080G.PinCode").Column
    PanofDoneeColNo = Sheet4.Range("PerNO5080G.DoneePAN").Column
    AmountColNo = Sheet4.Range("PerNO5080G.DonationAmt").Column
    DonationColNo = Sheet4.Range("PerNO5080G.DonationAmtOther").Column
'Konda updated 10-03-2026--V0.5
    TransactionRefNum_80GB = Sheet4.Range("PerNO5080G.Traref").Column
    IFSCCode_80GB = Sheet4.Range("PerNO5080G.IFSC").Column
'============================================

    TotalExRow = Range("PerNO5080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet4.Range("PerNO5080G.DoneeName").ClearContents
        Sheet4.Range("PerNO5080G.AddrDetail").ClearContents
        Sheet4.Range("PerNO5080G.CityOrTownOrDistrict").ClearContents
        Sheet4.Range("PerNO5080G.StateCode").ClearContents
        Sheet4.Range("PerNO5080G.PinCode").ClearContents
        Sheet4.Range("PerNO5080G.DoneePAN").ClearContents
        Sheet4.Range("PerNO5080G.DonationAmt").ClearContents
        Sheet4.Range("PerNO5080G.DonationAmtOther").ClearContents
    'Konda updated 10-03-2026--V0.5
        Sheet4.Range("PerNO5080G.Traref").ClearContents
        Sheet4.Range("PerNO5080G.IFSC").ClearContents
'============================================
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_B (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet4.Range("PerNO5080G.DoneeName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
            Sheet4.Cells(rowcount, NameColNo).Value = node("DoneeWithPanName")
            Sheet4.Cells(rowcount, AddressColNo).Value = node("AddressDetail")("AddrDetail")
            Sheet4.Cells(rowcount, CityColNo).Value = node("AddressDetail")("CityOrTownOrDistrict")
            Dim iState As Variant
            iState = UCase(node("AddressDetail")("StateCode"))
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet4.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet4.Cells(rowcount, PincodeColNo).Value = node("AddressDetail")("PinCode")
            Sheet4.Cells(rowcount, PanofDoneeColNo).Value = node("DoneePAN")
            Sheet4.Cells(rowcount, AmountColNo).Value = node("DonationAmtCash")
            Sheet4.Cells(rowcount, DonationColNo).Value = node("DonationAmtOtherMode")
    'Konda updated on 10-03-2026--v0.5
            Sheet4.Cells(rowcount, TransactionRefNum_80GB).Value = node("TransactionRefNum")
            Sheet4.Cells(rowcount, IFSCCode_80GB).Value = node("IFSCCode")
    '============================

        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Function
Function ImportSchedule80G100Appr(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long

'Konda updated 10-03-2026--V0.5
Dim TransactionRefNum_80GC, IFSCCode_80GC
'==============

Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80G")("Don100PercentApprReqd")("DoneeWithPan")
    
    NameColNo = Sheet4.Range("PerYES10080G.DoneeWithPanName").Column
    AddressColNo = Sheet4.Range("PerYES10080G.AddrDetail").Column
    CityColNo = Sheet4.Range("PerYES10080G.CityOrTownOrDistrict").Column
    StateCodeColNo = Sheet4.Range("PerYES10080G.StateCode").Column
    PincodeColNo = Sheet4.Range("PerYES10080G.PinCode").Column
    PanofDoneeColNo = Sheet4.Range("PerYES10080G.DoneePAN").Column
    AmountColNo = Sheet4.Range("PerYES10080G.DonationAmt").Column
    DonationColNo = Sheet4.Range("PerYES10080G.DonationAmtOther").Column
  'Konda updated 10-03-2026--V0.5
    TransactionRefNum_80GC = Sheet4.Range("PerYES10080G.Traref").Column
    IFSCCode_80GC = Sheet4.Range("PerYES10080G.IFSC").Column
'============================================
 
    TotalExRow = Sheet4.Range("PerYES10080G.DoneeWithPanName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet4.Range("PerYES10080G.DoneeWithPanName").ClearContents
        Sheet4.Range("PerYES10080G.AddrDetail").ClearContents
        Sheet4.Range("PerYES10080G.CityOrTownOrDistrict").ClearContents
        Sheet4.Range("PerYES10080G.StateCode").ClearContents
        Sheet4.Range("PerYES10080G.PinCode").ClearContents
        Sheet4.Range("PerYES10080G.DoneePAN").ClearContents
        Sheet4.Range("PerYES10080G.DonationAmt").ClearContents
        Sheet4.Range("PerYES10080G.DonationAmtOther").ClearContents
'Konda updated 10-03-2026--V0.5
        Sheet4.Range("PerYES10080G.Traref").ClearContents
        Sheet4.Range("PerYES10080G.IFSC").ClearContents
'============================================
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_C (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet4.Range("PerYES10080G.DoneeWithPanName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
            Sheet4.Cells(rowcount, NameColNo).Value = node("DoneeWithPanName")
            Sheet4.Cells(rowcount, AddressColNo).Value = node("AddressDetail")("AddrDetail")
            Sheet4.Cells(rowcount, CityColNo).Value = node("AddressDetail")("CityOrTownOrDistrict")
            Dim iState As Variant
            iState = node("AddressDetail")("StateCode")
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            
            If iState = "99" Then
            iState = ""
            End If
            
            Sheet4.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet4.Cells(rowcount, PincodeColNo).Value = node("AddressDetail")("PinCode")
            Sheet4.Cells(rowcount, PanofDoneeColNo).Value = node("DoneePAN")
            Sheet4.Cells(rowcount, AmountColNo).Value = node("DonationAmtCash")
            Sheet4.Cells(rowcount, DonationColNo).Value = node("DonationAmtOtherMode")
        'Konda updated on 10-03-2026--v0.5
            Sheet4.Cells(rowcount, TransactionRefNum_80GC).Value = node("TransactionRefNum")
            Sheet4.Cells(rowcount, IFSCCode_80GC).Value = node("IFSCCode")
    '============================

            
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt

End Function
Function ImportSchedule80G50Appr(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary, init As Object
Dim node, Nodelist As Object
'Change.25.01.2023.102.80GC5
'Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, ArnNbrColNo, AmountColNo, DonationColNo As Variant
'End Change===
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long

'Konda updated 10-03-2026--V0.5
Dim TransactionRefNum_80GD, IFSCCode_80GD
'==============


Set jsonObject = ParseJson(jsonText)
Set init = jsonObject("ITR")("ITR1")("Schedule80G")("Don50PercentApprReqd")
If init.exists("DoneeWithPan") Then
    Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80G")("Don50PercentApprReqd")("DoneeWithPan")
    NameColNo = Sheet4.Range("Per5080G.DoneeWithPanName").Column
    AddressColNo = Sheet4.Range("Per5080G.AddrDetail").Column
    CityColNo = Sheet4.Range("Per5080G.CityOrTownOrDistrict").Column
    StateCodeColNo = Sheet4.Range("Per5080G.StateCode").Column
    PincodeColNo = Sheet4.Range("Per5080G.PinCode").Column
    PanofDoneeColNo = Sheet4.Range("Per5080G.DoneePAN").Column
'Change.25.01.2023.102.80GC4
    ArnNbrColNo = Sheet4.Range("Per5080G.ArnNbr").Column
'ENd Change===
    AmountColNo = Sheet4.Range("Per5080G.DonationAmt").Column
    DonationColNo = Sheet4.Range("Per5080G.DonationAmtOther").Column
    
    'Konda updated 10-03-2026--V0.5
    TransactionRefNum_80GD = Sheet4.Range("Per5080G.Traref").Column
    IFSCCode_80GD = Sheet4.Range("Per5080G.IFSC").Column
'============================================

    TotalExRow = Range("Per5080G.DoneeWithPanName").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet4.Range("Per5080G.DoneeWithPanName").ClearContents
        Sheet4.Range("Per5080G.AddrDetail").ClearContents
        Sheet4.Range("Per5080G.CityOrTownOrDistrict").ClearContents
        Sheet4.Range("Per5080G.StateCode").ClearContents
        Sheet4.Range("Per5080G.PinCode").ClearContents
        Sheet4.Range("Per5080G.DoneePAN").ClearContents
        
'Change.25.01.2023.102.80GC3
        Sheet4.Range("Per5080G.ArnNbr").ClearContents
'end Change====
        Sheet4.Range("Per5080G.DonationAmt").ClearContents
        Sheet4.Range("Per5080G.DonationAmtOther").ClearContents
    'Konda updated 10-03-2026--V0.5
        Sheet4.Range("Per5080G.Traref").ClearContents
         Sheet4.Range("Per5080G.IFSC").ClearContents
'============================================
    
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_D (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet4.Range("Per5080G.DoneeWithPanName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
            Sheet4.Cells(rowcount, NameColNo).Value = node("DoneeWithPanName")
            Sheet4.Cells(rowcount, AddressColNo).Value = node("AddressDetail")("AddrDetail")
            Sheet4.Cells(rowcount, CityColNo).Value = node("AddressDetail")("CityOrTownOrDistrict")
            Dim iState As Variant
            iState = node("AddressDetail")("StateCode")
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet4.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet4.Cells(rowcount, PincodeColNo).Value = node("AddressDetail")("PinCode")
            Sheet4.Cells(rowcount, PanofDoneeColNo).Value = node("DoneePAN")
'Change.25.01.2023.102.80GC1
            Sheet4.Cells(rowcount, ArnNbrColNo).Value = node("ArnNbr")
'End Change
            Sheet4.Cells(rowcount, AmountColNo).Value = node("DonationAmtCash")
            Sheet4.Cells(rowcount, DonationColNo).Value = node("DonationAmtOtherMode")
    'Konda updated on 10-03-2026--v0.5
            Sheet4.Cells(rowcount, TransactionRefNum_80GD).Value = node("TransactionRefNum")
            Sheet4.Cells(rowcount, IFSCCode_80GD).Value = node("IFSCCode")
    '============================

        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End If
End Function
Function ImportSchedule80GGA(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim Capacity As Variant
Dim RelevantClause, NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, DateofDonationNo, AmountColNo, DonationColNo, dateDonation As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, rowcount, Rec80GGA1 As Long
Dim YYYY, MM, DD, strDate As String

Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80GGA")("DonationDtlsSciRsrchRuralDev")

    RelevantClause = Sheet12.Range("RelevantClauseClaimed_80GGA").Column
    NameColNo = Sheet12.Range("Name_of_Donee_80GGA").Column
    AddressColNo = Sheet12.Range("Address_80GGA").Column
    CityColNo = Sheet12.Range("City_Town_District_80GGA").Column
    StateCodeColNo = Sheet12.Range("State_Code_80GGA").Column
    PincodeColNo = Sheet12.Range("Pincode_80GGA").Column
    PanofDoneeColNo = Sheet12.Range("PAN_of_donee_80GGA").Column
    'DateofDonationNo = Sheet12.Range("Date_Donation").Column
    AmountColNo = Sheet12.Range("Donation_cash_80GGA").Column
    DonationColNo = Sheet12.Range("Donation_other_80GGA").Column

    TotalExRow = Range("RelevantClauseClaimed_80GGA").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        Sheet12.Range("RelevantClauseClaimed_80GGA").ClearContents
        Sheet12.Range("Name_of_Donee_80GGA").ClearContents
        Sheet12.Range("Address_80GGA").ClearContents
        Sheet12.Range("City_Town_District_80GGA").ClearContents
        Sheet12.Range("State_Code_80GGA").ClearContents
        Sheet12.Range("Pincode_80GGA").ClearContents
        Sheet12.Range("PAN_of_donee_80GGA").ClearContents
        'Sheet12.Range("Date_Donation").ClearContents
        Sheet12.Range("Donation_cash_80GGA").ClearContents
        Sheet12.Range("Donation_other_80GGA").ClearContents
    End If

    If (TotalDiffRow > 0) Then
     AddDiffRows_80GGA (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet12.Range("RelevantClauseClaimed_80GGA").name)
    rowcount = rowcount - 1
    cnt = 0
    Dim Temp80GGA As Variant
    For Each node In Nodelist
        rowcount = rowcount + 1
         
        Temp80GGA = node("RelevantClauseUndrDedClaimed")
        
        If Temp80GGA = "80GGA2a" Then
        Temp80GGA = "80GGA(2)(a)"
        ElseIf Temp80GGA = "80GGA2aa" Then
        Temp80GGA = "80GGA(2)(aa)"
        ElseIf Temp80GGA = "80GGA2b" Then
        Temp80GGA = "80GGA(2)(b)"
        ElseIf Temp80GGA = "80GGA2bb" Then
        Temp80GGA = "80GGA(2)(bb)"
        ElseIf Temp80GGA = "80GGA2c" Then
        Temp80GGA = "80GGA(2)(c)"
        ElseIf Temp80GGA = "80GGA2cc" Then
        Temp80GGA = "80GGA(2)(cc)"
        ElseIf Temp80GGA = "80GGA2d" Then
        Temp80GGA = "80GGA(2)(d)"
        ElseIf Temp80GGA = "80GGA2e" Then
        Temp80GGA = "80GGA(2)(e)"
        End If
               
            Sheet12.Cells(rowcount, RelevantClause).Value = Findtext1(Temp80GGA, "RelevantClause80GGA")
            Sheet12.Cells(rowcount, NameColNo).Value = node("NameOfDonee")
            Sheet12.Cells(rowcount, AddressColNo).Value = node("AddressDetail")("AddrDetail")
            Sheet12.Cells(rowcount, CityColNo).Value = node("AddressDetail")("CityOrTownOrDistrict")

            Dim iState As Variant
            iState = node("AddressDetail")("StateCode")
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If


            If iState = "99" Then
            iState = ""
            End If

            Sheet12.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet12.Cells(rowcount, PincodeColNo).Value = node("AddressDetail")("PinCode")
            Sheet12.Cells(rowcount, PanofDoneeColNo).Value = node("DoneePAN")
             If node("DonationAmtOtherMode") <> "" And node("DonationAmtOtherMode") <> "0" Then
                Sheet12.Cells(rowcount, DonationColNo).Value = node("DonationAmtOtherMode")
            End If
            Sheet12.Cells(rowcount, AmountColNo).Value = node("DonationAmtCash")
            dateDonation = node("DonationAmtCashDate")
                YYYY = Mid(dateDonation, 1, 4)
                MM = Mid(dateDonation, 6, 2)
                DD = Mid(dateDonation, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            'Sheet12.Cells(rowcount, DateofDonationNo).Value = strDate
            
           
        cnt = cnt + 1
    Next node
    Rec80GGA1 = cnt
End Function

' ============================================= '
' Public Methods
' ============================================= '

''1
' Convert JSON string to object (Dictionary/Collection)
'
' @method ParseJson
' @param {String} json_String
' @return {Object} (Dictionary or Collection)
' @throws 10001 - JSON parse error
''
Public Function ParseJson(ByVal JsonString As String) As Object
    Dim json_Index As Long
    json_Index = 1

    ' Remove vbCr, vbLf, and vbTab from json_String
    JsonString = VBA.Replace(VBA.Replace(VBA.Replace(JsonString, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")

    json_SkipSpaces JsonString, json_Index
    Select Case VBA.Mid$(JsonString, json_Index, 1)
    Case "{"
        Set ParseJson = json_ParseObject(JsonString, json_Index)
    Case "["
        Set ParseJson = json_ParseArray(JsonString, json_Index)
    Case Else
        ' Error: Invalid JSON string
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, json_Index, "Expecting '{' or '['")
    End Select
End Function

Private Sub json_SkipSpaces(json_String As String, ByRef json_Index As Long)
    ' Increment index to skip over spaces
    Do While json_Index > 0 And json_Index <= VBA.Len(json_String) And VBA.Mid$(json_String, json_Index, 1) = " "
        json_Index = json_Index + 1
    Loop
End Sub

' ============================================= '
' Private Functions
' ============================================= '

Private Function json_ParseObject(json_String As String, ByRef json_Index As Long) As Object
    Dim json_Key As String
    Dim json_NextChar As String

    Set json_ParseObject = CreateObject("Scripting.Dictionary")
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "{" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '{'")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "}" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_Key = json_ParseKey(json_String, json_Index)
            json_NextChar = json_Peek(json_String, json_Index)
            If json_NextChar = "[" Or json_NextChar = "{" Then
                Set json_ParseObject.item(json_Key) = json_ParseValue(json_String, json_Index)
            Else
                json_ParseObject.item(json_Key) = json_ParseValue(json_String, json_Index)
            End If
        Loop
    End If
End Function
Private Function json_ParseArray(json_String As String, ByRef json_Index As Long) As Collection
    Set json_ParseArray = New Collection

    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> "[" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '['")
    Else
        json_Index = json_Index + 1

        Do
            json_SkipSpaces json_String, json_Index
            If VBA.Mid$(json_String, json_Index, 1) = "]" Then
                json_Index = json_Index + 1
                Exit Function
            ElseIf VBA.Mid$(json_String, json_Index, 1) = "," Then
                json_Index = json_Index + 1
                json_SkipSpaces json_String, json_Index
            End If

            json_ParseArray.add json_ParseValue(json_String, json_Index)
        Loop
    End If
End Function
Private Function json_ParseErrorMessage(json_String As String, ByRef json_Index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing JSON:
    ' {"abcde":True}
    '          ^
    ' Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['

    Dim json_StartIndex As Long
    Dim json_StopIndex As Long

    ' Include 10 characters before and after error (if possible)
    json_StartIndex = json_Index - 10
    json_StopIndex = json_Index + 10
    If json_StartIndex <= 0 Then
        json_StartIndex = 1
    End If
    If json_StopIndex > VBA.Len(json_String) Then
        json_StopIndex = VBA.Len(json_String)
    End If

    json_ParseErrorMessage = "Error parsing JSON:" & VBA.vbNewLine & _
                             VBA.Mid$(json_String, json_StartIndex, json_StopIndex - json_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(json_Index - json_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Function json_ParseKey(json_String As String, ByRef json_Index As Long) As String
    ' Parse key with single or double quotes
    If VBA.Mid$(json_String, json_Index, 1) = """" Or VBA.Mid$(json_String, json_Index, 1) = "'" Then
        json_ParseKey = json_ParseString(json_String, json_Index)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim json_Char As String
        Do While json_Index > 0 And json_Index <= Len(json_String)
            json_Char = VBA.Mid$(json_String, json_Index, 1)
            If (json_Char <> " ") And (json_Char <> ":") Then
                json_ParseKey = json_ParseKey & json_Char
                json_Index = json_Index + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting '""' or '''")
    End If

    ' Check for colon and skip if present or throw if not present
    json_SkipSpaces json_String, json_Index
    If VBA.Mid$(json_String, json_Index, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting ':'")
    Else
        json_Index = json_Index + 1
    End If
End Function
Private Function json_Peek(json_String As String, ByVal json_Index As Long, Optional json_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing json_Index (ByVal instead of ByRef)
    json_SkipSpaces json_String, json_Index
    json_Peek = VBA.Mid$(json_String, json_Index, json_NumberOfCharacters)
End Function
Private Function json_ParseValue(json_String As String, ByRef json_Index As Long) As Variant
    json_SkipSpaces json_String, json_Index
    Select Case VBA.Mid$(json_String, json_Index, 1)
    Case "{"
        Set json_ParseValue = json_ParseObject(json_String, json_Index)
    Case "["
        Set json_ParseValue = json_ParseArray(json_String, json_Index)
    Case """", "'"
        json_ParseValue = json_ParseString(json_String, json_Index)
    Case Else
        If VBA.Mid$(json_String, json_Index, 4) = "true" Then
            json_ParseValue = True
            json_Index = json_Index + 4
        ElseIf VBA.Mid$(json_String, json_Index, 5) = "false" Then
            json_ParseValue = False
            json_Index = json_Index + 5
        ElseIf VBA.Mid$(json_String, json_Index, 4) = "null" Then
            json_ParseValue = Null
            json_Index = json_Index + 4
        ElseIf VBA.Instr("+-0123456789", VBA.Mid$(json_String, json_Index, 1)) Then
            json_ParseValue = json_ParseNumber(json_String, json_Index)
        Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(json_String, json_Index, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
        End If
    End Select
End Function
Private Function json_ParseString(json_String As String, ByRef json_Index As Long) As String
    Dim json_Quote As String
    Dim json_Char As String
    Dim json_Code As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    json_SkipSpaces json_String, json_Index

    ' Store opening quote to look for matching closing quote
    json_Quote = VBA.Mid$(json_String, json_Index, 1)
    json_Index = json_Index + 1

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        Select Case json_Char
        Case "\"
            ' Escaped string, \\, or \/
            json_Index = json_Index + 1
            json_Char = VBA.Mid$(json_String, json_Index, 1)

            Select Case json_Char
            Case """", "\", "/", "'"
                json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "b"
                json_BufferAppend json_Buffer, vbBack, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "f"
                json_BufferAppend json_Buffer, vbFormFeed, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "n"
                json_BufferAppend json_Buffer, vbCrLf, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "r"
                json_BufferAppend json_Buffer, vbCr, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "t"
                json_BufferAppend json_Buffer, vbTab, json_BufferPosition, json_BufferLength
                json_Index = json_Index + 1
            Case "u"
                ' Unicode character escape (e.g. \u00a9 = Copyright)
                json_Index = json_Index + 1
                json_Code = VBA.Mid$(json_String, json_Index, 4)
                json_BufferAppend json_Buffer, VBA.ChrW(VBA.val("&h" + json_Code)), json_BufferPosition, json_BufferLength
                json_Index = json_Index + 4
            End Select
        Case json_Quote
            json_ParseString = json_BufferToString(json_Buffer, json_BufferPosition)
            json_Index = json_Index + 1
            Exit Function
        Case Else
            json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
            json_Index = json_Index + 1
        End Select
    Loop
End Function
Private Sub json_BufferAppend(ByRef json_Buffer As String, _
                              ByRef json_Append As Variant, _
                              ByRef json_BufferPosition As Long, _
                              ByRef json_BufferLength As Long)
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Put "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
    '
    ' and clsStringAppend from Philip Swannell
    ' https://github.com/VBA-tools/VBA-JSON/pull/82

    Dim json_AppendLength As Long
    Dim json_LengthPlusPosition As Long

    json_AppendLength = VBA.Len(json_Append)
    json_LengthPlusPosition = json_AppendLength + json_BufferPosition

    If json_LengthPlusPosition > json_BufferLength Then
        ' Appending would overflow buffer, add chunk
        ' (double buffer length or append length, whichever is bigger)
        Dim json_AddedLength As Long
        json_AddedLength = IIf(json_AppendLength > json_BufferLength, json_AppendLength, json_BufferLength)

        json_Buffer = json_Buffer & VBA.Space$(json_AddedLength)
        json_BufferLength = json_BufferLength + json_AddedLength
    End If

    ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
    ' Function call on left-hand side of assignment must return Variant or Object
    Mid$(json_Buffer, json_BufferPosition + 1, json_AppendLength) = CStr(json_Append)
    json_BufferPosition = json_BufferPosition + json_AppendLength
End Sub

Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function
Private Function json_ParseNumber(json_String As String, ByRef json_Index As Long) As Variant
    Dim json_Char As String
    Dim json_Value As String
    Dim json_IsLargeNumber As Boolean

    json_SkipSpaces json_String, json_Index

    Do While json_Index > 0 And json_Index <= Len(json_String)
        json_Char = VBA.Mid$(json_String, json_Index, 1)

        If VBA.Instr("+-0123456789.eE", json_Char) Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            json_Value = json_Value & json_Char
            json_Index = json_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
            ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
            json_IsLargeNumber = IIf(InStr(json_Value, "."), Len(json_Value) >= 17, Len(json_Value) >= 16)
            If Not JsonOptions.UseDoubleForLargeNumbers And json_IsLargeNumber Then
                json_ParseNumber = json_Value
            Else
                ' VBA.Val does not use regional settings, so guard for comma is not needed
                json_ParseNumber = VBA.val(json_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Function ImportSchedule80GGC(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim Capacity As Variant
Dim DateofDonation_80GGC1, NatureTransColNo, ChequenoColNo, IFSCColNo, NameofDonorColNo, AccountofDonorColNo, PanofDoneeColNo, DonationCashColNo, DonationOtherColNo, dateDonation As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, cnt, rowcount, Rec80GGC1 As Long
Dim YYYY, MM, DD, strDate As String

'Konda updated on 17-03-2026---V0.6
Dim PoliticalPartyNameColNo, PoliticalPartyPANColNo
'========================
Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80GGC")("Schedule80GGCDetails")

    DateofDonation_80GGC1 = Sheet13.Range("DateofDonation_80GGC").Column
    'NatureTransColNo = Sheet13.Range("NatureofTransaction_80GGC").Column
    ChequenoColNo = Sheet13.Range("Chequeno_80GGC").Column
    IFSCColNo = Sheet13.Range("IFSC_80GGC").Column
   ' NameofDonorColNo = Sheet13.Range("NameofDonor_80GGC").Column
    'AccountofDonorColNo = Sheet13.Range("AccountofDonor_80GGC").Column
    DonationCashColNo = Sheet13.Range("Donationincash_80GGC").Column
    DonationOtherColNo = Sheet13.Range("Donationinothermode_80GGC").Column
'Konda updated on 17-03-2026--V0.6
    PoliticalPartyNameColNo = Sheet13.Range("Name_80GGC").Column
    PoliticalPartyPANColNo = Sheet13.Range("PAN_80GGC").Column
'========================
    TotalExRow = Range("DateofDonation_80GGC").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        Sheet13.Range("DateofDonation_80GGC").ClearContents
        'Sheet13.Range("NatureofTransaction_80GGC").ClearContents
        Sheet13.Range("Chequeno_80GGC").ClearContents
        Sheet13.Range("IFSC_80GGC").ClearContents
        'Sheet13.Range("NameofDonor_80GGC").ClearContents
        'Sheet13.Range("AccountofDonor_80GGC").ClearContents
        Sheet13.Range("Donationincash_80GGC").ClearContents
        Sheet13.Range("Donationinothermode_80GGC").ClearContents
'Konda updated on 17-03-2026--V0.6
        Sheet13.Range("Name_80GGC").ClearContents
        Sheet13.Range("PAN_80GGC").ClearContents
'=====================
    End If

    If (TotalDiffRow > 0) Then
     AddDiffRows_80GGC (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet13.Range("DateofDonation_80GGC").name)
    rowcount = rowcount - 1
    cnt = 0
    Dim Temp80GGC As Variant
    For Each node In Nodelist
        rowcount = rowcount + 1
         
        dateDonation = node("DonationDate")
        YYYY = Mid(dateDonation, 1, 4)
        MM = Mid(dateDonation, 6, 2)
        DD = Mid(dateDonation, 9, 2)
        strDate = DD & "/" & MM & "/" & YYYY
        Sheet13.Cells(rowcount, DateofDonation_80GGC1).Value = strDate 'Date
        
        If node("DonationAmtCash") <> "" And Sheet13.Cells(rowcount, DonationCashColNo).Locked = False Then 'Donation in Cash
            Sheet13.Cells(rowcount, DonationCashColNo).Value = node("DonationAmtCash")
        End If
        
        If node("DonationAmtOtherMode") <> "" And Sheet13.Cells(rowcount, DonationOtherColNo).Locked = False Then 'Donation in Other mode
            Sheet13.Cells(rowcount, DonationOtherColNo).Value = node("DonationAmtOtherMode")
        End If
'Konda updated on 17-03-2026--V0.6
        If node("PoliticalPartyName") <> "" And Sheet13.Cells(rowcount, PoliticalPartyNameColNo).Locked = False Then 'Donation in Other mode
            Sheet13.Cells(rowcount, PoliticalPartyNameColNo).Value = node("PoliticalPartyName")
        End If
        
        If node("PoliticalPartyPAN") <> "" And Sheet13.Cells(rowcount, PoliticalPartyPANColNo).Locked = False Then 'Donation in Other mode
            Sheet13.Cells(rowcount, PoliticalPartyPANColNo).Value = node("PoliticalPartyPAN")
        End If
'==============================
'        Dim naturetrans
'
'        If Node("NatureOfTrans") = "UPI" Then
'           naturetrans = "UPI transfer"
'        ElseIf Node("NatureOfTrans") = "Cheque" Then
'           naturetrans = "Cheque issued"
'        ElseIf Node("NatureOfTrans") = "IMPS" Then
'           naturetrans = "IMPS"
'        ElseIf Node("NatureOfTrans") = "NEFT" Then
'           naturetrans = "NEFT"
'        ElseIf Node("NatureOfTrans") = "RTGS" Then
'           naturetrans = "RTGS"
'        End If
'
'
'
'        If Node("NatureOfTrans") <> "" And Sheet13.Cells(rowcount, NatureTransColNo).Locked = False Then 'Transaction
'            Sheet13.Cells(rowcount, NatureTransColNo).Value = naturetrans
'        End If
        
        If node("IFSCCode") <> "" And Sheet13.Cells(rowcount, IFSCColNo).Locked = False Then 'IFSC
            Sheet13.Cells(rowcount, IFSCColNo).Value = node("IFSCCode")
        End If
        
'        If Node("BankName") <> "" And Sheet13.Cells(rowcount, NameofDonorColNo).Locked = False Then 'Bank Name
'            Sheet13.Cells(rowcount, NameofDonorColNo).Value = Node("BankName")
'        End If
        
'        If Node("BankAccountNo") <> "" And Sheet13.Cells(rowcount, AccountofDonorColNo).Locked = False Then 'Bank Acc Number
'            Sheet13.Cells(rowcount, AccountofDonorColNo).Value = Node("BankAccountNo")
'        End If
        
        If node("TransactionRefNum") <> "" And Sheet13.Cells(rowcount, ChequenoColNo).Locked = False Then 'Transacation Ref
            Sheet13.Cells(rowcount, ChequenoColNo).Value = node("TransactionRefNum")
        End If
    
            
           
        cnt = cnt + 1
    Next node
    Rec80GGC1 = cnt
End Function

'Malli comented
Function ImportSchedule80DD_1(jsonText As String)
'On Error Resume Next
'Dim jsonObject, jsonDictionary As Object
'Dim node, Nodelist As Object
'Dim Nature As Variant
''Dim NatureOfDisability As Variant
'
'Set jsonObject = ParseJson(jsonText)
'Set node = jsonObject("ITR")("ITR1")("Schedule80DD")
'
'If Trim(node("NatureOfDisability")) <> "" Then
'If node("NatureOfDisability") = "1" Then
'   Sheet14.Range("Naturedisability_80DD").Value = "1-Dependent person with disability"
'
'   ElseIf node("NatureOfDisability") = "2" Then
'   Sheet14.Range("Naturedisability_80DD").Value = "2-Dependent person with severe disability"
'
'   Else: Sheet14.Range("Naturedisability_80DD").Value = "(Select)"
'
'End If
'End If
'   'Typedependent_80DD
'
'If Trim(node("DependentType")) <> "" Then
' 'node("DependentType") = 8
'If node("DependentType") = "1" Then
'   Sheet14.Range("Typedependent_80DD").Value = "1.Spouse"
'
'   ElseIf node("DependentType") = "2" Then
'   Sheet14.Range("Typedependent_80DD").Value = "2.Son"
'
'   ElseIf node("DependentType") = "3" Then
'   Sheet14.Range("Typedependent_80DD").Value = "3.Daughter"
'
'   ElseIf node("DependentType") = "4" Then
'   Sheet14.Range("Typedependent_80DD").Value = "4.Father"
'
'   ElseIf node("DependentType") = "5" Then
'   Sheet14.Range("Typedependent_80DD").Value = "5.Mother"
'
'   ElseIf node("DependentType") = "6" Then
'   Sheet14.Range("Typedependent_80DD").Value = "6.Brother"
'
'   ElseIf node("DependentType") = "7" Then
'   Sheet14.Range("Typedependent_80DD").Value = "7.Sister"
'
'   ElseIf node("DependentType") = "8" Then
'   Sheet14.Range("Typedependent_80DD").Value = "8.Member of the HUF (in case of HUF)"
'
'   Else
'   Sheet14.Range("Typedependent_80DD").Value = "(Select)"
'
'   End If
'
'End If
'
'If Trim(node("DependentPan")) <> "" Then
' 'Sheet14.Range("PANdependent_80DD").Value = UVCase(node("DependentPan"))
' Sheet14.Range("PANdependent_80DD").Value = UCase(node("DependentPan"))
'End If
'
'If Trim(node("DependentAadhaar")) <> "" Then
' 'Sheet14.Range("Aadhaardependent_80DD").Value = UVCase(node("DependentAadhaar"))
' Sheet14.Range("Aadhaardependent_80DD").Value = UCase(node("DependentAadhaar"))
'End If
'
'If Trim(node("Form10IAFilingDate")) <> "" Then
'
' dob = Mid(node("Form10IAFilingDate"), 9, 2) & "/" & Mid(node("Form10IAFilingDate"), 6, 2) & "/" & Mid(node("Form10IAFilingDate"), 1, 4)
' Sheet14.Range("DatefilingFm10IA_80DD").Value = dob
'End If
'
'If Trim(node("Form10IAAckNum")) <> "" Then
'
' 'Sheet14.Range("AckNoFm10IAfiled_80DD").Value = UVCase(node("Form10IAAckNum"))
' Sheet14.Range("AckNoFm10IAfiled_80DD").Value = UCase(node("Form10IAAckNum"))
'End If
'
'If Trim(node("UDIDNum")) <> "" Then
' 'Sheet14.Range("UDIDNum_80DD").Value = UVCase(node("UDIDNum"))
' Sheet14.Range("UDIDNum_80DD").Value = UCase(node("UDIDNum"))
'End If

End Function
'Malli comented
Function ImportSchedule80U_1(jsonText As String)
'On Error Resume Next
'Dim jsonObject, jsonDictionary As Object
'Dim node, Nodelist As Object
'
'Set jsonObject = ParseJson(jsonText)
'Set node = jsonObject("ITR")("ITR1")("Schedule80U")
'
'If Trim(node("NatureOfDisability")) <> "" Then
'
'
'
'If node("NatureOfDisability") = "1" Then
'   Sheet14.Range("Naturedisability_80U").Value = Trim("1-Self with disability")
'
'   ElseIf node("NatureOfDisability") = "2" Then
'   Sheet14.Range("Naturedisability_80U").Value = Trim("2-Self with severe disability")
'
'   Else: Sheet14.Range("Naturedisability_80U").Value = "(Select)"
'
'End If
'End If
'
'If Trim(node("Form10IAFilingDate")) <> "" Then
'  dob = Mid(node("Form10IAFilingDate"), 9, 2) & "/" & Mid(node("Form10IAFilingDate"), 6, 2) & "/" & Mid(node("Form10IAFilingDate"), 1, 4)
'  Sheet14.Range("DatefilingFm10IA_80U").Value = dob
'End If
'
'If Trim(node("Form10IAAckNum")) <> "" Then
' 'Sheet14.Range("AckNoFm10IAfiled_80U").Value = UVCase(node("Form10IAAckNum"))
' Sheet14.Range("AckNoFm10IAfiled_80U").Value = UCase(node("Form10IAAckNum"))
'End If
'
'If Trim(node("UDIDNum")) <> "" Then
' 'Sheet14.Range("UDIDNum_80U").Value = UVCase(node("UDIDNum"))
' Sheet14.Range("UDIDNum_80U").Value = UCase(node("UDIDNum"))
'End If


End Function


'Malli
Function ImportSchedule80DD(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim Nature As Variant
Dim Date10IA_2 As Variant


Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR1")

If Nodelist.exists("Schedule80DD") Then

Dim NatureOfDisability_80DD
NatureOfDisability_80DD = jsonObject("ITR")("ITR1")("Schedule80DD")("NatureOfDisability")
If NatureOfDisability_80DD <> "" Then
If NatureOfDisability_80DD = "1" Then
Sheet14.Range("Naturedisability_80DD").Value = "1-Dependent person with disability"
ElseIf NatureOfDisability_80DD = "2" Then
Sheet14.Range("Naturedisability_80DD").Value = "2-Dependent person with severe disability"
Else: Sheet14.Range("Naturedisability_80DD").Value = "(Select)"
End If

End If
   
 'Konda updated on 19-04-2025
'as AY_2025_26 enhancement
Dim TypeOfDisability_80DD
TypeOfDisability_80DD = jsonObject("ITR")("ITR1")("Schedule80DD")("TypeOfDisability")
    If TypeOfDisability_80DD <> "" Then

If TypeOfDisability_80DD = "1" Then
'Konda_06/05/2025_Commented and updated as per New Schema_v0.7 and V0.7.1
'Konda---------as per Clarification tracker-14/05/2025
    'Sheet14.Range("Disability_80DD").Value = "(i) autism, cerebral palsy, or multiple disabilities and"
    Sheet14.Range("Disability_80DD").Value = "(i) autism, cerebral palsy, or multiple disabilities"
    ElseIf TypeOfDisability_80DD = "2" Then
    Sheet14.Range("Disability_80DD").Value = "(ii) others"


'   Sheet14.Range("Disability_80DD").Value = "(i) blindness"
'   ElseIf TypeOfDisability_80DD = "2" Then
'   Sheet14.Range("Disability_80DD").Value = "(ii) low vision"
'   ElseIf TypeOfDisability_80DD = "3" Then
'   Sheet14.Range("Disability_80DD").Value = "(iii) leprosy-cured"
'   ElseIf TypeOfDisability_80DD = "4" Then
'   Sheet14.Range("Disability_80DD").Value = "(iv) hearing impairment"
'   ElseIf TypeOfDisability_80DD = "5" Then
'   Sheet14.Range("Disability_80DD").Value = "(v) locomotor disability"
'   ElseIf TypeOfDisability_80DD = "6" Then
'   Sheet14.Range("Disability_80DD").Value = "(vi) mental retardation"
'   ElseIf TypeOfDisability_80DD = "7" Then
'   Sheet14.Range("Disability_80DD").Value = "(vii) mental illness"
'   ElseIf TypeOfDisability_80DD = "8" Then
'   Sheet14.Range("Disability_80DD").Value = "(viii) autism"
'   ElseIf TypeOfDisability_80DD = "9" Then
'   Sheet14.Range("Disability_80DD").Value = "(ix) cerebral palsy"
'   ElseIf TypeOfDisability_80DD = "10" Then
'   Sheet14.Range("Disability_80DD").Value = "(x) multiple disability"
   Else
   Sheet14.Range("Disability_80DD").Value = "(Select)"
   End If
End If

'-----------------------------------
Dim dependentType_80DD
dependentType_80DD = jsonObject("ITR")("ITR1")("Schedule80DD")("DependentType")
 
If dependentType_80DD <> "" Then

If dependentType_80DD = "1" Then

   Sheet14.Range("Typedependent_80DD").Value = "1.Spouse"
   
   ElseIf dependentType_80DD = "2" Then
   Sheet14.Range("Typedependent_80DD").Value = "2.Son"
   
   ElseIf dependentType_80DD = "3" Then
   Sheet14.Range("Typedependent_80DD").Value = "3.Daughter"
   
   ElseIf dependentType_80DD = "4" Then
   Sheet14.Range("Typedependent_80DD").Value = "4.Father"
   
   ElseIf dependentType_80DD = "5" Then
   Sheet14.Range("Typedependent_80DD").Value = "5.Mother"
   
   ElseIf dependentType_80DD = "6" Then
   Sheet14.Range("Typedependent_80DD").Value = "6.Brother"
   
   ElseIf dependentType_80DD = "7" Then
   Sheet14.Range("Typedependent_80DD").Value = "7.Sister"
   
   ElseIf dependentType_80DD = "8" Then
   Sheet14.Range("Typedependent_80DD").Value = "8.Member of the HUF (in case of HUF)"
   
   Else
   Sheet14.Range("Typedependent_80DD").Value = "(Select)"
   
   End If
   
End If

Dim DependentPan_80DD
 DependentPan_80DD = jsonObject("ITR")("ITR1")("Schedule80DD")("DependentPan")


If DependentPan_80DD <> "" Then
 
 Sheet14.Range("PANdependent_80DD").Value = UCase(DependentPan_80DD)
End If

Dim DependentAadhaar_80DD
DependentAadhaar_80DD = jsonObject("ITR")("ITR1")("Schedule80DD")("DependentAadhaar")
If DependentAadhaar_80DD <> "" Then

 Sheet14.Range("Aadhaardependent_80DD").Value = UCase(DependentAadhaar_80DD)
End If

'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'Dim filingdate1_80DD
'
'filingdate1_80DD = jsonObject("ITR")("ITR1")("Schedule80DD")("Form10IAFilingDate")
'If Trim(filingdate1_80DD) <> "" Then
'  Date10IA_2 = Mid(filingdate1_80DD, 9, 2) & "/" & Mid(filingdate1_80DD, 6, 2) & "/" & Mid(filingdate1_80DD, 1, 4)
'  Sheet14.Range("DatefilingFm10IA_80DD").Value = Date10IA_2
'End If

Dim acknum_80DD
 
acknum_80DD = jsonObject("ITR")("ITR1")("Schedule80DD")("Form10IAAckNum")
 
If Trim(acknum_80DD) <> "" Then
Sheet14.Range("AckNoFm10IAfiled_80DD").Value = Trim(acknum_80DD)
End If

'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'Konda updated
'as AY_2025_26 enhancement
'Dim AcknlNum11A2_80DD
'
'AcknlNum11A2_80DD = jsonObject("ITR")("ITR1")("Schedule80DD")("FormAckNum11A")
'
'If Trim(AcknlNum11A2_80DD) <> "" Then
'Sheet14.Range("AcknowledgeNum11A2_80DD").Value = Trim(AcknlNum11A2_80DD)
'End If
'-----------------------

Dim UDIDnum_80DD
 
UDIDnum_80DD = jsonObject("ITR")("ITR1")("Schedule80DD")("UDIDNum")
 
If Trim(UDIDnum_80DD) <> "" Then
 Sheet14.Range("UDIDNum_80DD").Value = Trim(UDIDnum_80DD)
End If

End If
End Function
'Malli
Function ImportSchedule80U(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim Date_New As Variant

Set jsonObject = ParseJson(jsonText)
 
Set Nodelist = jsonObject("ITR")("ITR1")
 
If Nodelist.exists("Schedule80U") Then
 
Dim NatureOfDisability_80U

NatureOfDisability_80U = jsonObject("ITR")("ITR1")("Schedule80U")("NatureOfDisability")

If NatureOfDisability_80U <> "" Then



If Trim(NatureOfDisability_80U) = "1" Then
   Sheet14.Range("Naturedisability_80U").Value = Trim("1-Self with disability")
   
   ElseIf Trim(NatureOfDisability_80U) = "2" Then
   Sheet14.Range("Naturedisability_80U").Value = Trim("2-Self with severe disability")
   
   Else: Sheet14.Range("Naturedisability_80U").Value = "(Select)"
   
End If
End If
Dim TypeOfDisability_80U
TypeOfDisability_80U = jsonObject("ITR")("ITR1")("Schedule80U")("TypeOfDisability")
    If TypeOfDisability_80U <> "" Then

If TypeOfDisability_80U = "1" Then
'Konda_06/05/2025_Commented and updated as per New Schema_v0.7 and V0.7.1
'Konda---------as per Clarification tracker-14/05/2025
   'Sheet14.Range("Disability_80U").Value = "(i) autism, cerebral palsy, or multiple disabilities and"
   Sheet14.Range("Disability_80U").Value = "(i) autism, cerebral palsy, or multiple disabilities"
   ElseIf TypeOfDisability_80U = "2" Then
   Sheet14.Range("Disability_80U").Value = "(ii) others"

'   Sheet14.Range("Disability_80U").Value = "(i) blindness"
'   ElseIf TypeOfDisability_80U = "2" Then
'   Sheet14.Range("Disability_80U").Value = "(ii) low vision"
'   ElseIf TypeOfDisability_80U = "3" Then
'   Sheet14.Range("Disability_80U").Value = "(iii) leprosy-cured"
'   ElseIf TypeOfDisability_80U = "4" Then
'   Sheet14.Range("Disability_80U").Value = "(iv) hearing impairment"
'   ElseIf TypeOfDisability_80U = "5" Then
'   Sheet14.Range("Disability_80U").Value = "(v) locomotor disability"
'   ElseIf TypeOfDisability_80U = "6" Then
'   Sheet14.Range("Disability_80U").Value = "(vi) mental retardation"
'   ElseIf TypeOfDisability_80U = "7" Then
'   Sheet14.Range("Disability_80U").Value = "(vii) mental illness"
'   ElseIf TypeOfDisability_80U = "8" Then
'   Sheet14.Range("Disability_80U").Value = "(viii) autism"
'   ElseIf TypeOfDisability_80U = "9" Then
'   Sheet14.Range("Disability_80U").Value = "(ix) cerebral palsy"
'   ElseIf TypeOfDisability_80U = "10" Then
'   Sheet14.Range("Disability_80U").Value = "(x) multiple disability"
   Else
   Sheet14.Range("Disability_80U").Value = "(Select)"
   End If
End If

'End-----------------------------------on 19-04-2025
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'Dim filingdate1_80U
'
'filingdate1_80U = jsonObject("ITR")("ITR1")("Schedule80U")("Form10IAFilingDate")
'
'If Trim(filingdate1_80U) <> "" Then
'
'  Date_New = Mid(filingdate1_80U, 9, 2) & "/" & Mid(filingdate1_80U, 6, 2) & "/" & Mid(filingdate1_80U, 1, 4)
'
'   Sheet14.Range("DatefilingFm10IA_80U").Value = Date_New
'
'End If

Dim acknum_80U

acknum_80U = Trim(jsonObject("ITR")("ITR1")("Schedule80U")("Form10IAAckNum"))


If Trim(acknum_80U) <> "" Then

Sheet14.Range("AckNoFm10IAfiled_80U").Value = UCase(acknum_80U)

End If
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
''Konda updated on 19-04-2025
''as AY_2025_26 enhancement
' Dim AcknlNum11A2_80U
'
'AcknlNum11A2_80U = Trim(jsonObject("ITR")("ITR1")("Schedule80U")("FormAckNum11A"))
'
'If Trim(AcknlNum11A2_80U) <> "" Then
'
'Sheet14.Range("AcknowledgeNum11A2_80U").Value = UCase(AcknlNum11A2_80U)
'
'End If
'end----------------------on 19-04-2025
Dim UDIDnum_80U
 
UDIDnum_80U = Trim(jsonObject("ITR")("ITR1")("Schedule80U")("UDIDNum"))


If Trim(UDIDnum_80U) <> "" Then

Sheet14.Range("UDIDNum_80U").Value = UCase(UDIDnum_80U)

End If

End If
End Function
'Konda_80C_AY_2025_26-------09/04/2025
Function ImportSchedule_80C(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object

Dim NaturePayment_80C, Identification_Number_80C, Amount_80C As Variant
Dim TotalXMLRow, rowcount, cnt, Rec80C  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long

    
 Set jsonObject = ParseJson(jsonText)
 
 'Malli------80C
 'Malli-25/04/2025
Dim Nodelist80C As Object
Set Nodelist80C = jsonObject("ITR")("ITR1")
If Nodelist80C.exists("Schedule80C") Then
'--------------------------------
  
   Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80C")("Schedule80CDtls")
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
    'NaturePayment_80C = Sheet15.Range("NaturePayment.80C").Column
    Identification_Number_80C = Sheet15.Range("Identification_Number.80C").Column
     Amount_80C = Sheet15.Range("Amount.80C").Column
    
    
    'Konda----08_05_2025
    'TotalExRow = Sheet15.Range("NaturePayment.80C").Rows.count
    TotalExRow = Sheet15.Range("Amount.80C").Rows.count
    '--------------------------
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet15.Range("NaturePayment.80C").Locked = False Then
'            Sheet15.Range("NaturePayment.80C").ClearContents
'        End If
        If Sheet15.Range("Identification_Number.80C").Locked = False Then
            Sheet15.Range("Identification_Number.80C").ClearContents
        End If
        If Sheet15.Range("Amount.80C").Locked = False Then
            Sheet15.Range("Amount.80C").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80C (TotalDiffRow)
    End If
    'Konda--------------08_05_2025
    'rowcount = getRowNo(Sheet15.Range("NaturePayment.80C").name)
     rowcount = getRowNo(Sheet15.Range("Amount.80C").name)
     '---------------------------
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet15.Cells(rowcount, NaturePayment_80C).Locked = False Then
'            Sheet15.Cells(rowcount, NaturePayment_80C).Value = node("NatureOfPayment")
'        End If
        If Sheet15.Cells(rowcount, Identification_Number_80C).Locked = False Then
            Sheet15.Cells(rowcount, Identification_Number_80C).Value = node("IdentificationNo")
        End If
        If Sheet15.Cells(rowcount, Amount_80C).Locked = False Then
            Sheet15.Cells(rowcount, Amount_80C).Value = node("Amount")
        End If
        
        cnt = cnt + 1
    Next node
    Rec80C = cnt
End If
End Function

'Konda_80CCC_AY_2025_26-------09/04/2025
'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Function ImportSchedule_80CCC(jsonText As String)
'On Error Resume Next
'Dim jsonObject, jsonDictionary As Object
'Dim node, Nodelist As Object
'
'Dim NaturePayment_80CCC, Policy_Document_Number_80CCC, Amount_80CCC As Variant
'Dim TotalXMLRow, rowcount, cnt, Rec80CCC  As Long
'Dim TotalDiffRow As Long
'Dim TotalExRow As Long
'
'
' Set jsonObject = ParseJson(jsonText)
'
''Malli-25/04/2025
'Dim Nodelist80CCC As Object
'Set Nodelist80CCC = jsonObject("ITR")("ITR1")
'If Nodelist80CCC.EXISTS("Schedule80CCC") Then
''-------------------------
'
'   Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80CCC")("Schedule80CCCDtls")
'
'    NaturePayment_80CCC = Sheet15.Range("Name_of_insurer.80CCC").Column
'    Policy_Document_Number_80CCC = Sheet15.Range("Policy_Document_Number.80CCC").Column
'    Amount_80CCC = Sheet15.Range("Amount.80CCC").Column
'
'
'
'    TotalExRow = Sheet15.Range("Name_of_insurer.80CCC").Rows.count
'    TotalXMLRow = Nodelist.count
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'    If (TotalXMLRow > 0) Then
'
'        If Sheet15.Range("Name_of_insurer.80CCC").Locked = False Then
'            Sheet15.Range("Name_of_insurer.80CCC").ClearContents
'        End If
'
'        If Sheet15.Range("Policy_Document_Number.80CCC").Locked = False Then
'            Sheet15.Range("Policy_Document_Number.80CCC").ClearContents
'        End If
'        If Sheet15.Range("Amount.80CCC").Locked = False Then
'            Sheet15.Range("Amount.80CCC").ClearContents
'        End If
'
'    End If
'
'    If (TotalDiffRow > 0) Then
'     AddDiffRows_80CCC (TotalDiffRow)
'    End If
'
'    rowcount = getRowNo(Sheet15.Range("Name_of_insurer.80CCC").name)
'    rowcount = rowcount - 1
'    cnt = 0
'
'    For Each node In Nodelist
'        rowcount = rowcount + 1
'        If Sheet15.Cells(rowcount, NaturePayment_80CCC).Locked = False Then
'            Sheet15.Cells(rowcount, NaturePayment_80CCC).Value = node("InsurerName")
'        End If
'        If Sheet15.Cells(rowcount, Policy_Document_Number_80CCC).Locked = False Then
'            Sheet15.Cells(rowcount, Policy_Document_Number_80CCC).Value = node("PolicyDocNo")
'        End If
'        If Sheet15.Cells(rowcount, Amount_80CCC).Locked = False Then
'            Sheet15.Cells(rowcount, Amount_80CCC).Value = node("Amount")
'        End If
'
'        cnt = cnt + 1
'    Next node
'    Rec80CCC = cnt
'End If
'End Function
'Konda_80E_AY_2025_26-------09/04/2025
Function ImportSchedule_80E(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80E, IFSC_80E, BankName_80E, PANof_80E, LoanAcctNum_80E, LoanDate_80E, TotalLoanAmt_80E, LoanOutStanding_80E, Intrst_80E As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80E, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

Set jsonObject = ParseJson(jsonText)

'Malli-25/04/2025
Dim NodeList80E As Object
Set NodeList80E = jsonObject("ITR")("ITR1")
If NodeList80E.exists("Schedule80E") Then
'---------------------------------

Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80E")("Schedule80EDtls")

    LoanfrmBankOrInstitute_80E = Sheet17.Range("LoanfrmBankOrInstitute.80E").Column
'    IFSC_80E = Sheet17.Range("IFSC.80E").Column  'Ankita_05/05/2025_Commented as per DESheet_v0.7
    BankName_80E = Sheet17.Range("bankName.80E").Column
'    PANof_80E = Sheet17.Range("PAN.80E").Column  'Ankita_05/05/2025_Commented as per DESheet_v0.7
    LoanAcctNum_80E = Sheet17.Range("loanAccNum.80E").Column
    LoanDate_80E = Sheet17.Range("loanDate.80E").Column
    TotalLoanAmt_80E = Sheet17.Range("loanAmt.80E").Column
    LoanOutStanding_80E = Sheet17.Range("loanOutstanding.80E").Column
    Intrst_80E = Sheet17.Range("Intrst.80E").Column
    
    
    TotalExRow = Sheet17.Range("LoanfrmBankOrInstitute.80E").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet17.Range("LoanfrmBankOrInstitute.80E").Locked = False Then
            Sheet17.Range("LoanfrmBankOrInstitute.80E").ClearContents
        End If
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'        If Sheet17.Range("IFSC.80E").Locked = False Then
'            Sheet17.Range("IFSC.80E").ClearContents
'        End If
        If Sheet17.Range("bankName.80E").Locked = False Then
            Sheet17.Range("bankName.80E").ClearContents
        End If
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'        If Sheet17.Range("PAN.80E").Locked = False Then
'            Sheet17.Range("PAN.80E").ClearContents
'        End If
        If Sheet17.Range("loanAccNum.80E").Locked = False Then
            Sheet17.Range("loanAccNum.80E").ClearContents
        End If
        If Sheet17.Range("loanDate.80E").Locked = False Then
            Sheet17.Range("loanDate.80E").ClearContents
        End If
        If Sheet17.Range("loanAmt.80E").Locked = False Then
            Sheet17.Range("loanAmt.80E").ClearContents
        End If
        If Sheet17.Range("loanOutstanding.80E").Locked = False Then
            Sheet17.Range("loanOutstanding.80E").ClearContents
        End If
        If Sheet17.Range("Intrst.80E").Locked = False Then
            Sheet17.Range("Intrst.80E").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80E (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet17.Range("LoanfrmBankOrInstitute.80E").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80E).Locked = False Then
            Dim LoanTknFrom_80E
            If node("LoanTknFrom") = "B" Then
                LoanTknFrom_80E = "Bank"
            ElseIf node("LoanTknFrom") = "I" Then
                LoanTknFrom_80E = "Institution"
            End If
            Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80E).Value = LoanTknFrom_80E
        End If
 'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet17.Cells(rowcount, IFSC_80E).Locked = False Then
'            Sheet17.Cells(rowcount, IFSC_80E).Value = node("IFSCCode")
'        End If
        If Sheet17.Cells(rowcount, BankName_80E).Locked = False Then
            Sheet17.Cells(rowcount, BankName_80E).Value = node("BankOrInstnName")
        End If
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet17.Cells(rowcount, PANof_80E).Locked = False Then
'            Sheet17.Cells(rowcount, PANof_80E).Value = node("PAN")
'        End If
        
        If Sheet17.Cells(rowcount, LoanAcctNum_80E).Locked = False Then
            Sheet17.Cells(rowcount, LoanAcctNum_80E).Value = node("LoanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("DateofLoan")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            End If
        
        If Sheet17.Cells(rowcount, LoanDate_80E).Locked = False Then
            Sheet17.Cells(rowcount, LoanDate_80E).Value = strDate
            strDate = ""
        End If
        If Sheet17.Cells(rowcount, TotalLoanAmt_80E).Locked = False Then
            Sheet17.Cells(rowcount, TotalLoanAmt_80E).Value = node("TotalLoanAmt")
        End If
        If Sheet17.Cells(rowcount, LoanOutStanding_80E).Locked = False Then
            Sheet17.Cells(rowcount, LoanOutStanding_80E).Value = node("LoanOutstndngAmt")
        End If
        
        If Sheet17.Cells(rowcount, Intrst_80E).Locked = False Then
            Sheet17.Cells(rowcount, Intrst_80E).Value = node("Interest80E")
        End If
        cnt = cnt + 1
    Next node
    Rec80E = cnt
End If
End Function
'Konda_80EE_AY_2025_26-------09/04/2025
Function ImportSchedule_80EE(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EE, IFSC_80EE, BankName_80EE, PANof_80EE, LoanAcctNum_80EE, LoanDate_80EE, TotalLoanAmt_80EE, LoanOutStanding_80EE, Intrst_80EE As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EE, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

Set jsonObject = ParseJson(jsonText)

'Malli-25/04/2025
Dim NodeList80EE As Object
Set NodeList80EE = jsonObject("ITR")("ITR1")
If NodeList80EE.exists("Schedule80EE") Then
'-----------------------------------
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'Dim ResHPVal_80EE
'ResHPVal_80EE = jsonObject("ITR")("ITR1")("Schedule80EE")("ResHPVal")
'
'If ResHPVal_80EE <> "" Then
'    If Sheet17.Range("ResidentialHP.80E").Locked = False Then
'            Sheet17.Range("ResidentialHP.80E").Value = ResHPVal_80EE
'        End If
'End If

Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80EE")("Schedule80EEDtls")

    LoanfrmBankOrInstitute_80EE = Sheet17.Range("LoanfrmBankOrInstitute.80EE").Column
'   IFSC_80EE = Sheet17.Range("IFSC.80EE").Column   'Ankita_05/05/2025_Commented as per DESheet_v0.7
    BankName_80EE = Sheet17.Range("bankName.80EE").Column
'   PANof_80EE = Sheet17.Range("PAN.80EE").Column   'Ankita_05/05/2025_Commented as per DESheet_v0.7
    LoanAcctNum_80EE = Sheet17.Range("loanAccNum.80EE").Column
    LoanDate_80EE = Sheet17.Range("loanDate.80EE").Column
    TotalLoanAmt_80EE = Sheet17.Range("loanAmt.80EE").Column
    LoanOutStanding_80EE = Sheet17.Range("loanOutstanding.80EE").Column
    
    Intrst_80EE = Sheet17.Range("Intrst.80EE").Column
    
    TotalExRow = Sheet17.Range("LoanfrmBankOrInstitute.80EE").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet17.Range("LoanfrmBankOrInstitute.80EE").Locked = False Then
            Sheet17.Range("LoanfrmBankOrInstitute.80EE").ClearContents
        End If
'        If Sheet17.Range("IFSC.80EE").Locked = False Then
'            Sheet17.Range("IFSC.80EE").ClearContents
'        End If
        If Sheet17.Range("bankName.80EE").Locked = False Then
            Sheet17.Range("bankName.80EE").ClearContents
        End If
'        If Sheet17.Range("PAN.80EE").Locked = False Then
'            Sheet17.Range("PAN.80EE").ClearContents
'        End If
        If Sheet17.Range("loanAccNum.80EE").Locked = False Then
            Sheet17.Range("loanAccNum.80EE").ClearContents
        End If
        If Sheet17.Range("loanDate.80EE").Locked = False Then
            Sheet17.Range("loanDate.80EE").ClearContents
        End If
        If Sheet17.Range("loanAmt.80EE").Locked = False Then
            Sheet17.Range("loanAmt.80EE").ClearContents
        End If
        If Sheet17.Range("loanOutstanding.80EE").Locked = False Then
            Sheet17.Range("loanOutstanding.80EE").ClearContents
        End If
        
        If Sheet17.Range("Intrst.80EE").Locked = False Then
            Sheet17.Range("Intrst.80EE").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80EE (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet17.Range("LoanfrmBankOrInstitute.80EE").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80EE).Locked = False Then
            Dim LoanTknFrom_80EE
            If node("LoanTknFrom") = "B" Then
                LoanTknFrom_80EE = "Bank"
            ElseIf node("LoanTknFrom") = "I" Then
                LoanTknFrom_80EE = "Institution"
            End If
            Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80EE).Value = LoanTknFrom_80EE
        End If
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet17.Cells(rowcount, IFSC_80EE).Locked = False Then
'            Sheet17.Cells(rowcount, IFSC_80EE).Value = node("IFSCCode")
'        End If
        If Sheet17.Cells(rowcount, BankName_80EE).Locked = False Then
            Sheet17.Cells(rowcount, BankName_80EE).Value = node("BankOrInstnName")
        End If
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet17.Cells(rowcount, PANof_80EE).Locked = False Then
'            Sheet17.Cells(rowcount, PANof_80EE).Value = node("PAN")
'        End If
        
        If Sheet17.Cells(rowcount, LoanAcctNum_80EE).Locked = False Then
            Sheet17.Cells(rowcount, LoanAcctNum_80EE).Value = node("LoanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("DateofLoan")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            End If
        
        If Sheet17.Cells(rowcount, LoanDate_80EE).Locked = False Then
            Sheet17.Cells(rowcount, LoanDate_80EE).Value = strDate
            strDate = ""
        End If
        If Sheet17.Cells(rowcount, TotalLoanAmt_80EE).Locked = False Then
            Sheet17.Cells(rowcount, TotalLoanAmt_80EE).Value = node("TotalLoanAmt")
        End If
        If Sheet17.Cells(rowcount, LoanOutStanding_80EE).Locked = False Then
            Sheet17.Cells(rowcount, LoanOutStanding_80EE).Value = node("LoanOutstndngAmt")
        End If
        
        If Sheet17.Cells(rowcount, Intrst_80EE).Locked = False Then
            Sheet17.Cells(rowcount, Intrst_80EE).Value = node("Interest80EE")
        End If
        cnt = cnt + 1
    Next node
    Rec80EE = cnt
End If
End Function
'Konda_80EEA_AY_2025_26-------09/04/2025
Function ImportSchedule_80EEA(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEA, IFSC_80EEA, BankName_80EEA, PANof_80EEA, LoanAcctNum_80EEA, LoanDate_80EEA, TotalLoanAmt_80EEA, LoanOutStanding_80EEA, Intrst_80EEA As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEA, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

Set jsonObject = ParseJson(jsonText)

'Malli-25/04/2025
Dim NodeList80EEA As Object
Set NodeList80EEA = jsonObject("ITR")("ITR1")
If NodeList80EEA.exists("Schedule80EEA") Then
'-----------------------------------------

Dim PropStmpDtyVal_80EEA
PropStmpDtyVal_80EEA = jsonObject("ITR")("ITR1")("Schedule80EEA")("PropStmpDtyVal")

If PropStmpDtyVal_80EEA <> "" Then
    If Sheet17.Range("Stampduty.80EEA").Locked = False Then
            Sheet17.Range("Stampduty.80EEA").Value = PropStmpDtyVal_80EEA
        End If
End If

Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80EEA")("Schedule80EEADtls")

    LoanfrmBankOrInstitute_80EEA = Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Column
'    IFSC_80EEA = Sheet17.Range("IFSC.80EEA").Column  'Ankita_05/05/2025_Commented as per DESheet_v0.7
    BankName_80EEA = Sheet17.Range("bankName.80EEA").Column
'    PANof_80EEA = Sheet17.Range("PAN.80EEA").Column  'Ankita_05/05/2025_Commented as per DESheet_v0.7
    LoanAcctNum_80EEA = Sheet17.Range("loanAccNum.80EEA").Column
    LoanDate_80EEA = Sheet17.Range("loanDate.80EEA").Column
    TotalLoanAmt_80EEA = Sheet17.Range("loanAmt.80EEA").Column
    LoanOutStanding_80EEA = Sheet17.Range("loanOutstanding.80EEA").Column
    
    Intrst_80EEA = Sheet17.Range("Intrst.80EEA").Column
    
    TotalExRow = Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Locked = False Then
            Sheet17.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
        End If
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'        If Sheet17.Range("IFSC.80EEA").Locked = False Then
'            Sheet17.Range("IFSC.80EEA").ClearContents
'        End If
        If Sheet17.Range("bankName.80EEA").Locked = False Then
            Sheet17.Range("bankName.80EEA").ClearContents
        End If
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'        If Sheet17.Range("PAN.80EEA").Locked = False Then
'            Sheet17.Range("PAN.80EEA").ClearContents
'        End If
        If Sheet17.Range("loanAccNum.80EEA").Locked = False Then
            Sheet17.Range("loanAccNum.80EEA").ClearContents
        End If
        If Sheet17.Range("loanDate.80EEA").Locked = False Then
            Sheet17.Range("loanDate.80EEA").ClearContents
        End If
        If Sheet17.Range("loanAmt.80EEA").Locked = False Then
            Sheet17.Range("loanAmt.80EEA").ClearContents
        End If
        If Sheet17.Range("loanOutstanding.80EEA").Locked = False Then
            Sheet17.Range("loanOutstanding.80EEA").ClearContents
        End If
        
        If Sheet17.Range("Intrst.80EEA").Locked = False Then
            Sheet17.Range("Intrst.80EEA").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80EEA (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet17.Range("LoanfrmBankOrInstitute.80EEA").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80EEA).Locked = False Then
            Dim LoanTknFrom_80EEA
            If node("LoanTknFrom") = "B" Then
                LoanTknFrom_80EEA = "Bank"
            ElseIf node("LoanTknFrom") = "I" Then
                LoanTknFrom_80EEA = "Institution"
            End If
            Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80EEA).Value = LoanTknFrom_80EEA
        End If
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet17.Cells(rowcount, IFSC_80EEA).Locked = False Then
'            Sheet17.Cells(rowcount, IFSC_80EEA).Value = node("IFSCCode")
'        End If
        If Sheet17.Cells(rowcount, BankName_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, BankName_80EEA).Value = node("BankOrInstnName")
        End If
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet17.Cells(rowcount, PANof_80EEA).Locked = False Then
'            Sheet17.Cells(rowcount, PANof_80EEA).Value = node("PAN")
'        End If
        
        If Sheet17.Cells(rowcount, LoanAcctNum_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, LoanAcctNum_80EEA).Value = node("LoanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("DateofLoan")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            End If
        
        If Sheet17.Cells(rowcount, LoanDate_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, LoanDate_80EEA).Value = strDate
            strDate = ""
        End If
        If Sheet17.Cells(rowcount, TotalLoanAmt_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, TotalLoanAmt_80EEA).Value = node("TotalLoanAmt")
        End If
        If Sheet17.Cells(rowcount, LoanOutStanding_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, LoanOutStanding_80EEA).Value = node("LoanOutstndngAmt")
        End If
        
        If Sheet17.Cells(rowcount, Intrst_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, Intrst_80EEA).Value = node("Interest80EEA")
        End If
        cnt = cnt + 1
    Next node
    Rec80EEA = cnt
End If
End Function
'Konda_80EEB_AY_2025_26-------09/04/2025
Function ImportSchedule_80EEB(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEB, IFSC_80EEB, BankName_80EEB, Vehicle_value_80EEB, VehicleRegNum_80EEB, PANof_80EEB, LoanAcctNum_80EEB, LoanDate_80EEB, TotalLoanAmt_80EEB, LoanOutStanding_80EEB, Intrst_80EEB As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEB, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

Set jsonObject = ParseJson(jsonText)

'Malli-25/04/2025
Dim NodeList80EEB As Object
Set NodeList80EEB = jsonObject("ITR")("ITR1")
If NodeList80EEB.exists("Schedule80EEB") Then
'-----------------------------------------------

Set Nodelist = jsonObject("ITR")("ITR1")("Schedule80EEB")("Schedule80EEBDtls")

    LoanfrmBankOrInstitute_80EEB = Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Column
'    IFSC_80EEB = Sheet17.Range("IFSC.80EEB").Column  'Ankita_05/05/2025_Commented as per DESheet_v0.7
    BankName_80EEB = Sheet17.Range("bankName.80EEB").Column
'    PANof_80EEB = Sheet17.Range("PAN.80EEB").Column  'Ankita_05/05/2025_Commented as per DESheet_v0.7
    LoanAcctNum_80EEB = Sheet17.Range("loanAccNum.80EEB").Column
    LoanDate_80EEB = Sheet17.Range("loanDate.80EEB").Column
    TotalLoanAmt_80EEB = Sheet17.Range("loanAmt.80EEB").Column
    LoanOutStanding_80EEB = Sheet17.Range("loanOutstanding.80EEB").Column
'    Vehicle_value_80EEB = Sheet17.Range("Vehicle_value.80EEB").Column  'Ankita_05/05/2025_Commented as per DESheet_v0.7
    VehicleRegNum_80EEB = Sheet17.Range("VehicleRegNum.80EEB").Column
    Intrst_80EEB = Sheet17.Range("Intrst.80EEB").Column
    
    
    
    TotalExRow = Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Locked = False Then
            Sheet17.Range("LoanfrmBankOrInstitute.80EEB").ClearContents
        End If
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'        If Sheet17.Range("IFSC.80EEB").Locked = False Then
'            Sheet17.Range("IFSC.80EEB").ClearContents
'        End If
        If Sheet17.Range("bankName.80EEB").Locked = False Then
            Sheet17.Range("bankName.80EEB").ClearContents
        End If
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'        If Sheet17.Range("PAN.80EEB").Locked = False Then
'            Sheet17.Range("PAN.80EEB").ClearContents
'        End If
        If Sheet17.Range("loanAccNum.80EEB").Locked = False Then
            Sheet17.Range("loanAccNum.80EEB").ClearContents
        End If
        If Sheet17.Range("loanDate.80EEB").Locked = False Then
            Sheet17.Range("loanDate.80EEB").ClearContents
        End If
        If Sheet17.Range("loanAmt.80EEB").Locked = False Then
            Sheet17.Range("loanAmt.80EEB").ClearContents
        End If
        If Sheet17.Range("loanOutstanding.80EEB").Locked = False Then
            Sheet17.Range("loanOutstanding.80EEB").ClearContents
        End If
        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'        If Sheet17.Range("Vehicle_value.80EEB").Locked = False Then
'            Sheet17.Range("Vehicle_value.80EEB").ClearContents
'        End If
        If Sheet17.Range("VehicleRegNum.80EEB").Locked = False Then
            Sheet17.Range("VehicleRegNum.80EEB").ClearContents
        End If
        If Sheet17.Range("Intrst.80EEB").Locked = False Then
            Sheet17.Range("Intrst.80EEB").ClearContents
        End If
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80EEB (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet17.Range("LoanfrmBankOrInstitute.80EEB").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80EEB).Locked = False Then
            Dim LoanTknFrom_80EEB
            If node("LoanTknFrom") = "B" Then
                LoanTknFrom_80EEB = "Bank"
            ElseIf node("LoanTknFrom") = "I" Then
                LoanTknFrom_80EEB = "Institution"
            End If
            Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80EEB).Value = LoanTknFrom_80EEB
        End If
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet17.Cells(rowcount, IFSC_80EEB).Locked = False Then
'            Sheet17.Cells(rowcount, IFSC_80EEB).Value = node("IFSCCode")
'        End If
        If Sheet17.Cells(rowcount, BankName_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, BankName_80EEB).Value = node("BankOrInstnName")
        End If
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet17.Cells(rowcount, PANof_80EEB).Locked = False Then
'            Sheet17.Cells(rowcount, PANof_80EEB).Value = node("PAN")
'        End If
        
        If Sheet17.Cells(rowcount, LoanAcctNum_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, LoanAcctNum_80EEB).Value = node("LoanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("DateofLoan")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            End If
        
        If Sheet17.Cells(rowcount, LoanDate_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, LoanDate_80EEB).Value = strDate
            strDate = ""
        End If
        If Sheet17.Cells(rowcount, TotalLoanAmt_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, TotalLoanAmt_80EEB).Value = node("TotalLoanAmt")
        End If
        If Sheet17.Cells(rowcount, LoanOutStanding_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, LoanOutStanding_80EEB).Value = node("LoanOutstndngAmt")
        End If
'Konda_06/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'        If Sheet17.Cells(rowcount, Vehicle_value_80EEB).Locked = False Then
'            Sheet17.Cells(rowcount, Vehicle_value_80EEB).Value = node("VehicleValue")
'        End If
        If Sheet17.Cells(rowcount, VehicleRegNum_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, VehicleRegNum_80EEB).Value = node("VehicleRegNo")
        End If
        If Sheet17.Cells(rowcount, Intrst_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, Intrst_80EEB).Value = node("Interest80EEB")
        End If
        cnt = cnt + 1
    Next node
    Rec80EEB = cnt
End If

End Function


'Konda_Int24B_AY_2026_27-------29/01/2026

'Konda_Int24B_AY_2025_26-------09/04/2025
'Function ImportSchedule_Int24B(jsonText As String)
'On Error Resume Next
'Dim jsonObject, jsonDictionary As Object
'Dim node, Nodelist As Object
'Dim LoanfrmBankOrInstitute_24B, IFSC_24B, BankName_24B, PANof_24B, LoanAcctNum_24B, LoanDate_24B, TotalLoanAmt_24B, LoanOutStanding_24B, Intrst_24B As Variant
'Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecInt24B, rowcount, cnt As Long
'Dim strDate As String
'Dim YYYY, MM, DD As String
'
'Set jsonObject = ParseJson(jsonText)
'
''Malli-25/04/2025
'Dim Nodelist24B As Object
'Set Nodelist24B = jsonObject("ITR")("ITR1")
'If Nodelist24B.exists("ScheduleUs24B") Then
''----------------------------
'
'
'Set Nodelist = jsonObject("ITR")("ITR1")("ScheduleUs24B")("ScheduleUs24BDtls")
'
'    LoanfrmBankOrInstitute_24B = Sheet16.Range("LoanfrmBankOrInstitute.24b").Column
''    IFSC_24B = Sheet16.Range("IFSC.24b").Column  'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    BankName_24B = Sheet16.Range("bankName.24b").Column
''    PANof_24B = Sheet16.Range("PAN.24b").Column    'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    LoanAcctNum_24B = Sheet16.Range("loanAccNum.24b").Column
'    LoanDate_24B = Sheet16.Range("loanDate.24b").Column
'    TotalLoanAmt_24B = Sheet16.Range("loanAmt.24b").Column
'    LoanOutStanding_24B = Sheet16.Range("loanOutstanding.24b").Column
'    Intrst_24B = Sheet16.Range("Intrst.24b").Column
'
'
'    TotalExRow = Sheet16.Range("LoanfrmBankOrInstitute.24b").Rows.count
'
'    TotalXMLRow = Nodelist.count
'    TotalDiffRow = TotalXMLRow - TotalExRow
'
'    If (TotalXMLRow > 0) Then
'        If Sheet16.Range("LoanfrmBankOrInstitute.24b").Locked = False Then
'            Sheet16.Range("LoanfrmBankOrInstitute.24b").ClearContents
'        End If
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
''        If Sheet16.Range("IFSC.24b").Locked = False Then
''            Sheet16.Range("IFSC.24b").ClearContents
''        End If
'        If Sheet16.Range("bankName.24b").Locked = False Then
'            Sheet16.Range("bankName.24b").ClearContents
'        End If
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
''        If Sheet16.Range("PAN.24b").Locked = False Then
''            Sheet16.Range("PAN.24b").ClearContents
''        End If
'        If Sheet16.Range("loanAccNum.24b").Locked = False Then
'            Sheet16.Range("loanAccNum.24b").ClearContents
'        End If
'        If Sheet16.Range("loanDate.24b").Locked = False Then
'            Sheet16.Range("loanDate.24b").ClearContents
'        End If
'        If Sheet16.Range("loanAmt.24b").Locked = False Then
'            Sheet16.Range("loanAmt.24b").ClearContents
'        End If
'        If Sheet16.Range("loanOutstanding.24b").Locked = False Then
'            Sheet16.Range("loanOutstanding.24b").ClearContents
'        End If
'        If Sheet16.Range("Intrst.24b").Locked = False Then
'            Sheet16.Range("Intrst.24b").ClearContents
'        End If
'
'    End If
'
'    If (TotalDiffRow > 0) Then
'     AddDiffRows_Int24B (TotalDiffRow)
'    End If
'
'    rowcount = getRowNo(Sheet16.Range("LoanfrmBankOrInstitute.24b").name)
'    rowcount = rowcount - 1
'    cnt = 0
'    For Each node In Nodelist
'        rowcount = rowcount + 1
'        If Sheet16.Cells(rowcount, LoanfrmBankOrInstitute_24B).Locked = False Then
'            Dim LoanTknFrom_24B
'            'Ankita_05/05/2025_Commented as per DESheet_v0.7
'            If node("LoanTknFrom") = "B" Then
'                LoanTknFrom_24B = "Bank"
'            ElseIf node("LoanTknFrom") = "I" Then
'                LoanTknFrom_24B = "Other than bank"
'            End If
'            Sheet16.Cells(rowcount, LoanfrmBankOrInstitute_24B).Value = LoanTknFrom_24B
'        End If
'
''        If Sheet16.Cells(rowcount, IFSC_24B).Locked = False Then
''            Sheet16.Cells(rowcount, IFSC_24B).Value = node("IFSCCode")
''        End If
'
''        If Sheet16.Cells(rowcount, PANof_24B).Locked = False Then
''            Sheet16.Cells(rowcount, PANof_24B).Value = node("PAN")
''        End If
'
'        If Sheet16.Cells(rowcount, BankName_24B).Locked = False Then
'            Sheet16.Cells(rowcount, BankName_24B).Value = node("BankOrInstnName")
'        End If
'
'
'        If Sheet16.Cells(rowcount, LoanAcctNum_24B).Locked = False Then
'            Sheet16.Cells(rowcount, LoanAcctNum_24B).Value = node("LoanAccNoOfBankOrInstnRefNo")
'        End If
'
'                strDate = node("DateofLoan")
'            If strDate <> "" Then
'                YYYY = Mid(strDate, 1, 4)
'                MM = Mid(strDate, 6, 2)
'                DD = Mid(strDate, 9, 2)
'                strDate = DD & "/" & MM & "/" & YYYY
'            End If
'
'        If Sheet16.Cells(rowcount, LoanDate_24B).Locked = False Then
'            Sheet16.Cells(rowcount, LoanDate_24B).Value = strDate
'            strDate = ""
'        End If
'        If Sheet16.Cells(rowcount, TotalLoanAmt_24B).Locked = False Then
'            Sheet16.Cells(rowcount, TotalLoanAmt_24B).Value = node("TotalLoanAmt")
'        End If
'        If Sheet16.Cells(rowcount, LoanOutStanding_24B).Locked = False Then
'            Sheet16.Cells(rowcount, LoanOutStanding_24B).Value = node("LoanOutstndngAmt")
'        End If
'
'        If Sheet16.Cells(rowcount, Intrst_24B).Locked = False Then
'            Sheet16.Cells(rowcount, Intrst_24B).Value = node("InterestUs24B")
'        End If
'        cnt = cnt + 1
'    Next node
'    RecInt24B = cnt
'End If
'End Function

'Malli----------10/04/2025

Sub AddDiffRows_80C(DiffRows As Long)
    'DefinedgridNameRange = "NaturePayment.80C||Identification_Number.80C||Amount.80C"
    DefinedgridNameRange = "Amount.80C||Identification_Number.80C"
    Sheet15.Activate
    'EfilingCommon.searchLastRow ("NaturePayment.80C")
    EfilingCommon.searchLastRow ("Amount.80C")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub

'Konda----------10/04/2025
'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Sub AddDiffRows_80CCC(DiffRows As Long)
'    DefinedgridNameRange = "Name_of_insurer.80CCC||Policy_Document_Number.80CCC||Amount.80CCC"
'    Sheet15.Activate
'    EfilingCommon.searchLastRow ("Name_of_insurer.80CCC")
'    insertRowUnderSectionWithFormula (DiffRows)
'End Sub
'Konda----------10/04/2025
Sub AddDiffRows_80E(DiffRows As Long)
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    DefinedgridNameRange = "LoanfrmBankOrInstitute.80E||bankName.80E||loanAccNum.80E||loanDate.80E||loanAmt.80E||loanOutstanding.80E||Intrst.80E"
    DefinedgridNameRange = "LoanfrmBankOrInstitute.80E||bankName.80E||loanAccNum.80E||loanDate.80E||loanAmt.80E||loanOutstanding.80E||Intrst.80E"
Sheet17.Activate
    EfilingCommon.searchLastRow ("LoanfrmBankOrInstitute.80E")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Konda----------10/04/2025
Sub AddDiffRows_80EE(DiffRows As Long)
'Ankita_05/05/2025_Commented as per DESheet_v0.7
    'DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||IFSC.80EE||bankName.80EE||PAN.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE"
    DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||bankName.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE||Combination_80EE"
    Sheet17.Activate
    EfilingCommon.searchLastRow ("LoanfrmBankOrInstitute.80EE")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Konda----------10/04/2025
Sub AddDiffRows_80EEA(DiffRows As Long)
'Ankita_05/05/2025_Commented as per DESheet_v0.7
    'DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||IFSC.80EEA||bankName.80EEA||PAN.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA"
    DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||bankName.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA||Combination_80EEA"
    Sheet17.Activate
    EfilingCommon.searchLastRow ("LoanfrmBankOrInstitute.80EEA")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Konda----------10/04/2025
Sub AddDiffRows_80EEB(DiffRows As Long)
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||IFSC.80EEB||bankName.80EEB||PAN.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEB||loanOutstanding.80EEB||Vehicle_value.80EEB||VehicleRegNum.80EEB||Intrst.80EEB"
    DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||bankName.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEB||loanOutstanding.80EEB||VehicleRegNum.80EEB||Intrst.80EEB"
    Sheet17.Activate
    EfilingCommon.searchLastRow ("LoanfrmBankOrInstitute.80EEB")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub

'Konda----------10/04/2025

'Konda----------29/01/2026
'Sub AddDiffRows_Int24B(DiffRows As Long)
''    DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||IFSC.24b||bankName.24b||PAN.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b"
'    'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b"
'    Sheet16.Activate
'    EfilingCommon.searchLastRow ("LoanfrmBankOrInstitute.24b")
'    insertRowUnderSectionWithFormula (DiffRows)
'End Sub
 'AddDiffRows_Sec80DSelfFamHIDtls
'Konda----------17/04/2024


Sub AddDiffRows_Sec80DSelfFamHIDtls(DiffRows As Long)
    DefinedgridNameRange = "NameInsurerA1.80D||PolicyNumA1.80D||AmtA1.80D" 'Ankita_06/05/2025_Commented as per DESheet_v0.7

    Sheet9.Activate
    EfilingCommon.searchLastRow ("NameInsurerA1.80D")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Sec80DSelfFamSrCtznHIDtls
'Konda----------17/04/2024
Sub AddDiffRows_Sec80DSelfFamSrCtznHIDtls(DiffRows As Long)
    DefinedgridNameRange = "NameInsurerB1.80D||PolicyNumB1.80D||AmtB1.80D"  'Ankita_06/05/2025_Commented as per DESheet_v0.7

    Sheet9.Activate
    EfilingCommon.searchLastRow ("NameInsurerB1.80D")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Sec80DParentsHIDtls
'Konda----------17/04/2024
Sub AddDiffRows_Sec80DParentsHIDtls(DiffRows As Long)
    DefinedgridNameRange = "NameInsurerA2.80D||PolicyNumA2.80D||AmtA2.80D" 'Ankita_06/05/2025_Commented as per DESheet_v0.7

    Sheet9.Activate
    EfilingCommon.searchLastRow ("NameInsurerA2.80D")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Sec80DParentsSrCtznHIDtls
'Konda----------17/04/2024
Sub AddDiffRows_Sec80DParentsSrCtznHIDtls(DiffRows As Long)
    DefinedgridNameRange = "NameInsurerB2.80D||PolicyNumB2.80D||AmtB2.80D"  'Ankita_06/05/2025_Commented as per DESheet_v0.7

    Sheet9.Activate
    EfilingCommon.searchLastRow ("NameInsurerB2.80D")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
'Konda_EA10_13A__AY_2025_26-------06-05-2025
Function ImportScheduleEA10_13A(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim Nodelist As Object

Set jsonObject = ParseJson(jsonText)
 
Set Nodelist = jsonObject("ITR")("ITR1")
 
If Nodelist.exists("ScheduleEA10_13A") Then
 
Dim Placeofwork_EA10

Placeofwork_EA10 = jsonObject("ITR")("ITR1")("ScheduleEA10_13A")("Placeofwork")

If Placeofwork_EA10 <> "" Then


If Trim(Placeofwork_EA10) = "1" Then
   Sheet18.Range("Sch10of13A_PlaceofWrk").Value = Trim("1. Metro")
   
   ElseIf Trim(Placeofwork_EA10) = "2" Then
   Sheet18.Range("Sch10of13A_PlaceofWrk").Value = Trim("2. Non-Metro")
   
   Else: Sheet18.Range("Sch10of13A_PlaceofWrk").Value = "(Select)"
   
End If
End If


Dim ActlHRARecv_EA10

ActlHRARecv_EA10 = Trim(jsonObject("ITR")("ITR1")("ScheduleEA10_13A")("ActlHRARecv"))


If Trim(ActlHRARecv_EA10) <> "" Then

Sheet18.Range("Sch10of13A_ActlHRArecivedA").Value = UCase(ActlHRARecv_EA10)

End If


Dim ActlRentPaid_EA10

ActlRentPaid_EA10 = Trim(jsonObject("ITR")("ITR1")("ScheduleEA10_13A")("ActlRentPaid"))


If Trim(ActlRentPaid_EA10) <> "" Then

Sheet18.Range("Sch10of13A_ActlRentpaid").Value = UCase(ActlRentPaid_EA10)

End If


Dim BasicSalary_EA10

BasicSalary_EA10 = Trim(jsonObject("ITR")("ITR1")("ScheduleEA10_13A")("BasicSalary"))

If Sheet18.Range("Sch10of13A_BasicSalary").Locked = False Then
Sheet18.Range("Sch10of13A_BasicSalary").Value = UCase(BasicSalary_EA10)

End If

'Konda----------08/05/2025
Dim DearnessAllwnc_EA10

DearnessAllwnc_EA10 = Trim(jsonObject("ITR")("ITR1")("ScheduleEA10_13A")("DearnessAllwnc"))

If Sheet18.Range("Sch10of13A_DearAllowance").Locked = False Then
Sheet18.Range("Sch10of13A_DearAllowance").Value = DearnessAllwnc_EA10
End If
'--------------------------
'Konda---------as per Clarification tracker-14/05/2025

Dim Sal40Or50Per_EA10

Sal40Or50Per_EA10 = Trim(jsonObject("ITR")("ITR1")("ScheduleEA10_13A")("Sal40Or50Per"))

If Sheet18.Range("Sch10of13A_50Por40Pofsalary").Locked = False Then
Sheet18.Range("Sch10of13A_50Por40Pofsalary").Value = Sal40Or50Per_EA10
End If
'-------------------------


End If
End Function
'Konda updated on 22-01-2026
Function ImportScheduleHP(ByVal jsonObject As Object)
On Error Resume Next
    Dim XpathOfHP As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, nodeList1, Nodelist2, init As Object
    Dim node, Node1, Node2 As Object
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim iCountry, sCountry As Variant
    Dim iLetOut, sLetOut As Variant
    Dim iTotalCoRow, sTotalCoRow, TotalDiffCoRow As Variant
    Dim iTotalTenRow, sTotalTenRow, TotalDiffTenRow As Variant

    Dim rowcount, cnt As Long
    Sheet19.Activate
Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")
If init.exists("PropertyDetails") Then

    Set Nodelist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("PropertyDetails")

    TotalExRow = Sheet19.Range("PropertySectionCOunt").Value
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    

    Dim i As Long
'    For i = 1 To TotalDiffRow
'        Sheet19.Activate
'        'AddBlockCall_hprptfrm
'    Next

    rowcount = 0
    cnt = 0

    
    For Each node In Nodelist
        rowcount = rowcount + 1
        'Co-Owner
        iTotalCoRow = Sheet19.Range("HP.Co.Name" & rowcount).Rows.count
         
        sTotalCoRow = node("CoOwners").count
        TotalDiffCoRow = WorksheetFunction.Max((sTotalCoRow - iTotalCoRow), 0)
        If TotalDiffCoRow > 0 Then
            Sheet19.Activate
            AddPropertyCoOWners (TotalDiffCoRow)
        End If
        
        'Tenants
        iTotalTenRow = Sheet19.Range("HP.NameofTenant" & rowcount).Rows.count
        sTotalTenRow = node("TenantDetails").count
        TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)
        If TotalDiffTenRow > 0 Then
            Sheet19.Activate
            AddPropertyTenant (TotalDiffTenRow)
        End If
            
            If Sheet19.Range("HP.AddrDetail" & rowcount).Locked = False Then
                Sheet19.Range("HP.AddrDetail" & rowcount).Value = UCase(node("AddressDetailWithZipCode")("AddrDetail"))
            End If
            If Sheet19.Range("HP.CityOrTownOrDistrict" & rowcount).Locked = False Then
                Sheet19.Range("HP.CityOrTownOrDistrict" & rowcount).Value = UCase(node("AddressDetailWithZipCode")("CityOrTownOrDistrict"))
            End If
            
            iState = UCase(node("AddressDetailWithZipCode")("StateCode"))
            
'Konda updated on 06-02-2026--SIT-109675
'            sState = Findtext(iState, "State")
             sState = Findtext(iState, "StateList")
            
            Sheet19.Range("HP.StateCode" & rowcount).Value = sState
            iCountry = UCase(node("AddressDetailWithZipCode")("CountryCode"))
            
'Konda updated on 06-02-2026--SIT-109675
'            sCountry = Findtext(iCountry, "Country")
             sCountry = Findtext(iCountry, "CountList")
        '===============================
            If Sheet19.Range("HP.CountryCode" & rowcount).Locked = False Then
                Sheet19.Range("HP.CountryCode" & rowcount).Value = sCountry
            End If
            If Sheet19.Range("HP.PinCode" & rowcount).Locked = False Then
                Sheet19.Range("HP.PinCode" & rowcount).Value = UCase(node("AddressDetailWithZipCode")("PinCode"))
            End If
            If Sheet19.Range("HP.ZipCode" & rowcount).Locked = False Then
                Sheet19.Range("HP.ZipCode" & rowcount).Value = UCase(node("AddressDetailWithZipCode")("ZipCode"))
            End If
            
            Dim OwnerProperty_HP
            
            OwnerProperty_HP = UCase(node("PropertyOwner"))
            
            If OwnerProperty_HP = "SE" Then
               OwnerProperty_HP = "Self"
            ElseIf OwnerProperty_HP = "MI" Then
               OwnerProperty_HP = "Minor"
            ElseIf OwnerProperty_HP = "SP" Then
               OwnerProperty_HP = "Spouse"
            ElseIf OwnerProperty_HP = "OT" Then
               OwnerProperty_HP = "Others"
            Else
                
            End If
            
            If Sheet19.Range("HP.OwnerProperty" & rowcount).Locked = False Then
                Sheet19.Range("HP.OwnerProperty" & rowcount).Value = OwnerProperty_HP
            End If
            Application.EnableEvents = True
            
            If Sheet19.Range("HP.OwnerPropertyDescription" & rowcount).Locked = False Then
                Sheet19.Range("HP.OwnerPropertyDescription" & rowcount).Value = node("PropertyOwnerOther")
            End If
            
            Dim Co_Ownedflag As Variant
            
            Co_Ownedflag = UCase(node("PropCoOwnedFlg"))
            If UCase(Co_Ownedflag) = "YES" Or UCase(Mid(Co_Ownedflag, 1, 1) = "Y") Then
            Co_Ownedflag = "Yes"
            ElseIf UCase(Co_Ownedflag) = "NO" Or UCase(Mid(Co_Ownedflag, 1, 1) = "N") Then
            Co_Ownedflag = "No"
            Else
            Co_Ownedflag = "(Select)"
            End If
            
            Sheet19.Range("HP.CoOwnedYN" & rowcount).Value = Co_Ownedflag
             
            Dim hpshare As Variant
            hpshare = UCase(node("AsseseeShareProperty"))
            
            If Sheet19.Range("HP.SharePercent" & rowcount).Locked = False Then
                Sheet19.Range("HP.SharePercent" & rowcount).Value = hpshare
            End If
            
            If node.exists("CoOwners") Then
       
            Set nodeList1 = Nothing
            Set nodeList1 = node("CoOwners")
            cnt = 0
            cnt = getRowNo(Sheet19.Range("HP.Co.Name" & rowcount).name)
            cnt = cnt - 1
            
            For Each Node1 In nodeList1
                    cnt = cnt + 1

                If Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Name" & rowcount).Column).Locked = False Then
                    Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Name" & rowcount).Column).Value = UCase(Node1("NameCoOwner"))
                End If
                If Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Pan" & rowcount).Column).Locked = False Then
                    Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Pan" & rowcount).Column).Value = UCase(Node1("PAN_CoOwner"))
                End If
                If Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Aadhaar" & rowcount).Column).Locked = False Then
                    Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Aadhaar" & rowcount).Column).Value = UCase(Node1("Aadhaar_CoOwner"))
                End If
                If Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Share" & rowcount).Column).Locked = False Then
                    Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Share" & rowcount).Column).Value = UCase(Node1("PercentShareProperty"))
                End If
                    
            Next Node1
            End If
            
            iLetOut = UCase(node("ifLetOut"))
            
            If iLetOut = "L" Then
                sLetOut = "Let Out"
            ElseIf iLetOut = "D" Then
                sLetOut = "Deemed Let Out"
            ElseIf iLetOut = "S" Then
                sLetOut = "Self Occupied"
            Else
                sLetOut = "(Select)"
            End If
            
            Sheet19.Range("HP.ifLetOut" & rowcount).Value = sLetOut
            Application.EnableEvents = True
            If node.exists("TenantDetails") Then
            If Not IsEmpty(node("TenantDetails")) Then
                Set Nodelist2 = node("TenantDetails")
                cnt = 0
                cnt = getRowNo(Sheet19.Range("HP.NameofTenant" & rowcount).name)
                cnt = cnt - 1
                For Each Node2 In Nodelist2
                    cnt = cnt + 1
   
                    If Sheet19.Cells(cnt, Sheet19.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                        Sheet19.Cells(cnt, Sheet19.Range("HP.NameofTenant" & rowcount).Column).Value = UCase(Node2("NameofTenant"))
                    End If
                    If Sheet19.Cells(cnt, Sheet19.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                        Sheet19.Cells(cnt, Sheet19.Range("HP.PANofTenant" & rowcount).Column).Value = UCase(Node2("PANofTenant"))
                    End If
                    If Sheet19.Cells(cnt, Sheet19.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                        Sheet19.Cells(cnt, Sheet19.Range("HP.TANofTenant" & rowcount).Column).Value = UCase(Node2("PANTANofTenant"))
                    End If
                        If Node2.Exsist("AadhaarofTenant") Then
                            Sheet19.Cells(cnt, Sheet19.Range("HP.AadharofTenant" & rowcount).Column).Value = UCase(Node2("AadhaarofTenant"))
                        End If
                Next Node2
            End If
            End If
            
                If Sheet19.Range("HP.AnnualLetableValue" & rowcount).Locked = False Then
                    Sheet19.Range("HP.AnnualLetableValue" & rowcount).Value = UCase(node("Rentdetails")("AnnualLetableValue"))
                End If
                
                If Sheet19.Range("HP.RentNotRealized" & rowcount).Locked = False Then
                    Sheet19.Range("HP.RentNotRealized" & rowcount).Value = UCase(node("Rentdetails")("RentNotRealized"))
                End If
                If Sheet19.Range("HP.LocalTaxes" & rowcount).Locked = False Then
                    Sheet19.Range("HP.LocalTaxes" & rowcount).Value = UCase(node("Rentdetails")("LocalTaxes"))
                End If
                
            
 Dim Nodelist3, Node3
If node("Rentdetails").exists("Section24B") Then

       
        Dim iTotalCoRow_24b, sTotalCoRow_24b, TotalDiffCoRow_24b
        iTotalCoRow_24b = Sheet19.Range("LoanfrmBankOrInstitute.24b" & rowcount).Rows.count
         
      
        sTotalCoRow_24b = node("Rentdetails")("Section24B")("Section24BDtls").count
    
        TotalDiffCoRow_24b = WorksheetFunction.Max((sTotalCoRow_24b - iTotalCoRow_24b), 0)
        If TotalDiffCoRow_24b > 0 Then
            Sheet19.Activate
            AddSection24b (TotalDiffCoRow_24b)
        End If
        
 Set Nodelist3 = node("Rentdetails")("Section24B")("Section24BDtls")
                 cnt = 0
                cnt = getRowNo(Sheet19.Range("LoanfrmBankOrInstitute.24b" & rowcount).name)
                cnt = cnt - 1
                For Each Node3 In Nodelist3
                    cnt = cnt + 1
                If Node3("LoanTknFrom") <> "" And Sheet19.Cells(cnt, Sheet19.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).Locked = False Then
          
                    Dim LoanTknFrom_24B As Variant
                    LoanTknFrom_24B = UCase(Node3("LoanTknFrom"))
                    If LoanTknFrom_24B = UCase("B") Then
                     Sheet19.Cells(cnt, Sheet19.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).Value = "Bank "
                    
              
                    ElseIf LoanTknFrom_24B = UCase("I") Then
               
                     Sheet19.Cells(cnt, Sheet19.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).Value = "Other than Bank"
                    End If
                End If
                
                If Node3("BankOrInstnName") <> "" And Sheet19.Cells(cnt, Sheet19.Range("bankName.24b" & rowcount).Column).Locked = False Then
                    Sheet19.Cells(cnt, Sheet19.Range("bankName.24b" & rowcount).Column).Value = UCase(Node3("BankOrInstnName"))
                End If
                
                If Node3("LoanAccNoOfBankOrInstnRefNo") <> "" And Sheet19.Cells(cnt, Sheet19.Range("loanAccNum.24b" & rowcount).Column).Locked = False Then
                    Sheet19.Cells(cnt, Sheet19.Range("loanAccNum.24b" & rowcount).Column).Value = UCase(Node3("LoanAccNoOfBankOrInstnRefNo"))
                End If
                
                If Node3("DateofLoan") <> "" And Sheet19.Cells(cnt, Sheet19.Range("loanDate.24b" & rowcount).Column).Locked = False Then
                    
                    Sheet19.Cells(cnt, Sheet19.Range("loanDate.24b" & rowcount).Column).Value = Mid(Node3("DateofLoan"), 9, 2) & "/" & Mid(Node3("DateofLoan"), 6, 2) & "/" & Mid(Node3("DateofLoan"), 1, 4)
                End If
                
                If Node3("TotalLoanAmt") <> "" And Sheet19.Cells(cnt, Sheet19.Range("loanAmt.24b" & rowcount).Column).Locked = False Then
                    Sheet19.Cells(cnt, Sheet19.Range("loanAmt.24b" & rowcount).Column).Value = UCase(Node3("TotalLoanAmt"))
                End If
                
                If Node3("LoanOutstndngAmt") <> "" And Sheet19.Cells(cnt, Sheet19.Range("loanOutstanding.24b" & rowcount).Column).Locked = False Then
                    Sheet19.Cells(cnt, Sheet19.Range("loanOutstanding.24b" & rowcount).Column).Value = UCase(Node3("LoanOutstndngAmt"))
                End If
                
                If Node3("InterestUs24B") <> "" And Sheet19.Cells(cnt, Sheet19.Range("Intrst.24b" & rowcount).Column).Locked = False Then
                    Sheet19.Cells(cnt, Sheet19.Range("Intrst.24b" & rowcount).Column).Value = UCase(Node3("InterestUs24B"))
                End If
                
                 
                     
                Next Node3
             
End If
            

                
                
                If Sheet19.Range("HP.RentOfEarlierYrSec_AandAA" & rowcount).Locked = False Then
                    Sheet19.Range("HP.RentOfEarlierYrSec_AandAA" & rowcount).Value = UCase(node("Rentdetails")("ArrearsUnrealizedRentRcvd"))
                End If
            cnt = cnt + 1
        Next node
End If
    Set Nodelist = jsonObject("ScheduleHP")
    rowcount = 0
        
    If ActiveWorkbook.Sheets("House Property").Visible = xlSheetVisible Then
    
        For Each node In Nodelist
        rowcount = rowcount + 1
        If Sheet19.Range("HP.PassTroughIncome").Locked = False Then
            Sheet19.Range("HP.PassTroughIncome").Value = UCase(node("PassThroghIncome"))
        End If
        Next node
    End If

End Function

'added by Konda for AY 2026-27
'start--
Sub AddSection24b(Optional iRows As Long = 0)
On Error GoTo endline
    Dim newrngname As Variant
    Dim numberofrows As Long
    Dim newfrmsize, i, te As Long


    Application.EnableEvents = False
    Sheets("HP").Activate

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    numberofrows = iRows
    For te = 1 To Sheet19.Range("PropertySectionCOunt").Value


        EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b"
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


'Malli_AY_2026_27 17/04/2026
Sub AddDiffRows_PRANNum(DiffRows As Long)
    'setDiffTblinfo_80CCCTable_1
    DefinedgridNameRange = "Pran_Sl||pran_new"
    Sheet1.Activate
    searchLastRow1 ("Pran_Sl")
    insertRowUnderSectionWithFormula (DiffRows)
    
     
End Sub





























