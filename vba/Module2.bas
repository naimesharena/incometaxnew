Attribute VB_Name = "Module2"
Option Explicit

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

fmsgbox "Import personal/tax details from downloaded Pre-filled JSON or Import from already generated JSON of the current assessment year."

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

ImportPersonalInfo (jsonText)
ImportFilingStatus (jsonText)
ImportITR1_IncomeDeductions (jsonText)
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
Else
    Sheet4.Visible = False
    Sheet12.Visible = False
    Sheet9.Visible = False
End If

End Sub
Function ImportPersonalInfo(jsonText As String)
 On Error Resume Next
Dim jsonObject As Object
Dim jsonDictionary As Object
Dim PAN, firstName, middleName, LastName, residenceNo, residenceName, roadOrStreet, localityOrArea, cityOrTownOrDistrict, StateCode, CountryCode, PinCode, zipCode, countryCodeMobile, mobileNo, emailAddress, dob, iEmpCat, sEmpCat, aadhaarCardNo, AadhaarEnrolmentIds As Variant
Dim sCountry, iCountry As Variant
Dim sState, iState As Variant
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
dob = jsonObject("ITR")("ITR1")("PersonalInfo")("DOB")
    YYYY = Mid(dob, 1, 4)
    MM = Mid(dob, 6, 2)
    DD = Mid(dob, 9, 2)
    strDate = DD & "/" & MM & "/" & YYYY
iEmpCat = jsonObject("ITR")("ITR1")("PersonalInfo")("EmployerCategory")
aadhaarCardNo = jsonObject("ITR")("ITR1")("PersonalInfo")("AadhaarCardNo")
'Enhancement
'AadhaarEnrolmentIds = jsonObject("ITR")("ITR1")("PersonalInfo")("AadhaarEnrolmentId")

If iEmpCat = "CGOV" Then
    sEmpCat = "Central Government"
ElseIf iEmpCat = "SGOV" Then
    sEmpCat = "State Government"
ElseIf iEmpCat = "PSU" Then
    sEmpCat = "Public Sector Undertaking"
ElseIf iEmpCat = "OTH" Then
    sEmpCat = "Others"
ElseIf iEmpCat = "PE" Then
    sEmpCat = "Pensioners"
ElseIf sEmpCat = "NA" Then
    sEmpCat = "Not Applicable (eg. Family pension etc)"
End If

    If firstName <> "" Then
        Sheet1.Range("sheet1.FirstName").Value = firstName
    End If
    If middleName <> "" Then
        Sheet1.Range("sheet1.MiddleName").Value = middleName
    End If
    If LastName <> "" Then
        Sheet1.Range("sheet1.SurNameOrOrgName").Value = LastName
    End If
    If PAN <> "" Then
        Sheet1.Range("sheet1.PAN").Value = PAN
    End If
    If aadhaarCardNo <> "" Then
        Sheet1.Range("Sheet1.Aadhaar").Value = aadhaarCardNo
    End If
'    Enhancement
'    If CStr(AadhaarEnrolmentIds) <> "" Then
'        Sheet1.Range("Sheet1.AadhaarEnrol").Value = CStr(AadhaarEnrolmentIds)
'    End If
    If strDate <> "" Then
        Sheet1.Range("sheet1.DOB").Value = strDate
    End If
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
    If sCountry <> "" Then
        Sheet1.Range("sheet1.Country").Value = sCountry
    End If
    If sState <> "" Then
        Sheet1.Range("sheet1.StateCode1").Value = sState
    End If
    If PinCode <> "" Then
        Sheet1.Range("sheet1.PinCode").Value = PinCode
    End If
    If zipCode <> "" Then
        Sheet1.Range("sheet1.ZipCode").Value = zipCode
    End If
    If sEmpCat <> "" Then
        Sheet1.Range("sheet1.EmployerCategory1").Value = sEmpCat
    End If
    If emailAddress <> "" Then
        Sheet1.Range("sheet1.EmailAddress").Value = emailAddress
    End If
    If countryCodeMobile <> "" Then
        Sheet1.Range("sheet1.MobileCountryCode").Value = countryCodeMobile
    End If
    If mobileNo <> "" Then
        Sheet1.Range("sheet1.Mobileno").Value = mobileNo
    End If
End Function
Function ImportFilingStatus(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim ReturnFileSec, SeventhProvisio139, NewTaxRegime, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec As Variant
Dim sReturnFile, iReturnFile As Variant
Dim iProvisoFlag, sProvisoFlag As Variant
Dim iDepositAmountFlag, sDepositAmountFlag As Variant
Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
Dim DateofOriginalfile, DateofOriginalfile1, NoticeDateussec, Filingtype As Variant
Dim YYYY, MM, DD, strDate As String

Set jsonObject = ParseJson(jsonText)

iReturnFile = jsonObject("ITR")("ITR1")("FilingStatus")("ReturnFileSec")
    If iReturnFile = "11" Then
       Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date"
    ElseIf iReturnFile = "12" Then
       Range("sheet1.ReturnFileSec").Value = "139(4)-Belated"
    ElseIf iReturnFile = "13" Then
       Range("sheet1.ReturnFileSec").Value = "142(1)"
    ElseIf iReturnFile = "14" Then
       Range("sheet1.ReturnFileSec").Value = "148"
    ElseIf iReturnFile = "15" Then
       Range("sheet1.ReturnFileSec").Value = "153A"
    ElseIf iReturnFile = "16" Then
       Range("sheet1.ReturnFileSec").Value = "153C"
    ElseIf iReturnFile = "17" Then
       Range("sheet1.ReturnFileSec").Value = "139(5)-Revised"
    ElseIf iReturnFile = "18" Then
       Range("sheet1.ReturnFileSec").Value = "139(9)"
    ElseIf iReturnFile = "20" Then
       Range("sheet1.ReturnFileSec").Value = "139(9A) - After condonation of delay u/s 119(2)(b)"
End If

iProvisoFlag = jsonObject("ITR")("ITR1")("FilingStatus")("SeventhProvisio139")
    If iProvisoFlag = "Y" Then
    sProvisoFlag = "Yes"
    ElseIf iProvisoFlag = "N" Then
    sProvisoFlag = "No"
End If
Sheet1.Range("sheet1.SeventhProvisoFlag").Value = sProvisoFlag

If jsonObject("ITR")("ITR1")("FilingStatus")("NewTaxRegime") = 1 Then
    ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 1
   Sheet5.Unprotect Password:=sPassword
   Sheet5.Range("BacValue").Value = 1
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
  resetBacYes
  Sheet1.Protect Password:=sPassword
 
  Sheet1.Activate
  ThisWorkbook.Protect Password:=sPassword
ElseIf jsonObject("ITR")("ITR1")("FilingStatus")("NewTaxRegime") = 2 Then
 Sheet5.Unprotect Password:=sPassword
   Sheet5.Range("BacValue").Value = 2
    ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 1
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
  resetBacNo
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
'Change-22.11.2022.102.16H
'Sheet1.Range("Sheet1.DepositAmountFlag").Value = sDepositAmountFlag
'
'If sDepositAmountFlag = "Yes" Then
'    Sheet1.Range("Sheet1.DepositAmount").Value = jsonObject("ITR")("ITR1")("FilingStatus")("AmtSeventhProvisio139i")
'End If
'---end

iAggrigateAmountFlag = jsonObject("ITR")("ITR1")("FilingStatus")("IncrExpAggAmt2LkTrvFrgnCntryFlg")
If iAggrigateAmountFlag = "Y" Then
sAggrigateAmountFlag = "Yes"
ElseIf iAggrigateAmountFlag = "N" Then
sAggrigateAmountFlag = "No"
End If
Sheet1.Range("Sheet1.AggrigateAmountFlag").Value = sAggrigateAmountFlag
If sAggrigateAmountFlag = "Yes" Then
    Sheet1.Range("Sheet1.AggrigateAmount").Value = jsonObject("ITR")("ITR1")("FilingStatus")("AmtSeventhProvisio139ii")
End If

iAggrigateAmountFlag1 = jsonObject("ITR")("ITR1")("FilingStatus")("IncrExpAggAmt1LkElctrctyPrYrFlg")
If iAggrigateAmountFlag1 = "Y" Then
    sAggrigateAmountFlag1 = "Yes"
ElseIf iAggrigateAmountFlag1 = "N" Then
    sAggrigateAmountFlag1 = "No"
End If
Sheet1.Range("Sheet1.AggrigateAmountFlag1").Value = sAggrigateAmountFlag1

If sAggrigateAmountFlag1 = "Yes" Then
    Sheet1.Range("Sheet1.AggrigateAmount1").Value = jsonObject("ITR")("ITR1")("FilingStatus")("AmtSeventhProvisio139iii")
End If

ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1")
ReturnFileSec = Mid(ReturnFileSec, 1, 2)
If ReturnFileSec = "17" Or ReturnFileSec = "18" Then
    Sheet1.Range("Sheet1.ReceiptNo").Value = jsonObject("ITR")("ITR1")("FilingStatus")("ReceiptNo")
    DateofOriginalfile = jsonObject("ITR")("ITR1")("FilingStatus")("OrigRetFiledDate")
        YYYY = Mid(DateofOriginalfile, 1, 4)
        MM = Mid(DateofOriginalfile, 6, 2)
        DD = Mid(DateofOriginalfile, 9, 2)
        strDate = DD & "/" & MM & "/" & YYYY
    Sheet1.Range("Sheet1.OrigRetFiledDate").Value = strDate
Else
    If jsonObject("ITR")("ITR1")("FilingStatus")("ReceiptNo") <> "" Then
        Sheet1.Range("Sheet1.ReceiptNo").Value = jsonObject("ITR")("ITR1")("FilingStatus")("ReceiptNo")
    End If
    If jsonObject("ITR")("ITR1")("FilingStatus")("OrigRetFiledDate") <> "" Then
        DateofOriginalfile = jsonObject("ITR")("ITR1")("FilingStatus")("OrigRetFiledDate")
            YYYY = Mid(DateofOriginalfile, 1, 4)
            MM = Mid(DateofOriginalfile, 6, 2)
            DD = Mid(DateofOriginalfile, 9, 2)
            strDate = DD & "/" & MM & "/" & YYYY
        Sheet1.Range("Sheet1.OrigRetFiledDate").Value = strDate
    End If
    If jsonObject("ITR")("ITR1")("FilingStatus")("NoticeNo") <> "" Then
        Sheet1.Range("sheet1.NoticeNo").Value = jsonObject("ITR")("ITR1")("FilingStatus")("NoticeNo")
    End If
    If jsonObject("ITR")("ITR1")("FilingStatus")("NoticeDateUnderSec") <> "" Then
        NoticeDateussec = jsonObject("ITR")("ITR1")("FilingStatus")("NoticeDateUnderSec")
            YYYY = Mid(NoticeDateussec, 1, 4)
            MM = Mid(NoticeDateussec, 6, 2)
            DD = Mid(NoticeDateussec, 9, 2)
            strDate = DD & "/" & MM & "/" & YYYY
        Sheet1.Range("sheet1.NoticeDate").Value = strDate
    End If
End If

End Function
Function ImportITR1_IncomeDeductions(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary, node, Nodelist As Object
Set jsonObject = ParseJson(jsonText)
Dim NatureColNo, DescriptionColNo, AmtColNo, TotalExRow, TotalDiffRow, TotalXMLRow, rowcount, cnt, RecTDS1 As Long
Dim SalNatureDesc, DeductionUs57iia, IncomeFromSal, TotalIncomeOfHP, IncomeOthSrc, GrossTotIncome, Section80C, Section80CCC, Section80CCDEmployeeOrSE, Section80CCD1B, Section80CCDEmployer, Section80D, Section80DD, Section80DDB As Variant
Dim InterestPayable, Section80E, Section80EE, Section80G, Section80GG, Section80GGA, Section80GGC, Section80U, Section80TTA, Section80TTB, Section80DDUsrType, Section80DDBUsrType, Section80EEA, Section80EEB, Section80UUsrType As Variant
Dim ArrearsUnrealizedRentRcvd, TotalIncome, GrossSalary, Salary, PerquisitesValue, ProfitsInSalary, DeductionUs16, DeductionUs16ia, EntertainmentAlw16ii, ProfessionalTaxUs16iii, TypeOfHP, GrossRentReceived, TaxPaidlocalAuth As Variant
Dim test As Variant
Dim init As Variant

GrossSalary = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("GrossSalary")
Salary = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("Salary")
PerquisitesValue = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("PerquisitesValue")
ProfitsInSalary = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("ProfitsInSalary")
DeductionUs16 = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("DeductionUs16")
DeductionUs16ia = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("DeductionUs16ia")
EntertainmentAlw16ii = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("EntertainmentAlw16ii")
ProfessionalTaxUs16iii = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("ProfessionalTaxUs16iii")
TypeOfHP = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("TypeOfHP")
GrossRentReceived = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("GrossRentReceived")
TaxPaidlocalAuth = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("TaxPaidlocalAuth")
InterestPayable = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("InterestPayable")
ArrearsUnrealizedRentRcvd = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("ArrearsUnrealizedRentRcvd")
DeductionUs57iia = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("DeductionUs57iia")
IncomeOthSrc = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("IncomeOthSrc")
Section80C = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80C")
Section80CCC = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCC")
Section80CCDEmployeeOrSE = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCDEmployeeOrSE")
Section80CCD1B = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCD1B")
Section80CCDEmployer = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80CCDEmployer")
Section80D = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80D")
Section80DD = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80DD")
Section80DDB = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80DDB")
Section80E = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80E")
Section80EE = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80EE")
Section80G = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80G")
Section80GG = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80GG")
Section80GGA = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80GGA")
Section80GGC = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80GGC")
Section80U = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80U")
Section80TTA = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80TTA")
Section80TTB = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80TTB")
Section80DDUsrType = Findtext(jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80DDUsrType"), "Selection80DD")
Section80DDBUsrType = Findtext(jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80DDBUsrType"), "Selection80DDB")
Section80EEA = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80EEA")
Section80EEB = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80EEB")
Section80UUsrType = Findtext(jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("UsrDeductUndChapVIA")("Section80UUsrType"), "Selection80U")
    
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.IncomeFromSal").Value = GrossSalary
    Sheet1.Range("IncD.Allowances").Value = Salary
    Sheet1.Range("IncD.Perquisites").Value = PerquisitesValue
    Sheet1.Range("IncD.Profits").Value = ProfitsInSalary
    
   ' Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.GrossRentRecieved").Value = GrossRentReceived
    'Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.TaxPaidLocalAuthorities").Value = TaxPaidlocalAuth
    Sheet1.Range("IncD.InterestBorrowedCapital").Value = InterestPayable
    Sheet1.Range("IncD.Arrears").Value = ArrearsUnrealizedRentRcvd
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.IncomeFromOS").Value = IncomeOthSrc
    Sheet1.Range("IncD.Section80C").Value = Section80C
    Sheet1.Range("IncD.Section80CCC").Value = Section80CCC
    Sheet1.Range("IncD.Section80CCD_SE").Value = Section80CCDEmployeeOrSE
    Sheet1.Range("IncD.Section80CCD1B_SE").Value = Section80CCD1B
    Sheet1.Range("IncD.Section80CCD").Value = Section80CCDEmployer
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Section80DValue").Value = Section80D
    'Sheet1.Range("SELECT80DD").Value = Section80DDUsrType
    Sheet1.Range("IncD.Section80DD").Value = Section80DD
    Sheet1.Range("SELECT80DDS").Value = Section80DDBUsrType
    Sheet1.Range("IncD.Section80DDB").Value = Section80DDB
    Sheet1.Range("IncD.Section80E").Value = Section80E
    Sheet1.Range("IncD.Section80EE").Value = Section80EE
   ' Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80EEA").Value = Section80EEA
    Sheet1.Range("IncD.Section80EEB").Value = Section80EEB
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Section80G").Value = Section80G
    Sheet1.Range("IncD.Section80GG").Value = Section80GG
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Section80GGA").Value = Section80GGA
    Sheet1.Range("IncD.Section80TTA").Value = Section80TTA
   ' Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Section80TTB").Value = Section80TTB
    'Sheet1.Range("IncD.Section80GGC").Value = Section80GGC
   ' Sheet1.Range("SELECT80U").Value = Section80UUsrType
    Sheet1.Range("IncD.Section80U").Value = Section80U
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
    Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")
    If init.exists("AllwncExemptUs10") Then
        Set Nodelist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("AllwncExemptUs10")("AllwncExemptUs10Dtls")
    
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
         AddDiffRows_Exempt2 (TotalDiffRow)
        End If
    
       rowcount = getRowNo(Sheet1.Range("Others.NOI_1").name)
       rowcount = rowcount - 1
       cnt = 0
    
        For Each node In Nodelist
       ' Sheet1.Unprotect Password:=getmsgstate
            rowcount = rowcount + 1
                 
                If UCase(node("SalNatureDesc")) = "10(10B)(i)" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette"
                ElseIf UCase(node("SalNatureDesc")) = "10(10B)(ii)" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government"
                ElseIf UCase(node("SalNatureDesc")) = "OTH" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
                Else
                    test = Findtext("Sec " & (node("SalNatureDesc")), "PART_Nature_1")
                    Sheet1.Range("J" & rowcount).Value = test
                End If
               If Not node("SalOthNatOfInc") = "" Then
                Sheet1.Range("Z" & rowcount).Value = node("SalOthNatOfInc")
               End If
                Sheet1.Range("Z" & rowcount).Value = node("SalOthAmount")
            cnt = cnt + 1
        Next node
        RecTDS1 = cnt
    End If
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Deductions_16").Value = DeductionUs16
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("IncD.Deduction16ia").Value = DeductionUs16ia
  '  Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Deduction16").Value = EntertainmentAlw16ii
 '   Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.Deduction16ic").Value = ProfessionalTaxUs16iii
 
    If TypeOfHP <> "" Then
        If TypeOfHP = "L" Then
            Sheet1.Range("IncD.TypeOfHP").Value = "Let Out"
        ElseIf TypeOfHP = "S" Then
            Sheet1.Range("IncD.TypeOfHP").Value = "Self Occupied"
        ElseIf TypeOfHP = "D" Then
            Sheet1.Range("IncD.TypeOfHP").Value = "Deemed Let Out"
        End If
    End If
    
     Set init = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")
    If init.exists("OthersInc") Then
        Set Nodelist = jsonObject("ITR")("ITR1")("ITR1_IncomeDeductions")("OthersInc")("OthersIncDtlsOthSrc")
    
        NatureColNo = Sheet1.Range("Others.NOI_2").Column
        DescriptionColNo = Sheet1.Range("Nature_Others_2").Column
        AmtColNo = Sheet1.Range("Others.Amount_2").Column
        
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
                
                If UCase(node("OthSrcNatureDesc")) = "DIV" Then
                    Sheet1.Range("IncD_q4div").Value = node("DividendInc")("DateRange")("Up16Of12To15Of3")
                    Sheet1.Range("IncD_q5div").Value = node("DividendInc")("DateRange")("Up16Of3To31Of3")
                    Sheet1.Range("IncD_q3div").Value = node("DividendInc")("DateRange")("Up16Of9To15Of12")
                    Sheet1.Range("IncD_q1div").Value = node("DividendInc")("DateRange")("Upto15Of6")
                    Sheet1.Range("IncD_q2div").Value = node("DividendInc")("DateRange")("Upto15Of9")
                End If
                
                
                If UCase(node("OthSrcNatureDesc")) = "OTH" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
                End If
              '  Sheet1.Unprotect Password:=getmsgstate
               If UCase(node("OthSrcNatureDesc")) <> "DIV" Then
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
        End If
    
       If (TotalDiffRow > 0) Then
           AddDiffRows_Exempt (TotalDiffRow)
       End If
    
       rowcount = getRowNo(Sheet1.Range("Others.NOI").name)
       rowcount = rowcount - 1
       cnt = 0
        
        For Each node In Nodelist
            rowcount = rowcount + 1
                
                If UCase(node("NatureDesc")) = "OTH" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
                End If
                
                If UCase(node("NatureDesc")) = "10(34)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(34) (Exempted Dividend Income)"
                End If
                
                If UCase(node("NatureDesc")) = "10(26AAA)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(26AAA) Any income as referred to in section 10(26AAA)"
                End If
                
                If UCase(node("NatureDesc")) = "10(26)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(26) Any income as referred to in section 10(26)"
                End If
                
                If UCase(node("NatureDesc")) = "10(19)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(19) Armed Forces Family pension in case of death during operational duty"
                End If
                
                If UCase(node("NatureDesc")) = "10(10D)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10D) Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(1)"
                End If
                
                If UCase(node("NatureDesc")) = "10(11)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(11) Statutory Provident Fund received"
                End If
                
                If UCase(node("NatureDesc")) = "10(12)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(12) Recognized Provident Fund received"
                End If
                
                If UCase(node("NatureDesc")) = "10(13)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(13) Approved superannuation fund received"
                End If
                
                If UCase(node("NatureDesc")) = "10(16)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(16) Scholarships granted to meet the cost of education"
                End If
                
                If UCase(node("NatureDesc")) = "10(17)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(17)-Allowance MP/MLA/MLC" 'Ankita
                End If
                
                If UCase(node("NatureDesc")) = "10(17A)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(17A) Award instituted by Government"
                End If
                
                If UCase(node("NatureDesc")) = "10(18)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(18) Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award"
                End If
                
                If UCase(node("NatureDesc")) = "10(10BC)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10BC) Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster"
                End If
                
                If UCase(node("NatureDesc")) = "DMDP" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Defense Medical Disability Pension"
                End If
            
                If UCase(node("NatureDesc")) = "AGRI" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Agriculture Income (less than equal to Rs.5000)"
                End If
               ' Sheet1.Unprotect Password:=getmsgstate
                Sheet1.Cells(rowcount, DescriptionColNo).Value = node("OthNatOfInc")
                Sheet1.Cells(rowcount, AmtColNo).Value = node("OthAmount")
    
            cnt = cnt + 1
        Next node
        RecTDS1 = cnt
    End If
    
    'Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("IncD.LessDeduction57").Value = DeductionUs57iia
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
Sheet1.Range("IncD.Section89").Value = jsonObject("ITR")("ITR1")("ITR1_TaxComputation")("Section89")
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
Function ImportTDSonOthThanSals(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim TANColNo, DEDNameColNo, DEDAmountDeducted, UTNColNo, FYColNo, TaxColNo, ClaimColNo, BroughtFwdTDSAmt As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, cnt, rowcount As Long

Set jsonObject = ParseJson(jsonText)

Set Nodelist = jsonObject("ITR")("ITR1")("TDSonOthThanSals")("TDSonOthThanSal")
    
    TANColNo = Sheet2.Range("TDSoth.TAN").Column
    DEDNameColNo = Sheet2.Range("TDSoth.EmployerOrDeductorOrCollecterName").Column
    DEDAmountDeducted = Sheet2.Range("TDSoth.AmountDeducted").Column
    FYColNo = Sheet2.Range("TDSoth.DeductedYear").Column
    BroughtFwdTDSAmt = Sheet2.Range("TDSoth.TotTDSOnAmtPaid").Column
    TaxColNo = Sheet2.Range("TDSoth.6income").Column
    
    TotalExRow = Range("TDSoth.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet2.Range("TDSoth.TAN").ClearContents
        Sheet2.Range("TDSoth.EmployerOrDeductorOrCollecterName").ClearContents
        Sheet2.Range("TDSoth.AmountDeducted").ClearContents
        Sheet2.Range("TDSoth.DeductedYear").ClearContents
        Sheet2.Range("TDSoth.TotTDSOnAmtPaid").ClearContents
        Sheet2.Range("TDSoth.6income").ClearContents
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
            Sheet2.Cells(rowcount, FYColNo).Value = node("DeductedYr")
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
        Range("TDSal.TAN").ClearContents
        Range("TDSal.EmployerOrDeductorOrCollecterName").ClearContents
        Range("TDSal.IncChrgSalary").ClearContents
        Range("TDSal.TotalTDSSalary").ClearContents
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
Dim TANColNo1, TDSAADhaar, DEDNameColNo1, DEDAmountDeducted1, UTNColNo1, FYColNo1, TaxColNo1, ClaimColNo1, BroughtFwdTDSAmt1 As Variant
Dim TotalExRow1, TotalXMLRow1, TotalDiffRow1, cnt, RecTDS11, rowcount As Long

Set jsonObject = ParseJson(jsonText)

    Set nodeList1 = jsonObject("ITR")("ITR1")("ScheduleTDS3Dtls")("TDS3Details")
    
    TANColNo1 = Sheet2.Range("TDS26QB.PAN").Column
    TDSAADhaar = Sheet2.Range("TDS26QB.Aadhar_Number").Column
    DEDNameColNo1 = Sheet2.Range("TDS26QB.EmployerOrDeductorName").Column
    DEDAmountDeducted1 = Sheet2.Range("TDS26QB.AmountDeducted").Column
    FYColNo1 = Sheet2.Range("TDS26QB.DeductedYear").Column
    BroughtFwdTDSAmt1 = Sheet2.Range("TDS26QB.TotTDSOnAmtPaid").Column
    TaxColNo1 = Sheet2.Range("TDS26QB.6income").Column
    
    TotalExRow1 = Range("TDS26QB.PAN").Rows.count
    
    TotalXMLRow1 = nodeList1.count
    TotalDiffRow1 = TotalXMLRow1 - TotalExRow1
    
    If (TotalXMLRow1 > 0) Then
        Sheet2.Range("TDS26QB.PAN").ClearContents
        Sheet2.Range("TDS26QB.Aadhar_Number").ClearContents
        Sheet2.Range("TDS26QB.EmployerOrDeductorName").ClearContents
        Sheet2.Range("TDS26QB.AmountDeducted").ClearContents
        Sheet2.Range("TDS26QB.DeductedYear").ClearContents
        Sheet2.Range("TDS26QB.TotTDSOnAmtPaid").ClearContents
        Sheet2.Range("TDS26QB.6income").ClearContents
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
            Sheet2.Cells(rowcount, DEDAmountDeducted1).Value = Node1("GrsRcptToTaxDeduct")
            Sheet2.Cells(rowcount, FYColNo1).Value = Node1("DeductedYr")
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
        Sheet2.Range("TaxP.BSRCode").ClearContents
        Sheet2.Range("TaxP.DateDep").ClearContents
        Sheet2.Range("TaxP.SrlNoOfChaln").ClearContents
        Sheet2.Range("TaxP.Amt").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_IT (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet2.Range("TaxP.BSRCode").name)
    rowcount = rowcount - 1
    cnt = 0
    
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
            Sheet11.Cells(rowcount, CollectedYr).Value = node("CollectedYr")
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
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long

Set jsonObject = ParseJson(jsonText)
Set Nodelist = jsonObject("ITR")("ITR1")("Refund")("BankAccountDtls")("AddtnlBankDetails")

    IFSC = Range("SchBA.IFSC").Column
    BankName = Range("SchBA.BankName").Column
    ACCNO = Range("SchBA.AcntNo").Column
    CheckBox = Range("tempxml").Column
    
    TotalExRow = Range("SchBA.IFSC").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Range("SchBA.IFSC").ClearContents
        Range("SchBA.BankName").ClearContents
        Range("SchBA.AcntNo").ClearContents
        Range("tempxml").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_BANK (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet3.Range("SchBA.IFSC").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
            Sheet3.Cells(rowcount, IFSC).Value = node("IFSCCode")
            Sheet3.Cells(rowcount, BankName).Value = node("BankName")
            Sheet3.Cells(rowcount, ACCNO).Value = node("BankAccountNo")
            Sheet3.Cells(rowcount, CheckBox).Value = node("UseForRefund")
            
            If CheckBox = True Then
            CheckBox = "true"
            ElseIf CheckBox = False Then
            CheckBox = "false"
            ElseIf CheckBox = "" Then
            CheckBox = "false"
            End If
            
        LinkCheckBoxes
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
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "No Claiming for Self/Family"
                End If
            End If
                                            
            If Trim(node("HealthInsPremSlfFam")) <> "" Then
                Sheet9.Range("Health_Insurance_80D").Value = node("HealthInsPremSlfFam")
            End If
            If Trim(node("PrevHlthChckUpSlfFam")) <> "" Then
                Sheet9.Range("Preventive_Health_80D").Value = node("PrevHlthChckUpSlfFam")
            End If
            If Trim(node("HlthInsPremSlfFamSrCtzn")) <> "" Then
                Sheet9.Range("Health_InsuranceSC_80D").Value = node("HlthInsPremSlfFamSrCtzn")
            End If
            If Trim(node("PrevHlthChckUpSlfFamSrCtzn")) <> "" Then
                Sheet9.Range("Preventive_Health_SC_80D").Value = node("PrevHlthChckUpSlfFamSrCtzn")
            End If
            If Trim(node("MedicalExpSlfFamSrCtzn")) <> "" Then
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
            
            If Trim(node("HlthInsPremParents")) <> "" Then
                Sheet9.Range("Health_Insurance2_80D").Value = node("HlthInsPremParents")
            End If
            If Trim(node("PrevHlthChckUpParents")) <> "" Then
                Sheet9.Range("Preventive_Health2_80D").Value = node("PrevHlthChckUpParents")
            End If
            If Trim(node("HlthInsPremParentsSrCtzn")) <> "" Then
                Sheet9.Range("Health_Insurance3_80D").Value = node("HlthInsPremParentsSrCtzn")
            End If
            If Trim(node("PrevHlthChckUpParentsSrCtzn")) <> "" Then
                Sheet9.Range("Preventive_Health3_80D").Value = node("PrevHlthChckUpParentsSrCtzn")
            End If
            If Trim(node("MedicalExpParentsSrCtzn")) <> "" Then
                Sheet9.Range("Medical_Expenditure2_80D").Value = node("MedicalExpParentsSrCtzn")
            End If

End Function
Function ImportSchedule80G100NoAppr(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long

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
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_80G_A (TotalDiffRow)
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
            
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt

End Function
Function ImportSchedule80G50Appr(jsonText As String)
 On Error Resume Next
Dim jsonObject, jsonDictionary, init As Object
Dim node, Nodelist As Object
Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, rowcount, cnt As Long

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
    AmountColNo = Sheet4.Range("Per5080G.DonationAmt").Column
    DonationColNo = Sheet4.Range("Per5080G.DonationAmtOther").Column
    
    TotalExRow = Range("Per5080G.DoneeWithPanName").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet4.Range("Per5080G.DoneeWithPanName").ClearContents
        Sheet4.Range("Per5080G.AddrDetail").ClearContents
        Sheet4.Range("Per5080G.CityOrTownOrDistrict").ClearContents
        Sheet4.Range("Per5080G.StateCode").ClearContents
        Sheet4.Range("Per5080G.PinCode").ClearContents
        Sheet4.Range("Per5080G.DoneePAN").ClearContents
        Sheet4.Range("Per5080G.DonationAmt").ClearContents
        Sheet4.Range("Per5080G.DonationAmtOther").ClearContents
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
            Sheet4.Cells(rowcount, AmountColNo).Value = node("DonationAmtCash")
            Sheet4.Cells(rowcount, DonationColNo).Value = node("DonationAmtOtherMode")
            
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
            dateDonation = node("DateOfDonationCash")
                YYYY = Mid(dateDonation, 1, 4)
                MM = Mid(dateDonation, 6, 2)
                DD = Mid(dateDonation, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            Sheet12.Cells(rowcount, DateofDonationNo).Value = strDate
            Sheet12.Cells(rowcount, AmountColNo).Value = node("DonationAmtCash")
            If node("DonationAmtOtherMode") <> "" And node("DonationAmtOtherMode") <> "0" Then
                Sheet12.Cells(rowcount, DonationColNo).Value = node("DonationAmtOtherMode")
            End If
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



