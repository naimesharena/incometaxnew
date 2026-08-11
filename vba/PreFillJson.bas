Attribute VB_Name = "PreFillJson"

Option Explicit
'01/12/2021
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

'Malli_30/05/2025
Public incDeductionsOthIncCPC_TAX_chk As Boolean
Public incDeductionsOthIncCPC_TAX_RowCnt As Variant
'---------------------------------------------------

Sub PreFillJson()
Dim filepath, jsonText, bacVal As String

Unload UserForm3
fmsgbox "Import personal/tax details from downloaded Pre-filled JSON or Import from already generated JSON of the current assessment year."

With Application.FileDialog(msoFileDialogFilePicker)
    'Makes sure the user can select only one file
    .AllowMultiSelect = False
    .Title = "Please select a Json file."
    'Filter to just the following types of files to narrow down selection options
    .Filters.add "Json File", "*.json", 1
    'Show the dialog box
    If .Show = True Then
        imported = 1
        filepath = .SelectedItems.item(1)
    Else
        imported = 0
        Exit Sub
    End If
    On Error Resume Next
End With

Open filepath For Input As #1
jsonText = Input$(LOF(1), 1)
Close #1

Dim jsonObject As Object

Set jsonObject = ParseJson(jsonText)

If jsonObject.exists("ITR") Then
imported = 0
fmsgbox ("Please select a valid prefill json.")
Exit Sub
End If





ImportPersonalInfo_pfl (jsonText)
ImportFilingStatus_pfl (jsonText)
ImportRefund_pfl (jsonText)

'Added by Shrutika- 22/01/2026 AY-26-27
 ImportScheduleHP_Pfl (jsonText)  'Malli_29/01/2026
'----------------------------

ImportScheduleTCS_pfl (jsonText)
ImportScheduleIT_pfl (jsonText)
ImportTDSonSalary_pfl (jsonText)
ImportTDSOthThanSals_pfl (jsonText)
ImportScheduleTDS3Dtls_pfl (jsonText)
Import_IncomeDeductions (jsonText)

'Change.25.01.2023.102.80GC8A.Prefil



'End Change====
ImportVerification_pfl (jsonText)
'ImportTaxPaid (jsonText)
'ImportVerification (jsonText)




'/* SIT-111718 Malli_AY_2026_27  prefill posibility removed for Old tax regime fields
'ImportSchedule80GD_pfl (jsonText)
'ImportSchedule80D (jsonText)
'ImportSchedule80DD_80U_pfl (jsonText)
'ImportscheduleEA10_13A_pfl (jsonText)
'ImportSchedule_80E_Pfl (jsonText)
'ImportSchedule_80EE_pfl (jsonText)
'ImportSchedule_80EEA_pfl (jsonText)
'ImportSchedule_80EEB_pfl (jsonText)
'----------------------------------------------------------------------

End Sub

Private Function DecodeBase64(ByVal strData As String) As Byte()

 

    Dim objXML As Object
    Dim objNode As Object
   
    ' help from MSXML
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.text = strData
    DecodeBase64 = objNode.nodeTypedValue
   
    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing

 

End Function


Function ImportPersonalInfo_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Dim jsonDictionary As Object
Dim sCountry, iCountry As Variant
Dim sState, iState, status As Variant
Dim YYYY, MM, DD, strDate As String

Dim DateOFFormOrIncorp, localityOrArea, countryCodeMobileNoSec, mobileNoSec
Dim residenceNo, cityOrTownOrDistrict, countryCodeMobile, mobileNo, emailAddress, emailAddressSec, stDcode, phoneNo, roadOrStreet, PinCode, residenceName, zipCode
Dim residentialStatus, firstName, surNameOrOrgName, middleName, natureOfEmp, TypeOfHP
Dim dob, aadhaarCardNo, PAN

Set jsonObject = ParseJson(jsonText)

'MsgBox jsonObject
'MsgBox jsonText


'DateOFFormOrIncorp = jsonObject("personalInfo")("orgFirmInfo")("DateOFFormOrIncorp")

'bac value added
If jsonObject("lastFiledITR") <> Null Then
 If jsonObject("lastFiledITR").exists("NewTaxRegime") Then
 'Malli----------------------
 'PAG1_C1_AY_2024-25 'Malli
'If jsonObject("lastFiledITR")("NewTaxRegime") = "Y" Then
 If jsonObject("lastFiledITR")("NewTaxRegime") = "N" Or jsonObject("lastFiledITR")("NewTaxRegime") = "Y" Or jsonObject("lastFiledITR")("NewTaxRegime") = "" Then
 
 '--------------------------
 Sheet1.Unprotect Password:=sPassword
  ActiveSheet.Shapes("BacNo").OLEFormat.Object.Value = 1 'Newly changed by Bindu
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
  
  'Malli----------
  '80GGC_C1 2024-25
  Sheet13.Unprotect Password:=sPassword
  Sheet13.Visible = xlSheetHidden
  Sheet13.Protect Password:=sPassword
  '80U-DD_C1 2024-25
  Sheet14.Unprotect Password:=sPassword
  Sheet14.Visible = xlSheetHidden
  Sheet14.Protect Password:=sPassword
  
  
  '---------------
  
  
  sPassword = EfilingCommon.getmsgstate
  Sheet5.Unprotect Password:=sPassword
  Sheet5.Range("BacValue").Value = 1
  Sheet5.Protect Password:=sPassword
  Sheet1.Unprotect Password:=sPassword
  resetBacYes
  Sheet1.Protect Password:=sPassword
  Sheet1.Activate
  ThisWorkbook.Protect Password:=sPassword
  'Malli--------
  'PAG1_C1_AY_2024-25 'Malli
'ElseIf jsonObject("lastFiledITR")("NewTaxRegime") = "N" Then
'ElseIf jsonObject("lastFiledITR")("NewTaxRegime") = "Y" Then
''---------
' Sheet5.Unprotect Password:=sPassword
'   Sheet5.Range("BacValue").Value = 2
'    ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 1
'    ThisWorkbook.Unprotect Password:=sPassword
'    Sheet4.Unprotect Password:=sPassword
'  Sheet4.Visible = xlSheetVisible
'  Sheet4.Protect Password:=sPassword
'  Sheet12.Unprotect Password:=sPassword
'  Sheet12.Visible = xlSheetVisible
'  Sheet12.Protect Password:=sPassword
'  Sheet9.Unprotect Password:=sPassword
'  Sheet9.Visible = xlSheetVisible
'  Sheet9.Protect Password:=sPassword
'
'
'  'Malli------
'  'PAG1_C1_AY_2024-25 'Malli
'  '80GGC_C1 2024-25
'  Sheet13.Unprotect Password:=sPassword
'  Sheet13.Visible = xlSheetVisible
'  Sheet13.Protect Password:=sPassword
'  '80U-DD_C1 2024-25
'  Sheet14.Unprotect Password:=sPassword
'  Sheet14.Visible = xlSheetVisible
'  Sheet14.Protect Password:=sPassword
'  '---------
'
'  sPassword = EfilingCommon.getmsgstate
'  Sheet5.Unprotect Password:=sPassword
'  Sheet5.Range("BacValue").Value = 2
'  Sheet5.Protect Password:=sPassword
'  Sheet1.Unprotect Password:=sPassword
'  resetBacNo
'  Sheet1.Protect Password:=sPassword
'
'  Sheet1.Activate
'  ThisWorkbook.Protect Password:=sPassword
'Else
 '   ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 0
  '  ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 0
End If
End If
End If


'end bac value change
localityOrArea = jsonObject("personalInfo")("address")("localityOrArea")
countryCodeMobileNoSec = jsonObject("personalInfo")("address")("countryCodeMobileNoSec") 'Added by Shrutika- 23/01/2026 AY-26-27
mobileNoSec = jsonObject("personalInfo")("address")("mobileNoSec")  'Added by Shrutika- 23/01/2026 AY-26-27
residenceNo = jsonObject("personalInfo")("address")("residenceNo")

cityOrTownOrDistrict = jsonObject("personalInfo")("address")("cityOrTownOrDistrict")
countryCodeMobile = jsonObject("personalInfo")("address")("countryCodeMobile")
mobileNo = jsonObject("personalInfo")("address")("mobileNo")

emailAddress = jsonObject("personalInfo")("address")("emailAddress")
stDcode = jsonObject("personalInfo")("address")("phone")("stDcode")
phoneNo = jsonObject("personalInfo")("address")("phone")("phoneNo")
iCountry = jsonObject("personalInfo")("address")("countryCode")
sCountry = Findtext(CStr(iCountry), "CountList")
roadOrStreet = jsonObject("personalInfo")("address")("roadOrStreet")
PinCode = jsonObject("personalInfo")("address")("pinCode")
iState = jsonObject("personalInfo")("address")("stateCode")
sState = Findtext(CStr(iState), "StateList")
zipCode = jsonObject("personalInfo")("address")("zipCode")
residenceName = jsonObject("personalInfo")("address")("residenceName")
emailAddressSec = jsonObject("personalInfo")("address")("emailAddressSecondary") 'Added by Shrutika- 20/01/2026 AY-26-27

firstName = jsonObject("personalInfo")("assesseeName")("firstName")
surNameOrOrgName = jsonObject("personalInfo")("assesseeName")("surNameOrOrgName")
middleName = jsonObject("personalInfo")("assesseeName")("middleName")

'Malli comented _AY_2025_26
'Chandru

'TypeOfHP = jsonObject("lastFiledITR")("typeOfHP")

'--------------------------

natureOfEmp = jsonObject("lastFiledITR")("employerCategory")
dob = jsonObject("personalInfo")("dob")
If dob <> "" Then
    YYYY = Mid(dob, 1, 4)
    MM = Mid(dob, 6, 2)
    DD = Mid(dob, 9, 2)
    strDate = DD & "/" & MM & "/" & YYYY
End If

aadhaarCardNo = jsonObject("personalInfo")("aadhaarCardNo")
aadhaarCardNo = StrConv(DecodeBase64(aadhaarCardNo), vbUnicode)
PAN = jsonObject("personalInfo")("pan")

'residentialStatus = jsonObject("personalInfo")("filingStatus")("residentialStatus")
'residentialStatus = Findtext(residentialStatus & " ", "ResStatus")


'    Sheet1.Range("sheet1.ResidentialStatus1").value = residentialStatus

status = jsonObject("personalInfo")("status")

    If status = "I" Then
        status = "I - INDIVIDUAL"
    ElseIf status = "H" Then
        status = "H - HUF"
    ElseIf status = "F" Then
        status = "F - FIRM(Other than LLP)"
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
    If surNameOrOrgName <> "" Then
        Sheet1.Range("sheet1.SurNameOrOrgName").Value = surNameOrOrgName
    End If
    Sheet1.Unprotect Password:=getmsgstate
    If PAN <> "" Then
        Sheet1.Range("sheet1.PAN").Value = PAN
    End If
    If aadhaarCardNo <> "" And Sheet1.Range("Sheet1.Aadhaar").Locked = False Then
    
        Sheet1.Range("Sheet1.Aadhaar").Value = aadhaarCardNo
    End If
    Sheet1.Unprotect Password:=getmsgstate
    If strDate <> "" Then
        Sheet1.Range("sheet1.DOB").Value = strDate
    End If
    Sheet1.Protect Password:=getmsgstate
    If residenceNo <> "" And Sheet1.Range("sheet1.ResidenceNo").Locked = False Then
        Sheet1.Range("sheet1.ResidenceNo").Value = residenceNo
    End If
    If residenceName <> "" And Sheet1.Range("sheet1.ResidenceName").Locked = False Then
        Sheet1.Range("sheet1.ResidenceName").Value = residenceName
    End If
    If roadOrStreet <> "" And Sheet1.Range("sheet1.RoadOrStreet").Locked = False Then
        Sheet1.Range("sheet1.RoadOrStreet").Value = roadOrStreet
    End If
    If localityOrArea <> "" And Sheet1.Range("sheet1.LocalityOrArea").Locked = False Then
        Sheet1.Range("sheet1.LocalityOrArea").Value = localityOrArea
    End If
    If cityOrTownOrDistrict <> "" And Sheet1.Range("sheet1.CityOrTownOrDistrict").Locked = False Then
        Sheet1.Range("sheet1.CityOrTownOrDistrict").Value = cityOrTownOrDistrict
    End If
    'Konda edited-SIT-67322-----25/06/2024
    If sState <> "" And Sheet1.Range("sheet1.StateCode1").Locked = False Then
        Sheet1.Range("sheet1.StateCode1").Value = sState
    End If
    'end-SIT-67322-----25/06/2024
    If sCountry <> "" And Sheet1.Range("sheet1.Country").Locked = False Then
        Sheet1.Range("sheet1.Country").Value = sCountry
    End If
    
'    If sState <> "" And Sheet1.Range("sheet1.StateCode1").Locked = False Then
'        Sheet1.Range("sheet1.StateCode1").Value = sState
'    End If
'       end ----------
    If PinCode <> "" And Sheet1.Range("sheet1.PinCode").Locked = False Then
        Sheet1.Range("sheet1.PinCode").Value = PinCode
    End If
    If zipCode <> "" And Sheet1.Range("sheet1.ZipCode").Locked = False Then
        Sheet1.Range("sheet1.ZipCode").Value = zipCode
    End If
    If emailAddress <> "" And Sheet1.Range("sheet1.EmailAddress").Locked = False Then
        Sheet1.Range("sheet1.EmailAddress").Value = emailAddress
    End If
    If countryCodeMobile <> "" And Sheet1.Range("sheet1.MobileCountryCode").Locked = False Then
        Sheet1.Range("sheet1.MobileCountryCode").Value = countryCodeMobile
    End If
    If mobileNo <> "" And Sheet1.Range("sheet1.Mobileno").Locked = False Then
        Sheet1.Range("sheet1.Mobileno").Value = mobileNo
    End If
'    If status <> "" Then
'        Sheet1.Range("sheet1.Status").Value = status
'    End If
    If stDcode <> "" And Sheet1.Range("sheet1.STDcode").Locked = False Then
        Sheet1.Range("sheet1.STDcode").Value = stDcode
    End If
    If phoneNo <> "" And Sheet1.Range("sheet1.PhoneNo").Locked = False Then
        Sheet1.Range("sheet1.PhoneNo").Value = phoneNo
    End If
    '-------------------------------------------------
'    If countryCodeMobileNoSec <> "" And Sheet1.Range("sheet1.mobileCountryCode2").Locked = False Then
'        Sheet1.Range("sheet1.mobileCountryCode2").Value = countryCodeMobileNoSec
'    End If
'    If mobileNoSec <> "" And Sheet1.Range("sheet1.MobileNoSec").Locked = False Then
'        Sheet1.Range("sheet1.MobileNoSec").Value = mobileNoSec
'    End If
   
'    If emailAddressSec <> "" And Sheet1.Range("sheet1.EmailAddress2").Locked = False Then
'        Sheet1.Range("sheet1.EmailAddress2").Value = emailAddressSec
'    End If
     '---------------------
     
    'Added by Shrutika- 23/01/2026 AY-26-27
   
     If countryCodeMobileNoSec <> "" And Sheet1.Range("sheet1.MobileCountryCode1").Locked = False Then
        Sheet1.Range("sheet1.MobileCountryCode1").Value = countryCodeMobileNoSec
    End If
    If mobileNoSec <> "" And Sheet1.Range("sheet1.Mobileno1").Locked = False Then
        Sheet1.Range("sheet1.Mobileno1").Value = mobileNoSec
    End If
    
    If emailAddressSec <> "" And Sheet1.Range("sheet1.EmailAddress1").Locked = False Then
        Sheet1.Range("sheet1.EmailAddress1").Value = emailAddressSec
    End If
'------------------------------------------------------------------------

    'issue 1 solved
    If natureOfEmp <> "" And Sheet1.Range("sheet1.EmployerCategory1").Locked = False Then
        If natureOfEmp = "CGOV" Then
            natureOfEmp = "Central Government"
        ElseIf natureOfEmp = "SGOV" Then
            natureOfEmp = "State Government"
        ElseIf natureOfEmp = "PSU" Then
            natureOfEmp = "Public Sector Undertaking"
        ElseIf natureOfEmp = "OTH" Then
            natureOfEmp = "Others"
        ElseIf natureOfEmp = "PE" Then
            natureOfEmp = "Pensioners - Central Government"
        ElseIf natureOfEmp = "PESG" Then
            natureOfEmp = "Pensioners - State Government"
        ElseIf natureOfEmp = "PEPS" Then
            natureOfEmp = "Pensioners - Public sector undertaking "
        ElseIf natureOfEmp = "PEO" Then
            natureOfEmp = "Pensioners - Others"
        ElseIf natureOfEmp = "NA" Then
            natureOfEmp = "Not Applicable (eg. Family pension etc)"
        End If

        Sheet1.Range("sheet1.EmployerCategory1").Value = natureOfEmp
    End If
     'issue 1 solved
     
     
     
    '---------------'Added by Shrutika- 23/01/2026 AY-26-27

    Dim RepName, RepEmail, RepMobile, RepCountryCode
    
    RepName = jsonObject("assesseeRep")("repName")
    
    'commented by Malli
'        If RepName <> "" And Sheet1.Range("sheet1.NameRepAssessee").Locked = False Then
'            'Sheet1.Range("sheet1.RepAssessee").Value = "Y - Yes"
'            Sheet1.Range("sheet1.RepAssessee").Value = "Yes"
'            ElseIf RepName = "" And Sheet1.Range("sheet1.NameRepAssessee").Locked = False Then
'            'Sheet1.Range("sheet1.RepAssessee").Value = "N -No"
'            Sheet1.Range("sheet1.RepAssessee").Value = "No"
'        End If
        
        'Malli_27/01/2026
        If RepName <> "" And Sheet1.Range("sheet1.RepAssessee").Locked = False Then
            Sheet1.Range("sheet1.RepAssessee").Value = "Yes"
        ElseIf RepName = "" And Sheet1.Range("sheet1.RepAssessee").Locked = False Then
            Sheet1.Range("sheet1.RepAssessee").Value = "No"
        End If
         '-------------------------------------------
    
        If RepName <> "" And Sheet1.Range("sheet1.NameRepAssessee").Locked = False Then
            Sheet1.Range("sheet1.NameRepAssessee").Value = RepName
        End If
             
     
     RepEmail = jsonObject("assesseeRep")("repEmailID")
    If RepEmail <> "" And Sheet1.Range("sheet1.EmailRepAssessee").Locked = False Then
        Sheet1.Range("sheet1.EmailRepAssessee").Value = RepEmail
    End If
     
    RepMobile = jsonObject("assesseeRep")("repMobileNo")
     If RepMobile <> "" And Sheet1.Range("sheet1.ContactRepAssessee").Locked = False Then
        Sheet1.Range("sheet1.ContactRepAssessee").Value = RepMobile
    End If
    
   
    RepCountryCode = jsonObject("assesseeRep")("countryCodeRepMobileNo")

        If RepCountryCode <> "" And Sheet1.Range("sheet1.CountryCodeRepAssessee").Locked = False Then
          Sheet1.Range("sheet1.CountryCodeRepAssessee").Value = RepCountryCode
        End If
     '-----------------------------------------------------------'Added by Shrutika- 23/01/2026 AY-26-27

   
   
'Malli commented--------27/01/2026
    
'Malli-------AY_2025_26

'If jsonObject("lastFiledITR") <> Null Then
''Chandru
'    TypeOfHP = jsonObject("lastFiledITR")("typeOfHP")
'    If TypeOfHP <> "" And Sheet1.Range("IncD.TypeOfHP").Locked = False Then
'
'        If TypeOfHP = "Y" Or TypeOfHP = "L" Then
'            TypeOfHP = "Let Out"
'            ElseIf TypeOfHP = "S" Or TypeOfHP = "N" Then
'            TypeOfHP = "Self Occupied"
'            ElseIf TypeOfHP = "D" Then
'            TypeOfHP = "Deemed Let Out"
'    End If
'
'    Sheet1.Range("IncD.TypeOfHP").Value = TypeOfHP
'    End If
'ElseIf jsonObject("form26as") <> Null Then
'
''form26as.typeOfHP  'Y
'
'    TypeOfHP = jsonObject("form26as")("typeOfHP")
'
'    If TypeOfHP <> "" And Sheet1.Range("IncD.TypeOfHP").Locked = False Then
'
'    If TypeOfHP = "Y" Or TypeOfHP = "L" Then
'            TypeOfHP = "Let Out"
'            ElseIf TypeOfHP = "S" Or TypeOfHP = "N" Then
'            TypeOfHP = "Self Occupied"
'            ElseIf TypeOfHP = "D" Then
'            TypeOfHP = "Deemed Let Out"
'    End If
'
'End If
'     Sheet1.Range("IncD.TypeOfHP").Value = TypeOfHP
'End If
'   '--------------------------------
   
End Function

Function ImportFilingStatus_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Dim jsonDictionary As Object
Dim ReturnFileSec, AsseseeRepFlg, node, Nodelist, NewTaxRegime, Form10IEDate, Form10IEAckNo, SeventhProvisio139, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, RepName, Repcapacity, RepAddress, RepPAN, RepAadhaar
Dim section139, depAmtAggAmt, uniqueNo
Set jsonObject = ParseJson(jsonText)

    Dim Filingtype As Variant
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
    Dim DateofOriginalfile As Variant
    Dim NoticeDateussec As Variant
    Dim iProvisoFlag, sProvisoFlag As Variant
    Dim iProvisoFlag1, sProvisoFlag1 As Variant
    Dim iDepositAmountFlag, sDepositAmountFlag As Variant
    Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
    Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant
    
'    iProvisoFlag = jsonObject("filingStatus")("SeventhProvisio139")
'    If iProvisoFlag = "Y" Then
'        sProvisoFlag = "Yes"
'    ElseIf iProvisoFlag = "N" Then
'        sProvisoFlag = "No"
'    End If
'    If Sheet1.Range("Sheet1.SeventhProvisoFlag").Locked = False Then
'    Sheet1.Range("sheet1.SeventhProvisoFlag").Value = sProvisoFlag
'    End If
    
'Chandru
       iProvisoFlag = jsonObject("filingStatus")("SeventhProvisio139")
    If iProvisoFlag <> "" Then
        If iProvisoFlag = "Y" Then
            iProvisoFlag = "Yes"
        ElseIf iProvisoFlag = "N" Then
            iProvisoFlag = "No"
        End If
'    Else
'        iProvisoFlag = "(Select)"
    End If
     If Sheet1.Range("Sheet1.SeventhProvisoFlag").Locked = False Then
        Sheet1.Range("Sheet1.SeventhProvisoFlag").Value = iProvisoFlag
    End If
    
    
    ' some change in aggregate flag 1crore and 2Lakhs
    If iProvisoFlag = "Y" Then
        iDepositAmountFlag = jsonObject("filingStatus")("DepAmtAggAmtExcd1CrPrYrFlg")
        If iDepositAmountFlag = "Y" Then
            sDepositAmountFlag = "Yes"
        ElseIf iDepositAmountFlag = "N" Then
            sDepositAmountFlag = "No"
     End If
        
  End If 'Malli
'******
'

'
   
    '&&&&&&&&&&&

'Change-22.11.2022.102.16J
'        If sDepositAmountFlag <> "" And Sheet1.Range("Sheet1.DepositAmountFlag").Locked = False Then
'            Sheet1.Range("Sheet1.DepositAmountFlag").Value = sDepositAmountFlag
'        End If
'        If sDepositAmountFlag = "Yes" And jsonObject("filingStatus")("AmtSeventhProvisio139i") <> "" Then
'            Sheet1.Range("Sheet1.DepositAmount").Value = jsonObject("filingStatus")("AmtSeventhProvisio139i")
'        End If
'---end

        iAggrigateAmountFlag = jsonObject("form26as")("IncrExpAggAmt2LkTrvFrgnCntryFlg")
        If iAggrigateAmountFlag = "Y" Then
            sAggrigateAmountFlag = "Yes"
        ElseIf iAggrigateAmountFlag = "N" Then
            sAggrigateAmountFlag = "No"
        End If
         If sAggrigateAmountFlag <> "" And Sheet1.Range("Sheet1.AggrigateAmountFlag").Locked = False Then
            Sheet1.Range("Sheet1.AggrigateAmountFlag").Value = sAggrigateAmountFlag
        End If
        
        If sAggrigateAmountFlag = "Yes" And jsonObject("form26as")("AmtSeventhProvisio139ii") <> "" Then
        'Ankita 08/11/2024
            If Sheet1.Range("Sheet1.AggrigateAmount").Locked = False Then
            Sheet1.Range("Sheet1.AggrigateAmount").Value = jsonObject("form26as")("AmtSeventhProvisio139ii")
            End If
        End If
   ' End If  'Malli
    'end change


      
    iProvisoFlag1 = jsonObject("filingStatus")("clauseiv7provisio139i")
    If iProvisoFlag1 = "Y" Then
        sProvisoFlag1 = "Yes"
    ElseIf iProvisoFlag1 = "N" Then
    sProvisoFlag1 = "No"
    End If
    
    
 If Sheet1.Range("clauseiv7provisio139iFlg").Locked = False Then
    Sheet1.Range("clauseiv7provisio139iFlg").Value = sProvisoFlag1
 End If


    'For Each node In jsonObject("filingStatus")("clauseiv7provisio139iDtls")
                
'       If node("clauseiv7provisio139iNature") = "1" Then
'          Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes"
'          Sheet1.Range("clauseiv7provisio139iAmount_3") = node("clauseiv7provisio139iAmount")
'       ElseIf node("clauseiv7provisio139iNature") = "2" Then
'          Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes"
'          Sheet1.Range("clauseiv7provisio139iAmount_4") = node("clauseiv7provisio139iAmount")
'       End If
       
       
    'Next node


'Newly added by Bindu

For Each node In jsonObject("filingStatus")("clauseiv7provisio139iDtls")


    If UCase(node("clauseiv7provisio139iNature")) = UCase("Yes") Then
         
         'Ankita 08/11/2024
         If Sheet1.Range("clauseiv7provisio139iFlg_3").Locked = False Then
         Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes"
         End If
         If Sheet1.Range("clauseiv7provisio139iAmount_3").Locked = False Then
         Sheet1.Range("clauseiv7provisio139iAmount_3") = node("clauseiv7provisio139iAmount")
         End If
         '---------------------
         
    End If
Next

'FilingStatus.clauseiv7provisio139i
'FilingStatus.clauseiv7provisio139iDtls.clauseiv7provisio139iNature
'FilingStatus.clauseiv7provisio139iDtls.clauseiv7provisio139iAmount


    uniqueNo = jsonObject("filingStatus")("uniqueNo")
    
    ReturnFileSec = jsonObject("filingStatus")("returnFileSec")
    ReceiptNo = jsonObject("filingStatus")("receiptNo")
    OrigRetFiledDate = jsonObject("filingStatus")("origRetFiledDate")
    NoticeDateUnderSec = jsonObject("filingStatus")("noticeDateUnderSec")
If ReturnFileSec <> "" And Sheet1.Range("sheet1.ReturnFileSec").Locked = False Then
        iReturnFile = ReturnFileSec
            
'        If iReturnFile = "11" Then
'           sReturnFile = "139(1)-On or before due date"
'        ElseIf iReturnFile = "12" Then
'           sReturnFile = "139(4)-After due date"
'        ElseIf iReturnFile = "13" Then
'            Sheet1.Shapes("RadioButton2").OLEFormat.Object.Value = 1
'        RadioButton2_Click
'           sReturnFile = "142(1)"
'        ElseIf iReturnFile = "14" Then
'
'           Sheet1.Shapes("RadioButton2").OLEFormat.Object.Value = 1
'        RadioButton2_Click
'        sReturnFile = "148"
'        ElseIf iReturnFile = "15" Then
'           sReturnFile = "153A"
'        ElseIf iReturnFile = "16" Then
'           sReturnFile = "153C"
'        ElseIf iReturnFile = "17" Then
'           sReturnFile = "139(5)-Revised Return"
'        ElseIf iReturnFile = "18" Then
'        Sheet1.Shapes("RadioButton2").OLEFormat.Object.Value = 1
'        RadioButton2_Click
'           sReturnFile = "139(9)"
'        ElseIf iReturnFile = "20" Then
'           sReturnFile = "139(9A) - After condonation of delay u/s 119(2)(b)"
'        End If


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
        'Konda Updated on 23-12-2025
               'Range("sheet1.ReturnFileSec").Value = "139(9A) - After condonation of delay u/s 119(2)(b)"
               '119(2)(b)- After condonation of delay
                Range("sheet1.ReturnFileSec").Value = "119(2)(b)- After condonation of delay"
            End If
'        Sheet1.Range("sheet1.ReturnFileSec").Value = sReturnFile
End If
   
        ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1")
                
        ReturnFileSec = Mid(ReturnFileSec, 1, 2)
        
        If ReturnFileSec = "17" Then
            If Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
                Sheet1.Range("sheet1.ReceiptNo").Value = ReceiptNo
            End If
            
            DateofOriginalfile = OrigRetFiledDate
            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
                Sheet1.Range("sheet1.OrigRetFiledDate").Value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
            End If
        Else
            If Sheet1.Range("sheet1.ReceiptNo").Locked = False Then
                Sheet1.Range("Sheet1.ReceiptNo").Value = ReceiptNo
            End If
            
            DateofOriginalfile = OrigRetFiledDate
            If DateofOriginalfile <> "" And Sheet1.Range("sheet1.OrigRetFiledDate").Locked = False Then
                Dim datefile As String
                datefile = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
                Sheet1.Range("sheet1.OrigRetFiledDate").Value = datefile
            End If
            
            'issue resolved
            NoticeDateussec = NoticeDateUnderSec
            If NoticeDateussec <> "" And Sheet1.Range("sheet1.NoticeDate").Locked = False Then
                Sheet1.Range("sheet1.NoticeDate").Value = Mid(NoticeDateussec, 9, 2) & "/" & Mid(NoticeDateussec, 6, 2) & "/" & Mid(NoticeDateussec, 1, 4)
            End If
            'issue resolved
        End If
        'issue 2 solved
        If uniqueNo <> "" And Sheet1.Range("sheet1.NoticeNo").Locked = False Then
            Sheet1.Range("sheet1.NoticeNo").Value = uniqueNo
        End If
        'issue 2 solved
End Function

Function ImportVerification_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
Dim jsonDictionary As Object
Dim FatherName, AssesseeVerName, AssesseeVerPAN, Capacity, Place As Variant


Set jsonObject = ParseJson(jsonText)
 
 FatherName = jsonObject("verification")("declaration")("fatherName")
 AssesseeVerName = jsonObject("verification")("declaration")("assesseeVerName")
 AssesseeVerPAN = jsonObject("verification")("declaration")("assesseeVerPAN")
 
 Capacity = jsonObject("verification")("capacity")
 
    If Capacity = "S" Or Capacity = "Self" Then
        Capacity = "Self"
    ElseIf Capacity = "R" Or Capacity = "Representative" Then
        Capacity = "Representative"
    ElseIf Capacity = "K" Or Capacity = "Karta" Then
        Capacity = "Karta"
    ElseIf Capacity = "P" Or Capacity = "Partner" Then
        Capacity = "Partner"
    Else
        Capacity = ""
    End If

        If FatherName <> "" And Sheet3.Range("Ver.FatherName").Locked = False Then
            Sheet3.Range("Ver.FatherName").Value = FatherName
        End If
        
        If AssesseeVerName <> "" And Sheet3.Range("Ver.AssesseeVerName").Locked = False Then
            Sheet3.Range("Ver.AssesseeVerName").Value = AssesseeVerName
        End If
        
        If AssesseeVerPAN <> "" And Sheet3.Range("Ver.PAN").Locked = False Then
            Sheet3.Range("Ver.PAN").Value = AssesseeVerPAN
        End If
            
            
        If Capacity <> "" And Sheet3.Range("Ver.capacity").Locked = False Then
            Sheet3.Range("Ver.capacity").Value = Capacity
        End If
        
                                                
End Function
Function ImportRefund_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist, NodeTemp
Dim IFSCColNo, BankNameColNo, ACCNOColNo, CheckBox, AccountTyp As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim NodeMain, NodeListMain

Set jsonObject = ParseJson(jsonText)



Set Nodelist = New Collection
If jsonObject.exists("bankAccountDtls") Then
Set NodeListMain = jsonObject("bankAccountDtls")
If NodeListMain <> Empty Then
Dim flag
For Each NodeMain In NodeListMain
    For Each node In NodeMain("addtnlBankDetails")
        flag = 0
        For Each NodeTemp In Nodelist
            If NodeTemp("bankAccountNo") = node("bankAccountNo") Then
                flag = 1
            End If
        Next NodeTemp
        If flag = 0 Then
            Nodelist.add node
        End If
        
    Next node
Next NodeMain
End If
End If

'Konda------------------Comented--AY_2025-26
    'Set NodeListMain = jsonObject("lastFiledITR")("bankAccountDtls")
    'If NodeListMain <> Empty Then
    'For Each NodeMain In NodeListMain
    '    If NodeMain.exists("addtnlBankDetails") Then
    '    For Each node In NodeMain("addtnlBankDetails")
    '        flag = 0
    '        For Each NodeTemp In Nodelist
    '            If NodeTemp("bankAccountNo") = node("bankAccountNo") Then
    '                flag = 1
    '            End If
    '        Next NodeTemp
    '        If flag = 0 Then
    '            Nodelist.add node
    '        End If
    '
    '    Next node
    '    End If
    'Next NodeMain
    'End If
'end-----------------------






'    'issue resolved
'    Set NodeListMain = jsonObject("bankAccountDtls")
'    'issue resolved
'    For Each NodeMain In NodeListMain
'
'        Set Nodelist = NodeMain("addtnlBankDetails")
        
  'Malli-----AY_2025_26 Change
        
    Set NodeListMain = jsonObject("bankAccountDtls")
        For Each NodeMain In NodeListMain
            Set Nodelist = NodeMain("addtnlBankDetails")
        '-------------------------------------------------------
        IFSCColNo = Sheet3.Range("SchBA.IFSC").Column
        BankNameColNo = Sheet3.Range("SchBA.BankName").Column
        ACCNOColNo = Sheet3.Range("SchBA.AcntNo").Column
   'Konda------------------------Uncomented--AY_2025-26
        CheckBox = Sheet3.Range("tempXML").Column
   'End-----------------------------
        'Malli------------
        
        AccountTyp = Sheet3.Range("SchBA.AcntType").Column
        
        '-------------
        
    'Konda------------------------Uncomented--AY_2025-26
    Dim CheckBoxColNo
        CheckBoxColNo = Sheet3.Range("SchBA.CheckBox").Column
     'End-----------------------------
        TotalExRow = Range("SchBA.IFSC").Rows.count
        
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
        
        
        If (TotalXMLRow > 0) Then
            Sheet3.Range("SchBA.IFSC").ClearContents
            Sheet3.Range("SchBA.BankName").ClearContents
            Sheet3.Range("SchBA.AcntNo").ClearContents
         'Konda------------------------Uncomented--AY_2025-26
            Sheet3.Range("tempxml").ClearContents
         'End-----------------------------
             Sheet3.Range("SchBA.AcntType").ClearContents  'Malli
            
        End If
        
  'Konda---------------Updated AY_2025-26
'        If (TotalDiffRow > 0) Then
'         AddDiffRows_BANK (TotalDiffRow)
'        End If

        If (TotalDiffRow > 0) Then
            Dim idi As Long
            idi = 0
            For idi = 1 To TotalDiffRow
             AddDiffRows_BANK (1)
             Next
        End If
    'End--------------------Updated AY_2025-26
        rowcount = getRowNo(Sheet3.Range("SchBA.IFSC").name)
        rowcount = rowcount - 1
        cnt = 0
        
        For Each node In Nodelist
            rowcount = rowcount + 1
                
                If Trim(node("ifsccode")) <> "" And Sheet3.Cells(rowcount, IFSCColNo).Locked = False Then
                    Sheet3.Cells(rowcount, IFSCColNo).Value = node("ifsccode")
                End If
                
                If Trim(node("bankName")) <> "" And Sheet3.Cells(rowcount, BankNameColNo).Locked = False Then
                    Sheet3.Cells(rowcount, BankNameColNo).Value = node("bankName")
                End If
                
                If Trim(node("bankAccountNo")) <> "" And Sheet3.Cells(rowcount, ACCNOColNo).Locked = False Then
                    Sheet3.Cells(rowcount, ACCNOColNo).Value = node("bankAccountNo")
                End If
                
             'Malli---------------------------------comented
    'Konda------------------------Uncomented--AY_2025-26
                Dim useForRefund
                useForRefund = node("useForRefund")
                If useForRefund = "" Then
                    Sheet3.Cells(rowcount, CheckBox).Value = False
                ElseIf useForRefund = "N" Or useForRefund = "false" Then
                    Sheet3.Cells(rowcount, CheckBox).Value = False
                Else
                    Sheet3.Cells(rowcount, CheckBox).Value = True
                End If

                If CheckBox = True Then
                CheckBox = "true"
                ElseIf CheckBox = False Then
                CheckBox = "false"
                ElseIf CheckBox = "" Then
                CheckBox = "false"
                End If

            LinkCheckBoxes
  'End-----------------------------
           '--------------------------------------------------
           'Malli-----------------------
                Dim AcType As Variant
                AcType = node("AccountType")
                
               ' AcType = "OTH"
'
                If AcType = "SB" Then
                    Sheet3.Cells(rowcount, AccountTyp).Value = "Savings Account"

                    ElseIf AcType = "CA" Then
                    Sheet3.Cells(rowcount, AccountTyp).Value = "Current Account"

                    ElseIf AcType = "CC" Then
                    Sheet3.Cells(rowcount, AccountTyp).Value = "Cash Credit Account"

                    ElseIf AcType = "OD" Then
                    Sheet3.Cells(rowcount, AccountTyp).Value = "Over draft account"

                    ElseIf AcType = "NRO" Then
                    Sheet3.Cells(rowcount, AccountTyp).Value = "Non Resident Account"

                    ElseIf AcType = "OTH" Then
                    Sheet3.Cells(rowcount, AccountTyp).Value = "Other"
                End If
                    
          
            
            cnt = cnt + 1
        Next node
        RecTDS1 = cnt
   Next NodeMain
End Function




'Added by Shrutika- 23/01/2026 AY-26-27

'lastFiledITR.schedule80E.schedule80EDtls

Function ImportSchedule_80E_Pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80E As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80E, IFSC_80E, BankName_80E, PANof_80E, LoanAcctNum_80E, LoanDate_80E, TotalLoanAmt_80E, LoanOutStanding_80E, Intrst_80E As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80E, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String
 
 
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80E = jsonObject("lastFiledITR")
 
If NodeList80E.exists("schedule80E") Then
 
Set Nodelist = jsonObject("lastFiledITR")("schedule80E")("schedule80EDtls")

    LoanfrmBankOrInstitute_80E = Sheet17.Range("LoanfrmBankOrInstitute.80E").Column
    BankName_80E = Sheet17.Range("bankName.80E").Column
    LoanAcctNum_80E = Sheet17.Range("loanAccNum.80E").Column
    LoanDate_80E = Sheet17.Range("loanDate.80E").Column
    TotalLoanAmt_80E = Sheet17.Range("loanAmt.80E").Column
     
    
    TotalExRow = Sheet17.Range("LoanfrmBankOrInstitute.80E").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet17.Range("LoanfrmBankOrInstitute.80E").Locked = False Then
            Sheet17.Range("LoanfrmBankOrInstitute.80E").ClearContents
        End If
        
        If Sheet17.Range("bankName.80E").Locked = False Then
            Sheet17.Range("bankName.80E").ClearContents
        End If
        
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
            Dim LoanTknFrom_80E, LoanTknFrom_80E_pfl
            
            LoanTknFrom_80E_pfl = node("loanTknFrom")
            If UCase(LoanTknFrom_80E_pfl) = UCase("B") Then
                LoanTknFrom_80E = "Bank"
            ElseIf UCase(LoanTknFrom_80E_pfl) = UCase("I") Then
                LoanTknFrom_80E = "Institution"
            Else
                LoanTknFrom_80E = "(Select)"
            End If
            Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80E).Value = LoanTknFrom_80E
        End If
        
        
        If Sheet17.Cells(rowcount, BankName_80E).Locked = False Then
            Sheet17.Cells(rowcount, BankName_80E).Value = node("bankOrInstnName")
        End If
         
        If Sheet17.Cells(rowcount, LoanAcctNum_80E).Locked = False Then
            Sheet17.Cells(rowcount, LoanAcctNum_80E).Value = node("loanAccNoOfBankOrInstnRefNo")
        End If
                
            strDate = node("dateofLoan")
            If strDate <> "" Then
               strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2) & "/" & Mid(strDate, 1, 4)
            End If
        
        If Sheet17.Cells(rowcount, LoanDate_80E).Locked = False Then
            Sheet17.Cells(rowcount, LoanDate_80E).Value = strDate
            strDate = ""
        End If
        
        If Sheet17.Cells(rowcount, TotalLoanAmt_80E).Locked = False Then
            Sheet17.Cells(rowcount, TotalLoanAmt_80E).Value = node("totalLoanAmt")
        End If
        
         
        
        cnt = cnt + 1
    Next node
    Rec80E = cnt
 End If
End Function

'Added by Shrutika- 23/01/2026 AY-26-27
'lastFiledITR.Schedule80EE.Schedule80EEDtls
Function ImportSchedule_80EE_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80EE As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EE, IFSC_80EE, BankName_80EE, PANof_80EE, LoanAcctNum_80EE, LoanDate_80EE, TotalLoanAmt_80EE, LoanOutStanding_80EE, Intrst_80EE As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EE, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

 
 
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80EE = jsonObject("lastFiledITR")
 
If NodeList80EE.exists("schedule80EE") Then
 
Set Nodelist = jsonObject("lastFiledITR")("schedule80EE")("schedule80EEDtls")

    LoanfrmBankOrInstitute_80EE = Sheet17.Range("LoanfrmBankOrInstitute.80EE").Column
    BankName_80EE = Sheet17.Range("bankName.80EE").Column
    LoanAcctNum_80EE = Sheet17.Range("loanAccNum.80EE").Column
    LoanDate_80EE = Sheet17.Range("loanDate.80EE").Column
    TotalLoanAmt_80EE = Sheet17.Range("loanAmt.80EE").Column
     
    
    TotalExRow = Sheet17.Range("LoanfrmBankOrInstitute.80EE").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet17.Range("LoanfrmBankOrInstitute.80EE").Locked = False Then
            Sheet17.Range("LoanfrmBankOrInstitute.80EE").ClearContents
        End If
        If Sheet17.Range("IFSC.80EE").Locked = False Then
            Sheet17.Range("IFSC.80EE").ClearContents
        End If
        If Sheet17.Range("bankName.80EE").Locked = False Then
            Sheet17.Range("bankName.80EE").ClearContents
        End If
        If Sheet17.Range("PAN.80EE").Locked = False Then
            Sheet17.Range("PAN.80EE").ClearContents
        End If
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
            Dim LoanTknFrom_80EE, LoanTknFrom_80EE_pfl
            
            LoanTknFrom_80EE_pfl = node("loanTknFrom")
            If UCase(LoanTknFrom_80EE_pfl) = UCase("B") Then
                LoanTknFrom_80EE = "Bank"
            ElseIf UCase(LoanTknFrom_80EE_pfl) = UCase("I") Then
                LoanTknFrom_80EE = "Institution"
            Else
                LoanTknFrom_80EE = "(Select)"
            End If
            Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80EE).Value = LoanTknFrom_80EE
        End If
         
        If Sheet17.Cells(rowcount, BankName_80EE).Locked = False Then
            Sheet17.Cells(rowcount, BankName_80EE).Value = node("bankOrInstnName")
        End If
         
        If Sheet17.Cells(rowcount, LoanAcctNum_80EE).Locked = False Then
            Sheet17.Cells(rowcount, LoanAcctNum_80EE).Value = node("loanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("dateofLoan")
            If strDate <> "" Then
                strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2) & "/" & Mid(strDate, 1, 4)
            End If
        
        If Sheet17.Cells(rowcount, LoanDate_80EE).Locked = False Then
            Sheet17.Cells(rowcount, LoanDate_80EE).Value = strDate
            strDate = ""
        End If
        If Sheet17.Cells(rowcount, TotalLoanAmt_80EE).Locked = False Then
            Sheet17.Cells(rowcount, TotalLoanAmt_80EE).Value = node("totalLoanAmt")
        End If
        
         
        cnt = cnt + 1
    Next node
    Rec80EE = cnt
End If
End Function


'Added by Shrutika- 23/01/2026 AY-26-27
'lastFiledITR.schedule80EEA.schedule80EEADtls

Function ImportSchedule_80EEA_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80EEA As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEA, IFSC_80EEA, BankName_80EEA, PANof_80EEA, LoanAcctNum_80EEA, LoanDate_80EEA, TotalLoanAmt_80EEA, LoanOutStanding_80EEA, Intrst_80EEA As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEA, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

Set jsonObject = ParseJson(jsonText)
Set NodeList80EEA = jsonObject("lastFiledITR")
 
If NodeList80EEA.exists("schedule80EEA") Then
  
        
Dim PropStmpDtyVal_80EEA
PropStmpDtyVal_80EEA = jsonObject("lastFiledITR")("schedule80EEA")("propStmpDtyVal")

If PropStmpDtyVal_80EEA <> "" Then
    If Sheet17.Range("Stampduty.80EEA").Locked = False Then
            Sheet17.Range("Stampduty.80EEA").Value = PropStmpDtyVal_80EEA
        End If
End If



Set Nodelist = jsonObject("lastFiledITR")("schedule80EEA")("schedule80EEADtls")

    LoanfrmBankOrInstitute_80EEA = Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Column
    BankName_80EEA = Sheet17.Range("bankName.80EEA").Column
    LoanAcctNum_80EEA = Sheet17.Range("loanAccNum.80EEA").Column
    LoanDate_80EEA = Sheet17.Range("loanDate.80EEA").Column
    TotalLoanAmt_80EEA = Sheet17.Range("loanAmt.80EEA").Column
     
    
    TotalExRow = Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Locked = False Then
            Sheet17.Range("LoanfrmBankOrInstitute.80EEA").ClearContents
        End If
        
        If Sheet17.Range("bankName.80EEA").Locked = False Then
            Sheet17.Range("bankName.80EEA").ClearContents
        End If
        
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
            Dim LoanTknFrom_80EEA, LoanTknFrom_80EEA_pfl
            
            LoanTknFrom_80EEA_pfl = node("loanTknFrom")
            If UCase(LoanTknFrom_80EEA_pfl) = UCase("B") Then
                LoanTknFrom_80EEA = "Bank"
            ElseIf UCase(LoanTknFrom_80EEA_pfl) = UCase("I") Then
                LoanTknFrom_80EEA = "Institution"
            Else
                LoanTknFrom_80EEA = "(Select)"
            End If
            Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80EEA).Value = LoanTknFrom_80EEA
        End If
        
        If Sheet17.Cells(rowcount, BankName_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, BankName_80EEA).Value = node("bankOrInstnName")
        End If
         
        If Sheet17.Cells(rowcount, LoanAcctNum_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, LoanAcctNum_80EEA).Value = node("loanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("dateofLoan")
            If strDate <> "" Then
                strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2) & "/" & Mid(strDate, 1, 4)
            End If
        
        If Sheet17.Cells(rowcount, LoanDate_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, LoanDate_80EEA).Value = strDate
            strDate = ""
        End If
        If Sheet17.Cells(rowcount, TotalLoanAmt_80EEA).Locked = False Then
            Sheet17.Cells(rowcount, TotalLoanAmt_80EEA).Value = node("totalLoanAmt")
        End If
         
        cnt = cnt + 1
    Next node
    Rec80EEA = cnt
 End If
End Function

'Added by Shrutika- 23/01/2026 AY-26-27
'lastFiledITR.schedule80EEB.schedule80EEBDtls

Function ImportSchedule_80EEB_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary, NodeList80EEB As Object
Dim node, Nodelist As Object
Dim LoanfrmBankOrInstitute_80EEB, IFSC_80EEB, BankName_80EEB, Vehicle_value_80EEB, VehicleRegNum_80EEB, PANof_80EEB, LoanAcctNum_80EEB, LoanDate_80EEB, TotalLoanAmt_80EEB, LoanOutStanding_80EEB, Intrst_80EEB As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, Rec80EEB, rowcount, cnt As Long
Dim strDate As String
Dim YYYY, MM, DD As String

 
 
    
Set jsonObject = ParseJson(jsonText)
Set NodeList80EEB = jsonObject("lastFiledITR")
 
If NodeList80EEB.exists("schedule80EEB") Then
  
Set Nodelist = jsonObject("lastFiledITR")("schedule80EEB")("schedule80EEBDtls")

    LoanfrmBankOrInstitute_80EEB = Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Column
    BankName_80EEB = Sheet17.Range("bankName.80EEB").Column
    LoanAcctNum_80EEB = Sheet17.Range("loanAccNum.80EEB").Column
    LoanDate_80EEB = Sheet17.Range("loanDate.80EEB").Column
    TotalLoanAmt_80EEB = Sheet17.Range("loanAmt.80EEB").Column
    VehicleRegNum_80EEB = Sheet17.Range("VehicleRegNum.80EEB").Column
     
    
    
    TotalExRow = Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        If Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Locked = False Then
            Sheet17.Range("LoanfrmBankOrInstitute.80EEB").ClearContents
        End If
         
        If Sheet17.Range("bankName.80EEB").Locked = False Then
            Sheet17.Range("bankName.80EEB").ClearContents
        End If
        
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
            Dim LoanTknFrom_80EEB, LoanTknFrom_80EEB_pfl
            LoanTknFrom_80EEB_pfl = node("loanTknFrom")
            
            If UCase(LoanTknFrom_80EEB_pfl) = UCase("B") Then
                LoanTknFrom_80EEB = "Bank"
            ElseIf UCase(LoanTknFrom_80EEB_pfl) = UCase("I") Then
                LoanTknFrom_80EEB = "Institution"
            Else
                LoanTknFrom_80EEB = "(Select)"
            End If
            Sheet17.Cells(rowcount, LoanfrmBankOrInstitute_80EEB).Value = LoanTknFrom_80EEB
        End If
          
        If Sheet17.Cells(rowcount, BankName_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, BankName_80EEB).Value = node("bankOrInstnName")
        End If
         
        If Sheet17.Cells(rowcount, LoanAcctNum_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, LoanAcctNum_80EEB).Value = node("loanAccNoOfBankOrInstnRefNo")
        End If
                
                strDate = node("dateofLoan")
            If strDate <> "" Then
                strDate = Trim(Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2) & "/" & Mid(strDate, 1, 4))
            End If
        
        If Sheet17.Cells(rowcount, LoanDate_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, LoanDate_80EEB).Value = strDate
            strDate = ""
        End If
        If Sheet17.Cells(rowcount, TotalLoanAmt_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, TotalLoanAmt_80EEB).Value = node("totalLoanAmt")
        End If
         
        If Sheet17.Cells(rowcount, VehicleRegNum_80EEB).Locked = False Then
            Sheet17.Cells(rowcount, VehicleRegNum_80EEB).Value = node("vehicleRegNo")
        End If
         
        cnt = cnt + 1
    Next node
    Rec80EEB = cnt
End If
End Function

'Added by Shrutika- 20/01/2026 AY-26-27

Function ImportscheduleEA10_13A_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary   As Object
Dim EA10_13A_placeofwork, EA10_13A_placeofwork_pfl As Variant
Dim EA10_13A_actlHRARecv, EA10_13A_actlHRARecv_pfl As Variant

Set jsonObject = ParseJson(jsonText)

'Place of Residence
'lastFiledITR.scheduleEA10_13A.placeofwork
EA10_13A_placeofwork = jsonObject("lastFiledITR")("scheduleEA10_13A")("placeofwork")

If EA10_13A_placeofwork <> "" And Sheet18.Range("Sch10of13A_PlaceofWrk").Locked = False Then
        If EA10_13A_placeofwork = "1" Then
               EA10_13A_placeofwork_pfl = "1. Metro"
        ElseIf EA10_13A_placeofwork = "2" Then
               EA10_13A_placeofwork_pfl = "2. Non-Metro"
        Else
               EA10_13A_placeofwork_pfl = "(Select)"
        End If
    Sheet18.Range("Sch10of13A_PlaceofWrk").Value = EA10_13A_placeofwork_pfl
End If

'Actual HRA received (A)
'form24q.ScheduleEA10_13A.actlHRARecv
 EA10_13A_actlHRARecv_pfl = jsonObject("form24q")("scheduleEA10_13A")("actlHRARecv")
If EA10_13A_actlHRARecv_pfl <> "" And Sheet18.Range("Sch10of13A_PlaceofWrk").Locked = False Then
Sheet18.Range("Sch10of13A_ActlHRArecivedA").Value = EA10_13A_actlHRARecv_pfl
End If

End Function




Function ImportScheduleTCS_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object

Dim TANColNo, DEDNameColNo, Amtfrom26ASColNo, TaxColNo, ClaimColNo, CollectionYearColNo As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt, RecTCS  As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim strDate As String
Dim YYYY, MM, DD As String
    
 Set jsonObject = ParseJson(jsonText)
  
   Set Nodelist = jsonObject("form26as")("scheduleTCS")("tcs")
    
    TANColNo = Sheet11.Range("TCS.TAN").Column
    DEDNameColNo = Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").Column
    Amtfrom26ASColNo = Sheet11.Range("TCS.AmountCollected").Column
    CollectionYearColNo = Sheet11.Range("TCS.CollectionYear").Column
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
       
            If Trim(node("employerOrDeductorOrCollectDetl")("tan")) <> "" And Sheet11.Cells(rowcount, TANColNo).Locked = False Then
                Sheet11.Cells(rowcount, TANColNo).Value = node("employerOrDeductorOrCollectDetl")("tan")
            End If
            
            If Trim(node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")) <> "" And Sheet11.Cells(rowcount, DEDNameColNo).Locked = False Then
                Sheet11.Cells(rowcount, DEDNameColNo).Value = node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")
            End If
            
            If Trim(node("amtfrm26AS")) <> "" And Sheet11.Cells(rowcount, Amtfrom26ASColNo).Locked = False Then
                Sheet11.Cells(rowcount, Amtfrom26ASColNo).Value = node("amtfrm26AS")
            End If
            
            If Trim(node("CollectedYr")) <> "" And Sheet11.Cells(rowcount, CollectionYearColNo).Locked = False Then
                Sheet11.Cells(rowcount, CollectionYearColNo).Value = node("CollectedYr")
            End If
            
            If Trim(node("amtTCSClaimedThisYear")) <> "" And Sheet11.Cells(rowcount, TaxColNo).Locked = False Then
                Sheet11.Cells(rowcount, TaxColNo).Value = node("amtTCSClaimedThisYear")
            End If
            
            If Trim(node("amtTCSClaimedThisYear")) <> "" And Sheet11.Cells(rowcount, ClaimColNo).Locked = False Then
                Sheet11.Cells(rowcount, ClaimColNo).Value = node("amtTCSClaimedThisYear")
            End If
                
        cnt = cnt + 1
    Next node
End If
    RecTCS = cnt
End Function

Function ImportScheduleIT_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim BSRCodeColNo, DateDepColNo, SrlNoChallanColNo, AmtColNo As Variant
Dim TotalXMLRow, RecTDS1, rowcount, cnt As Long
Dim TotalDiffRow As Long
Dim TotalExRow As Long
Dim strDate As String
Dim YYYY, MM, DD As String


Set jsonObject = ParseJson(jsonText)
'change2 start
    Set Nodelist = jsonObject("form26as")("taxPayments")("taxPayment")
'change2 end

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
        
            If Trim(node("bsrCode")) <> "" And Sheet2.Cells(rowcount, BSRCodeColNo).Locked = False Then
                Sheet2.Cells(rowcount, BSRCodeColNo).Value = node("bsrCode")
            End If
            
            strDate = node("dateDep")
            If strDate <> "" Then
                YYYY = Mid(strDate, 1, 4)
                MM = Mid(strDate, 6, 2)
                DD = Mid(strDate, 9, 2)
                strDate = DD & "/" & MM & "/" & YYYY
            End If
            
            If strDate <> "" And Sheet2.Cells(rowcount, DateDepColNo).Locked = False Then
                Sheet2.Cells(rowcount, DateDepColNo).Value = strDate
                strDate = ""
            End If
            
            If Trim(node("srlNoOfChaln")) <> "" And Sheet2.Cells(rowcount, SrlNoChallanColNo).Locked = False Then
                Sheet2.Cells(rowcount, SrlNoChallanColNo).Value = node("srlNoOfChaln")
            End If
            
            If Trim(node("amt")) <> "" And Sheet2.Cells(rowcount, AmtColNo).Locked = False Then
                Sheet2.Cells(rowcount, AmtColNo).Value = node("amt")
            End If
      
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Function

Function ImportTDSonSalary_pfl(jsonText As String)
On Error Resume Next
    
    Dim jsonObject, jsonDictionary As Object
    Dim node, Nodelist As Object
    Dim rowcount, cnt As Variant
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    
    Dim TANNoEmployer, TDSNameOfEmployer, TDSIncomeCharge, TDSTotalTax As Long
    Dim TotalExRow As Long

    Set jsonObject = ParseJson(jsonText)
    
    Set Nodelist = jsonObject("form26as")("tdsOnSalaries")("tdsOnSalary")
    
    
    TANNoEmployer = Sheet2.Range("TDSal.TAN").Column
    TDSNameOfEmployer = Sheet2.Range("TDSal.EmployerOrDeductorOrCollecterName").Column
    TDSIncomeCharge = Sheet2.Range("TDSal.IncChrgSalary").Column
    TDSTotalTax = Sheet2.Range("TDSal.TotalTDSSalary").Column
    
    TotalExRow = Range("TDSal.TAN").Rows.count
    
    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet2.Range("TDSal.TAN").ClearContents
        Sheet2.Range("TDSal.EmployerOrDeductorOrCollecterName").ClearContents
        Sheet2.Range("TDSal.IncChrgSalary").ClearContents
        Sheet2.Range("TDSal.TotalTDSSalary").ClearContents
    End If
    
    If (TotalDiffRow > 0) Then
        AddDiffRows_TDS1 (TotalDiffRow)
    End If
    
    rowcount = getRowNo(Sheet2.Range("TDSal.TAN").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
            rowcount = rowcount + 1
            
            If Trim(node("employerOrDeductorOrCollectDetl")("tan")) <> "" And Sheet2.Cells(rowcount, TANNoEmployer).Locked = False Then
                Sheet2.Cells(rowcount, TANNoEmployer).Value = node("employerOrDeductorOrCollectDetl")("tan")
            End If
            
            If Trim(node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")) <> "" And Sheet2.Cells(rowcount, TDSNameOfEmployer).Locked = False Then
                Sheet2.Cells(rowcount, TDSNameOfEmployer).Value = node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")
            End If
            
            If Trim(node("incChrgSal")) <> "" And Sheet2.Cells(rowcount, TDSIncomeCharge).Locked = False Then
                Sheet2.Cells(rowcount, TDSIncomeCharge).Value = node("incChrgSal")
            End If
            
            If Trim(node("totalTDSSal")) <> "" And Sheet2.Cells(rowcount, TDSTotalTax).Locked = False Then
                Sheet2.Cells(rowcount, TDSTotalTax).Value = node("totalTDSSal")
            End If
            
        
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Function

'import json for TDSonOthThanSals
Function ImportTDSOthThanSals_pfl(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object
Dim TANColNo, FYColNo, TDSOtherColNo, TDSDeductedColNo, TDSClaimedColNo, TDSGrossAmountColNo, TDSHeadColNo, DeductedYearColNo As Variant
Dim TotalExRow, TotalXMLRow, TotalDiffRow, RecTDS1, cnt, rowcount As Long
Dim SecTDSColNo, TDSSection_TDS2, TDSSection As Variant '29/04/2025


Set jsonObject = ParseJson(jsonText)

'change1 start
Set Nodelist = jsonObject("form26as")("tdsOnOthThanSals")("tdSonOthThanSal")
    TANColNo = Sheet2.Range("TDSoth.TAN").Column
    FYColNo = Sheet2.Range("TDSoth.EmployerOrDeductorOrCollecterName").Column
    'Malli--V0.3_AY_2025_26_29/04/2025
     SecTDSColNo = Sheet2.Range("TDsOthr.SectionTDS").Column
    '-------------------------------
    TDSOtherColNo = Sheet2.Range("TDSoth.AmountDeducted").Column
    DeductedYearColNo = Sheet2.Range("TDSoth.DeductedYear").Column
    TDSDeductedColNo = Sheet2.Range("TDSoth.TotTDSOnAmtPaid").Column
    TDSClaimedColNo = Sheet2.Range("TDSoth.6income").Column
'    TDSGrossAmountColNo = Sheet2.Range("TDsOthr.grossamount").Column
'    TDSHeadColNo = Sheet2.Range("TDsOthr.headincome").Column
    
    TotalExRow = Range("TDSoth.TAN").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
     Sheet2.Range("TDSoth.TAN").ClearContents
     Sheet2.Range("TDSoth.EmployerOrDeductorOrCollecterName").ClearContents
     Sheet2.Range("TDsOthr.SectionTDS").ClearContents '29/04/2025
     Sheet2.Range("TDSoth.AmountDeducted").ClearContents
     Sheet2.Range("TDSoth.DeductedYear").ClearContents
     Sheet2.Range("TDSoth.TotTDSOnAmtPaid").ClearContents
     Sheet2.Range("TDSoth.6income").ClearContents
'     Sheet2.Range("TDsOthr.grossamount").ClearContents
'     Sheet2.Range("TDsOthr.headincome").ClearContents
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
            
            If Trim(node("employerOrDeductorOrCollectDetl")("tan")) <> "" And Sheet2.Cells(rowcount, TANColNo).Locked = False Then
                Sheet2.Cells(rowcount, TANColNo).Value = node("employerOrDeductorOrCollectDetl")("tan")
            End If
            
            If Trim(node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")) <> "" And Sheet2.Cells(rowcount, FYColNo).Locked = False Then
                Sheet2.Cells(rowcount, FYColNo).Value = node("employerOrDeductorOrCollectDetl")("employerOrDeductorOrCollecterName")
            End If
            
     'Malli-----29/04/2025
     'form26as.tdsOnOthThanSals.tdSonOthThanSal.sectionCode
     
     TDSSection_TDS2 = node("sectionCode")
     'Debug.Print TDSSection_TDS2
     If TDSSection_TDS2 <> "" Then
            If UCase(TDSSection_TDS2) = UCase("92A") Then
            TDSSection = "192-Salary-Payment to Government employees other than Indian Government employees"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("92B") Then
            TDSSection = "192-Salary-Payment to employees other than Government employees"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("92C") Then
            TDSSection = "192-Salary-Payment to Indian Government employees"
            
            '13/05/2025
            'ElseIf TDSSection_TDS2 = "192A" Then
            'ElseIf TDSSection_TDS2 = "2AA" Or TDSSection_TDS2 = "2aa" Or UCase(TDSSection_TDS2) = "2AA" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("2AA") Then
            TDSSection = "192A-TDS on PF withdrawal"
            '-----------------------------------
            ElseIf UCase(TDSSection_TDS2) = UCase("193") Then
            TDSSection = "193-Interest on Securities"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("194") Then
            TDSSection = "194-Dividends"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94A") Then
            TDSSection = "194A-Interest other than 'Interest on securities'"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94B") Then
            TDSSection = "194B-Winning from lottery or crossword puzzle"
        
            '13/05/2025
            'ElseIf TDSSection_TDS2 = "94BA" Then
            'ElseIf TDSSection_TDS2 = "9BA" Or TDSSection_TDS2 = "9ba" Or UCase(TDSSection_TDS2) = "9BA" Then
            ElseIf UCase(TDSSection_TDS2) = UCase("9BA") Then
            TDSSection = "194BA-Winnings from online games"
            '-----------------------------
            
            ElseIf UCase(TDSSection_TDS2) = UCase("4BB") Then
            TDSSection = "194BB-Winning from horse race"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94C") Then
            TDSSection = "194C-Payments to contractors and sub-contractors"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94D") Then
            TDSSection = "194D-Insurance commission"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("4DA") Then
            TDSSection = "194DA-Payment in respect of life insurance policy"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94E") Then
            TDSSection = "194E-Payments to non-resident sportsmen or sports associations"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("4EE") Then
            TDSSection = "194EE-Payments in respect of deposits under National Savings"
             
            
            '13/05/2025---------------------
            'ElseIf TDSSection_TDS2 = "4F" Then
            'ElseIf TDSSection_TDS2 = "94F" Or TDSSection_TDS2 = "94f" Or UCase(TDSSection_TDS2) = "94F" Then
            ElseIf UCase(TDSSection_TDS2) = UCase("94F") Then
            TDSSection = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
            
            'ElseIf TDSSection_TDS2 = "4G" Then
            'ElseIf TDSSection_TDS2 = "94G" Or TDSSection_TDS2 = "94g" Or UCase(TDSSection_TDS2) = "94G" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("94G") Then
            TDSSection = "194G-Commission, price, etc. on sale of lottery tickets"
            
            'ElseIf TDSSection_TDS2 = "4H" Then
             'ElseIf TDSSection_TDS2 = "94H" Or TDSSection_TDS2 = "94h" Or UCase(TDSSection_TDS2) = "94H" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("94H") Then
            TDSSection = "194H-Commission or brokerage"
            
            'ElseIf TDSSection_TDS2 = "4-IA" Then
             'ElseIf TDSSection_TDS2 = "4IA" Or TDSSection_TDS2 = "4ia" Or UCase(TDSSection_TDS2) = "4IA" Then
              ElseIf UCase(TDSSection_TDS2) = UCase("4IA") Then
            TDSSection = "194I(a)-Rent on hiring of plant and machinery"
            
            'ElseIf TDSSection_TDS2 = "4-IB" Then
             'ElseIf TDSSection_TDS2 = "4IB" Or TDSSection_TDS2 = "4ib" Or UCase(TDSSection_TDS2) = "4IB" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("4IB") Then
            TDSSection = "194I(b)-Rent on other than plant and machinery"
            
            'ElseIf TDSSection_TDS2 = "4IA" Then
            'ElseIf TDSSection_TDS2 = "9IA" Or TDSSection_TDS2 = "9ia" Or UCase(TDSSection_TDS2) = "9IA" Then
            ElseIf UCase(TDSSection_TDS2) = UCase("9IA") Then
            TDSSection = "194IA-TDS on Sale of immovable property"
            
            'ElseIf TDSSection_TDS2 = "4IB" Then
            'ElseIf TDSSection_TDS2 = "9IB" Or TDSSection_TDS2 = "9ib" Or UCase(TDSSection_TDS2) = "9IB" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("9IB") Then
            TDSSection = "194IB-Payment of rent by certain individuals or Hindu undivided"
            '----------------------------------------------
            
            ElseIf UCase(TDSSection_TDS2) = UCase("4IC") Then
            TDSSection = "194IC-Payment under specified agreement"
           
            
            'ElseIf TDSSection_TDS2 = "94J-A" Then
            'ElseIf TDSSection_TDS2 = "4JA" Or TDSSection_TDS2 = "4ja" Or UCase(TDSSection_TDS2) = "4JA" Then
            ElseIf UCase(TDSSection_TDS2) = UCase("4JA") Then
            TDSSection = "194J(a)-Fees for technical services"
            
            'ElseIf TDSSection_TDS2 = "94J-B" Then
            'ElseIf TDSSection_TDS2 = "4JB" Or TDSSection_TDS2 = "4jb" Or UCase(TDSSection_TDS2) = "4JB" Then
            ElseIf UCase(TDSSection_TDS2) = UCase("4JB") Then
            TDSSection = "194J(b)-Fees for professional  services or royalty etc"
            '---------------------------------------
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94K") Then
            TDSSection = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("4LA") Then
            TDSSection = "194LA-Payment of compensation on acquisition of certain immovable"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("4LB") Then
            TDSSection = "194LB-Income by way of Interest from Infrastructure Debt fund"
            
            '13/05/2025
            'ElseIf TDSSection_TDS2 = "4LC1" Then
            'ElseIf TDSSection_TDS2 = "LC1" Or TDSSection_TDS2 = "lc1" Or UCase(TDSSection_TDS2) = "LC1" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("LC1") Then
            TDSSection = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
            
            'ElseIf TDSSection_TDS2 = "4LC2" Then
             'ElseIf TDSSection_TDS2 = "LC2" Or TDSSection_TDS2 = "lc2" Or UCase(TDSSection_TDS2) = "LC2" Then
              ElseIf UCase(TDSSection_TDS2) = UCase("LC2") Then
            TDSSection = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
            
            'ElseIf TDSSection_TDS2 = "4LC3" Then
             'ElseIf TDSSection_TDS2 = "LC3" Or TDSSection_TDS2 = "lc3" Or UCase(TDSSection_TDS2) = "LC3" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("LC3") Then
            TDSSection = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
            
            ''Malli_AY_2026_27  SIT-117189 09/04/2026
'            ElseIf TDSSection_TDS2 = "4BA1" Then
            ElseIf TDSSection_TDS2 = "BA1" Then
            TDSSection = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"

            ''Malli_AY_2026_27  SIT-117189 09/04/2026
'            ElseIf TDSSection_TDS2 = "4BA2" Then
            ElseIf TDSSection_TDS2 = "BA2" Then
            TDSSection = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"




'            ElseIf TDSSection_TDS2 = "LBA1" Then
'            TDSSection = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
'
'            ElseIf TDSSection_TDS2 = "LBA2" Then
'            TDSSection = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
'
           ' ElseIf TDSSection_TDS2 = "LBA3" Then
            'ElseIf TDSSection_TDS2 = "BA3" Or TDSSection_TDS2 = "ba3" Or UCase(TDSSection_TDS2) = "BA3" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("BA3") Then
            TDSSection = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
            '-----------------------------
            ElseIf UCase(TDSSection_TDS2) = UCase("LBB") Then
            TDSSection = "194LBB-Income in respect of units of investment fund"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94R") Then
            TDSSection = "194R-Benefits or perquisites of business or profession"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94S") Then
            TDSSection = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
            
           '13/05/2025
           ' ElseIf TDSSection_TDS2 = "94B-P" Then
            'ElseIf TDSSection_TDS2 = "4BP" Or TDSSection_TDS2 = "4bp" Or UCase(TDSSection_TDS2) = "4BP" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("4BP") Then
            TDSSection = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
            
            'ElseIf TDSSection_TDS2 = "94R-P" Then
           ' ElseIf TDSSection_TDS2 = "4RP" Or TDSSection_TDS2 = "4rp" Or UCase(TDSSection_TDS2) = "4RP" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("4RP") Then
            TDSSection = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
            
            'ElseIf TDSSection_TDS2 = "94S-P" Then
            'ElseIf TDSSection_TDS2 = "4SP" Or TDSSection_TDS2 = "4sp" Or UCase(TDSSection_TDS2) = "4SP" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("4SP") Then
            TDSSection = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
            '-------------------------------
            
            ElseIf UCase(TDSSection_TDS2) = UCase("LBC") Then
            TDSSection = "194LBC-Income in respect of investment in securitization trust"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("4LD") Then
            TDSSection = "194LD-TDS on interest on bonds / government securities"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94M") Then
            TDSSection = "194M-Payment of certain sums by certain individuals or HUF"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94N") Then
            TDSSection = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
            
            '13/05/2025
            'ElseIf TDSSection_TDS2 = "94N-F" Then
            'ElseIf TDSSection_TDS2 = "4NF" Or TDSSection_TDS2 = "4nf" Or UCase(TDSSection_TDS2) = "4NF" Then
            ElseIf UCase(TDSSection_TDS2) = UCase("4NF") Then
            TDSSection = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
            
            'ElseIf TDSSection_TDS2 = "94N-C" Then
            'ElseIf TDSSection_TDS2 = "4NC" Or TDSSection_TDS2 = "4nc" Or UCase(TDSSection_TDS2) = "4NC" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("4NC") Then
            TDSSection = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
            
            'ElseIf TDSSection_TDS2 = "94N-FT" Then
            'ElseIf TDSSection_TDS2 = "NFT" Or TDSSection_TDS2 = "nft" Or UCase(TDSSection_TDS2) = "NFT" Then
             ElseIf UCase(TDSSection_TDS2) = UCase("NFT") Then
            TDSSection = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
            '-----------------------------------------
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94O") Then
            TDSSection = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94P") Then
            TDSSection = "194P-Deduction of tax in case of specified senior citizen"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("94Q") Then
            TDSSection = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("195") Then
            TDSSection = "195-Other sums payable to a non-resident"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("96A") Then
            TDSSection = "196A-Income in respect of units of non-residents"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("96B") Then
            TDSSection = "196B-Payments in respect of units to an offshore fund"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("96C") Then
            TDSSection = "196C-Income from foreign currency bonds or shares of Indian"
            
            ElseIf UCase(TDSSection_TDS2) = UCase("96D") Then
            TDSSection = "196D-Income of foreign institutional investors from securities"
            
              ' '
            
            '13/05/2025
           'ElseIf TDSSection_TDS2 = "96DA" Then
            'ElseIf TDSSection_TDS2 = "6DA" Or TDSSection_TDS2 = "6da" Or UCase(TDSSection_TDS2) = "6DA" Then
            ElseIf UCase(TDSSection_TDS2) = UCase("6DA") Then
            TDSSection = "196D(1A)-Income of specified fund from securities"
            
            'ElseIf TDSSection_TDS2 = "94BA-P" Then
            'ElseIf TDSSection_TDS2 = "BAP" Or TDSSection_TDS2 = "bap" Or UCase(TDSSection_TDS2) = "BAP" Then
            ElseIf UCase(TDSSection_TDS2) = UCase("BAP") Then
            TDSSection = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
            '-------------------------------------------
            Else
               'TDSSection = ""
               TDSSection = "(Select)"
        End If
        
            If TDSSection <> "" And Sheet2.Cells(rowcount, SecTDSColNo).Locked = False Then
                Sheet2.Cells(rowcount, SecTDSColNo).Value = TDSSection
                TDSSection_TDS2 = ""
            End If
       
     End If
     '--------------------
            
            
            If Trim(node("grossAmount")) <> "" And Sheet2.Cells(rowcount, TDSOtherColNo).Locked = False Then
                Sheet2.Cells(rowcount, TDSOtherColNo).Value = node("grossAmount")
            End If
            
            'Malli 03/07/2026 updated as per enhancement tds deducted year should by default 2025
'            If Trim(node("deductedYr")) <> "" And Sheet2.Cells(rowcount, DeductedYearColNo).Locked = False Then
'                Sheet2.Cells(rowcount, DeductedYearColNo).Value = node("deductedYr")
'            End If
            
            
            If Sheet2.Cells(rowcount, DeductedYearColNo).Locked = False Then
                     Dim deductedYr_tds2, pfl_deductedYr_tds2
                         deductedYr_tds2 = Trim(node("deductedYr"))
            If deductedYr_tds2 <> "" Then
            
                    If deductedYr_tds2 = "2008" Then
                       pfl_deductedYr_tds2 = "2008-09"
                       
                    ElseIf deductedYr_tds2 = "2009" Then
                       pfl_deductedYr_tds2 = "2009-10"
                       
                    ElseIf deductedYr_tds2 = "2010" Then
                       pfl_deductedYr_tds2 = "2010-11"
                       
                    ElseIf deductedYr_tds2 = "2011" Then
                       pfl_deductedYr_tds2 = "2011-12"
                       
                    ElseIf deductedYr_tds2 = "2012" Then
                       pfl_deductedYr_tds2 = "2012-13"
                       
                    ElseIf deductedYr_tds2 = "2013" Then
                       pfl_deductedYr_tds2 = "2013-14"
                       
                    ElseIf deductedYr_tds2 = "2014" Then
                       pfl_deductedYr_tds2 = "2014-15"
                       
                    ElseIf deductedYr_tds2 = "2015" Then
                       pfl_deductedYr_tds2 = "2015-16"
                       
                    ElseIf deductedYr_tds2 = "2016" Then
                       pfl_deductedYr_tds2 = "2016-17"
                       
                    ElseIf deductedYr_tds2 = "2017" Then
                       pfl_deductedYr_tds2 = "2017-18"
                       
                    ElseIf deductedYr_tds2 = "2018" Then
                       pfl_deductedYr_tds2 = "2018-19"
                       
                    ElseIf deductedYr_tds2 = "2019" Then
                       pfl_deductedYr_tds2 = "2019-20"
                       
                    ElseIf deductedYr_tds2 = "2020" Then
                       pfl_deductedYr_tds2 = "2020-21"
                       
                    ElseIf deductedYr_tds2 = "2021" Then
                       pfl_deductedYr_tds2 = "2021-22"
                       
                    ElseIf deductedYr_tds2 = "2022" Then
                       pfl_deductedYr_tds2 = "2022-23"
                       
                    ElseIf deductedYr_tds2 = "2023" Then
                       pfl_deductedYr_tds2 = "2023-24"
                       
                    ElseIf deductedYr_tds2 = "2024" Then
                       pfl_deductedYr_tds2 = "2024-25"
                       
                    ElseIf deductedYr_tds2 = "2025" Then
                       pfl_deductedYr_tds2 = "2025-26"
                       
                    Else
                        pfl_deductedYr_tds2 = "2025-26"
                    End If
               
                   Sheet2.Cells(rowcount, DeductedYearColNo).Value = pfl_deductedYr_tds2
            Else
                    Sheet2.Cells(rowcount, DeductedYearColNo).Value = "2025-26"
            End If
            End If
            
            
            '------------
            
            If Trim(node("taxDeductCreditDtls")("taxDeductedOwnHands")) <> "" And Sheet2.Cells(rowcount, TDSDeductedColNo).Locked = False Then
                Sheet2.Cells(rowcount, TDSDeductedColNo).Value = node("taxDeductCreditDtls")("taxDeductedOwnHands")
            End If
            
            If Trim(node("taxDeductCreditDtls")("taxClaimedOwnHands")) <> "" And Sheet2.Cells(rowcount, TDSClaimedColNo).Locked = False Then
                Sheet2.Cells(rowcount, TDSClaimedColNo).Value = node("taxDeductCreditDtls")("taxClaimedOwnHands")
            End If
'change1 end
'            Sheet2.Cells(rowcount, TDSGrossAmountColNo).Value = Node("grossAmount")
'
'            Dim HeadOfIncome As Variant
'            HeadOfIncome = Node("headOfIncome")
'
'            If HeadOfIncome = "BP" Then
'            HeadOfIncome = "Income from Business & Profession"
'            ElseIf HeadOfIncome = "HP" Then
'            HeadOfIncome = "Income from House Property"
'            ElseIf HeadOfIncome = "OS" Then
'            HeadOfIncome = "Income from Other Source"
'            ElseIf HeadOfIncome = "EI" Then
'            HeadOfIncome = "Exempt Income"
'            ElseIf HeadOfIncome = "NA" Then
'            HeadOfIncome = "Not applicable (only in case TDS is deducted u/s 194N)"
'            End If
'
'
'            Sheet2.Cells(rowcount, TDSHeadColNo).Value = HeadOfIncome
            
            
        cnt = cnt + 1
    Next node
  End If
    RecTDS1 = cnt
End Function
Sub ImportScheduleTDS3Dtls_pfl(jsonText As String)
On Error Resume Next

    Dim jsonObject, jsonDictionary As Object
    Dim RecTDS11, RecTDS21, cnt As Long
    Dim Nodelist, node As Object
    Dim TotalXMLRow, TotalExRow, TotalDiffRow, rowcount As Long
    Dim PANColNo, AadhaarColNo, FYColNo, TDSOtherColNo, TDSDeductedColNo, TDSClaimedColNo, TDSyearColNo, TDSGrossAmountColNo, TDSHeadColNo As Variant
    Dim SecTDSColNo, TDSSection_TDS3, TDSSection As Variant '29/04/2025


    Set jsonObject = ParseJson(jsonText)
    Set Nodelist = jsonObject("form26as")("scheduleTDS3Dtls")("tds3Details")
  
    
    'TDSCreditColNo = Sheet2.Range("TDS2.TdsCredit").Column
    PANColNo = Sheet2.Range("TDS26QB.PAN").Column
    AadhaarColNo = Sheet2.Range("TDS26QB.Aadhar_Number").Column
    FYColNo = Sheet2.Range("TDS26QB.EmployerOrDeductorName").Column
    
    SecTDSColNo = Sheet2.Range("TDsOthr2.SectionTDSDeducted").Column  'Malli-29/04/205
    
    TDSOtherColNo = Sheet2.Range("TDS26QB.AmountDeducted").Column
    TDSyearColNo = Sheet2.Range("TDS26QB.DeductedYear").Column
    TDSDeductedColNo = Sheet2.Range("TDS26QB.TotTDSOnAmtPaid").Column
    TDSClaimedColNo = Sheet2.Range("TDS26QB.6income").Column
    'TDSGrossAmountColNo = Sheet2.Range("TDsOthr2.grossamount").Column
'    TDSHeadColNo = Sheet2.Range("TDsOthr2.headincome").Column
    
    TotalExRow = Range("TDS26QB.PAN").Rows.count

    TotalXMLRow = Nodelist.count
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
     Sheet2.Range("TDS26QB.PAN").ClearContents
     Sheet2.Range("TDS26QB.Aadhar_Number").ClearContents
     Sheet2.Range("TDS26QB.EmployerOrDeductorName").ClearContents
     Sheet2.Range("TDsOthr2.SectionTDSDeducted").ClearContents  'Malli-29/04/205
     Sheet2.Range("TDS26QB.AmountDeducted").ClearContents
     Sheet2.Range("TDS26QB.DeductedYear").ClearContents
     Sheet2.Range("TDS26QB.TotTDSOnAmtPaid").ClearContents
     Sheet2.Range("TDS26QB.6income").ClearContents
     'Sheet2.Range("TDsOthr2.grossamount").ClearContents
    End If
    If (TotalDiffRow > 0) Then
     AddDiffRows_TDSoth1 (TotalDiffRow)
    End If


    rowcount = getRowNo(Sheet2.Range("TDS26QB.PAN").name)
    rowcount = rowcount - 1
    cnt = 0
If TotalXMLRow > 0 Then
    For Each node In Nodelist
        rowcount = rowcount + 1
        
            If Trim(node("panOfTenant")) <> "" And Sheet2.Cells(rowcount, PANColNo).Locked = False Then
                Sheet2.Cells(rowcount, PANColNo).Value = UCase(node("panOfTenant"))
            End If
                'not in json
            
            If Trim(node("AadhaarofTenant")) <> "" And Sheet2.Cells(rowcount, AadhaarColNo).Locked = False Then
                Sheet2.Cells(rowcount, AadhaarColNo).Value = UCase(node("AadhaarofTenant"))
            End If
            
            If Trim(node("nameOfTenant")) <> "" And Sheet2.Cells(rowcount, FYColNo).Locked = False Then
                Sheet2.Cells(rowcount, FYColNo).Value = UCase(node("nameOfTenant"))
            End If
            
   'Malli----------29/04/2025
   'form26as.scheduleTDS3Dtls.tds3Details.sectionCode
   
    TDSSection_TDS3 = node("sectionCode")
       'Debug.Print TDSSection_TDS3
     If TDSSection_TDS3 <> "" Then
            If UCase(TDSSection_TDS3) = UCase("92A") Then
            TDSSection = "192-Salary-Payment to Government employees other than Indian Government employees"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("92B") Then
            TDSSection = "192-Salary-Payment to employees other than Government employees"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("92C") Then
            TDSSection = "192-Salary-Payment to Indian Government employees"
            
            ''13/05/2025
            'ElseIf TDSSection_TDS3 = "192A" Then
            'ElseIf TDSSection_TDS3 = "2AA" Or TDSSection_TDS3 = "2aa" Or UCase(TDSSection_TDS3) = "2AA" Then
             ElseIf UCase(TDSSection_TDS3) = UCase("2AA") Then
            TDSSection = "192A-TDS on PF withdrawal"
            '------------------------------------------------------
            
            ElseIf UCase(TDSSection_TDS3) = UCase("193") Then
            TDSSection = "193-Interest on Securities"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("194") Then
            TDSSection = "194-Dividends"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94A") Then
            TDSSection = "194A-Interest other than 'Interest on securities'"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94B") Then
            TDSSection = "194B-Winning from lottery or crossword puzzle"
             '13/05/2025
            'ElseIf TDSSection_TDS3 = "94BA" Then
            'ElseIf TDSSection_TDS3 = "9BA" Or TDSSection_TDS3 = "9ba" Or UCase(TDSSection_TDS3) = "9BA" Then
            ElseIf UCase(TDSSection_TDS3) = UCase("9BA") Then
            TDSSection = "194BA-Winnings from online games"
            '----------------------------------------
            ElseIf UCase(TDSSection_TDS3) = UCase("4BB") Then
            TDSSection = "194BB-Winning from horse race"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94C") Then
            TDSSection = "194C-Payments to contractors and sub-contractors"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94D") Then
            TDSSection = "194D-Insurance commission"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("4DA") Then
            TDSSection = "194DA-Payment in respect of life insurance policy"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94E") Then
            TDSSection = "194E-Payments to non-resident sportsmen or sports associations"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("4EE") Then
            TDSSection = "194EE-Payments in respect of deposits under National Savings"
            '13/05/2025---------------------
              'ElseIf TDSSection_TDS3 = "4F" Then
             ' ElseIf TDSSection_TDS3 = "94F" Or TDSSection_TDS3 = "94f" Or UCase(TDSSection_TDS3) = "94F" Then
                ElseIf UCase(TDSSection_TDS3) = UCase("94F") Then
            TDSSection = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India"
            
            'ElseIf TDSSection_TDS3 = "4G" Then
            'ElseIf TDSSection_TDS3 = "94G" Or TDSSection_TDS3 = "94g" Or UCase(TDSSection_TDS3) = "94G" Then
             ElseIf UCase(TDSSection_TDS3) = UCase("94G") Then
            TDSSection = "194G-Commission, price, etc. on sale of lottery tickets"
            
            'ElseIf TDSSection_TDS3 = "4H" Then
             'ElseIf TDSSection_TDS3 = "94H" Or TDSSection_TDS3 = "94h" Or UCase(TDSSection_TDS3) = "94H" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("94H") Then
            TDSSection = "194H-Commission or brokerage"
            
           ' ElseIf TDSSection_TDS3 = "4-IA" Then
            ' ElseIf TDSSection_TDS3 = "4IA" Or TDSSection_TDS3 = "4ia" Or UCase(TDSSection_TDS3) = "4IA" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("4IA") Then
              ' 'As per Ashish and Madhuconfirmation dropdown changed  06/06/2025
            'TDSSection = "194I(a)-Rent on hiring of plant and machinery"
            TDSSection = "194IA-TDS on Sale of immovable property"
            
            'ElseIf TDSSection_TDS3 = "4-IB" Then
            'ElseIf TDSSection_TDS3 = "4IB" Or TDSSection_TDS3 = "4ib" Or UCase(TDSSection_TDS3) = "4IB" Then
             ElseIf UCase(TDSSection_TDS3) = UCase("4IB") Then
              'As per Ashish and Madhuconfirmation dropdown changed  06/06/2025
            'TDSSection = "194I(b)-Rent on other than plant and machinery"
            TDSSection = "194IB-Payment of rent by certain individuals or Hindu undivided"
            
            'ElseIf TDSSection_TDS3 = "4IA" Then
           ' ElseIf TDSSection_TDS3 = "9IA" Or TDSSection_TDS3 = "9ia" Or UCase(TDSSection_TDS3) = "9IA" Then
             ElseIf UCase(TDSSection_TDS3) = UCase("9IA") Then
            TDSSection = "194IA-TDS on Sale of immovable property"
            
            'ElseIf TDSSection_TDS3 = "4IB" Then
            'ElseIf TDSSection_TDS3 = "9IB" Or TDSSection_TDS3 = "9ib" Or UCase(TDSSection_TDS3) = "9IB" Then
             ElseIf UCase(TDSSection_TDS3) = UCase("9IB") Then
            TDSSection = "194IB-Payment of rent by certain individuals or Hindu undivided"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("4IC") Then
            TDSSection = "194IC-Payment under specified agreement"
            
            'ElseIf TDSSection_TDS3 = "94J-A" Then
            'ElseIf TDSSection_TDS3 = "4JA" Or TDSSection_TDS3 = "4ja" Or UCase(TDSSection_TDS3) = "4JA" Then
            ElseIf UCase(TDSSection_TDS3) = UCase("4JA") Then
            TDSSection = "194J(a)-Fees for technical services"
            
            'ElseIf TDSSection_TDS3 = "94J-B" Then
            'ElseIf TDSSection_TDS3 = "4JB" Or TDSSection_TDS3 = "4jb" Or UCase(TDSSection_TDS3) = "4JB" Then
            ElseIf UCase(TDSSection_TDS3) = UCase("4JB") Then
            TDSSection = "194J(b)-Fees for professional  services or royalty etc"
            '---------
            ElseIf UCase(TDSSection_TDS3) = UCase("94K") Then
            TDSSection = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("4LA") Then
            TDSSection = "194LA-Payment of compensation on acquisition of certain immovable"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("4LB") Then
            TDSSection = "194LB-Income by way of Interest from Infrastructure Debt fund"
            
            '13/05/2025
            'ElseIf TDSSection_TDS3 = "4LC1" Then
             'ElseIf TDSSection_TDS3 = "LC1" Or TDSSection_TDS3 = "lc1" Or UCase(TDSSection_TDS3) = "LC1" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("LC1") Then
            TDSSection = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC"
            
            'ElseIf TDSSection_TDS3 = "4LC2" Then
             'ElseIf TDSSection_TDS3 = "LC2" Or TDSSection_TDS3 = "lc2" Or UCase(TDSSection_TDS3) = "LC2" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("LC2") Then
            TDSSection = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC"
            
            'ElseIf TDSSection_TDS3 = "4LC3" Then
            ' ElseIf TDSSection_TDS3 = "LC3" Or TDSSection_TDS3 = "lc3" Or UCase(TDSSection_TDS3) = "LC3" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("LC3") Then
            TDSSection = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC"
            
            'Malli_AY_2026_27  SIT-117189 09/04/2026
            ElseIf TDSSection_TDS3 = "BA1" Then
            TDSSection = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder"
            
            'Malli_AY_2026_27  SIT-117189 09/04/2026
            ElseIf TDSSection_TDS3 = "BA2" Then
            TDSSection = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder"
'
'            ElseIf TDSSection_TDS3 = "LBA1" Then
'            TDSSection = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR"
'
'            ElseIf TDSSection_TDS3 = "LBA2" Then
'            TDSSection = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR"
            
           ' ElseIf TDSSection_TDS3 = "LBA3" Then
             'ElseIf TDSSection_TDS3 = "BA3" Or TDSSection_TDS3 = "ba3" Or UCase(TDSSection_TDS3) = "BA3" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("BA3") Then
            TDSSection = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR"
            '-----------------------------------------------------
            
            ElseIf UCase(TDSSection_TDS3) = UCase("LBB") Then
            TDSSection = "194LBB-Income in respect of units of investment fund"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94R") Then
            TDSSection = "194R-Benefits or perquisites of business or profession"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94S") Then
            TDSSection = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons"
            
             '13/05/2025
            'ElseIf TDSSection_TDS3 = "94B-P" Then
            'ElseIf TDSSection_TDS3 = "4BP" Or TDSSection_TDS3 = "4bp" Or UCase(TDSSection_TDS3) = "4BP" Then
             ElseIf UCase(TDSSection_TDS3) = UCase("4BP") Then
            TDSSection = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released"
            
            'ElseIf TDSSection_TDS3 = "94R-P" Then
            'ElseIf TDSSection_TDS3 = "4RP" Or TDSSection_TDS3 = "4rp" Or UCase(TDSSection_TDS3) = "4RP" Then
             ElseIf UCase(TDSSection_TDS3) = UCase("4RP") Then
            TDSSection = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released"
            
            'ElseIf TDSSection_TDS3 = "94S-P" Then
            'ElseIf TDSSection_TDS3 = "4SP" Or TDSSection_TDS3 = "4sp" Or UCase(TDSSection_TDS3) = "4SP" Then
            ElseIf UCase(TDSSection_TDS3) = UCase("4SP") Then
            TDSSection = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released"
            '----------------------------------------------
            
            ElseIf UCase(TDSSection_TDS3) = UCase("LBC") Then
            TDSSection = "194LBC-Income in respect of investment in securitization trust"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("4LD") Then
            TDSSection = "194LD-TDS on interest on bonds / government securities"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94M") Then
            TDSSection = "194M-Payment of certain sums by certain individuals or HUF"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94N") Then
            TDSSection = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso"
            
            '13/05/2025
            'ElseIf TDSSection_TDS3 = "94N-F" Then
             'ElseIf TDSSection_TDS3 = "4NF" Or TDSSection_TDS3 = "4nf" Or UCase(TDSSection_TDS3) = "4NF" Then
             ElseIf UCase(TDSSection_TDS3) = UCase("4NF") Then
            TDSSection = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties"
            
            'ElseIf TDSSection_TDS3 = "94N-C" Then
             'ElseIf TDSSection_TDS3 = "4NC" Or TDSSection_TDS3 = "4nc" Or UCase(TDSSection_TDS3) = "4NC" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("4NC") Then
            TDSSection = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso"
            
            'ElseIf TDSSection_TDS3 = "94N-FT" Then
             'ElseIf TDSSection_TDS3 = "NFT" Or TDSSection_TDS3 = "nft" Or UCase(TDSSection_TDS3) = "NFT" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("NFT") Then
            TDSSection = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies"
            '----------------------------------------------------
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94O") Then
            TDSSection = "194O-Payment of certain sums by e-commerce operator to e-commerce participant."
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94P") Then
            TDSSection = "194P-Deduction of tax in case of specified senior citizen"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("94Q") Then
            TDSSection = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("195") Then
            TDSSection = "195-Other sums payable to a non-resident"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("96A") Then
            TDSSection = "196A-Income in respect of units of non-residents"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("96B") Then
            TDSSection = "196B-Payments in respect of units to an offshore fund"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("96C") Then
            TDSSection = "196C-Income from foreign currency bonds or shares of Indian"
            
            ElseIf UCase(TDSSection_TDS3) = UCase("96D") Then
            TDSSection = "196D-Income of foreign institutional investors from securities"
            
            '13/05/2025
            'ElseIf TDSSection_TDS3 = "96DA" Then
            'ElseIf TDSSection_TDS3 = "6DA" Or TDSSection_TDS3 = "6da" Or UCase(TDSSection_TDS3) = "6DA" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("6DA") Then
            TDSSection = "196D(1A)-Income of specified fund from securities"
            
            'ElseIf TDSSection_TDS3 = "94BA-P" Then
            'ElseIf TDSSection_TDS3 = "BAP" Or TDSSection_TDS3 = "bap" Or UCase(TDSSection_TDS3) = "BAP" Then
              ElseIf UCase(TDSSection_TDS3) = UCase("BAP") Then
            TDSSection = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released"
            
            Else
               'TDSSection = ""
               TDSSection = "(Select)"
        End If
        
            If TDSSection <> "" And Sheet2.Cells(rowcount, SecTDSColNo).Locked = False Then
                Sheet2.Cells(rowcount, SecTDSColNo).Value = TDSSection
                TDSSection_TDS3 = ""
            End If
       
     End If
     '--------------------
            
            
            
   '--------------------------
            
            If Trim(node("grossAmount")) <> "" And Sheet2.Cells(rowcount, TDSOtherColNo).Locked = False Then
                Sheet2.Cells(rowcount, TDSOtherColNo).Value = UCase(node("grossAmount"))
            End If
                
            'Malli 03/07/2026 updated as per enhancement tds deducted year should by default 2025
'            If Trim(node("deductedYr")) <> "" And Sheet2.Cells(rowcount, TDSyearColNo).Locked = False Then
'                Sheet2.Cells(rowcount, TDSyearColNo).Value = UCase(node("deductedYr"))
'            End If
            
            
            If Sheet2.Cells(rowcount, TDSyearColNo).Locked = False Then
                     Dim deductedYr_tds3, pfl_deductedYr_tds3
                         deductedYr_tds3 = Trim(node("deductedYr"))
                    If deductedYr_tds3 <> "" Then
                            
                            If deductedYr_tds3 = "2017" Then
                               pfl_deductedYr_tds3 = "2017-18"
                               
                            ElseIf deductedYr_tds3 = "2018" Then
                               pfl_deductedYr_tds3 = "2018-19"
                               
                            ElseIf deductedYr_tds3 = "2019" Then
                               pfl_deductedYr_tds3 = "2019-20"
                               
                            ElseIf deductedYr_tds3 = "2020" Then
                               pfl_deductedYr_tds3 = "2020-21"
                               
                            ElseIf deductedYr_tds3 = "2021" Then
                               pfl_deductedYr_tds3 = "2021-22"
                               
                            ElseIf deductedYr_tds3 = "2022" Then
                               pfl_deductedYr_tds3 = "2022-23"
                               
                            ElseIf deductedYr_tds3 = "2023" Then
                               pfl_deductedYr_tds3 = "2023-24"
                               
                            ElseIf deductedYr_tds3 = "2024" Then
                               pfl_deductedYr_tds3 = "2024-25"
                               
                            ElseIf deductedYr_tds3 = "2025" Then
                               pfl_deductedYr_tds3 = "2025-26"
                               
                            Else
                                pfl_deductedYr_tds3 = "2025-26"
                            End If
                            
                            Sheet2.Cells(rowcount, TDSyearColNo).Value = pfl_deductedYr_tds3
                    Else
                            Sheet2.Cells(rowcount, TDSyearColNo).Value = "2025-26"
                    End If
            End If
            '-----------------
            
           
            
' change 3 start
            If Trim(node("taxDeductCreditDtls")("taxDeductedOwnHands")) <> "" And Sheet2.Cells(rowcount, TDSDeductedColNo).Locked = False Then
                Sheet2.Cells(rowcount, TDSDeductedColNo).Value = UCase(node("taxDeductCreditDtls")("taxDeductedOwnHands"))
            End If
            
            If Trim(node("taxDeductCreditDtls")("taxClaimedOwnHands")) <> "" And Sheet2.Cells(rowcount, TDSClaimedColNo).Locked = False Then
                Sheet2.Cells(rowcount, TDSClaimedColNo).Value = UCase(node("taxDeductCreditDtls")("taxClaimedOwnHands"))
            End If
' change 3 end
            
'            Dim HeadOfIncome As Variant
'            'not in json
'              HeadOfIncome = UCase(Node("HeadOfIncome"))
'            If HeadOfIncome = "HP" Then
'            HeadOfIncome = "Income from House property"
'            End If
'            Sheet2.Cells(rowcount, TDSHeadColNo).Value = HeadOfIncome
            
        cnt = cnt + 1
    Next node
End If
    RecTDS11 = cnt
End Sub
Function Import_IncomeDeductions(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary, node, Nodelist As Object
Dim EntertainmentAlw16ii, ProfessionalTaxUs16iii
Dim Salary, PerquisitesValue, ProfitsInSalary, TypeOfHP, GrossRentReceived As Variant
Dim Section80C, Section80CCC, Section80CCDEmployeeOrSE, Section80CCD1B, Section80CCDEmployer, Section80D, Section80DD, Section80DDB, Section80DDUsrType, Section80DDBUsrType, Section80UUsrType As Variant
Dim Section80E, Section80U, Section80TTA, section89 As Variant
Dim init, NatureColNo, DescriptionColNo, AmtColNo, TotalExRow, TotalXMLRow, TotalDiffRow, rowcount, cnt, test

Set jsonObject = ParseJson(jsonText)

'issue 3 solved
EntertainmentAlw16ii = jsonObject("form24q")("incomeDeductions")("entertainmentAlw16Ii")
ProfessionalTaxUs16iii = jsonObject("form24q")("incomeDeductions")("professionalTaxUs16Iii")
Salary = jsonObject("form24q")("incomeDeductions")("salary")
PerquisitesValue = jsonObject("form24q")("incomeDeductions")("perquisitesValue")
ProfitsInSalary = jsonObject("form24q")("incomeDeductions")("profitsInSalary")

'TypeOfHP = jsonObject("form26as")("scheduleHP")("propertyDetails")("typeOfHP") 'Newly updated by Bindu

GrossRentReceived = jsonObject("form26as")("grossRent")

'/* SIT-111718 Malli_AY_2026_27  prefill posibility removed for Old tax regime fields
'Section80C = jsonObject("form24q")("usrDeductUndChapVIAType")("section80C")
'Section80CCC = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCC")
'Section80CCDEmployeeOrSE = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCDEmployeeOrSE")
'Section80CCD1B = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCD1B")
'Section80DDB = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDB")
'--------------------------------------
Section80CCDEmployer = jsonObject("form24q")("usrDeductUndChapVIAType")("section80CCDEmployer")

'Comented as per the latest prefill AY_2024-25 'Malli
'Section80DD = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DD")



'-----------------------


'Comented as per the latest prefill AY_2024-25 'Malli
'Section80DDUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDUsrType")
'Section80UUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80UUsrType")
'Section80UUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80UUsrType") 'for 80U Sheet
 'lastFilledITR.usrDeductUndChapVIAType.section80UUsrType

'for 80U Sheet Malli--------------

'If Section80UUsrType = "1" Then
'    Section80UUsrType = "1-Self with disability"
'ElseIf Section80UUsrType = "2" Then
'    Section80UUsrType = "2-Self with severe disability"
'Else
'    Section80UUsrType = ""
'End If
'-------------------------
'-----------------------------------------------

'/* SIT-111718 Malli_AY_2026_27  prefill posibility removed for Old tax regime fields
'AY_2024_25 Change 'Malli 'Newly changed by Bindu
'Section80DDBUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80DDBUsrType")
'Section80E = jsonObject("form24q")("usrDeductUndChapVIAType")("section80E")
'Section80TTA = jsonObject("form24q")("usrDeductUndChapVIAType")("section80TTA")
'Section80TTB = jsonObject("form24q")("usrDeductUndChapVIAType")("section80TTB")

'---------------------------------------------------




'Malli---------------
'Section80U = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80U")

'------------

'section89 = jsonObject("form24q")("taxComputation")("section89")

'Malli SIT-67324

'
'section89 = jsonObject("form10E")("TaxComputation")("section89")

'comented by Malli 'AY_2025_26 as per Lavany confirmation
             section89 = jsonObject("form10E")("TaxComputation")("section89")
'Konda  SIT-67324 25/06/2024
     'section89 = jsonObject("form24q")("taxComputation")("section89")
'end  SIT-67324 25/06/2024

'-----------------------
'section89 = jsonObject("form24q")("taxComputation")("section89")

'-----------
'Malli--------------
'If Section80DDUsrType = "1" Then
'    Section80DDUsrType = "1-Dependent Person with Disability"
'ElseIf Section80DDUsrType = "2" Then
'    Section80DDUsrType = "2-Dependent person with Severe Disability"
'Else
'    Section80DDUsrType = ""
'End If
 

'/* SIT-111718 Malli_AY_2026_27  prefill posibility removed for Old tax regime fields
'If Section80DDBUsrType = "1" Then
'    Section80DDBUsrType = "1-Self or Dependent"
'ElseIf Section80DDBUsrType = "2" Then
'    Section80DDBUsrType = "2-Self or Dependent- Senior Citizen "
'Else
'    Section80DDBUsrType = ""
'End If
'---------------------------------------------
 
'If Section80UUsrType = "1" Then
'    Section80UUsrType = "1-Self with disability"
'ElseIf Section80UUsrType = "2" Then
'    Section80UUsrType = "2-Self with severe disability"
'Else
'    Section80UUsrType = ""
'End If
'------------------------

    If Salary <> "" And Sheet1.Range("IncD.Allowances").Locked = False Then
        Sheet1.Range("IncD.Allowances").Value = Salary
    End If
    If PerquisitesValue <> "" And Sheet1.Range("IncD.Perquisites").Locked = False Then
        Sheet1.Range("IncD.Perquisites").Value = PerquisitesValue
    End If
    If ProfitsInSalary <> "" And Sheet1.Range("IncD.Profits").Locked = False Then
        Sheet1.Range("IncD.Profits").Value = ProfitsInSalary
    End If
    
    
    Set init = jsonObject("form24q")
    If init <> Empty Then
    If init.exists("allwncExemptUs10DtlsType") Then
        Set Nodelist = jsonObject("form24q")("allwncExemptUs10DtlsType")
    
        NatureColNo = Sheet1.Range("Others.NOI_1").Column
        DescriptionColNo = Sheet1.Range("Nature_Others_1").Column
        AmtColNo = Sheet1.Range("Others.Amount_1").Column
        
        TotalExRow = Range("Others.NOI_1").Rows.count
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
    
        If (TotalXMLRow > 0) Then
            Sheet1.Range("Others.NOI_1").ClearContents
            Sheet1.Range("Nature_Others_1").ClearContents
            Sheet1.Range("Others.Amount_1").ClearContents
        End If
        
        If (TotalDiffRow > 0) Then
'            AddDiffRows_Exempt2 (TotalDiffRow)
'Konda add-------------------
            If UCase(node("SalNatureDesc")) = "10(13A)" Then
                TotalDiffRow = TotalDiffRow - 1
                AddDiffRows_Exempt2 (TotalDiffRow)
            Else
             AddDiffRows_Exempt2 (TotalDiffRow)
            End If
        End If
    
       rowcount = getRowNo(Sheet1.Range("Others.NOI_1").name)
       rowcount = rowcount - 1
       cnt = 0
    
        For Each node In Nodelist
        
        If UCase(node("salNatureDesc")) <> "10(13A)" Then 'Newly added by Bindu as per DE V7
            rowcount = rowcount + 1
                 
            If UCase(node("salNatureDesc")) = "10(10B)(i)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette"
            ElseIf UCase(node("salNatureDesc")) = "10(10B)(ii)" Then
                Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government"
'Konda updated on 12-03-2026--V0.5
'            ElseIf UCase(node("salNatureDesc")) = "OTH" Then
'                Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
            
                    ElseIf UCase(node("salNatureDesc")) = "10(17)" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(17)-Allowance MP/MLA/MLC " 'Ankita

            Else
                test = Findtext("Sec " & (node("salNatureDesc")), "PART_Nature_1")
                Sheet1.Range("J" & rowcount).Value = test
'====================================
            End If
'            If Not Node("SalOthNatOfInc") = "" Then
'                 Sheet1.Range("Z" & rowcount).Value = Node("SalOthNatOfInc")
'            End If

'Konda updated on 12-03-2026--V0.5
'            Sheet1.Range("AO" & rowcount).Value = node("salOthAmount")
            If UCase(node("salOthAmount")) <> "" Then
                Sheet1.Cells(rowcount, AmtColNo).Value = node("salOthAmount")
            End If
'=======================================
            cnt = cnt + 1
            End If ''Bindu
        Next node
    End If
    End If
    
    If EntertainmentAlw16ii <> "" And Sheet1.Range("IncD.Deduction16").Locked = False Then
        Sheet1.Range("IncD.Deduction16").Value = EntertainmentAlw16ii
    End If
    
    If ProfessionalTaxUs16iii <> "" And Sheet1.Range("IncD.Deduction16ic").Locked = False Then
        Sheet1.Range("IncD.Deduction16ic").Value = ProfessionalTaxUs16iii
    End If
    
'    TypeOfHP = UCase(node("TypeOfHP"))
'    If TypeOfHP <> "" Then
'        If TypeOfHP = "Y" Then
'            Sheet1.Range("IncD.TypeOfHP").Value = "Let Out"
'        ElseIf TypeOfHP = "S" Then
'            Sheet1.Range("IncD.TypeOfHP").Value = "Self Occupied"
'        ElseIf TypeOfHP = "D" Then
'            Sheet1.Range("IncD.TypeOfHP").Value = "Deemed Let Out"
'        End If
'    End If
    

  
    
    
    Sheet1.Range("IncD.GrossRentRecieved").Value = GrossRentReceived
    
 'Malli_30/05/2025
        incDeductionsOthIncCPC_TAX_chk = False
        incDeductionsOthIncCPC_TAX_RowCnt = 0
        '------------------------------------------
    
    Set init = jsonObject("form26as")
    If init <> Empty Then
    If init.exists("incomeDeductionsOthersInc") Then
        Set Nodelist = jsonObject("form26as")("incomeDeductionsOthersInc")
    
        NatureColNo = Sheet1.Range("Others.NOI_2").Column
        DescriptionColNo = Sheet1.Range("Nature_Others_2").Column
        AmtColNo = Sheet1.Range("Others.Amount_2").Column
        
        TotalExRow = Range("Others.NOI_2").Rows.count
        
        TotalXMLRow = Nodelist.count
        TotalDiffRow = TotalXMLRow - TotalExRow
    
        If (TotalXMLRow > 0) Then
            Sheet1.Range("Others.NOI_2").ClearContents
            Sheet1.Range("Nature_Others_2").ClearContents
            Sheet1.Range("Others.Amount_2").ClearContents
        End If
    
        If (TotalDiffRow > 0) Then
            AddDiffRows_Exempt1 (TotalDiffRow)
        End If
 
        rowcount = getRowNo(Sheet1.Range("Others.NOI_2").name)
        rowcount = rowcount - 1
        cnt = 0
        Sheet1.Activate
    
        For Each node In Nodelist
            rowcount = rowcount + 1
            incDeductionsOthIncCPC_TAX_RowCnt = rowcount   'Malli_30/05/2025
        
                If UCase(node("othSrcNatureDesc")) = "SAV" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Savings Bank Account"
                ElseIf UCase(node("othSrcNatureDesc")) = "IFD" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Deposit (Bank/Post Office/Cooperative Society)"
                ElseIf UCase(node("othSrcNatureDesc")) = "TAX" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Income Tax Refund"
                     incDeductionsOthIncCPC_TAX_chk = True  'Malli_30/05/2025
                ElseIf UCase(node("othSrcNatureDesc")) = "FAP" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Family pension"
               'Malli----------comented AY_2025_26 SIT-85920 and SIT-85923 on 08/01/2025
                'Newly added by Bindu
                ElseIf UCase(node("othSrcNatureDesc")) = "DIV" Then
                rowcount = rowcount - 1
                   Sheet1.Range("IncD.Div").Value = node("othSrcOthAmount")
                    '--end
               'Malli------------SIT_94085
                'Else
                ElseIf UCase(node("othSrcNatureDesc")) = "OTH" Then
                    Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
                '---------------------------------------------------------
'                 If UCase(node("othSrcNatureDesc")) <> "DIV" Then
'                    Sheet1.Cells(rowcount, AmtColNo).Value = node("othSrcOthAmount")
'                End If
                End If
               
                If UCase(node("othSrcNatureDesc")) <> "DIV" Then
                    Sheet1.Cells(rowcount, AmtColNo).Value = node("othSrcOthAmount")
                End If
            cnt = cnt + 1
        Next node
    End If
    End If
    
    
'Malli--------SIT-93728------------30/05/2025
If incDeductionsOthIncCPC_TAX_chk <> True Then
           
           Set init = jsonObject("incDeductionsOthIncCPC")
           If init <> Empty Then
               Set Nodelist = jsonObject("incDeductionsOthIncCPC")
               
                NatureColNo = Sheet1.Range("Others.NOI_2").Column
                DescriptionColNo = Sheet1.Range("Nature_Others_2").Column
                AmtColNo = Sheet1.Range("Others.Amount_2").Column
               
               TotalExRow = Range("Others.NOI_2").Rows.count
               TotalXMLRow = Nodelist.count
               TotalDiffRow = TotalXMLRow - TotalExRow
               
            If Not incDeductionsOthIncCPC_TAX_RowCnt > 0 Then
               If (TotalXMLRow > 0) Then
                   Sheet1.Range("Others.NOI_2").ClearContents
                   Sheet1.Range("Nature_Others_2").ClearContents
                   Sheet1.Range("Others.Amount_2").ClearContents
               End If
            
               If (TotalDiffRow > 0) Then
                     AddDiffRows_Exempt1 (TotalDiffRow)
               End If
               
               rowcount = getRowNo(Sheet1.Range("Others.NOI_2").name)
               rowcount = rowcount - 1
               cnt = 0
               Sheet1.Activate
            
            End If
               
               
           
               For Each node In Nodelist
               
                   If Not incDeductionsOthIncCPC_TAX_RowCnt > 0 Then
                   rowcount = rowcount + 1
                   End If
                   
                   If Sheet1.Cells(rowcount, NatureColNo).Locked = False Then
                       
                       If UCase(node("othSrcNatureDesc")) = "TAX" Then
                       
                            If incDeductionsOthIncCPC_TAX_RowCnt > 0 Then
                                 AddDiffRows_Exempt1 (1)
                                 rowcount = rowcount + 1
                            End If
                          
                           
                           Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Income Tax Refund"
                           
                           If Sheet1.Cells(rowcount, AmtColNo).Locked = False Then
                           Sheet1.Cells(rowcount, AmtColNo).Value = node("othSrcOthAmount")
                           End If
                           
                       End If
                   End If
                   
                    
                   cnt = cnt + 1
               Next node
           End If
           
    
End If
    
    '------------------------------------
    
    
    '/* SIT-111718 Malli_AY_2026_27  prefill posibility removed for Old tax regime fields
    
'    If Section80C <> "" And Sheet1.Range("IncD.Section80C").Locked = False Then
'        Sheet1.Range("IncD.Section80C").Value = Section80C
'    End If
    
    
'    If Section80CCC <> "" And Sheet1.Range("IncD.Section80CCC").Locked = False Then
'        Sheet1.Range("IncD.Section80CCC").Value = Section80CCC
'    End If
    
    
'    If Section80CCDEmployeeOrSE <> "" And Sheet1.Range("IncD.Section80CCD_SE").Locked = False Then
'        Sheet1.Range("IncD.Section80CCD_SE").Value = Section80CCDEmployeeOrSE
'    End If
  
    
'    If Section80CCD1B <> "" And Sheet1.Range("IncD.Section80CCD1B_SE").Locked = False Then
'        Sheet1.Range("IncD.Section80CCD1B_SE").Value = Section80CCD1B
'    End If

''Newly Uncommented by Bindu
'    If Section80DDB <> "" And Sheet1.Range("IncD.Section80DDB").Locked = False Then
'        Sheet1.Range("IncD.Section80DDB").Value = Section80DDB
'    End If
     
     
'     If Section80E <> "" And Sheet1.Range("IncD.Section80E").Locked = False Then
'        Sheet1.Range("IncD.Section80E").Value = Section80E
'    End If
'
    '-----------------------------------------
    
    If Section80CCDEmployer <> "" And Sheet1.Range("IncD.Section80CCD").Locked = False Then
        Sheet1.Range("IncD.Section80CCD").Value = Section80CCDEmployer
    End If
    
    'Comented as per the latest prefill AY_2024-25 'Malli
'    If Section80DD <> "" And Sheet1.Range("IncD.Section80DD").Locked = False Then
'        Sheet1.Range("IncD.Section80DD").Value = Section80DD
'    End If
    
    
    
    
    
   '/* SIT-111718 Malli_AY_2026_27  prefill posibility removed for Old tax regime fields
    
    'Malli----V0.3____wqAY_2025_26---29/04/2025

'    Dim Form10BAAckNum_Pfl As Variant
'    Form10BAAckNum_Pfl = jsonObject("Form10BA")("Form10BAAckNum")
'    If Form10BAAckNum_Pfl <> "" Then
'    If Sheet5.Range("BacValue").Value = 1 Then
'                    Sheet1.Unprotect Password:=getmsgstate
'                    Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = False
'                    Sheet1.Range("Sheet1.AckNum").MergeArea.Interior.Color = "&HCCFFCC"
'                    Sheet1.Range("Sheet1.AckNum").MergeArea.Value = Form10BAAckNum_Pfl
'                    Sheet1.Unprotect Password:=getmsgstate
'                    Sheet1.Range("Sheet1.AckNum").MergeArea.Font.Color = "&HD8D8D8"
'                    Sheet1.Range("Sheet1.AckNum").MergeArea.Interior.Color = "&HD8D8D8"
'                    Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = True
'    Else
'            If Form10BAAckNum_Pfl <> "" And Sheet1.Range("Sheet1.AckNum").Locked = False Then
'                Sheet1.Range("Sheet1.AckNum").Value = Form10BAAckNum_Pfl
'            End If
'    End If
'
'    End If
    '------------------
    
    
    
'    If Section80TTA <> "" And Sheet1.Range("IncD.Section80TTA").Locked = False Then
'        Sheet1.Range("IncD.Section80TTA").Value = Section80TTA
'    End If
'    If Section80TTB <> "" And Sheet1.Range("IncD.Section80TTB").Locked = False Then
'        Sheet1.Range("IncD.Section80TTB").Value = Section80TTB
'    End If
    '----------------------------------
    
    
    'Malli---------------
'    If Section80U <> "" And Sheet1.Range("IncD.Section80U").Locked = False Then
'        Sheet1.Range("IncD.Section80U").Value = Section80U
'    End If
    '--------------------------------------
    If section89 <> "" And Sheet1.Range("IncD.Section89").Locked = False Then
        Sheet1.Range("IncD.Section89").Value = section89
    End If
    
    'Malli---------
'    If Section80DDUsrType <> "" And Sheet1.Range("SELECT80DD").Locked = False Then
'    Sheet1.Range("SELECT80DD").Value = Section80DDUsrType
'    End If
    
'    '/* SIT-111718 Malli_AY_2026_27  prefill posibility removed for Old tax regime fields
'    If Section80DDBUsrType <> "" And Sheet1.Range("SELECT80DDS").Locked = False Then
'    Sheet1.Range("SELECT80DDS").Value = Section80DDBUsrType
'    End If
'-----------------------
     
'    If Section80UUsrType <> "" And Sheet1.Range("SELECT80U").Locked = False Then
'    Sheet1.Range("SELECT80U").Value = Section80UUsrType
'    End If
    '--------------
    
    'Malli------ADDED-AY_2024-25-----
'    If Section80UUsrType <> "" And Sheet1.Range("Naturedisability_80U").Locked = False Then
'     Sheet14.Range("Naturedisability_80U").Value = Section80UUsrType
'    End If
    '--------------
'    If Sheet1.Range("IncD.AnyOtherDeductions").Locked = False Then
'        Sheet1.Range("IncD.AnyOtherDeductions").Value = AnyOthSec80CCH
'        Sheet1.Range("IncD.AnyOther").Value = "80CCH-Contribution to Agnipath Scheme"
'    End If
'issue 3 solved
End Function

Function ImportSchedule80D_old(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim node, Nodelist As Object

Set jsonObject = ParseJson(jsonText)

If Not jsonObject.exists("lastFiledITR") Then
    Exit Function
End If
If Not jsonObject("lastFiledITR").exists("schedule80D") Then
    Exit Function
End If
If Not jsonObject("lastFiledITR")("schedule80D").exists("Sec80DSelfFamSrCtznHealth") Then
    Exit Function
End If

'lastFiledITR.schedule80D.Sec80DSelfFamSrCtznHealth.SeniorCitizenFlag

Set node = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")
    
            If Trim(node("SeniorCitizenFlag")) <> "" And Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Locked = False Then
                If node("SeniorCitizenFlag") = "Y" Then
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Yes"
                ElseIf node("SeniorCitizenFlag") = "N" Then
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "No"
                ElseIf node("SeniorCitizenFlag") = "S" Then
                    'Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "No Claiming for Self/Family" 'Malli
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Not claiming for Self/ Family"
                Else
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "(Select)"
                End If
            End If
                                            
            If Trim(node("HealthInsPremSlfFam")) <> "" And Sheet9.Range("Health_Insurance_80D").Locked = False Then
                Sheet9.Range("Health_Insurance_80D").Value = node("HealthInsPremSlfFam")
            End If
            If Trim(node("PrevHlthChckUpSlfFam")) <> "" And Sheet9.Range("Preventive_Health_80D").Locked = False Then
                Sheet9.Range("Preventive_Health_80D").Value = node("PrevHlthChckUpSlfFam")
            End If
            If Trim(node("HlthInsPremSlfFamSrCtzn")) <> "" And Sheet9.Range("Health_InsuranceSC_80D").Locked = False Then
                Sheet9.Range("Health_InsuranceSC_80D").Value = node("HlthInsPremSlfFamSrCtzn")
            End If
            If Trim(node("PrevHlthChckUpSlfFamSrCtzn")) <> "" And Sheet9.Range("Preventive_Health_SC_80D").Locked = False Then
                Sheet9.Range("Preventive_Health_SC_80D").Value = node("PrevHlthChckUpSlfFamSrCtzn")
            End If
            If Trim(node("MedicalExpSlfFamSrCtzn")) <> "" And Sheet9.Range("Medical_Expenditure_SC_80D").Locked = False Then
                Sheet9.Range("Medical_Expenditure_SC_80D").Value = node("MedicalExpSlfFamSrCtzn")
            End If
            
'    Set Node = jsonObject("lastFiledITR")("schedule80D")("Sec80DSelfFamSrCtznHealth")

            If Trim(node("ParentSeniorCitizenFlag")) <> "" And Sheet9.Range("DropDown_ValueOf_SC_80D").Locked = False Then
                If node("ParentSeniorCitizenFlag") = "Y" Then
                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Yes"
                ElseIf node("ParentSeniorCitizenFlag") = "N" Then
                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "No"
                ElseIf node("ParentSeniorCitizenFlag") = "P" Then
                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Not claiming for Parents"
                End If
            End If
            
            If Trim(node("HlthInsPremParents")) <> "" And Sheet9.Range("Health_Insurance2_80D").Locked = False Then
                Sheet9.Range("Health_Insurance2_80D").Value = node("HlthInsPremParents")
            End If
            If Trim(node("PrevHlthChckUpParents")) <> "" And Sheet9.Range("Preventive_Health2_80D").Locked = False Then
                Sheet9.Range("Preventive_Health2_80D").Value = node("PrevHlthChckUpParents")
            End If
            If Trim(node("HlthInsPremParentsSrCtzn")) <> "" And Sheet9.Range("Health_Insurance3_80D").Locked = False Then
                Sheet9.Range("Health_Insurance3_80D").Value = node("HlthInsPremParentsSrCtzn")
            End If
            If Trim(node("PrevHlthChckUpParentsSrCtzn")) <> "" And Sheet9.Range("Preventive_Health3_80D").Locked = False Then
                Sheet9.Range("Preventive_Health3_80D").Value = node("PrevHlthChckUpParentsSrCtzn")
            End If
            If Trim(node("MedicalExpParentsSrCtzn")) <> "" And Sheet9.Range("Medical_Expenditure2_80D").Locked = False Then
                Sheet9.Range("Medical_Expenditure2_80D").Value = node("MedicalExpParentsSrCtzn")
            End If

End Function
Function ImportSchedule80D(jsonText As String)
On Error Resume Next
Dim init, jsonObject As Object
Set jsonObject = ParseJson(jsonText)
Set init = jsonObject("lastFiledITR")

If init <> Empty Then
If init.exists("schedule80D") Then
    Dim schedule80DObject As Object
    Set schedule80DObject = init("schedule80D")
        If schedule80DObject.exists("Sec80DSelfFamSrCtznHealth") Then
            Dim Sec80DSelfFamSrCtznHealthobject, SeniorCitizenFlag, parentsSeniorCitizenFlag As Object
            Set Sec80DSelfFamSrCtznHealthobject = schedule80DObject("Sec80DSelfFamSrCtznHealth")
            SeniorCitizenFlag = Sec80DSelfFamSrCtznHealthobject("SeniorCitizenFlag")
            parentsSeniorCitizenFlag = Sec80DSelfFamSrCtznHealthobject("ParentSeniorCitizenFlag")
            
            If Trim(Sec80DSelfFamSrCtznHealthobject("SeniorCitizenFlag")) <> "" Then
                If Trim(Sec80DSelfFamSrCtznHealthobject("SeniorCitizenFlag")) = "Y" Then
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Yes"
                ElseIf Trim(Sec80DSelfFamSrCtznHealthobject("SeniorCitizenFlag")) = "N" Then
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "No"
                ElseIf Trim(Sec80DSelfFamSrCtznHealthobject("SeniorCitizenFlag")) = "S" Then
                    Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Not claiming for Self/ Family"
                End If
            Else
                Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "(Select)"
            End If
            
             
            'lastFiledITR.schedule80D.Sec80DSelfFamSrCtznHealth.ParentSeniorCitizenFlag
            
            If Trim(Sec80DSelfFamSrCtznHealthobject("ParentSeniorCitizenFlag")) <> "" Then
                If Trim(Sec80DSelfFamSrCtznHealthobject("ParentSeniorCitizenFlag")) = "Y" Then
                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Yes"
                ElseIf Trim(Sec80DSelfFamSrCtznHealthobject("ParentSeniorCitizenFlag")) = "N" Then
                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "No"
                ElseIf Trim(Sec80DSelfFamSrCtznHealthobject("ParentSeniorCitizenFlag")) = "P" Then
                    Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Not claiming for Parents"
                End If
            Else
                Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "(Select)"
            End If
        End If
End If
End If

End Function
Function ImportTaxPaid(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim item As Object

Set jsonObject = ParseJson(jsonText)

    If Trim(Sheet3.Range("IncD.AdvanceTax").Value) = "" And Sheet3.Range("IncD.AdvanceTax").Locked = False Then
        Sheet3.Range("IncD.AdvanceTax").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("AdvanceTax")
    End If

    If Trim(Sheet3.Range("IncD.TDS").Value) = "" And Sheet3.Range("IncD.TDS").Locked = False Then
        Sheet3.Range("IncD.TDS").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("TDS")
    End If

    If Trim(Sheet3.Range("IncD.TCS").Value) = "" And Sheet3.Range("IncD.TCS").Locked = False Then
        Sheet3.Range("IncD.TCS").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("TCS")
    End If

    If Trim(Sheet3.Range("IncD.SelfAssessmentTax").Value) = "" And Sheet3.Range("IncD.SelfAssessmentTax").Locked = False Then
        Sheet3.Range("IncD.SelfAssessmentTax").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("SelfAssessmentTax")
    End If

'    If Trim(Sheet3.Range("IncD.TotalTaxesPaid").Value) = "" Then
'        Sheet3.Range("IncD.TotalTaxesPaid").Value = jsonObject("ITR")("ITR1")("TaxPaid")("TaxesPaid")("TotalTaxesPaid")
'    End If

'    If Trim(Sheet3.Range("IncD.BalTaxPayable").Value) = "" Then
'        Sheet3.Range("IncD.BalTaxPayable").Value = jsonObject("ITR")("ITR1")("TaxPaid")("BalTaxPayable")
'    End If

End Function

Function ImportVerification(jsonText As String)
On Error Resume Next

Dim jsonObject, jsonDictionary As Object
Dim Capacity As Variant
Dim Place, AssesseeVerName, FatherName, PAN As Variant

Set jsonObject = ParseJson(jsonText)

Capacity = jsonObject("verification")("capacity")
Place = jsonObject("verification")("place")
AssesseeVerName = jsonObject("verification")("declaration")("assesseeVerName")
FatherName = jsonObject("verification")("declaration")("fatherName")
PAN = jsonObject("verification")("declaration")("assesseeVerPAN")

If Capacity <> "" And Sheet3.Range("Ver.capacity").Locked = False Then
    If Capacity = "S" Or Capacity = "Self" Then
        Capacity = "Self"
    ElseIf Capacity = "R" Or Capacity = "Representative" Then
        Capacity = "Representative"
    Else
        Capacity = ""
    End If
    Sheet3.Range("Ver.capacity").Value = Capacity
End If

If Place <> "" And Sheet3.Range("Ver.Place").Locked = False Then
    Sheet3.Range("Ver.Place").Value = Place
End If

If AssesseeVerName <> "" And Sheet3.Range("Ver.AssesseeVerName").Locked = False Then
    Sheet3.Range("Ver.AssesseeVerName").Value = AssesseeVerName
End If

If FatherName <> "" And Sheet3.Range("Ver.FatherName").Locked = False Then
    Sheet3.Range("Ver.FatherName").Value = FatherName
End If

If PAN <> "" And Sheet3.Range("Ver.PAN").Locked = False Then
    Sheet3.Range("Ver.PAN").Value = PAN
End If

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
'Change.25.01.2023.102.80GC8.Prefil
Sub ImportSchedule80GD_pfl(jsonText As String)
On Error Resume Next
Dim jsonObject, jsonDictionary As Object
Dim Nodelist, node As Object
Dim Rect80G11, Rect80G12, cnt As Long
Dim TotalXMLRow, TotalExRow, TotalDiffRow, rowcount As Long
Dim Sch80GDoneeWithPanName, Sch80GAddrDetail, Sch80GcityOrTownOrDistrict, Sch80GStateCode, Sch80GpinCode, Sch80GDoneePAN, Sch80GDoneeARN, Sch80GDonationAmtCash, Sch80GDonationAmtOtherMode As Variant

        Set jsonObject = ParseJson(jsonText)
        Set Nodelist = jsonObject("Schedule80G")("Don50PercentApprReqd")("DoneeWithPan")
        
        Sch80GDoneeWithPanName = Sheet4.Range("Per5080G.DoneeWithPanName").Column
        Sch80GAddrDetail = Sheet4.Range("Per5080G.AddrDetail").Column
        Sch80GcityOrTownOrDistrict = Sheet4.Range("Per5080G.CityOrTownOrDistrict").Column
        Sch80GStateCode = Sheet4.Range("Per5080G.StateCode").Column
        Sch80GpinCode = Sheet4.Range("Per5080G.PinCode").Column
        Sch80GDoneePAN = Sheet4.Range("Per5080G.DoneePAN").Column
        Sch80GDoneeARN = Sheet4.Range("Per5080G.ArnNbr").Column
        Sch80GDonationAmtCash = Sheet4.Range("Per5080G.DonationAmt").Column
        Sch80GDonationAmtOtherMode = Sheet4.Range("Per5080G.DonationAmtOther").Column


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
            Sheet4.Range("Per5080G.ArnNbr").ClearContents
            Sheet4.Range("Per5080G.DonationAmt").ClearContents
            Sheet4.Range("Per5080G.DonationAmtOther").ClearContents
        End If
        If (TotalDiffRow > 0) Then
        'Here i need a function for addrow
        AddDiffRows_80G_D (TotalDiffRow)
        End If

rowcount = getRowNo(Sheet4.Range("Per5080G.DoneeWithPanName").name)
rowcount = rowcount - 1
        cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        
            If Sheet4.Cells(rowcount, Sch80GDoneeWithPanName).Locked = False Then
                Sheet4.Cells(rowcount, Sch80GDoneeWithPanName).Value = UCase(node("DoneeWithPanName"))
            End If
            
            If Sheet4.Cells(rowcount, Sch80GAddrDetail).Locked = False Then
                Sheet4.Cells(rowcount, Sch80GAddrDetail).Value = UCase(node("AddressDetail")("AddrDetail"))
            End If
            
            If Sheet4.Cells(rowcount, Sch80GcityOrTownOrDistrict).Locked = False Then
                Sheet4.Cells(rowcount, Sch80GcityOrTownOrDistrict).Value = UCase(node("AddressDetail")("CityOrTownOrDistrict"))
            End If
            
            If Sheet4.Cells(rowcount, Sch80GStateCode).Locked = False Then
                Dim iState As Variant
                iState = node("AddressDetail")("StateCode")
                If Len(iState) = "1" Then
                iState = "0" & iState
                End If
                If iState = "99" Then
                iState = ""
                End If
                Sheet4.Cells(rowcount, Sch80GStateCode).Value = Findtext(iState, "StateList")
            End If
            
            If Sheet4.Cells(rowcount, Sch80GpinCode).Locked = False Then
                Sheet4.Cells(rowcount, Sch80GpinCode).Value = UCase(node("AddressDetail")("PinCode"))
            End If
            
            If Sheet4.Cells(rowcount, Sch80GDoneePAN).Locked = False Then
                Sheet4.Cells(rowcount, Sch80GDoneePAN).Value = UCase(node("DoneePAN"))
            End If
            
            If Sheet4.Cells(rowcount, Sch80GDoneeARN).Locked = False Then
                Sheet4.Cells(rowcount, Sch80GDoneeARN).Value = UCase(node("DoneeARN"))
            End If
            
            If Sheet4.Cells(rowcount, Sch80GDonationAmtCash).Locked = False Then
                Sheet4.Cells(rowcount, Sch80GDonationAmtCash).Value = (node("DonationAmtCash"))
            End If
            
            If Sheet4.Cells(rowcount, Sch80GDonationAmtOtherMode).Locked = False Then
                Sheet4.Cells(rowcount, Sch80GDonationAmtOtherMode).Value = (node("DonationAmtOtherMode"))
            End If

        cnt = cnt + 1
    Next node
Rect80G11 = cnt
End Sub
'End Change=======

'Malli

'Form10IA.dateOfFiling 'Form10IA.acknNumber
Function ImportSchedule80DD_80U_pfl(jsonText As String)
On Error Resume Next
Dim init, jsonObject As Object
Set jsonObject = ParseJson(jsonText)
'Set init = jsonObject("Form10IA")
'If init <> Empty Then

'If init.exists("dateOfFiling") Then

Dim schedule80DDObject As Object
'Set schedule80DDObject = init("schedule80D")

'Malli AY_2025_26
'Dim dateOfFiling, acknNumber As Object
Dim dateOfFiling, acknNumber As Variant
'------------------------------------------

'dateOfFiling = jsonObject("Form10IA")("dateOfFiling")
acknNumber = jsonObject("Form10IA")("acknNumber")

'If dateOfFiling <> "" Then
''Dim dob As Variant
'dob = Mid(dateOfFiling, 9, 2) & "/" & Mid(dateOfFiling, 6, 2) & "/" & Mid(dateOfFiling, 1, 4)
'If Sheet14.Range("DatefilingFm10IA_80DD").Locked = False Then
'   Sheet14.Range("DatefilingFm10IA_80DD").Value = dob      'DatefilingFm10IA_80DD
'End If

'If Sheet14.Range("DatefilingFm10IA_80U").Locked = False Then
'   Sheet14.Range("DatefilingFm10IA_80U").Value = dob       'DatefilingFm10IA_80U
'End If
'
'End If

          'Commented by Shrutika- 23/01/2026 AY-26-27

'If acknNumber <> "" And Sheet14.Range("AckNoFm10IAfiled_80DD").Locked = False Then
'        Sheet14.Range("AckNoFm10IAfiled_80DD").Value = acknNumber   'AckNoFm10IAfiled_80DD
'End If
'If acknNumber <> "" And Sheet14.Range("AckNoFm10IAfiled_80U").Locked = False Then
'Sheet14.Range("AckNoFm10IAfiled_80U").Value = acknNumber            'AckNoFm10IAfiled_80U
'End If
'-------------------------------------------------------'Commented by Shrutika- 23/01/2026 AY-26-27

End Function

Function ImportSchedule80U_pfl(jsonText As String)
On Error Resume Next
Dim init, jsonObject As Object
Set jsonObject = ParseJson(jsonText)
'Set init = jsonObject("Form10IA")
'If init <> Empty Then

'If init.exists("dateOfFiling") Then
Dim dateOfFiling, acknNumber As Variant

'Dim schedule80DDObject As Object
'Set schedule80DDObject = init("schedule80D")


'dateOfFiling = jsonObject("Form10IA")("dateOfFiling")
'acknNumber = jsonObject("Form10IA")("acknNumber")

'Section80UUsrType = jsonObject("lastFiledITR")("usrDeductUndChapVIAType")("section80UUsrType")

'If dateOfFiling <> "" Then
''Dim dob As Variant
'dob = Mid(dateOfFiling, 9, 2) & "/" & Mid(dateOfFiling, 6, 2) & "/" & Mid(dateOfFiling, 1, 4)
'If Sheet14.Range("DatefilingFm10IA_80U").Locked = False Then
'Sheet14.Range("DatefilingFm10IA_80U").Value = dob  'DatefilingFm10IA_80DD
'End If

'If acknNumber <> "" And Sheet14.Range("AckNoFm10IAfiled_80U").Locked = False Then
'Sheet14.Range("AckNoFm10IAfiled_80U").Value = acknNumber   'AckNoFm10IAfiled_80DD
'End If

End Function

'Form10IA.dateOfFiling

Sub alerts_c()
Application.EnableEvents = True
End Sub


 
'Malli_AY_2026-27    29/01/2026
Function ImportScheduleHP_Pfl(jsonText As String)
On Error Resume Next

Dim jsonObject As Object
    Dim XpathOfHP As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist, nodeList1, Nodelist2, Nodelist3, init As Object
    Dim node, Node1, Node2, Node3 As Object
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    Dim iState, sState As Variant
    Dim iCountry, sCountry As Variant
    Dim iLetOut, sLetOut As Variant
    Dim iTotalCoRow, sTotalCoRow, TotalDiffCoRow As Variant
    Dim iTotalTenRow, sTotalTenRow, TotalDiffTenRow As Variant
    Dim iTotal24BRow, sTotal24BRow, TotalDiff24BRow As Variant
    Dim rowcount, cnt As Long
    
    
    Dim Node_1, Node_1M, Nodelist_2, Node_2, Node_1MM, Node_3, Nodelist_3 As Object
    Dim panOfTenant_count As Integer
 
 
 '30/01/2026
 Dim ais_Tenchk_pfl As Boolean
 ais_Tenchk_pfl = False
 '--------------------
 
 
Set jsonObject = ParseJson(jsonText)

'Primary source

Dim init_1 As Object
panOfTenant_count = 0

If jsonObject("form26as") <> Null Then
Set init_1 = jsonObject("form26as")("scheduleHP")

If init_1 <> Empty Then


Set Node_1M = init_1("propertyDetails")
     panOfTenant_count = 0
     For Each Node_1 In Node_1M
     sTotalCoRow = 0
  
       Set Node_1MM = Node_1("tenantDetails")
          For Each Node_2 In Node_1MM
            Dim panOfTenant_Nodechk As Variant
            panOfTenant_Nodechk = Node_2("panOfTenant")
                 If panOfTenant_Nodechk <> "" Then
                 panOfTenant_count = panOfTenant_count + 1
                 End If
       Next Node_2
  Next Node_1
End If
End If

If panOfTenant_count > 0 Then

rowcount = 0

'Malli_AY_2026-27  25/03/2026

  'Debug.Print init_1("propertyDetails").count

  If init_1("propertyDetails").count > 2 Then

  MsgBox "* ""Since you have more than two house properties, please proceed with filing ITR-2 or ITR-3.""", vbOKOnly, "Alert:"

  End If

'------------------------------
 
  
For Each Node_1 In Node_1M
       
                Dim typeOfHP_tre As Boolean
                typeOfHP_tre = True
                rowcount = rowcount + 1
                
                
                iTotalTenRow = Sheet19.Range("HP.NameofTenant" & rowcount).Rows.count
                sTotalTenRow = Node_1("tenantDetails").count
                TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)

                        If TotalDiffTenRow > 0 Then
                        Sheet19.Activate
                        AddPropertyTenant (TotalDiffTenRow)
                        End If
               
                cnt = 0
                cnt = getRowNo(Sheet19.Range("HP.NameofTenant" & rowcount).name)
                cnt = cnt - 1
            
            
iLetOut = Node_1("typeOfHP")
                               
        If iLetOut = "Y" Or iLetOut = "L" Then
            sLetOut = "Let Out"
        ElseIf iLetOut = "D" Then
            sLetOut = "Deemed Let Out"
        ElseIf iLetOut = "N" Or iLetOut = "S" Then
            sLetOut = "Self Occupied"
        Else
            sLetOut = "(Select)"
        End If
         
        If Sheet19.Range("HP.ifLetOut" & rowcount).Locked = False Then
            Sheet19.Range("HP.ifLetOut" & rowcount).Value = sLetOut
        End If
  
Set Nodelist_2 = Node_1("tenantDetails")

For Each Node_2 In Nodelist_2
              
     If UCase(Node_2("panOfTenant")) <> "" Then
            
            cnt = cnt + 1

                    If Sheet19.Cells(cnt, Sheet19.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                        Sheet19.Cells(cnt, Sheet19.Range("HP.NameofTenant" & rowcount).Column).Value = UCase(Node_2("nameOfTenant"))
                    End If
                    If Sheet19.Cells(cnt, Sheet19.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                        Sheet19.Cells(cnt, Sheet19.Range("HP.PANofTenant" & rowcount).Column).Value = UCase(Node_2("panOfTenant"))
                        ais_Tenchk_pfl = True
                    End If
                    If Sheet19.Cells(cnt, Sheet19.Range("HP.AadharofTenant" & rowcount).Column).Locked = False Then
                        Sheet19.Cells(cnt, Sheet19.Range("HP.AadharofTenant" & rowcount).Column).Value = UCase(Node_2("aadhaarofTenant"))
                    End If
 
                    If Sheet19.Cells(cnt, Sheet19.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                        Sheet19.Cells(cnt, Sheet19.Range("HP.TANofTenant" & rowcount).Column).Value = UCase(Node_2("panTANofTenant"))
                        ais_Tenchk_pfl = True
                    End If
                    
            
             End If
             
Next Node_2
                
            
                If Sheet19.Range("HP.AnnualLetableValue" & rowcount).Locked = False Then
                Sheet19.Range("HP.AnnualLetableValue" & rowcount).Value = UCase(Node_1("grossRent"))
                End If
                
                
Next Node_1







Else:
   'Secoundary source
   
   '\\* AY_2026-27 onwards for ITR-1 & ITR-4 HP schedule Type of House property? field has two types of prefill sources under last field ITR which is _
   '1  lastFiledITR.typeOfHP   ( it should prefill only single block)
   '2  lastFiledITR.scheduleHP.PropertyDetails.ifLetOut  ( it should prefill only multiple block)
   '*//

    If panOfTenant_count <= 0 Then
    
    
    Dim typeOfHP_1stLFIsorce, typeOfHP_1stLFIsorcepfl As Variant
    
    typeOfHP_1stLFIsorce = jsonObject("lastFiledITR")("typeOfHP")
    
    If typeOfHP_1stLFIsorce <> "" Then
    'LFI Primary source
     
                    If UCase(typeOfHP_1stLFIsorce) = UCase("Y") Or UCase(typeOfHP_1stLFIsorce) = UCase("L") Then
                        typeOfHP_1stLFIsorcepfl = "Let Out"
                    ElseIf UCase(typeOfHP_1stLFIsorce) = UCase("D") Then
                        typeOfHP_1stLFIsorcepfl = "Deemed Let Out"
                    ElseIf UCase(typeOfHP_1stLFIsorce) = UCase("N") Or UCase(typeOfHP_1stLFIsorce) = UCase("S") Then
                        typeOfHP_1stLFIsorcepfl = "Self Occupied"
                    Else
                        typeOfHP_1stLFIsorcepfl = "(Select)"
                    End If
                     
                    If Sheet19.Range("HP.ifLetOut1").Locked = False Then
                        Sheet19.Range("HP.ifLetOut1").Value = typeOfHP_1stLFIsorcepfl
                    End If
    
     Else:
    
        
        'LFI Secoundary source
         Set init = jsonObject("lastFiledITR")("scheduleHP")
        
        
                    If init <> Empty Then
                    If init.exists("propertyDetails") Then
                    
                        Set Nodelist = jsonObject("lastFiledITR")("scheduleHP")("propertyDetails")
                     
                        rowcount = 0
                        cnt = 0
                        'Malli_AY_2026-27  25/03/2026

                     'Debug.Print Nodelist.count

                    If Nodelist.count > 2 Then

                         MsgBox "* ""Since you have more than two house properties, please proceed with filing ITR-2 or ITR-3.""", vbOKOnly, "Alert:"

                    End If

                    '--------------------------
 
                    
                        
                        For Each node In Nodelist
                            rowcount = rowcount + 1
                            
                                If Sheet19.Range("HP.AddrDetail" & rowcount).Locked = False Then
                                    Sheet19.Range("HP.AddrDetail" & rowcount).Value = UCase(node("addressDetailWithZipCode")("addrDetail"))
                                End If
                                If Sheet19.Range("HP.CityOrTownOrDistrict" & rowcount).Locked = False Then
                                    Sheet19.Range("HP.CityOrTownOrDistrict" & rowcount).Value = UCase(node("addressDetailWithZipCode")("cityOrTownOrDistrict"))
                                End If
                                
                                iState = UCase(node("addressDetailWithZipCode")("stateCode"))
                               'SIT-109896  'Malli 06/02/2026
                                'sState = Findtext(iState, "State")
                                sState = Findtext(iState, "StateList")
                                '-------------------------------------
                                 Sheet19.Range("HP.StateCode" & rowcount).Value = sState
                                iCountry = UCase(node("addressDetailWithZipCode")("countryCode"))
                                'SIT-109896  'Malli 06/02/2026
                                'sCountry = Findtext(iCountry, "Country")
                                sCountry = Findtext(iCountry, "CountList")
                                '-----------------------
                                 
                                If Sheet19.Range("HP.CountryCode" & rowcount).Locked = False Then
                                    Sheet19.Range("HP.CountryCode" & rowcount).Value = sCountry
                                End If
                                If Sheet19.Range("HP.PinCode" & rowcount).Locked = False Then
                                    Sheet19.Range("HP.PinCode" & rowcount).Value = UCase(node("addressDetailWithZipCode")("pinCode"))
                                End If
                                If Sheet19.Range("HP.ZipCode" & rowcount).Locked = False Then
                                    Sheet19.Range("HP.ZipCode" & rowcount).Value = UCase(node("addressDetailWithZipCode")("zipCode"))
                                End If
                                
                                Dim OwnerProperty_HP
                                
                                OwnerProperty_HP = UCase(node("propertyOwner"))
                                
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
                                
                                If Sheet19.Range("HP.OwnerPropertyDescription" & rowcount).Locked = False Then
                                    Sheet19.Range("HP.OwnerPropertyDescription" & rowcount).Value = node("propertyOwnerOther")
                                End If
                                
                                Dim Co_Ownedflag As Variant
                                
                                Co_Ownedflag = UCase(node("propCoOwnedFlg"))
                                If UCase(Co_Ownedflag) = "YES" Or UCase(Mid(Co_Ownedflag, 1, 1) = "Y") Then
                                Co_Ownedflag = "Yes"
                                ElseIf UCase(Co_Ownedflag) = "NO" Or UCase(Mid(Co_Ownedflag, 1, 1) = "N") Then
                                Co_Ownedflag = "No"
                                Else
                                Co_Ownedflag = "(Select)"
                                End If
                                
                                Sheet19.Range("HP.CoOwnedYN" & rowcount).Value = Co_Ownedflag
                                 
                                Dim hpshare As Variant
                                hpshare = UCase(node("asseseeShareProperty"))
                                
                                If Sheet19.Range("HP.SharePercent" & rowcount).Locked = False Then
                                    Sheet19.Range("HP.SharePercent" & rowcount).Value = hpshare
                                End If
                                
                        
                        
                        If node.exists("coOwners") Then
                        
                        
                            'Co-Owner
                            iTotalCoRow = Sheet19.Range("HP.Co.Name" & rowcount).Rows.count
                             
                            sTotalCoRow = node("coOwners").count
                            
                            TotalDiffCoRow = WorksheetFunction.Max((sTotalCoRow - iTotalCoRow), 0)
                            If TotalDiffCoRow > 0 Then
                                Sheet19.Activate
                                AddPropertyCoOWners (TotalDiffCoRow)
                            End If
                            
                            
                                Set nodeList1 = node("coOwners")
                                cnt = 0
                                cnt = getRowNo(Sheet19.Range("HP.Co.Name" & rowcount).name)
                                cnt = cnt - 1
                                
                                For Each Node1 In nodeList1
                                        cnt = cnt + 1
                                            If Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Name" & rowcount).Column).Locked = False Then
                                                Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Name" & rowcount).Column).Value = UCase(Node1("nameCoOwner"))
                                            End If
                                            If Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Pan" & rowcount).Column).Locked = False Then
                                                Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Pan" & rowcount).Column).Value = UCase(Node1("panCoOwner"))
                                            End If
                                            If Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Aadhaar" & rowcount).Column).Locked = False Then
                                                Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Aadhaar" & rowcount).Column).Value = UCase(Node1("aadhaarCoOwner"))
                                            End If
                                            If Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Share" & rowcount).Column).Locked = False Then
                                                Sheet19.Cells(cnt, Sheet19.Range("HP.Co.Share" & rowcount).Column).Value = UCase(Node1("percentShareProperty"))
                                            End If
                                        
                                Next Node1
                        End If
                                
                                iLetOut = UCase(node("ifLetOut"))
                                
                                
                                If iLetOut = "Y" Or iLetOut = "L" Then
                                    sLetOut = "Let Out"
                                ElseIf iLetOut = "D" Then
                                    sLetOut = "Deemed Let Out"
                                ElseIf iLetOut = "N" Or iLetOut = "S" Then
                                    sLetOut = "Self Occupied"
                                Else
                                    sLetOut = "(Select)"
                                End If
                                
                                
                                Sheet19.Range("HP.ifLetOut" & rowcount).Value = sLetOut
                        
                    If node.exists("tenantDetails") Then
                    
                                'Tenant
                                iTotalTenRow = Sheet19.Range("HP.NameofTenant" & rowcount).Rows.count
                                sTotalTenRow = node("tenantDetails").count
                                TotalDiffTenRow = WorksheetFunction.Max((sTotalTenRow - iTotalTenRow), 0)
                                
                                If TotalDiffTenRow > 0 Then
                                    Sheet19.Activate
                                    AddPropertyTenant (TotalDiffTenRow)
                                End If
                                
                                Set Nodelist2 = node("tenantDetails")
                                cnt = 0
                                cnt = getRowNo(Sheet19.Range("HP.NameofTenant" & rowcount).name)
                                cnt = cnt - 1
                                For Each Node2 In Nodelist2
                                    cnt = cnt + 1
                                                If Sheet19.Cells(cnt, Sheet19.Range("HP.NameofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet19.Cells(cnt, Sheet19.Range("HP.NameofTenant" & rowcount).Column).Value = UCase(Node2("nameOfTenant"))
                                                End If
                                                If Sheet19.Cells(cnt, Sheet19.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet19.Cells(cnt, Sheet19.Range("HP.PANofTenant" & rowcount).Column).Value = UCase(Node2("panOfTenant"))
                                                    ais_Tenchk_pfl = True
                                                End If
                                                If Sheet19.Cells(cnt, Sheet19.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet19.Cells(cnt, Sheet19.Range("HP.TANofTenant" & rowcount).Column).Value = UCase(Node2("panTANofTenant"))
                                                    ais_Tenchk_pfl = True
                                                End If
                                                If Sheet19.Cells(cnt, Sheet19.Range("HP.AadharofTenant" & rowcount).Column).Locked = False Then
                                                        If Node2.Exsist("AadhaarofTenant") Then
                                                            Sheet19.Cells(cnt, Sheet19.Range("HP.AadharofTenant" & rowcount).Column).Value = UCase(Node2("AadhaarofTenant"))
                                                        End If
                                                End If
                                Next Node2
                            End If
                                
                                
   
                                
     '24B<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
     
 
     If node.exists("rentdetails") Then

       'Schedule24(B)
       iTotal24BRow = Sheet19.Range("LoanfrmBankOrInstitute.24b" & rowcount).Rows.count
       sTotal24BRow = node("rentdetails")("section24B")("section24BDtls").count
       TotalDiff24BRow = WorksheetFunction.Max((sTotal24BRow - iTotal24BRow), 0)

       If TotalDiff24BRow > 0 Then
           Sheet19.Activate
           AddSection24b (TotalDiff24BRow)
       End If


       Set Nodelist3 = node("rentdetails")("section24B")("section24BDtls")
           cnt = 0
           cnt = getRowNo(Sheet19.Range("LoanfrmBankOrInstitute.24b" & rowcount).name)
           cnt = cnt - 1

           For Each Node3 In Nodelist3
               cnt = cnt + 1

        If Sheet19.Cells(cnt, Sheet19.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).Locked = False Then
            Dim LoanTknFrom_24B, LoanTknFrom_24B_pfl

            LoanTknFrom_24B_pfl = Node3("loanTknFrom")
                If UCase(LoanTknFrom_24B_pfl) = UCase("B") Then
                        LoanTknFrom_24B = "Bank "
                ElseIf UCase(LoanTknFrom_24B_pfl) = UCase("I") Then
                        LoanTknFrom_24B = "Other than Bank"
                Else
                        LoanTknFrom_24B = "(Select)"
                End If

            Sheet19.Cells(cnt, Sheet19.Range("LoanfrmBankOrInstitute.24b" & rowcount).Column).Value = LoanTknFrom_24B
        End If

        If Sheet19.Cells(cnt, Sheet19.Range("bankName.24b" & rowcount).Column).Locked = False Then
            Sheet19.Cells(cnt, Sheet19.Range("bankName.24b" & rowcount).Column).Value = UCase(Node3("bankOrInstnName"))
        End If

        If Sheet19.Cells(cnt, Sheet19.Range("loanAccNum.24b" & rowcount).Column).Locked = False Then
            Sheet19.Cells(cnt, Sheet19.Range("loanAccNum.24b" & rowcount).Column).Value = UCase(Node3("loanAccNoOfBankOrInstnRefNo"))
        End If

        If Sheet19.Cells(cnt, Sheet19.Range("loanDate.24b" & rowcount).Column).Locked = False Then
            Dim str24BDate As Variant
            str24BDate = Node3("dateofLoan")
            If str24BDate <> "" Then
                    str24BDate = Mid(str24BDate, 9, 2) & "/" & Mid(str24BDate, 6, 2) & "/" & Mid(str24BDate, 1, 4)
                    Sheet19.Cells(cnt, Sheet19.Range("loanDate.24b" & rowcount).Column).Value = str24BDate
                       str24BDate = ""
            End If
        End If

        If Sheet19.Cells(cnt, Sheet19.Range("loanAmt.24b" & rowcount).Column).Locked = False Then
            Sheet19.Cells(cnt, Sheet19.Range("loanAmt.24b" & rowcount).Column).Value = UCase(Node3("totalLoanAmt"))
        End If



           Next Node3
       End If
  
 
     
     '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
     
     
     



           cnt = cnt + 1
       Next node
       
       
End If
End If


End If
End If
End If


'Malli_30/01/2026
Dim init_ais As Object
Dim Node_1ais, Node_ais, Node_2ais, Node_1MMais As Object
Dim iTotalTenRow_ais, sTotalTenRow_ais, TotalDiffTenRow_ais, panOfTenant_counts_ais As Variant

If ais_Tenchk_pfl <> True Then

        If jsonObject("ais") <> Null Then
        Set init_ais = jsonObject("ais")("ScheduleHP")
        
                If init_ais <> Empty Then
                Set Node_ais = init_ais("PropertyDetails")
                     panOfTenant_counts_ais = 0
                     
                  For Each Node_1ais In Node_ais
                     sTotalCoRow = 0
                  
                       Set Node_1MMais = Node_1ais("TenantDetails")
                          For Each Node_2ais In Node_1MMais
                            Dim ais_panOfTenant_Nodechk, ais_PANTANofTenant_Nodechk As Variant
                            
                            ais_panOfTenant_Nodechk = Node_2ais("PANOfTenant")
                            ais_PANTANofTenant_Nodechk = Node_2ais("PANTANofTenant")
                            
                                 If ais_panOfTenant_Nodechk <> "" Or ais_PANTANofTenant_Nodechk <> "" Then
                                 panOfTenant_counts_ais = panOfTenant_counts_ais + 1
                                 End If
                                 
                       Next Node_2ais
                  Next Node_1ais
                  
                End If
                
                
                If panOfTenant_counts_ais > 0 Then
                rowcount = 0
                
                For Each Node_1ais In Node_ais
                rowcount = rowcount + 1
                
                        iTotalTenRow_ais = Sheet19.Range("HP.NameofTenant" & rowcount).Rows.count
                        sTotalTenRow_ais = Node_1ais("TenantDetails").count
                        TotalDiffTenRow_ais = WorksheetFunction.Max((sTotalTenRow_ais - iTotalTenRow_ais), 0)
        
                                If TotalDiffTenRow_ais > 0 Then
                                Sheet19.Activate
                                AddPropertyTenant (TotalDiffTenRow_ais)
                                End If
                       
                        cnt = 0
                        cnt = getRowNo(Sheet19.Range("HP.NameofTenant" & rowcount).name)
                        cnt = cnt - 1
                
                        If Sheet19.Range("HP.ifLetOut" & rowcount).Locked = False Then
                        Sheet19.Range("HP.ifLetOut" & rowcount).Value = "Let Out"
                        End If
                
                
                            Set Node_1MMais = Node_1ais("TenantDetails")
                                  For Each Node_2ais In Node_1MMais
                                        cnt = cnt + 1

                                                If Sheet19.Cells(cnt, Sheet19.Range("HP.PANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet19.Cells(cnt, Sheet19.Range("HP.PANofTenant" & rowcount).Column).Value = UCase(Node_2ais("PANOfTenant"))
                                                End If
                                                
                                                If Sheet19.Cells(cnt, Sheet19.Range("HP.TANofTenant" & rowcount).Column).Locked = False Then
                                                    Sheet19.Cells(cnt, Sheet19.Range("HP.TANofTenant" & rowcount).Column).Value = UCase(Node_2ais("PANTANofTenant"))
                                                End If
 
                                  Next Node_2ais
                
                
                      cnt = cnt + 1
                Next Node_1ais
                
                End If
                
                
        End If
End If
'----------------

End Function



















