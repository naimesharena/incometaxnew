Attribute VB_Name = "GenerateJson"
' 09/11/2022
Option Explicit
Public xmlLinesCount As Variant
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

Function getHashIteration() As String
'Changes from 1243 to 1977
'    getHashIteration = "1977"
'     getHashIteration = "1978"   'Changed from 1977 to 1978 by Ankita on 02/01/2025
     getHashIteration = "1988"   'Ankita_29/01/2026
End Function

Function getHashKey() As String
'Changes from zcKkZH7jksLZWLYT to RoduFpNqMzQlWO9Q
'    getHashKey = "RoduFpNqMzQlWO9Q" 'changed by Chetan C M on 03/01/2025, Commented on 18/02/2025
'     getHashKey = "LO3QtH59fGuVaETa"  'Changed by Ankita on 02/01/2025
     getHashKey = "HZX4oKH11zARYIb2"  ''Ankita_29/01/2026
End Function

Sub Generate_JSON()

On Error GoTo Endline1

    Dim JSONFileName As String
    Dim AssesseePan
'    sPassword = EfilingCommon.getmsgstate
'    Sheet1.Unprotect Password:=sPassword
    'Dim FSO, FSO1, jsonFileObject, jsonFileExport As Object
    
    'Set jsonFileObject = CreateObject("Scripting.FileSystemObject")
    'EfilingCommon.InitializeAllPublicVariables
       
    ValidateSchedulePI
    
    
    'Validate80G_All
    'ValidatePartBTI_BTTI_Verification
    'ValidateBA
      
    If (((Sheet3.Range("IncD.TotalTaxesPaid").Value - Sheet3.Range("IncD.TCS").Value) > 0) And _
            (IIf(Sheet1.Range("IncD.TotalHeadSalaries").Value = "", 0, Sheet1.Range("IncD.TotalHeadSalaries").Value) And _
            IIf(Sheet1.Range("IncD.IncomeFromSal").Value = "", 0, Sheet1.Range("IncD.IncomeFromSal").Value) And _
            IIf(Sheet1.Range("IncD.Allowances").Value = "", 0, Sheet1.Range("IncD.Allowances").Value) And _
            IIf(Sheet1.Range("IncD.Perquisites").Value = "", 0, Sheet1.Range("IncD.Perquisites").Value) And _
            IIf(Sheet1.Range("IncD.Profits").Value = "", 0, Sheet1.Range("IncD.Profits").Value) And _
            IIf(Sheet1.Range("IncomeNotified89A").Value = "", 0, Sheet1.Range("IncomeNotified89A").Value) And _
            IIf(Sheet1.Range("IncomeNotifiedOther89A").Value = "", 0, Sheet1.Range("IncomeNotifiedOther89A").Value) And _
            IIf(Sheet1.Range("IncD.Deduction16").Value = "", 0, Sheet1.Range("IncD.Deduction16").Value) And _
            IIf(Sheet1.Range("IncD.Deduction16ia").Value = "", 0, Sheet1.Range("IncD.Deduction16ia").Value) And _
            IIf(Sheet1.Range("IncD.IncomeHeadHouseProperty").Value = "", 0, Sheet1.Range("IncD.IncomeHeadHouseProperty").Value) And _
            IIf(Sheet1.Range("IncD.GrossRentRecieved").Value = "", 0, Sheet1.Range("IncD.GrossRentRecieved").Value) And _
            IIf(Sheet1.Range("IncD.TaxPaidLocalAuthorities").Value = "", 0, Sheet1.Range("IncD.TaxPaidLocalAuthorities").Value) And _
            IIf(Sheet1.Range("IncD.AnnualValue").Value = "", 0, Sheet1.Range("IncD.AnnualValue").Value) And _
            IIf(Sheet1.Range("IncD.StandardDeduction").Value = "", 0, Sheet1.Range("IncD.StandardDeduction").Value) And _
            IIf(Sheet1.Range("IncD.InterestBorrowedCapital").Value = "", 0, Sheet1.Range("IncD.InterestBorrowedCapital").Value) And _
            IIf(Sheet1.Range("IncD.IncomeFromOS").Value = "", 0, Sheet1.Range("IncD.IncomeFromOS").Value))) Then

            fmsgbox "No Income details or tax computation is provided in ITR but details regarding taxes paid is provided"
            CloseMsg
        End If





   'Newly updated by Bindu on 4th Feb 2025 as per DE V3
   
    'If (Sheet1.Range("IncD.TotalIncome").Value > 5000000) Then
    If ((Sheet1.Range("IncD.TotalIncome_New").Value - Sheet1.Range("IncD.CG_LTCG")) > 5000000) Then
        fmsgbox "ITR-1 is for individuals having Income from Salaries, One House Property, Other Sources(Interest  etc.)and having taxable income upto Rs.50 lakh. Please fill other ITR"
        CloseMsg
    End If
    
       
    

    firstName = Sheet1.Range("sheet1.FirstName").Value
    middleName = Sheet1.Range("sheet1.MiddleName").Value
    residenceName = Sheet1.Range("Sheet1.ResidenceName").Value
    'DesigOfficerWardorCircle = Sheet1.Range("sheet1.DesigOfficerWardorCircle").Value
    'SpousePAN = Sheet1.Range("sheet1.PANOFSPOUSE").Value
    TotalTaxPayable = Sheet1.Range("IncD.TotalTaxPayable").Value
    
    
    TypeOfHP = Sheet1.Range("IncD.TypeOfHP").Value
    'IncomeFromHP = Sheet1.Range("IncD.IncomeFromHP").Value
    IncomeFromOS = Sheet1.Range("IncD.IncomeFromOS").Value
    'GrossTotIncome = Sheet1.Range("IncD.GrossTotIncome").Value
     'Konda Updated AY-2025-26 V0.3.2 on 07-02-2025
    GrossTotIncomeIncLTCG112A = Sheet1.Range("IncD.GrossTotIncome_New").Value
    
    'Newly updated by Bindu as Permission DE V3 on 4th feb 2024
    GrossTotIncome = Sheet1.Range("IncD.GrossTotIncome").Value
    

'    InitProgBar
'    ProgressFrameCaption = "Generating XML"
'    mainProcCaption = "Writing XML"
'    noOfProcessMain = 11
    AssesseePan = Sheet1.Range("sheet1.PAN").Value
    JSONFileName = ThisWorkbook.Path & "\ITR1_" & AssesseePan & ".json"
'        Set jsonFileExport = jsonFileObject.CreateTextFile(JSONFileName, True)
'        jsonFileExport.WriteLine (ToJsonFormat())
'Open JSONFileName For Output As #1
Dim str
str = ToJsonFormat()
'Print #1, str
'Close #1
'MsgBox ToJsonFormat()
        
ConvertJSONToString2 (str)
 
ProgressBarHide

Sheet1.Protect Password:=sPassword
Sheet1.Protect Password:=sPassword
fmsgbox "Json is saved in the path :  " & JSONFileName & Chr(13) & Chr(13) & "Please upload the return(json) on below link" & Chr(13) & Chr(13) & "To upload the saved JSON, Go to https://incometaxindiaefiling.gov.in -> Login -> e-File -> Income Tax Return." ', vbOKOnly, "ITR-1 AY 2020-2120"
ThisWorkbook.Unprotect Password:=sPassword
Sheet7.Visible = xlSheetHidden
ThisWorkbook.Protect Password:=sPassword
'Newly added by Bindu on 10/01/2025
Sheet3.Activate
Sheet1.Activate

Endline1:

End Sub

Public Function ConvertJSONToString2(ByVal str As String)
On Error GoTo Endline2


Dim ts(1) As Object
    Dim count As Variant
    Dim FSO, jsonObject As Object
    Dim strXml, hashcode As String
    Dim SecretKey As String
    Dim iteration As Long
    Dim FileContents As String
    Dim SEARCH_FOR As String
    Dim REPLACE_WITH As String
    Dim arrSearchThis() As Variant
    Dim objTS As Variant
    Dim i, DigestFoundat As Variant
    Dim jsonText As String
    iteration = getHashIteration
    SecretKey = getHashKey
    
'    Set FSO = CreateObject("Scripting.FileSystemObject")
'    Open JSONFileName For Input As #1
'        jsonText = Input$(LOF(1), 1)
'    Close #1
    i = 0
    InitProgBar
    ProgressFrameCaption = "JSON Generation"
    mainProcCaption = "Encrypting JSON"
    noOfProcessMain = 3
    UpdateProgressBar
    Set jsonObject = ImportJson.ParseJson(str)
    UpdateProgressBar
'    hashcode = Base64_HMACSHA256_JSON(iteration, str, SecretKey)
     hashcode = HMACSHA256A(str, StrConv(SecretKey, vbFromUnicode), iteration) ' added by Chetan C M on 03/01/2025
    jsonObject("ITR")("ITR1")("CreationInfo")("Digest") = hashcode
'    Dim JSONFileName1
'    JSONFileName1 = ThisWorkbook.Path & "\" & AssesseePan & "1.json"
    Dim filepath
    Dim AssesseePan
    AssesseePan = Sheet1.Range("sheet1.PAN").Value
    AssesseePan = "ITR1_" & AssesseePan
    
    filepath = ThisWorkbook.Path & "\" & AssesseePan & ".json"
    Open filepath For Output As #2
    subProcCaption = "Publishing JSON"
    Print #2, ToJson(jsonObject);
            UpdateProgressBar
    Close #2
    
Endline2:
    
End Function

Public Function Base64_HMACSHA256_JSON(ByVal iteration As Long, sTextToHash As String, ByVal sSharedSecretKey As String)

    Dim asc As Object, enc As Object
    Dim TextToHash() As Byte
    Dim SharedSecretKey() As Byte
    
    Dim tempTextToHash() As Byte
    Dim tempString, CountTextToHash As Variant
    Dim i, j As Variant
    
    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")
    ReDim TextToHash(0)
    subProcCaption = "Encrypting JSON"
    TextToHash = StrConv(sTextToHash, vbFromUnicode)
        UpdateProgressBar
    SharedSecretKey = asc.Getbytes_4(sSharedSecretKey)
    enc.Key = SharedSecretKey

    Dim byteS() As Byte
    byteS = enc.ComputeHash_2((TextToHash))
    
    For i = 1 To iteration
        byteS = enc.ComputeHash_2((byteS))
    Next
     
    Base64_HMACSHA256_JSON = EncodeBase64json(byteS)
    Set asc = Nothing
    Set enc = Nothing

End Function


Function EncodeBase64json(ByRef arrData() As Byte) As String

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64json = objNode.text

    Set objNode = Nothing
    Set objXML = Nothing

End Function

Function ToJsonFormat() As String
Dim str As String
Dim init, itr, itr1

Set init = CreateObject("Scripting.Dictionary")
Set itr = CreateObject("Scripting.Dictionary")
Set itr1 = CreateObject("Scripting.Dictionary")

itr1.add "CreationInfo", Form01Header()
itr1.add "Form_ITR1", Form_ITR1()
UpdateProgressBar
itr1.add "PersonalInfo", PersonalInfo()
UpdateProgressBar
itr1.add "FilingStatus", FilingStatus()
UpdateProgressBar
If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
itr1.add "PartA_139_8A", PartA_139_8A()
itr1.add "PartB-ATI", PartB_ATI()
End If
UpdateProgressBar
itr1.add "ITR1_IncomeDeductions", IncomeDeductions()
UpdateProgressBar
itr1.add "ITR1_TaxComputation", TaxComputation()
'konda added AY 2025-25 on 04-02-2025
itr1.add "LTCG112A", LTCG112A_New()
itr1.add "TaxPaid", TaxPaid()
itr1.add "Refund", Refund()
itr1.add "Verification", Verification()

If (Sheet5.Range("BacValue").Value) = 2 Then
'itr1.add "Schedule80D", Schedule80D() 'malli comented for replacing position
itr1.add "Schedule80G", Schedule80G()
itr1.add "Schedule80GGA", Schedule80GGA()
'80GGC 2024-25 Bindu
itr1.add "Schedule80GGC", Schedule80GGC()

itr1.add "Schedule80D", Schedule80D()
'Malli 80U_DD
itr1.add "Schedule80DD", Schedule80DD_1()
itr1.add "Schedule80U", Schedule80U_1()


'Malli_____________AY_2025_26 09/04/2025
itr1.add "Schedule80E", Schedule80E()
itr1.add "Schedule80EE", Schedule80EE()
itr1.add "Schedule80EEA", Schedule80EEA()
itr1.add "Schedule80EEB", Schedule80EEB()

itr1.add "Schedule80C", Schedule80C()
'Konda_05/05/2025_Commented as per New Schema_v0.7 and V0.7.1
'itr1.add "Schedule80CCC", Schedule80CCC()
'-----------------------------------------
'itr1.add "ScheduleUs24B", Schedule24B()
'---------------------------------------
End If

'Konda_29/01/2026========
'If Sheet5.Range("BacValue").Value = 1 And Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "S" And Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" Then
'      itr1.add "ScheduleUs24B", Schedule24B()
'ElseIf Sheet5.Range("BacValue").Value = 2 And Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" Then
'      itr1.add "ScheduleUs24B", Schedule24B()
'End If

'Konda AY_2025_26 06-05-2025------------
If Sheet5.Range("BacValue").Value = 2 And Mid(Range("sheet1.EmployerCategory1").Value, 1, 1) <> "N" Then
      itr1.add "ScheduleEA10_13A", ScheduleEA10_13A()

End If
'-------------------




itr1.add "TDSonSalaries", TDSonSalaries()
itr1.add "TDSonOthThanSals", TDSonOthThanSals()
itr1.add "ScheduleTDS3Dtls", ScheduleTDS3Dtls()
itr1.add "ScheduleTCS", ScheduleTCS()
itr1.add "TaxPayments", TaxPayments()
itr1.add "TaxReturnPreparer", TaxReturnPreparer()
itr.add "ITR1", itr1
init.add "ITR", itr

str = ToJson(init)
ToJsonFormat = str
End Function
Function Form01Header() As Object
    Dim jsonDictionary
    Dim CreationDate, SWVersionNo, SWCreatedBy, JSONCreatedBy, IntermediaryCity, Digest, JSONCreationDate As Variant

Set jsonDictionary = CreateObject("Scripting.Dictionary")

    SWVersionNo = getSWVersionNo
    SWCreatedBy = getSWCreatedBy
    JSONCreatedBy = getJSONCreatedBy
    JSONCreationDate = Range("DateOfProcessing").Value
'Changed by Riyaz on 31/12/2024 for SIT-85556
'    IntermediaryCity = getJSONCreatedBy
    IntermediaryCity = getIntermediaryCity
    Digest = "-"
    CreationDate = Range("DateOfProcessing").Value

    
    jsonDictionary("SWVersionNo") = SWVersionNo
    jsonDictionary("SWCreatedBy") = SWCreatedBy
    jsonDictionary("JSONCreatedBy") = JSONCreatedBy
    jsonDictionary("JSONCreationDate") = JSONCreationDate
    jsonDictionary("IntermediaryCity") = IntermediaryCity
    jsonDictionary("Digest") = Digest
    Set Form01Header = jsonDictionary
    
End Function
Function Form_ITR1() As Object
Dim jsonDictionary
Dim FormName, Description, AssessmentYear, SchemaVer, FormVer As Variant

Set jsonDictionary = CreateObject("Scripting.Dictionary")

FormName = getFormName
Description = getFormDescription
AssessmentYear = getAssessmentYear
SchemaVer = getSchemaVer
FormVer = getFormVer

    jsonDictionary("FormName") = FormName
    jsonDictionary("Description") = Description
    jsonDictionary("AssessmentYear") = AssessmentYear
    jsonDictionary("SchemaVer") = SchemaVer
    jsonDictionary("FormVer") = FormVer
    Set Form_ITR1 = jsonDictionary
End Function
Function PartA_139_8A() As Object
On Error Resume Next
subProcCaption = "PartA_139_8A"
noOfProcessSub = 10
UpdateProgressBar
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

If Sheet201.Range("U_AadhaarCardNo").Value <> "" Then
jsonDictionary("AadhaarCardNo") = Sheet201.Range("U_AadhaarCardNo").Value
End If

'If Sheet201.Range("U_AadhaarEnrolmentId").Value <> "" Then
'jsonDictionary("AadhaarEnrolmentId") = Sheet201.Range("U_AadhaarEnrolmentId").Value
'End If

If Sheet201.Range("U_Name").Value <> "" Then
jsonDictionary("Name") = Sheet201.Range("U_Name").Value
End If

If Sheet201.Range("U_PAN").Value <> "" Then
jsonDictionary("PAN") = Sheet201.Range("U_PAN").Value
End If

UpdateProgressBar

'Changed Year from 2023 to 2024
'Chetan C M Changed Year from 2024 to 2025
jsonDictionary("AssessmentYear") = "2025"

If Sheet201.Range("U_PreviouslyFiledForThisAY").Value <> "" Then
jsonDictionary("PreviouslyFiledForThisAY") = Mid(Sheet201.Range("U_PreviouslyFiledForThisAY").Value, 1, 1)
End If

If Mid(Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Value, 1, 1) = "1" Then
    jsonDictionary("PreviouslyFiledForThisAY_139_8A") = "1"
ElseIf Mid(Sheet201.Range("U_PreviouslyFiledForThisAY_139_8A").Value, 1, 1) = "O" Then
    jsonDictionary("PreviouslyFiledForThisAY_139_8A") = "2"
End If

Dim Applicable_139_8A As Object
Set Applicable_139_8A = CreateObject("Scripting.Dictionary")

If Sheet201.Range("U_ITRForm").Value <> "" Then
Applicable_139_8A("ITRForm") = Sheet201.Range("U_ITRForm").Value
End If

UpdateProgressBar
If Sheet201.Range("U_AcknowledgementNo").Value <> "" Then
Applicable_139_8A("AcknowledgementNo") = Sheet201.Range("U_AcknowledgementNo").Value
End If

If Trim(Sheet201.Range("U_OrigRetFiledDate").Value) <> "" Then
    Applicable_139_8A("OrigRetFiledDate") = Dformat(Sheet201.Range("U_OrigRetFiledDate").Value, "")
End If

If Applicable_139_8A.count > 0 Then
    jsonDictionary.add "Applicable_139_8A", Applicable_139_8A
End If
UpdateProgressBar

If Mid(Sheet201.Range("U_LaidOutIn_139_8A").Value, 1, 1) <> "" Then
jsonDictionary("LaidOutIn_139_8A") = Mid(Sheet201.Range("U_LaidOutIn_139_8A").Value, 1, 1)
End If


jsonDictionary("ITRFormUpdatingInc") = "ITR1"

Dim row As Range
Dim UpdatingInc As Collection
Dim UpdatingIncDtls, UpdatingIncObj As Object
Set UpdatingIncDtls = CreateObject("Scripting.Dictionary")
Set UpdatingInc = New Collection
Set UpdatingIncObj = CreateObject("Scripting.Dictionary")

For Each row In Sheet201.Range("U_ReasonsForUpdatingIncome").Rows
    If row.Cells(1).Value = "" Or row.Cells(1).Value = "(Select)" Then Exit For
    Select Case row.Cells(1).Value
    
    Case "Return previously not filed"
        
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "1"
    
    Case "Income not reported correctly"
        
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "2"
    
    Case "Wrong heads of income chosen"
    
    UpdatingIncDtls("ReasonsForUpdatingIncome") = "3"
    
    Case "Reduction of carried forward loss"
    
    UpdatingIncDtls("ReasonsForUpdatingIncome") = "4"
    
    Case "Reduction of unabsorbed depreciation"
    
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "5"
    
    Case "Reduction of tax credit u/s 115JB/115JC"
        
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "6"
        
    Case "Wrong rate of tax"
            
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "7"
            
    Case "Others "
        
        UpdatingIncDtls("ReasonsForUpdatingIncome") = "OTH"
    
    End Select
    
    UpdatingInc.add UpdatingIncDtls
    Set UpdatingIncDtls = CreateObject("Scripting.Dictionary")
Next row
If UpdatingInc.count > 0 Then
    UpdatingIncObj.add "ReasonsForUpdatingIncDtls", UpdatingInc
    jsonDictionary.add "UpdatingInc", UpdatingIncObj
End If
UpdateProgressBar

If Sheet201.Range("U_UpdatedReturnDuringPeriod").Value <> "" Then
    Select Case Sheet201.Range("U_UpdatedReturnDuringPeriod").Value
    Case "Up to 12 months from the end of Relevant Assessment Year"
        jsonDictionary("UpdatedReturnDuringPeriod") = "1"
    Case "Between 12 to 24 Months from the end of Relevant Assessment Year"
        jsonDictionary("UpdatedReturnDuringPeriod") = "2"
    Case "Between 24 to 36 Months from the end of Relevant Assessment Year"  'Ankita_04/11/2025
        jsonDictionary("UpdatedReturnDuringPeriod") = "3"
    Case "Between 36 to 48 Months from the end of Relevant Assessment Year"  'Ankita_04/11/2025
        jsonDictionary("UpdatedReturnDuringPeriod") = "4"
    End Select
End If
UpdateProgressBar

Dim RetrntoRedCarriedFL As Object
Set RetrntoRedCarriedFL = CreateObject("Scripting.Dictionary")
If Sheet201.Range("U_UnabsorbedDepreciation").Value <> "" Then
RetrntoRedCarriedFL("UnabsorbedDepreciation") = Mid(Sheet201.Range("U_UnabsorbedDepreciation").Value, 1, 1)
End If
Dim UDYear As Collection
Dim UDYearObj As Object
Dim UnabsorbedDepreciationYearDtls As Object
Set UDYear = New Collection
Set UDYearObj = CreateObject("Scripting.Dictionary")
Set UnabsorbedDepreciationYearDtls = CreateObject("Scripting.Dictionary")
UpdateProgressBar
For Each row In Sheet201.Range("U_UnabsorbedDepreciationYear").Rows
    If row.Cells(1).Value = "" Or row.Cells(1).Value = "(Select)" Then Exit For
    'Ankita_UR
'    If row.Cells(1).Value = "2024-25" Then
'        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2024"
'    ElseIf row.Cells(1).Value = "2025-26" Then
'        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2025"
'    End If
    
    
        If row.Cells(1).Value = "2026-27" Then
        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2026"
'    ElseIf row.Cells(1).Value = "2025-26" Then
'        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2025"
        ElseIf row.Cells(1).Value = "2027-28" Then  'Ankita_04/11/2025
        UnabsorbedDepreciationYearDtls("UnabsorbedDepreciationYear") = "2027"  'Ankita_04/11/2025
    End If
    
    If row.Cells(1).Offset(0, 1).Value <> "" Then
'        UnabsorbedDepreciationYearDtls("ReturnFiledEffectFlg") = Mid(row.Cells(1).Offset(0, 1).Value, 1, 1)
         UnabsorbedDepreciationYearDtls("RevisedReturnFile") = Mid(row.Cells(1).Offset(0, 1).Value, 1, 1) 'Ankita_04/11/2025==========
    End If
    
'    If row.Cells(1).Offset(0, 16).Value <> "" And row.Cells(1).Offset(0, 16).Value <> "(Select)" Then
'        If row.Cells(1).Offset(0, 16).Value = "Orginal Return 139(1)/139(4)" Then
'            UnabsorbedDepreciationYearDtls("OriginalReturn") = "Y"
'        ElseIf row.Cells(1).Offset(0, 16).Value = "Revised return" Then
'            UnabsorbedDepreciationYearDtls("RevisedReturnFile") = "Y"
'        ElseIf row.Cells(1).Offset(0, 16).Value = "Updated return" Then
'            UnabsorbedDepreciationYearDtls("UpdatedReturnFile") = "Y"
'        End If
'
'    End If

'    If row.Cells(1).Offset(0, 16).Value <> "" And row.Cells(1).Offset(0, 16).Value <> "(Select)" Then
'        If row.Cells(1).Offset(0, 16).Value = "Original Return 139(1)/139(4)" Then
'            UnabsorbedDepreciationYearDtls("ReturnType") = "1"
'        ElseIf row.Cells(1).Offset(0, 16).Value = "Revised return" Then
'            UnabsorbedDepreciationYearDtls("ReturnType") = "2"
'        ElseIf row.Cells(1).Offset(0, 16).Value = "Updated return" Then
'            UnabsorbedDepreciationYearDtls("ReturnType") = "3"
'        End If
'    End If
    
        'Ankita_04/11/2025
    If row.Cells(1).Offset(0, 16).Value <> "" Then
        UnabsorbedDepreciationYearDtls("UpdatedReturnFile") = Mid(row.Cells(1).Offset(0, 16).Value, 1, 1)
    End If
    '-----------------------------------

    UDYear.add UnabsorbedDepreciationYearDtls
    Set UnabsorbedDepreciationYearDtls = CreateObject("Scripting.Dictionary")
    UpdateProgressBar
Next row

If UDYear.count > 0 Then
    UDYearObj.add "UnabsorbedDepreciationYearDtls", UDYear
    RetrntoRedCarriedFL.add "UDYear", UDYearObj
End If

If RetrntoRedCarriedFL.count > 0 Then
    jsonDictionary.add "RetrntoRedCarriedFL", RetrntoRedCarriedFL
End If

UpdateProgressBar
Set PartA_139_8A = jsonDictionary

End Function

Function PartB_ATI() As Object
subProcCaption = "PartB_ATI"
noOfProcessSub = 10
Dim jsonDictionary, HeadOfInc As Object
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set HeadOfInc = CreateObject("Scripting.Dictionary")

UpdateProgressBar

HeadOfInc("Salaries") = UVCase(Sheet202.Range("U_Salaries").Value)
HeadOfInc("IncomeFromHP") = UVCase(Sheet202.Range("U_IncomeFromHP").Value)
'HeadOfInc("IncomeFromBP") = UVCase(Sheet202.Range("U_IncomeFromBP").Value)
'HeadOfInc("IncomeFromCG") = UVCase(Sheet202.Range("U_IncomeFromCG").Value)
HeadOfInc("IncomeFromOS") = UVCase(Sheet202.Range("U_IncomeFromOS").Value)
UpdateProgressBar
HeadOfInc("Total") = UVCase(Sheet202.Range("U_Total").Value)
jsonDictionary.add "HeadOfInc", HeadOfInc

jsonDictionary("LatestTotInc") = UVCase(Sheet202.Range("U_LatestTotInc").Value)
jsonDictionary("UpdatedTotInc") = UVCase(Sheet202.Range("U_UpdatedTotInc").Value)
UpdateProgressBar
jsonDictionary("AmtPayable") = UVCase(Sheet202.Range("U_AmtPayable").Value)
jsonDictionary("AmtRefundable") = UVCase(Sheet202.Range("U_AmtRefundable").Value)
jsonDictionary("LastAmtPayable") = UVCase(Sheet202.Range("U_LastAmtPayable").Value)
jsonDictionary("Refund") = UVCase(Sheet202.Range("U_Refund").Value)
UpdateProgressBar
jsonDictionary("TotRefund") = UVCase(Sheet202.Range("U_TotRefund").Value)
jsonDictionary("FeeIncUS234F") = UVCase(Sheet202.Range("U_FeeIncUS234F").Value)
jsonDictionary("RegAssessementTAX") = UVCase(Sheet202.Range("U_RegAssessementTAX").Value)
jsonDictionary("AggrLiabilityRefund") = UVCase(Sheet202.Range("U_AggrLiabilityRefund").Value)
jsonDictionary("AggrLiabilityNoRefund") = UVCase(Sheet202.Range("U_AggrLiabilityNoRefund").Value)
jsonDictionary("AddtnlIncTax") = UVCase(Sheet202.Range("U_AddtnlIncTax").Value)
jsonDictionary("NetPayable") = UVCase(Sheet202.Range("U_NetPayable").Value)
jsonDictionary("TaxUS140B") = UVCase(Sheet202.Range("U_TaxUS140B").Value)
UpdateProgressBar
jsonDictionary("TaxDue10_11") = UVCase(Sheet202.Range("U_TaxDue10_11").Value)
UpdateProgressBar
Dim cell As Range
Dim TaxPayment1 As Collection
Dim ITTaxPayments, ScheduleIT1 As Object
Dim TaxPayment1Obj As Object
Set TaxPayment1Obj = CreateObject("Scripting.Dictionary")
Set TaxPayment1 = New Collection
Set ITTaxPayments = CreateObject("Scripting.Dictionary")
Set ScheduleIT1 = CreateObject("Scripting.Dictionary")
Dim itemp
itemp = 0
For Each cell In Sheet202.Range("U_BSRCode1").Cells
    itemp = itemp + 1
    If cell.Value = "" Or cell.Value = "(Select)" Then Exit For
    ITTaxPayments("slno") = itemp
    ITTaxPayments("BSRCode") = cell.Value
    ITTaxPayments("DateDep") = Dformat(cell.Offset(0, 1).Value, "")
    ITTaxPayments("SrlNoOfChaln") = cell.Offset(0, 2).Value
    ITTaxPayments("Amt") = cell.Offset(0, 5).Value
    TaxPayment1.add ITTaxPayments
    Set ITTaxPayments = CreateObject("Scripting.Dictionary")
    
Next cell
UpdateProgressBar
If TaxPayment1.count > 0 Then
    
    TaxPayment1Obj.add "ITTaxPayments", TaxPayment1
    ScheduleIT1.add "TaxPayment1", TaxPayment1Obj
    ScheduleIT1("Total") = Sheet202.Range("U_Total1").Value
    jsonDictionary.add "ScheduleIT1", ScheduleIT1
End If

Dim TaxPayment2Obj As Object
Set TaxPayment2Obj = CreateObject("Scripting.Dictionary")
Dim TaxPayment2 As Collection
Set TaxPayment2 = New Collection
Dim ScheduleIT2 As Object
Set ScheduleIT2 = CreateObject("Scripting.Dictionary")
itemp = 0
For Each cell In Sheet202.Range("U_BSRCode2").Cells
    If cell.Value = "" Or cell.Value = "(Select)" Then Exit For
    itemp = itemp + 1
    ITTaxPayments("slno") = itemp
    ITTaxPayments("BSRCode") = cell.Value
    ITTaxPayments("DateDep") = Dformat(cell.Offset(0, 1).Value, "")
    ITTaxPayments("SrlNoOfChaln") = cell.Offset(0, 2).Value
    ITTaxPayments("Amt") = cell.Offset(0, 5).Value
    TaxPayment2.add ITTaxPayments
    Set ITTaxPayments = CreateObject("Scripting.Dictionary")
Next cell
UpdateProgressBar
If TaxPayment2.count > 0 Then
    TaxPayment2Obj.add "ITTaxPayments", TaxPayment2
    ScheduleIT2.add "TaxPayment2", TaxPayment2Obj
    ScheduleIT2("Total") = Sheet202.Range("U_Total2").Value
    jsonDictionary.add "ScheduleIT2", ScheduleIT2
End If


jsonDictionary("ReleifUS89") = UVCase(Sheet202.Range("U_ReleifUS89").Value)


Set PartB_ATI = jsonDictionary

End Function
'XML Generation Function For Personal Information
Function PersonalInfo() As Object
subProcCaption = "Personal Info"
noOfProcessSub = 10
Dim jsonDictionary, AssesseeName, Phone, Address
Dim residenceNo, localityOrArea, cityOrTownOrDistrict, countryCodeMobile, emailAddress, roadOrStreet, dob, EmployerCategory, AadhaarEnrolmentIds, aadhaarCardNo, StateCode As Variant
Dim CountryCode As Long

'Konda updated on 19-01-2026 Schema-V0.2
Dim residenceNo1, residenceName1, localityOrArea1, cityOrTownOrDistrict1, countryCodeMobile1, emailAddress1, roadOrStreet1, StateCode1 As Variant
Dim CountryCode_1 As Long
'Dim WantUpdateAdd, AlternateAddress
'V0.8
Dim SecondaryAdd, AlternateAddress
Dim sCountry, sCountry_Sec
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set AssesseeName = CreateObject("Scripting.Dictionary")
Set Phone = CreateObject("Scripting.Dictionary")
Set Address = CreateObject("Scripting.Dictionary")

'Konda updated on 19-01-2026 schema-V0.2
'Set WantUpdateAdd = CreateObject("Scripting.Dictionary")
'V0.8
Set SecondaryAdd = CreateObject("Scripting.Dictionary")
'+=
Set AlternateAddress = CreateObject("Scripting.Dictionary")

'Konda updated on 05-02-2026-SIT-109489,SIT-109486 and SIT-109480
MobileCountryCode2 = Sheet1.Range("sheet1.MobileCountryCode1").Value
mobileNo2 = Sheet1.Range("sheet1.Mobileno1").Value
Email_1 = Sheet1.Range("sheet1.EmailAddress1").Value
'=======

If empcat = "Central Government" Then
    empcat = "CGOV"
ElseIf empcat = "State Government" Then
    empcat = "SGOV"
'Konda updared Removed enum in Schema V0.4 27-02-2026
'    'Changed as per v0.3 by Konda_12/02/2026=====
'ElseIf empcat = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958" Then
'    empcat = "SCJ"
ElseIf empcat = "Public Sector Undertaking" Then
    empcat = "PSU"
ElseIf empcat = "Others" Then
    empcat = "OTH"
'ElseIf empcat = "Pensioners" Then
ElseIf empcat = "Pensioners - Central Government" Then
    empcat = "PE"
ElseIf empcat = "Pensioners - State Government" Then
    empcat = "PESG"
ElseIf empcat = "Pensioners - Public sector undertaking" Then
    empcat = "PEPS"
ElseIf empcat = "Pensioners - Others" Then
    empcat = "PEO"
ElseIf empcat = "Not Applicable (eg. Family pension etc)" Then
    empcat = "NA"
End If

residenceNo = UCase(Flat)
roadOrStreet = UCase(Sheet1.Range("sheet1.RoadOrStreet").Value)
localityOrArea = UCase(Area)
cityOrTownOrDistrict = UCase(City)
'CountryCode = Mid(sCountry, 1, InStr(sCountry, "-") - 1)
countryCodeMobile = MobileCountryCode
emailAddress = Email
dob = Mid((Range("sheet1.DOB").Value), 7, 4) & "-" & Mid((Range("sheet1.DOB").Value), 4, 2) & "-" & Mid((Range("sheet1.DOB").Value), 1, 2)
EmployerCategory = UCase(empcat)
aadhaarCardNo = Sheet1.Range("Sheet1.Aadhaar").Value

'Konda updated on 19-01-2026 schema-V0.2

'03/04/2026
'residenceNo1 = UCase(Flat1)
residenceNo1 = UCase(Sheet1.Range("sheet1.ResidenceNo1").Value)
'---------------
residenceName1 = UCase(Sheet1.Range("sheet1.ResidenceName1").Value)
roadOrStreet1 = UCase(Sheet1.Range("sheet1.RoadOrStreet1").Value)
'03/04/2026
'localityOrArea1 = UCase(Area1)
localityOrArea1 = UCase(Sheet1.Range("sheet1.LocalityOrArea1").Value)
state1 = Sheet1.Range("sheet1.StateCode2")  '03/04/2026
'cityOrTownOrDistrict1 = UCase(City1)
cityOrTownOrDistrict1 = UCase(Sheet1.Range("sheet1.CityOrTownOrDistrict1").Value)
'-------------------


'CountryCode_1 = Mid(sCountry1, 1, InStr(sCountry1, "-") - 1)
countryCodeMobile1 = MobileCountryCode2
emailAddress1 = Email_1
'--------------------------------
'CountryCode = Sheet1.Range("sheet1.Country").Value
'CountryCode_1 = Sheet1.Range("sheet1.Country1").Value

sCountry = Sheet1.Range("sheet1.Country").Value
If Mid(Sheet1.Range("sheet1.Country1").Value, 1, 2) <> "(S" Then
sCountry_Sec = Sheet1.Range("sheet1.Country1").Value
End If

    If firstName <> "" Then
        AssesseeName("FirstName") = UCase(firstName)
    End If
UpdateProgressBar
    If middleName <> "" Then
        AssesseeName("MiddleName") = UCase(middleName)
    End If

UpdateProgressBar
    If LastName <> "" Then
        AssesseeName("SurNameOrOrgName") = UCase(LastName)
    End If
jsonDictionary.add "AssesseeName", AssesseeName

UpdateProgressBar
    If PAN <> "" Then
       jsonDictionary("PAN") = UCase(PAN)
    End If

UpdateProgressBar

        Address("ResidenceNo") = residenceNo
 
 UpdateProgressBar
    If residenceName <> "" Then
        Address("ResidenceName") = UCase(residenceName)
    End If

UpdateProgressBar
    If Sheet1.Range("sheet1.RoadOrStreet").Value <> "" Then
        Address("RoadOrStreet") = roadOrStreet
    End If

UpdateProgressBar
    If Area <> "" Then
        Address("LocalityOrArea") = localityOrArea
    End If

UpdateProgressBar
    If City <> "" Then
        Address("CityOrTownOrDistrict") = cityOrTownOrDistrict
    End If

UpdateProgressBar
    If State <> "" Then
        Address("StateCode") = Mid((State), 1, 2)
    End If
 
UpdateProgressBar
    If sCountry <> "" Then
        Address("CountryCode") = Mid(sCountry, 1, InStr(sCountry, "-") - 1)
    End If
    
       
    If (Sheet1.Range("sheet1.PinCode").Value) <> "" Then
        Address("PinCode") = CDbl(Sheet1.Range("sheet1.PinCode").Value)
    End If

    If zipCode <> "" Then
        Address("ZipCode") = UCase(zipCode)
    End If
'Commented by Konda on 05-02-2026-SIT-109489,SIT-109486 and SIT-109480
'    If (Len(mobileNo) > 0) Then
'        Address("CountryCodeMobile") = countryCodeMobile
'        Address("MobileNo") = mobileNo
'    End If
' 'Konda updated on 19-01-2026 schema-V0.2.4
'    If (Len(mobileNo2) > 0) Then
'        Address("CountryCodeMobileNoSec") = countryCodeMobile1
'        Address("MobileNoSec") = mobileNo2
'    End If
'
'
'    If Email <> "" Then
'        Address("EmailAddress") = emailAddress
'    End If
' 'Konda updated on 19-01-2026 schema-V0.2.4
'    If Email_1 <> "" Then
'        Address("EmailAddressSec") = emailAddress1
'    End If
'SIT-109734
    If (Len(mobileNo) > 0) Then
        Address("CountryCodeMobile") = CDbl(countryCodeMobile)
         Address("MobileNo") = CDbl(mobileNo)
    End If
 
    If (Len(mobileNo2) > 0) Then
        Address("CountryCodeMobileNoSec") = CDbl(MobileCountryCode2)
         Address("MobileNoSec") = CDbl(mobileNo2)
    End If

    If Email <> "" Then
        Address("EmailAddress") = emailAddress
    End If
 
    If Email_1 <> "" Then
        Address("EmailAddressSec") = Email_1
    End If
  '============================
    jsonDictionary.add "Address", Address
    
'
' 'Konda updated on 17-03-2026--V0.6
''  'Konda updated on 19-01-2026 schema-V0.2
'    If Sheet1.Range("Secondary_Address") <> "" Then
'        If Sheet1.Range("Secondary_Address").Value = "Yes" Then
'            jsonDictionary("WantUpdateAdd") = "Y"
'       Else
'            jsonDictionary("WantUpdateAdd") = "N"
'       End If
'    End If
'
'    If Sheet1.Range("Secondary_Address").Value = "Yes" Then

'=====================================
If Sheet1.Range("Secondary_Address").Value <> "" Then
        If Sheet1.Range("Secondary_Address").Value = "Yes" Then
            jsonDictionary("SecondaryAdd") = "Y"
       Else
            jsonDictionary("SecondaryAdd") = "N"
       End If
    End If
    '===
    If Sheet1.Range("Secondary_Address").Value = "Yes" Then
    '====+ V0.8
'If residenceNo1 <> "" Then
'+=
        If residenceNo1 <> "" Then
            AlternateAddress("ResidenceNo") = residenceNo1
            Else
            AlternateAddress("ResidenceNo") = ""
        End If
        
        
        
       
        
         If residenceName1 <> "" Then
            AlternateAddress("ResidenceName") = residenceName1
        End If
        
        If roadOrStreet1 <> "" Then
            AlternateAddress("RoadOrStreet") = roadOrStreet1
        End If
        
        If localityOrArea1 <> "" Then
            AlternateAddress("LocalityOrArea") = localityOrArea1
            Else
             AlternateAddress("LocalityOrArea") = ""
        End If
        
        If cityOrTownOrDistrict1 <> "" Then
            AlternateAddress("CityOrTownOrDistrict") = cityOrTownOrDistrict1
            Else
            AlternateAddress("CityOrTownOrDistrict") = ""
        End If
        
        
        If state1 <> "" Then
           AlternateAddress("StateCode") = Mid((state1), 1, 2)
           Else
           AlternateAddress("StateCode") = ""
        End If
        
        
    If sCountry_Sec <> "" Then
        AlternateAddress("CountryCode") = Mid(sCountry_Sec, 1, InStr(sCountry_Sec, "-") - 1)
    End If
        

        If Sheet1.Range("sheet1.PinCode1").Value <> "" Then
            AlternateAddress("PinCode") = CDbl(Sheet1.Range("sheet1.PinCode1").Value)
        End If
        
 'Commented By konda on 02/02/2026-SIT-109363
'        If Sheet1.Range("sheet1.ZipCode1").Value <> "" Then
'            AlternateAddress("ZipCode") = Int(Sheet1.Range("sheet1.ZipCode1").Value)
'        End If
        
        'Konda_02/02/2026-SIT-109363
        If Sheet1.Range("sheet1.ZipCode1").Value <> "" Then
            AlternateAddress("ZipCode") = UCase(Sheet1.Range("sheet1.ZipCode1").Value)
        End If
        
   
    jsonDictionary.add "AlternateAddress", AlternateAddress
    'End If         'Konda updated on 17-03-2026--V0.6 V0.8
    '+====
    ElseIf Sheet1.Range("Secondary_Address").Value = "No" Then
    '====+ V0.8
'If residenceNo1 <> "" Then
'+=
        If residenceNo1 <> "" Then
            AlternateAddress("ResidenceNo") = residenceNo1
            Else
            AlternateAddress("ResidenceNo") = ""
        End If
        
         If residenceName1 <> "" Then
            AlternateAddress("ResidenceName") = residenceName1
        End If
        
        If roadOrStreet1 <> "" Then
            AlternateAddress("RoadOrStreet") = roadOrStreet1
        End If
        
        If localityOrArea1 <> "" Then
            AlternateAddress("LocalityOrArea") = localityOrArea1
            Else
            AlternateAddress("LocalityOrArea") = ""
        End If
        
        If cityOrTownOrDistrict1 <> "" Then
            AlternateAddress("CityOrTownOrDistrict") = cityOrTownOrDistrict1
            Else
            AlternateAddress("CityOrTownOrDistrict") = ""
        End If
        
        If state1 <> "" Then

            AlternateAddress("StateCode") = Mid((state1), 1, 2)
            Else
            AlternateAddress("StateCode") = ""
        End If
        
    If sCountry_Sec <> "" Then
        AlternateAddress("CountryCode") = Mid(sCountry_Sec, 1, InStr(sCountry_Sec, "-") - 1)
    End If
        

        If Sheet1.Range("sheet1.PinCode1").Value <> "" Then
            AlternateAddress("PinCode") = CDbl(Sheet1.Range("sheet1.PinCode1").Value)
        End If
        
        If Sheet1.Range("sheet1.ZipCode1").Value <> "" Then
            AlternateAddress("ZipCode") = UCase(Sheet1.Range("sheet1.ZipCode1").Value)
        End If
        
   
    jsonDictionary.add "AlternateAddress", AlternateAddress
    End If         'Konda updated on 17-03-2026--V0.6 V0.8
   ' +===
UpdateProgressBar
    
    
    
    If dob <> "" Then
        jsonDictionary("DOB") = dob
    End If
       
    If empcat <> "" Then
        jsonDictionary("EmployerCategory") = EmployerCategory
    End If


    If Sheet1.Range("Sheet1.Aadhaar").Value <> "" Then
        jsonDictionary("AadhaarCardNo") = aadhaarCardNo
    End If
'    Enhancement
'    If Sheet1.Range("Sheet1.AadhaarEnrol").Value <> "" Then
'        jsonDictionary("AadhaarEnrolmentId") = Sheet1.Range("Sheet1.AadhaarEnrol").Value
'    End If

Set PersonalInfo = jsonDictionary
End Function
'XML Format For Header
Function FilingStatus() As Object
subProcCaption = "Filling Status"
noOfProcessSub = 6
Dim jsonDictionary
Dim SeventhProvisio139, NewTaxRegime, DepAmtAggAmtExcd1CrPrYrFlg, AmtSeventhProvisio139i, IncrExpAggAmt2LkTrvFrgnCntryFlg, AmtSeventhProvisio139ii, IncrExpAggAmt1LkElctrctyPrYrFlg, AmtSeventhProvisio139iii, ReceiptNo, NoticeNo, OrigRetFiledDate, NoticeDateUnderSec, ItrFilingDueDate, OptOutNewTaxRegime As Variant
Dim clauseiv7provisio139i
Dim clauseiv7provisio139iDtls, clauseiv7provisio139iType As Variant

Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
Set clauseiv7provisio139iDtls = New Collection

Dim AssesseeRep
Set AssesseeRep = CreateObject("Scripting.Dictionary")

Set jsonDictionary = CreateObject("Scripting.Dictionary")

UpdateProgressBar

If ProvisoFlag = "Yes" Then
        ProvisoFlag = "Y"
    ElseIf ProvisoFlag = "No" Then
        ProvisoFlag = "N"
End If
    
If DepositAmountFlag = "Yes" Then
        DepositAmountFlag = "Y"
    ElseIf DepositAmountFlag = "No" Then
        DepositAmountFlag = "N"
End If

If AggrigateAmountFlag = "Yes" Then
        AggrigateAmountFlag = "Y"
    ElseIf AggrigateAmountFlag = "No" Then
        AggrigateAmountFlag = "N"
End If

If AggrigateAmountFlag1 = "Yes" Then
        AggrigateAmountFlag1 = "Y"
    ElseIf AggrigateAmountFlag1 = "No" Then
        AggrigateAmountFlag1 = "N"
End If
 
SeventhProvisio139 = UCase(ProvisoFlag)
NewTaxRegime = Range("BacValue").Value
DepAmtAggAmtExcd1CrPrYrFlg = UCase(DepositAmountFlag)
AmtSeventhProvisio139i = DepositAmount
IncrExpAggAmt2LkTrvFrgnCntryFlg = UCase(AggrigateAmountFlag)
AmtSeventhProvisio139ii = AggrigateAmount
IncrExpAggAmt1LkElctrctyPrYrFlg = UCase(AggrigateAmountFlag1)
AmtSeventhProvisio139iii = AggrigateAmount1
ReceiptNo = Sheet1.Range("sheet1.ReceiptNo").Value
NoticeNo = Sheet1.Range("sheet1.NoticeNo").Value

'Konda Updated on 13-11-2024 SIT-83156
'OrigRetFiledDate = Mid(Sheet1.Range("sheet1.OrigRetFiledDate").Value, 7, 4) & "-" & Mid((Sheet1.Range("sheet1.OrigRetFiledDate").Value), 4, 2) & "-" & Mid((Sheet1.Range("sheet1.OrigRetFiledDate").Value), 1, 2)
If Sheet1.Range("sheet1.OrigRetFiledDate").Value <> "" Then
OrigRetFiledDate = Dformat(Sheet1.Range("sheet1.OrigRetFiledDate").Value, "")
End If
'End changeson 13-11-2024 SIT-83156

NoticeDateUnderSec = Mid((Sheet1.Range("sheet1.NoticeDate").Value), 7, 4) & "-" & Mid((Sheet1.Range("sheet1.NoticeDate").Value), 4, 2) & "-" & Mid((Sheet1.Range("sheet1.NoticeDate").Value), 1, 2)
'PAG_C9 2024-25 Bindu
ItrFilingDueDate = Mid((Sheet1.Range("Due_DateITR1").Value), 7, 4) & "-" & Mid((Sheet1.Range("Due_DateITR1").Value), 4, 2) & "-" & Mid((Sheet1.Range("Due_DateITR1").Value), 1, 2)

    
    jsonDictionary("ReturnFileSec") = UVCase(CDbl(ReturnFileSec))
UpdateProgressBar
    
    If SeventhProvisio139 <> "" Then
        jsonDictionary("SeventhProvisio139") = SeventhProvisio139
    End If
        
    If Not IsEmpty(Range("BacValue").Value) Then
        
        
        'PAG_E4
'        If NewTaxRegime = 1 Then
'            NewTaxRegime = "Y"
'        Else
'            NewTaxRegime = "N"
'        End If
        'jsonDictionary("NewTaxRegime") = UCase(NewTaxRegime)
        'PAG_C4 2024-25 Bindu
        If NewTaxRegime = 2 Then
            NewTaxRegime = "Y"
        Else
            NewTaxRegime = "N"
        End If
        
        jsonDictionary("OptOutNewTaxRegime") = UCase(NewTaxRegime)
        
    End If
    If DepAmtAggAmtExcd1CrPrYrFlg <> "" Then
        jsonDictionary("DepAmtAggAmtExcd1CrPrYrFlg") = DepAmtAggAmtExcd1CrPrYrFlg
    End If
    
    jsonDictionary("AmtSeventhProvisio139i") = AmtSeventhProvisio139i

    If IncrExpAggAmt2LkTrvFrgnCntryFlg <> "" Then
        jsonDictionary("IncrExpAggAmt2LkTrvFrgnCntryFlg") = IncrExpAggAmt2LkTrvFrgnCntryFlg
    End If
    
    jsonDictionary("AmtSeventhProvisio139ii") = AmtSeventhProvisio139ii
    
    If IncrExpAggAmt1LkElctrctyPrYrFlg <> "" Then
        jsonDictionary("IncrExpAggAmt1LkElctrctyPrYrFlg") = IncrExpAggAmt1LkElctrctyPrYrFlg
    End If
    
    jsonDictionary("AmtSeventhProvisio139iii") = AmtSeventhProvisio139iii
    
    If UCase(ReturnFileSec) = 13 Or UCase(ReturnFileSec) = 14 Or UCase(ReturnFileSec) = 15 Or UCase(ReturnFileSec) = 16 Or UCase(ReturnFileSec) = 20 Then
        jsonDictionary("NoticeNo") = NoticeNo
        jsonDictionary("NoticeDateUnderSec") = NoticeDateUnderSec
    ElseIf UCase(ReturnFileSec) = 17 Then
        jsonDictionary("ReceiptNo") = ReceiptNo
'        If Trim(Sheet201.Range("U_OrigRetFiledDate").Value) <> "" Then
            jsonDictionary("OrigRetFiledDate") = OrigRetFiledDate
'        End If
    ElseIf UCase(ReturnFileSec) = 18 Then
        jsonDictionary("ReceiptNo") = ReceiptNo
        jsonDictionary("NoticeNo") = NoticeNo
'        If Trim(Sheet201.Range("U_OrigRetFiledDate").Value) <> "" Then
            jsonDictionary("OrigRetFiledDate") = OrigRetFiledDate
'        End If
        jsonDictionary("NoticeDateUnderSec") = NoticeDateUnderSec
    ElseIf UCase(ReturnFileSec) = 21 Then
        jsonDictionary("ReceiptNo") = ReceiptNo
'        If Trim(Sheet201.Range("U_OrigRetFiledDate").Value) <> "" Then
            jsonDictionary("OrigRetFiledDate") = OrigRetFiledDate
'        End If
    End If
    
    'PAG_C9 2024-25 Bindu
    jsonDictionary("ItrFilingDueDate") = ItrFilingDueDate
    
    
    '139 clause iv added
    If Sheet1.Range("clauseiv7provisio139iFlg") = "Yes" Then
        jsonDictionary("clauseiv7provisio139i") = "Y"
'Change-22.11.2022.102.19G
'       If Sheet1.Range("clauseiv7provisio139iFlg_1") = "Yes" Then
'         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "1"
'         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_1")
'         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
'         Set clauseiv7provisio139iType = Nothing
'         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
'       End If
'       If Sheet1.Range("clauseiv7provisio139iFlg_2") = "Yes" Then
'         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "2"
'         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_2")
'         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
'         Set clauseiv7provisio139iType = Nothing
'         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
'       End If
'---end

'       If Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes" Then
'         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "3"
'         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_3")
'         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
'         Set clauseiv7provisio139iType = Nothing
'         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
'       End If
'       If Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes" Then
'         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "4"
'         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_4")
'         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
'         Set clauseiv7provisio139iType = Nothing
'         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
'       End If
       
       If Sheet1.Range("clauseiv7provisio139iFlg_3") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "1"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_3")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       If Sheet1.Range("clauseiv7provisio139iFlg_4") = "Yes" Then
         clauseiv7provisio139iType("clauseiv7provisio139iNature") = "2"
         clauseiv7provisio139iType("clauseiv7provisio139iAmount") = Sheet1.Range("clauseiv7provisio139iAmount_4")
         clauseiv7provisio139iDtls.add clauseiv7provisio139iType
         Set clauseiv7provisio139iType = Nothing
         Set clauseiv7provisio139iType = CreateObject("Scripting.Dictionary")
       End If
       jsonDictionary.add "clauseiv7provisio139iDtls", clauseiv7provisio139iDtls
    Else
        jsonDictionary("clauseiv7provisio139i") = "N"
    End If
    
  'Konda updated on 19-01-2026 schema-V0.2
  
    If Sheet1.Range("sheet1.RepAssessee") = "Yes" Then
        jsonDictionary("AsseseeRepFlg") = "Y"
        If Sheet1.Range("sheet1.RepAssessee") = "Yes" Then
            AssesseeRep("RepName") = Sheet1.Range("sheet1.NameRepAssessee").Value
             AssesseeRep("RepEmailID") = Sheet1.Range("sheet1.EmailRepAssessee").Value
             AssesseeRep("CountryCodeRepMobileNo") = Sheet1.Range("sheet1.CountryCodeRepAssessee").Value
             AssesseeRep("RepMobileNo") = CDbl(Sheet1.Range("sheet1.ContactRepAssessee").Value)

        End If
        
        jsonDictionary.add "AssesseeRep", AssesseeRep
        
        Set AssesseeRep = Nothing
         Set AssesseeRep = CreateObject("Scripting.Dictionary")
        Else
        jsonDictionary("AsseseeRepFlg") = "N"
    End If

Set FilingStatus = jsonDictionary

End Function

'XML Generation Function For Income Deduction

Function IncomeDeductions() As Object
subProcCaption = "Income Deductions"
noOfProcessSub = 16
Dim i As Long
Dim jsonDictionary, UsrDeductUndChapVIA, DeductUndChapVIA, AllwncExemptUs10, AllwncExemptUs10Dtl, DividendInc, NOT89AInc, OthersInc, DateRange, ExemptIncAgriOthUs10, OthersIncDtlsOth, ExemptIncAgriOthUs10Dtl
Dim IncomeFromSal, TotalIncomeOfHP, IncomeOthSrc, Section80C, Section80CCC, Section80CCDEmployeeOrSE, Section80CCD1B, Section80CCDEmployer, Section80D, Section80DD, Section80DDB As Variant
Dim Section80E, Section80EE, Section80G, Section80GG, Section80GGA, Section80GGC, Section80U, Section80TTA, Section80TTB, TotalChapVIADeductions, Section80DDUsrType, Section80DDBUsrType, Section80EEA, Section80EEB, Section80UUsrType As Variant
Dim TotalIncome, GrossSalary, Salary, PerquisitesValue, ProfitsInSalary, IncomeNotified89A, IncomeNotifiedOther89A, AllwncExemptUs10Dtls, TotalAllwncExemptUs10, NetSalary, DeductionUs16, DeductionUs16ia, EntertainmentAlw16ii, ProfessionalTaxUs16iii, GrossRentReceived, TaxPaidlocalAuth As Variant
Dim AnnualValue, StandardDeduction, InterestPayable, ArrearsUnrealizedRentRcvd, OthersIncDtlsOthSrc, Upto15Of6, Upto15Of9, Up16Of9To15Of12, Up16Of12To15Of3, Up16Of3To31Of3, TotDividendInc, DeductionUs57iia, ExemptIncAgriOthUs10Dtls, ExemptIncAgriOthUs10Total As Variant
Dim Upto15Of6_89A, Upto15Of9_89A, Up16Of9To15Of12_89A, Up16Of12To15Of3_89A, Up16Of3To31Of3_89A, OSIncomeNotified89A As Variant
Dim Instr, test As String
Dim IncomeNotified89AType, NOT89AType, NOT89A As Variant


'New Schema updated by Konda as AY-2026-27 on 26-12-2025
Dim PensionContribution80CCC, PensionContribution80CCD1, PensionContribution80CCD1B, PensionContributionFund

Set PensionContribution80CCC = New Collection
Set PensionContribution80CCD1 = New Collection
Set PensionContribution80CCD1B = New Collection
Set PensionContributionFund = CreateObject("Scripting.Dictionary")

Set IncomeNotified89AType = New Collection
'Set NOT89A = New Collection
'Set NOT89AType = CreateObject("Scripting.Dictionary")
Set AllwncExemptUs10Dtls = New Collection
Set OthersIncDtlsOthSrc = New Collection
Set ExemptIncAgriOthUs10Dtls = New Collection

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set UsrDeductUndChapVIA = CreateObject("Scripting.Dictionary")
Set DeductUndChapVIA = CreateObject("Scripting.Dictionary")
Set AllwncExemptUs10 = CreateObject("Scripting.Dictionary")
Set AllwncExemptUs10Dtl = CreateObject("Scripting.Dictionary")
Set DividendInc = CreateObject("Scripting.Dictionary")
Set NOT89AInc = CreateObject("Scripting.Dictionary")
Set OthersInc = CreateObject("Scripting.Dictionary")
Set DateRange = CreateObject("Scripting.Dictionary")
Set ExemptIncAgriOthUs10 = CreateObject("Scripting.Dictionary")
Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
Set ExemptIncAgriOthUs10Dtl = CreateObject("Scripting.Dictionary")

Upto15Of6 = Sheet1.Range("IncD_q1div").Value
Upto15Of9 = Sheet1.Range("IncD_q2div").Value
Up16Of9To15Of12 = Sheet1.Range("IncD_q3div").Value
Up16Of12To15Of3 = Sheet1.Range("IncD_q4div").Value
Up16Of3To31Of3 = Sheet1.Range("IncD_q5div").Value
TotDividendInc = Sheet1.Range("IncD.Div").Value


'Commented by Ankita_28/01/2026===============
'Upto15Of6_89A = UVCase(Sheet1.Range("IncD_q1OS1").Value)
'Upto15Of9_89A = UVCase(Sheet1.Range("IncD_q2OS1").Value)
'Up16Of9To15Of12_89A = UVCase(Sheet1.Range("IncD_q3OS1").Value)
'Up16Of12To15Of3_89A = UVCase(Sheet1.Range("IncD_q4OS1").Value)
'Up16Of3To31Of3_89A = UVCase(Sheet1.Range("IncD_q5OS1").Value)

'Konda updated on 22=01-2026
'OSIncomeNotified89A = Sheet1.Range("OSIncomeNotified89A").Value

GrossSalary = Sheet1.Range("IncD.IncomeFromSal").Value
Salary = Sheet1.Range("IncD.Allowances").Value
PerquisitesValue = Sheet1.Range("IncD.Perquisites").Value
ProfitsInSalary = Sheet1.Range("IncD.Profits").Value
'IncomeNotified89A = Sheet1.Range("IncomeNotified89A").Value
IncomeNotifiedOther89A = Sheet1.Range("IncomeNotifiedOther89A").Value
TotalAllwncExemptUs10 = Sheet1.Range("Less_allowance").Value
NetSalary = Sheet1.Range("Net_salary").Value
DeductionUs16 = Sheet1.Range("Deductions_16").Value
DeductionUs16ia = Sheet1.Range("IncD.Deduction16ia").Value
EntertainmentAlw16ii = Sheet1.Range("IncD.Deduction16").Value
ProfessionalTaxUs16iii = Sheet1.Range("IncD.Deduction16ic").Value
IncomeFromSal = Sheet1.Range("IncD.TotalHeadSalaries").Value

'GrossRentReceived = Sheet1.Range("IncD.GrossRentRecieved").Value  'Konda_29/01/2026======
'TaxPaidlocalAuth = Sheet1.Range("IncD.TaxPaidLocalAuthorities").Value
'AnnualValue = Sheet1.Range("IncD.AnnualValue").Value
'StandardDeduction = Sheet1.Range("IncD.StandardDeduction").Value  'Konda_29/01/2026======
'InterestPayable = Sheet1.Range("IncD.InterestBorrowedCapital").Value
'ArrearsUnrealizedRentRcvd = Sheet1.Range("IncD.Arrears").Value

TotalIncomeOfHP = Sheet1.Range("IncD.IncomeHeadHouseProperty").Value
IncomeOthSrc = IncomeFromOS
DeductionUs57iia = Sheet1.Range("IncD.LessDeduction57").Value
GrossTotIncome = GrossTotIncome
GrossTotIncomeIncLTCG112A = GrossTotIncomeIncLTCG112A 'Konda added V0.3.2 AY2025-26 on 07-02-2025
Section80C = Sheet1.Range("IncD.Section80C").Value
Section80CCC = Sheet1.Range("IncD.Section80CCC").Value
Section80CCDEmployeeOrSE = Sheet1.Range("IncD.Section80CCD_SE").Value
Section80CCD1B = Sheet1.Range("IncD.Section80CCD1B_SE").Value
Section80CCDEmployer = Sheet1.Range("IncD.Section80CCD").Value
Section80D = Sheet1.Range("IncD.Section80DValue").Value
Section80DD = Sheet1.Range("IncD.Section80DD").Value
Section80DDB = Sheet1.Range("IncD.Section80DDB").Value
Section80E = Sheet1.Range("IncD.Section80E").Value
Section80EE = Sheet1.Range("IncD.Section80EE").Value
Section80G = Sheet1.Range("IncD.Section80G").Value
Section80GG = Sheet1.Range("IncD.Section80GG").Value
Section80GGA = Sheet1.Range("IncD.Section80GGA").Value
Section80GGC = Sheet1.Range("IncD.Section80GGC").Value
Section80U = Sheet1.Range("IncD.Section80U").Value
Section80TTA = Sheet1.Range("IncD.Section80TTA").Value
Section80TTB = Sheet1.Range("IncD.Section80TTB").Value
TotalChapVIADeductions = Sheet1.Range("IncD.TotalChapVIADeductions_Input").Value
'Section80DDUsrType = SELECT80DD
'Section80DDBUsrType = SELECT80DDB
Section80EEA = Sheet1.Range("IncD.Section80EEA").Value
Section80EEB = Sheet1.Range("IncD.Section80EEB").Value
'Section80UUsrType = SELECT80U
'TotalIncome = Sheet1.Range("IncD.TotalIncome").Value
'Newly updated by Bindu as per DE V3 on 4th Feb 2025
TotalIncome = Sheet1.Range("IncD.TotalIncome_New").Value

ExemptIncAgriOthUs10Total = Sheet1.Range("ExemptIncomeTotal").Value

    If Trim(Sheet1.Range("IncD.IncomeFromSal").Value) <> "" Then
        jsonDictionary("GrossSalary") = GrossSalary
     Else
        jsonDictionary("GrossSalary") = 0
    End If
    
    If Trim(Sheet1.Range("IncD.Allowances").Value) <> "" Then
        jsonDictionary("Salary") = Salary
     Else
        jsonDictionary("Salary") = 0
    End If
    
    If Trim(Sheet1.Range("IncD.Perquisites").Value) <> "" Then
        jsonDictionary("PerquisitesValue") = PerquisitesValue
     Else
        jsonDictionary("PerquisitesValue") = 0
    End If
                 
    If Trim(Sheet1.Range("IncD.Profits").Value) <> "" Then
        jsonDictionary("ProfitsInSalary") = ProfitsInSalary
     Else
        jsonDictionary("ProfitsInSalary") = 0
    End If
 'commented by Konda as per New Schema V0.2.2 on 20-01-2026
'    If Trim(Sheet1.Range("Increliefus89A").Value) <> "" Then
'        jsonDictionary("Increliefus89A") = Sheet1.Range("Increliefus89A").Value
'     Else
'        jsonDictionary("Increliefus89A") = 0
'    End If
    
'    If Trim(Sheet1.Range("OSIncreliefus89A").Value) <> "" Then
'        jsonDictionary("Increliefus89AOS") = Sheet1.Range("OSIncreliefus89A").Value
'     Else
'        jsonDictionary("Increliefus89AOS") = 0
'    End If
    
'commented by Konda as per New Schema V0.2.2 on 20-01-2026
'    If Trim(Sheet1.Range("IncomeNotified89A_AmountUS").Value) <> "" Then
'        NOT89AType("NOT89ACountrycode") = "US"
'        NOT89AType("NOT89AAmount") = Sheet1.Range("IncomeNotified89A_AmountUS").Value
'        IncomeNotified89AType.add NOT89AType
'        Set NOT89AType = Nothing
'        Set NOT89AType = CreateObject("Scripting.Dictionary")
'    End If
'    If Trim(Sheet1.Range("IncomeNotified89A_AmountUK").Value) <> "" Then
'        NOT89AType("NOT89ACountrycode") = "UK"
'        NOT89AType("NOT89AAmount") = Sheet1.Range("IncomeNotified89A_AmountUK").Value
'        IncomeNotified89AType.add NOT89AType
'        Set NOT89AType = Nothing
'        Set NOT89AType = CreateObject("Scripting.Dictionary")
'    End If
'    If Trim(Sheet1.Range("IncomeNotified89A_AmountCan").Value) <> "" Then
'        NOT89AType("NOT89ACountrycode") = "CA"
'        NOT89AType("NOT89AAmount") = Sheet1.Range("IncomeNotified89A_AmountCan").Value
'        IncomeNotified89AType.add NOT89AType
'        Set NOT89AType = Nothing
'        Set NOT89AType = CreateObject("Scripting.Dictionary")
'    End If
'    jsonDictionary.add "IncomeNotified89AType", IncomeNotified89AType
'    If Trim(Sheet1.Range("IncomeNotified89A").Value) <> "" Then
'        jsonDictionary("IncomeNotified89A") = IncomeNotified89A
'     Else
'        jsonDictionary("IncomeNotified89A") = 0
'    End If
'    If Trim(Sheet1.Range("IncomeNotifiedOther89A").Value) <> "" Then
'        jsonDictionary("IncomeNotifiedOther89A") = IncomeNotifiedOther89A
'     Else
'        jsonDictionary("IncomeNotifiedOther89A") = 0
'    End If
'============================================Schema V0.2.2 on 20-01-2026


    If (Not IsEmpty(Others_NOI4)) Or Sheet1.Range("Sheet1.HRA") > 0 Then
        If (UBound(Others_NOI4) > 0) Or Sheet1.Range("Sheet1.HRA") > 0 Then
'    If (Not IsEmpty(Others_NOI4)) Then
'        If (UBound(Others_NOI4) > 0) Then
            For i = 1 To UBound(Others_NOI4)
                If (Others_NOI4(i) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette") Then
                    Others_NOI4(i) = "10(10B)(i)"
                ElseIf (Others_NOI4(i) = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government") Then
                    Others_NOI4(i) = "10(10B)(ii)"
                ElseIf (Others_NOI4(i) = "Section 10(14)(i) - Allowances referred in sub-clauses (a) to (c) of sub-rule (1) in Rule 2BB") Then
                    Others_NOI4(i) = "10(14)(i)(115BAC)"
                ElseIf (Others_NOI4(i) = "Section 10(14)(ii) -  Transport allowance granted to certain physically handicapped assessee") Then
                    Others_NOI4(i) = "10(14)(ii)(115BAC)"
   'Konda updaed on 10-03-2026--V0.5
'                ElseIf (Others_NOI4(i) = "Any Other") Then
'                    Others_NOI4(i) = "OTH"
                ElseIf (Others_NOI4(i) = "Sec 10(17)-Allowance MP/MLA/MLC ") Then 'Ankita
                    Others_NOI4(i) = "10(17)"
    '================================
                ElseIf (Others_NOI4(i) = "Exempt income received by a judge covered under the payment of salaries to Supreme Court/High Court judges Act /Rules") Then
                    Others_NOI4(i) = "EIC"
                
                'Malli---
                 ElseIf (Others_NOI4(i) = "10(10B) First proviso - Compensation limit notified by CG in the Official Gazette") Then
                    Others_NOI4(i) = "10(10B)(i)"
                    
                     ElseIf (Others_NOI4(i) = "10(10B) Second proviso - Compensation under scheme approved by the Central Government") Then
                    Others_NOI4(i) = "10(10B)(ii)"
                
                '------
                
                
                Else
                    If InStr(Others_NOI4(i), "-") = 0 Then
                        GoTo NextIteration
                    Else
                       test = Mid(Others_NOI4(i), 5, InStr(Others_NOI4(i), "-") - 5)
                        Others_NOI4(i) = test
                    End If
                End If
NextIteration:
            Next
            
            'Newly added by Bindu as per DE V7
                If Sheet1.Range("Sheet1.HRA") > 0 Then
                'Malli_09/05/2025
                    AllwncExemptUs10Dtl("SalNatureDesc") = "10(13A)"
                    AllwncExemptUs10Dtl("SalOthAmount") = Sheet1.Range("Sheet1.HRA")
                    AllwncExemptUs10Dtls.add AllwncExemptUs10Dtl
                Set AllwncExemptUs10Dtl = Nothing
                Set AllwncExemptUs10Dtl = CreateObject("Scripting.Dictionary")
                
                End If
'                AllwncExemptUs10Dtls.add AllwncExemptUs10Dtl
'                Set AllwncExemptUs10Dtl = Nothing
'                Set AllwncExemptUs10Dtl = CreateObject("Scripting.Dictionary")
'
            For i = 1 To UBound(Others_NOI4)
                If Others_NOI4(i) <> "" Then
                    AllwncExemptUs10Dtl("SalNatureDesc") = Others_NOI4(i)
                End If

 'Konda updaed on 10-03-2026--V0.5
'                If Others_NOI5(i) <> "" And Others_NOI5(i) <> "Not Applicable" Then
'                    AllwncExemptUs10Dtl("SalOthNatOfInc") = UCase(Others_NOI5(i))
'                End If
'==========================
                If Others_Amt3(i) <> "" Then
                    AllwncExemptUs10Dtl("SalOthAmount") = Others_Amt3(i)
                End If
                AllwncExemptUs10Dtls.add AllwncExemptUs10Dtl
                Set AllwncExemptUs10Dtl = Nothing
                Set AllwncExemptUs10Dtl = CreateObject("Scripting.Dictionary")
            Next
        If Trim(Sheet1.Range("Less_allowance").Value) <> "" Then
            AllwncExemptUs10("TotalAllwncExemptUs10") = TotalAllwncExemptUs10
        Else
            AllwncExemptUs10("TotalAllwncExemptUs10") = 0
        End If
    AllwncExemptUs10.add "AllwncExemptUs10Dtls", AllwncExemptUs10Dtls
    jsonDictionary.add "AllwncExemptUs10", AllwncExemptUs10
    End If
    End If
    
    If Trim(Sheet1.Range("Net_salary").Value) <> "" Then
        jsonDictionary("NetSalary") = NetSalary
     Else
        jsonDictionary("NetSalary") = 0
    End If
    
    If Trim(Sheet1.Range("Deductions_16").Value) <> "" Then
        jsonDictionary("DeductionUs16") = DeductionUs16
     Else
        jsonDictionary("DeductionUs16") = 0
    End If
    
    If Trim(Sheet1.Range("IncD.Deduction16ia").Value) <> "" Then
        jsonDictionary("DeductionUs16ia") = DeductionUs16ia
     Else
        jsonDictionary("DeductionUs16ia") = 0
    End If
    
    If Trim(Sheet1.Range("IncD.Deduction16").Value) <> "" Then
        jsonDictionary("EntertainmentAlw16ii") = EntertainmentAlw16ii
     Else
        jsonDictionary("EntertainmentAlw16ii") = 0
    End If
    
    If Trim(Sheet1.Range("IncD.Deduction16ic").Value) <> "" Then
        jsonDictionary("ProfessionalTaxUs16iii") = ProfessionalTaxUs16iii
     Else
        jsonDictionary("ProfessionalTaxUs16iii") = 0
    End If
    
    If Trim(Sheet1.Range("IncD.TotalHeadSalaries").Value) <> "" Then
        jsonDictionary("IncomeFromSal") = IncomeFromSal
     Else
        jsonDictionary("IncomeFromSal") = 0
    End If
    
UpdateProgressBar

'commented by Konda as per New Schema V0.2.2 on 20-01-2026

'    If TypeOfHP <> "" And Mid(UCase(TypeOfHP), 1, 1) <> "(" Then
'        jsonDictionary("TypeOfHP") = Mid(UCase(TypeOfHP), 1, 1)
'    End If
'
'    If Trim(Sheet1.Range("IncD.GrossRentRecieved").Value) <> "" Then
'        jsonDictionary("GrossRentReceived") = GrossRentReceived
'    Else
'        jsonDictionary("GrossRentReceived") = 0
'    End If
'
'    If Trim(Sheet1.Range("IncD.TaxPaidLocalAuthorities").Value) <> "" Then
'        jsonDictionary("TaxPaidlocalAuth") = TaxPaidlocalAuth
'    Else
'        jsonDictionary("TaxPaidlocalAuth") = 0
'    End If
'
'    If Trim(Sheet1.Range("IncD.AnnualValue").Value) <> "" Then
'        jsonDictionary("AnnualValue") = AnnualValue
'    Else
'        jsonDictionary("AnnualValue") = 0
'    End If
'
'    If Trim(Sheet1.Range("IncD.StandardDeduction").Value) <> "" Then
'        jsonDictionary("StandardDeduction") = StandardDeduction
'    Else
'        jsonDictionary("StandardDeduction") = 0
'    End If
'
'    If Trim(Sheet1.Range("IncD.InterestBorrowedCapital").Value) <> "" Then
'        jsonDictionary("InterestPayable") = InterestPayable
'    Else
'        jsonDictionary("InterestPayable") = 0
'    End If
'
'    If Trim(Sheet1.Range("IncD.Arrears").Value) <> "" Then
'        jsonDictionary("ArrearsUnrealizedRentRcvd") = ArrearsUnrealizedRentRcvd
'     Else
'        jsonDictionary("ArrearsUnrealizedRentRcvd") = 0
'    End If
'
'    If Trim(Sheet1.Range("IncD.IncomeHeadHouseProperty").Value) <> "" Then
'        jsonDictionary("TotalIncomeOfHP") = TotalIncomeOfHP
'     Else
'        jsonDictionary("TotalIncomeOfHP") = 0
'    End If
    
'    If Trim(Sheet1.Range("IncD.IncomeHeadHousePropertyINTER").Value) <> "" Then
'      jsonDictionary("TotalIncomeOfHPNegative") = Sheet1.Range("IncD.IncomeHeadHousePropertyINTER").Value
'     Else
'        jsonDictionary("TotalIncomeOfHPNegative") = 0
'    End If
'    If IncomeFromHP <> "" Then
'        jsonDictionary("TotalIncomeOfHP") = IncomeFromHP
'     Else
'        jsonDictionary("TotalIncomeOfHP") = 0
'    End If
'UpdateProgressBar

'Konda updated HP on 22-01-2026

Dim m As Long
 Dim PropertyDetails, PropertyDetailsobj, AddressDetailWithZipCode, CoOwners, CoOwnersobj, TenantDetails, TenantDetailsobj, Rentdetails

 Set PropertyDetailsobj = CreateObject("Scripting.Dictionary")
 Set AddressDetailWithZipCode = CreateObject("Scripting.Dictionary")
 Set CoOwnersobj = CreateObject("Scripting.Dictionary")
 Set TenantDetailsobj = CreateObject("Scripting.Dictionary")
 Set Rentdetails = CreateObject("Scripting.Dictionary")
 Set PropertyDetails = New Collection
 Set CoOwners = New Collection
 Set TenantDetails = New Collection
 
 subProcCaption = "Schedule HP"
   
    If (Not IsEmpty(AddrDetail_HP)) Then
    If (UBound(AddrDetail_HP) > 0) Then
        noOfProcessSub = UBound(AddrDetail_HP)
        For m = 1 To UBound(AddrDetail_HP)
            If Len(AddrDetail_HP(m)) = 0 Then
                Exit For
            End If
            PropertyDetailsobj("HPSNo") = m
            If AddrDetail_HP(m) <> "" Then
                AddressDetailWithZipCode("AddrDetail") = UCase(AddrDetail_HP(m))
            Else
                AddressDetailWithZipCode("AddrDetail") = ""
            End If
            
            If CityOrTownOrDistrict_HP(m) <> "" Then
                AddressDetailWithZipCode("CityOrTownOrDistrict") = UCase(CityOrTownOrDistrict_HP(m))
            Else
                AddressDetailWithZipCode("CityOrTownOrDistrict") = ""
            End If
            
            If StateCode_HP(m) <> "" Then
                AddressDetailWithZipCode("StateCode") = UCase(StateCode_HP(m))
            End If
            
            If CountryCode_HP(m) <> "" Then
                AddressDetailWithZipCode("CountryCode") = UCase(CountryCode_HP(m))
            End If
            
            If CountryCode_HP(m) = "91" Then
                If PinCode_HP(m) <> "" Then
                    AddressDetailWithZipCode("PinCode") = UVCase(CDbl(PinCode_HP(m)))
                End If
            End If
            
            If CountryCode_HP(m) <> "91" Then
                If ZipCode_HP(m) <> "" Then
                    AddressDetailWithZipCode("ZipCode") = UCase(ZipCode_HP(m))
                Else
                    AddressDetailWithZipCode("ZipCode") = " "
                End If
            End If
            PropertyDetailsobj.add "AddressDetailWithZipCode", AddressDetailWithZipCode
            Set AddressDetailWithZipCode = Nothing
            Set AddressDetailWithZipCode = CreateObject("Scripting.Dictionary")

         If OwnerProperty_HP(m) <> "" Then
            If (OwnerProperty_HP(m)) = "Self" Then
               OwnerProperty_HP(m) = "SE"
            ElseIf (OwnerProperty_HP(m)) = "Minor" Then
               OwnerProperty_HP(m) = "MI"
            ElseIf (OwnerProperty_HP(m)) = "Spouse" Then
               OwnerProperty_HP(m) = "SP"
            ElseIf (OwnerProperty_HP(m)) = "Others" Then
               OwnerProperty_HP(m) = "OT"
            End If
           
            If OwnerProperty_HP(m) <> "" Then
            PropertyDetailsobj("PropertyOwner") = UCase(OwnerProperty_HP(m))
            End If
            
            If OwnerPropertyDescription_HP(m) <> "" Then
            PropertyDetailsobj("PropertyOwnerOther") = UCase(OwnerPropertyDescription_HP(m))
            End If
            
            If CoOwnedYN_HP(m) <> "" Then
            PropertyDetailsobj("PropCoOwnedFlg") = UCase(CoOwnedYN_HP(m))
            End If
            
            If CoOwnedShare_HP(m) <> "" Then
            PropertyDetailsobj("AsseseeShareProperty") = CoOwnedShare_HP(m)
            Else
            PropertyDetailsobj("AsseseeShareProperty") = "100"
            End If
            
            If UCase(CoOwnedYN_HP(m)) = "YES" Then
                Dim j As Long
                Dim rangecells1 As Range
                Dim rangecells2 As Range
                Dim rangecells3 As Range
                Dim rangecells4 As Range

                Set rangecells1 = Sheet19.Range("HP.Co.Name" & m).Cells
                Set rangecells2 = Sheet19.Range("HP.Co.PAN" & m).Cells
                Set rangecells3 = Sheet19.Range("HP.Co.Share" & m).Cells
                Set rangecells4 = Sheet19.Range("HP.Co.Aadhaar" & m).Cells
            
                setTblinfo_hpcoindex (m)
            
                ReDim CoName_HP(m, end_hpco)
                ReDim CoPAN_HP(m, end_hpco)
                ReDim CoAadhar_HP(m, end_hpco)
                ReDim CoShare_HP(m, end_hpco)

                For j = 1 To end_hpco
                    CoName_HP(m, j) = rangecells1.item(j).Value
                    CoPAN_HP(m, j) = rangecells2.item(j).Value
                    CoAadhar_HP(m, j) = rangecells4.item(j).Value
                    CoShare_HP(m, j) = rangecells3.item(j).text
                    CoShare_HP(m, j) = Round(CoShare_HP(m, j), 2)
      
                    CoOwnersobj("CoOwnersSNo") = j
                    If CoName_HP(m, j) <> "" Then
                        CoOwnersobj("NameCoOwner") = CoName_HP(m, j)
                    End If
                    If CoPAN_HP(m, j) <> "" Then
                        CoOwnersobj("PAN_CoOwner") = CoPAN_HP(m, j)
                    End If
                    If CoAadhar_HP(m, j) <> "" Then
                        CoOwnersobj("Aadhaar_CoOwner") = UCase(CoAadhar_HP(m, j))
                    End If
                    CoOwnersobj("PercentShareProperty") = CoShare_HP(m, j)
                    CoOwners.add CoOwnersobj
                    Set CoOwnersobj = Nothing
                    Set CoOwnersobj = CreateObject("Scripting.Dictionary")
                Next
                  PropertyDetailsobj.add "CoOwners", CoOwners
                Set CoOwners = Nothing
                Set CoOwners = New Collection
            End If
          
            
            If ifLetOut_HP(m) <> "" Then
                If ifLetOut_HP(m) = "L" Then
                    PropertyDetailsobj("ifLetOut") = "L"
                End If
                If ifLetOut_HP(m) = "S" Then
                    PropertyDetailsobj("ifLetOut") = "S"
                End If
                If ifLetOut_HP(m) = "D" Then
                    PropertyDetailsobj("ifLetOut") = "D"
                End If
            Else
                PropertyDetailsobj("ifLetOut") = UCase(DefaultifLetOut_HP)
            End If
            
            If ifLetOut_HP(m) = "S" Then
            Else

                Set rangecells1 = Sheet19.Range("HP.NameofTenant" & m).Cells
                Set rangecells2 = Sheet19.Range("HP.PANofTenant" & m).Cells
                Set rangecells3 = Sheet19.Range("HP.TANofTenant" & m).Cells
                Set rangecells4 = Sheet19.Range("HP.AadharofTenant" & m).Cells
            
                setTblinfo_hprptfrm
            
                ReDim NameofTenant_HP(rangecells1.Cells.count)
                ReDim PANofTenant_HP(rangecells2.Cells.count)
                ReDim AadharofTenant_HP(rangecells4.Cells.count)
                ReDim TANofTenant_HP(rangecells3.Cells.count)
                If UBound(NameofTenant_HP) > 0 Then
                    Dim setobject
                    setobject = False
                    For j = 1 To UBound(NameofTenant_HP)
                        NameofTenant_HP(j) = rangecells1.item(j).Value
                        PANofTenant_HP(j) = rangecells2.item(j).Value
                        TANofTenant_HP(j) = rangecells3.item(j).Value
                        AadharofTenant_HP(j) = rangecells4.item(j).Value
                        If NameofTenant_HP(j) <> "" Then
                            setobject = True
                            TenantDetailsobj("TenantSNo") = j
                            TenantDetailsobj("NameofTenant") = UCase(NameofTenant_HP(j))
                            If PANofTenant_HP(j) <> "" Then
                            TenantDetailsobj("PANofTenant") = UCase(PANofTenant_HP(j))
                            End If
                            If AadharofTenant_HP(j) <> "" Then
                            TenantDetailsobj("AadhaarofTenant") = UCase(AadharofTenant_HP(j))
                            End If
                            If TANofTenant_HP(j) <> "" Then
                            TenantDetailsobj("PANTANofTenant") = TANofTenant_HP(j)
                            End If
                       
                            TenantDetails.add TenantDetailsobj
                            Set TenantDetailsobj = Nothing
                            Set TenantDetailsobj = CreateObject("Scripting.Dictionary")
                         End If
                    Next
                    If setobject Then
                        PropertyDetailsobj.add "TenantDetails", TenantDetails
                        Set TenantDetails = Nothing
                        Set TenantDetails = New Collection
                    End If
                End If
            End If

            If AnnualLetableValue_HP(m) <> "" Then
                Rentdetails("AnnualLetableValue") = UVCase(CDbl(AnnualLetableValue_HP(m)))
            Else
                Rentdetails("AnnualLetableValue") = UVCase(CDbl(DefaultAnnualLetableValue_HP))
            End If

            If RentNotRealized_HP(m) <> "" Then
                Rentdetails("RentNotRealized") = UVCase(CDbl(RentNotRealized_HP(m)))
            Else
                Rentdetails("RentNotRealized") = UVCase(CDbl(DefaultRentNotRealized_HP))
            End If

            If LocalTaxes_HP(m) <> "" Then
                Rentdetails("LocalTaxes") = UVCase(CDbl(LocalTaxes_HP(m)))
            Else
                Rentdetails("LocalTaxes") = UVCase(CDbl(DefaultLocalTaxes_HP))
            End If

            If TotalUnrealizedAndTax_HP(m) <> "" Then
                Rentdetails("TotalUnrealizedAndTax") = UVCase(CDbl(TotalUnrealizedAndTax_HP(m)))
            Else
                Rentdetails("TotalUnrealizedAndTax") = UVCase(CDbl(DefaultTotalUnrealizedAndTax_HP))
            End If

            If BalanceALV_HP(m) <> "" Then
                Rentdetails("BalanceALV") = UVCase(CDbl(BalanceALV_HP(m)))
            Else
                Rentdetails("BalanceALV") = UVCase(CDbl(DefaultBalanceALV_HP))
            End If

            If IncomeOfHPInOwnHand_HP(m) <> "" Then
                Rentdetails("AnnualOfPropOwned") = UVCase(CDbl(IncomeOfHPInOwnHand_HP(m)))
            Else
                Rentdetails("AnnualOfPropOwned") = 0
            End If

            If ThirtyPercentOfBalance_HP(m) <> "" Then
                Rentdetails("ThirtyPercentOfBalance") = UVCase(CDbl(ThirtyPercentOfBalance_HP(m)))
            Else
                Rentdetails("ThirtyPercentOfBalance") = UVCase(CDbl(DefaultThirtyPercentOfBalance_HP))
            End If

            If IntOnBorwCap_HP(m) <> "" Then
                 Rentdetails("IntOnBorwCap") = UVCase(CDbl(IntOnBorwCap_HP(m)))
            Else
                 Rentdetails("IntOnBorwCap") = UVCase(CDbl(DefaultIntOnBorwCap_HP))
            End If
'Konda--------------AY_2026_27_V0.2.4--------------------------------
           Dim ScheduleUs24BDtls_G, ScheduleUs24BDtls, Section24B, Section24BDtls, mIntCells24B, k

Set Section24B = CreateObject("Scripting.Dictionary")
Set ScheduleUs24BDtls_G = CreateObject("Scripting.Dictionary")
Set Section24BDtls = CreateObject("Scripting.Dictionary")

Set ScheduleUs24BDtls = CreateObject("Scripting.Dictionary")
Set ScheduleUs24BDtls = New Collection

subProcCaption = "24B"

Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G

If Sheet19.Range("TotAmt.24b" & m).Value > 0 Then
 mIntCells24B = Sheet19.Range("LoanfrmBankOrInstitute.24b" & m).count
 
  For k = 1 To mIntCells24B
   If (Sheet19.Range("LoanfrmBankOrInstitute.24b" & m).Cells(k, 1).Value <> "" And Sheet19.Range("LoanfrmBankOrInstitute.24b" & m).Cells(k, 1).Value <> "(Select)") Then
   
   LoanTknFrom_G = Sheet19.Range("LoanfrmBankOrInstitute.24b" & m).Cells(k, 1).Value
   
   BankOrInstnName_G = Sheet19.Range("bankName.24b" & m).Cells(k, 1).Value
   
   LoanAccNoOfBankOrInstnRefNo_G = Sheet19.Range("loanAccNum.24b" & m).Cells(k, 1).Value
   DateofLoan_G = Sheet19.Range("loanDate.24b" & m).Cells(k, 1).Value
   TotalLoanAmt_G = Sheet19.Range("loanAmt.24b" & m).Cells(k, 1).Value
   LoanOutstndngAmt_G = Sheet19.Range("loanOutstanding.24b" & m).Cells(k, 1).Value
   InterestPaid_G = Sheet19.Range("Intrst.24b" & m).Cells(k, 1).Value
   
   
   If LoanTknFrom_G <> "" Then
            
             If LoanTknFrom_G = "Bank " Then
            LoanTknFrom_G = "B"
            
            ElseIf LoanTknFrom_G = "Other than Bank" Then
            LoanTknFrom_G = "I"
            End If
        ScheduleUs24BDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        ScheduleUs24BDtls_G("LoanTknFrom") = ""
   End If
    
   If BankOrInstnName_G <> "" Then
   ScheduleUs24BDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   ScheduleUs24BDtls_G("BankOrInstnName") = ""
   End If
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   ScheduleUs24BDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   ScheduleUs24BDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   ScheduleUs24BDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   ScheduleUs24BDtls_G("DateofLoan") = ""
   End If
   
   ScheduleUs24BDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   
   ScheduleUs24BDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
   
   ScheduleUs24BDtls_G("InterestUs24B") = UVCase(InterestPaid_G)
 
                ScheduleUs24BDtls.add ScheduleUs24BDtls_G
                Set ScheduleUs24BDtls_G = Nothing
                Set ScheduleUs24BDtls_G = CreateObject("Scripting.Dictionary")
  End If
  Next
  
               
               Section24B.add "Section24BDtls", ScheduleUs24BDtls
                Set ScheduleUs24BDtls = Nothing
                Set ScheduleUs24BDtls = CreateObject("Scripting.Dictionary")
 

 TotalInterestPaid_G = Sheet19.Range("TotAmt.24b" & m).Value
 Section24B("TotalInterestUs24B") = UVCase(TotalInterestPaid_G)

                Rentdetails.add "Section24B", Section24B
                Set Section24B = Nothing
                Set Section24B = CreateObject("Scripting.Dictionary")
End If
           
'---------------------------------------------------------------------
         
            If TotalDeduct_HP(m) <> "" Then
                Rentdetails("TotalDeduct") = UVCase(CDbl(TotalDeduct_HP(m)))
            Else
                Rentdetails("TotalDeduct") = UVCase(CDbl(DefaultTotalDeduct_HP))
            End If
            
            If Sheet19.Range("HP.RentOfEarlierYrSec_AandAA" & m).Value <> "" Then
                Rentdetails("ArrearsUnrealizedRentRcvd") = UVCase(CDbl(Sheet19.Range("HP.RentOfEarlierYrSec_AandAA" & m).Value))
            End If
            
            If IncomeOfHP_HP(m) <> "" Then
                Rentdetails("IncomeOfHP") = UVCase(CDbl(IncomeOfHP_HP(m)))
            Else
                Rentdetails("IncomeOfHP") = UVCase(CDbl(DefaultIncomeOfHP_HP))
            End If
            PropertyDetailsobj.add "Rentdetails", Rentdetails
            Set Rentdetails = Nothing
            Set Rentdetails = CreateObject("Scripting.Dictionary")
            
            PropertyDetails.add PropertyDetailsobj
            Set PropertyDetailsobj = Nothing
            Set PropertyDetailsobj = CreateObject("Scripting.Dictionary")
        End If
        Next
    jsonDictionary.add "PropertyDetails", PropertyDetails
            
'    If Sheet19.Range("HP.PassTroughIncome").Value <> "" Then
'        jsonDictionary("PassThroghIncome") = UVCase(CDbl(Sheet19.Range("HP.PassTroughIncome").Value))
'    End If
    
    
End If
End If  'Konda_29/01/2026========

    If TotalIncomeChargeableUnHP_HP <> "" Then
        jsonDictionary("TotalIncomeChargeableUnHP") = UVCase(CDbl(TotalIncomeChargeableUnHP_HP))
    Else
        jsonDictionary("TotalIncomeChargeableUnHP") = UVCase(CDbl(DefaultTotalIncomeChargeableUnHP_HP))
    End If
    

'End If
'End If



'=========================================Schema V0.2.2 on 20-01-2026
    'Imcome from Other Sources
    
    If IncomeFromOS <> "" Then
        jsonDictionary("IncomeOthSrc") = IncomeOthSrc
     Else
        jsonDictionary("IncomeOthSrc") = 0
    End If
UpdateProgressBar
                
     If (Not IsEmpty(Others_NOI2)) Then
     If (UBound(Others_NOI2) > 0) Then
            For i = 1 To UBound(Others_NOI2)
'                If Mid(Others_NOI2(i), 1, 1) = "D" Then
'                Others_NOI2(i) = "DIV"
                If Mid(Others_NOI2(i), 1, 15) = "Interest from S" Then
                            Others_NOI2(i) = "SAV"
                ElseIf Mid(Others_NOI2(i), 1, 15) = "Interest from D" Then
                            Others_NOI2(i) = "IFD"
                ElseIf Mid(Others_NOI2(i), 1, 15) = "Interest from I" Then
                            Others_NOI2(i) = "TAX"
                ElseIf Mid(Others_NOI2(i), 1, 11) = "Any Other I" Then
                            Others_NOI2(i) = "OII"
                ElseIf Mid(Others_NOI2(i), 1, 1) = "F" Then
                            Others_NOI2(i) = "FAP"
                ElseIf Others_NOI2(i) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(11)" Then
                            Others_NOI2(i) = "10(11)(iP)"
                ElseIf Others_NOI2(i) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(11)" Then
                            Others_NOI2(i) = "10(11)(iiP)"
                ElseIf Others_NOI2(i) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(12)" Then
                            Others_NOI2(i) = "10(12)(iP)"
                ElseIf Others_NOI2(i) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(12)" Then
                            Others_NOI2(i) = "10(12)(iiP)"
                ElseIf Mid(Others_NOI2(i), 1, 1) = "A" Then
                            Others_NOI2(i) = "OTH"
                Else
                   Others_NOI2(i) = Mid(Others_NOI2(i), 1, 9)
                End If
            Next i
                    
            For i = 1 To UBound(Others_NOI2)
                If Others_NOI2(i) <> "" Then
                    OthersIncDtlsOth("OthSrcNatureDesc") = Others_NOI2(i)
                End If
                            
                If Others_NOI3(i) <> "" And Others_NOI3(i) <> "Not Applicable" Then
                    OthersIncDtlsOth("OthSrcOthNatOfInc") = UCase(Others_NOI3(i))
                End If
                        
                If Others_Amt2(i) <> "" Then
                    OthersIncDtlsOth("OthSrcOthAmount") = Others_Amt2(i)
                End If
                
                OthersIncDtlsOthSrc.add OthersIncDtlsOth
                Set OthersIncDtlsOth = Nothing
                Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
            Next
             End If
        End If
        
    'Commented by Konda as per New Schema V0.2.2 on 20-01-2026
    '        If Sheet1.Range("OSIncomeNotifiedOther89A") <> "" Then
    '
    '                     Set OthersIncDtlsOth = Nothing
    '                    Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
    '                    OthersIncDtlsOth("OthSrcNatureDesc") = "OTHNOT89A"
    '                    OthersIncDtlsOth("OthSrcOthAmount") = Sheet1.Range("OSIncomeNotifiedOther89A")
    '                    OthersIncDtlsOthSrc.add OthersIncDtlsOth
    '            End If
                
            If Up16Of12To15Of3 <> "" Or Up16Of3To31Of3 <> "" Or Up16Of9To15Of12 <> "" Or Upto15Of6 <> "" Or Upto15Of9 <> "" Then
                    Set OthersIncDtlsOth = Nothing
                    Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
                    Set DateRange = Nothing
                    Set DateRange = CreateObject("Scripting.Dictionary")
                    
                    'Malli----AY_2025_26
'Konda updated on 13-04-2026
'                    If TotDividendInc <> "" Then
'                    OthersIncDtlsOth("OthSrcOthAmount") = TotDividendInc
'                    Else
'                    OthersIncDtlsOth("OthSrcOthAmount") = 0
'                    End If
                    OthersIncDtlsOth("OthSrcNatureDesc") = "DIV"
                    
                    If TotDividendInc <> "" Then
                    OthersIncDtlsOth("OthSrcOthAmount") = TotDividendInc
                    Else
                    OthersIncDtlsOth("OthSrcOthAmount") = 0
                    End If
                    
'                    DateRange("Up16Of12To15Of3") = Up16Of12To15Of3
'                    DateRange("Up16Of3To31Of3") = Up16Of3To31Of3
'                    DateRange("Up16Of9To15Of12") = Up16Of9To15Of12
'                    DateRange("Upto15Of6") = Upto15Of6
'                    DateRange("Upto15Of9") = Upto15Of9

                    DateRange("Upto15Of6") = Upto15Of6
                    DateRange("Upto15Of9") = Upto15Of9
                    DateRange("Up16Of9To15Of12") = Up16Of9To15Of12
                    DateRange("Up16Of12To15Of3") = Up16Of12To15Of3
                    DateRange("Up16Of3To31Of3") = Up16Of3To31Of3
                    
'=============================updated on 13-04-2026
                    DividendInc.add "DateRange", DateRange
                    OthersIncDtlsOth.add "DividendInc", DividendInc
                    OthersIncDtlsOthSrc.add OthersIncDtlsOth

                End If

            
            
'            If Trim(Sheet1.Range("OSIncomeNotified89A_AmountUS").Value) <> "" Or Trim(Sheet1.Range("OSIncomeNotified89A_AmountUK").Value) <> "" Or Trim(Sheet1.Range("OSIncomeNotified89A_AmountCan").Value) <> "" Then
'                    Set OthersIncDtlsOth = Nothing
'                    Set OthersIncDtlsOth = CreateObject("Scripting.Dictionary")
'                    Set DateRange = Nothing
'                    Set DateRange = CreateObject("Scripting.Dictionary")
''Commented by Konda as per New Schema V0.2.2 on 20-01-2026
''                    OthersIncDtlsOth("OthSrcOthAmount") = OSIncomeNotified89A
''                    OthersIncDtlsOth("OthSrcNatureDesc") = "NOT89A"
'                    DateRange("Up16Of12To15Of3") = Up16Of12To15Of3_89A
'                    DateRange("Up16Of3To31Of3") = Up16Of3To31Of3_89A
'                    DateRange("Up16Of9To15Of12") = Up16Of9To15Of12_89A
'                    DateRange("Upto15Of6") = Upto15Of6_89A
'                    DateRange("Upto15Of9") = Upto15Of9_89A
'
''Commented by Konda as per New Schema V0.2.2 on 20-01-2026
''                    If Trim(Sheet1.Range("OSIncomeNotified89A_AmountUS").Value) <> "" Then
''                        NOT89AType("NOT89ACountrycode") = "US"
''                        NOT89AType("NOT89AAmount") = Sheet1.Range("OSIncomeNotified89A_AmountUS").Value
''                        NOT89A.add NOT89AType
''                        Set NOT89AType = Nothing
''                        Set NOT89AType = CreateObject("Scripting.Dictionary")
''                    End If
''                    If Trim(Sheet1.Range("OSIncomeNotified89A_AmountUK").Value) <> "" Then
''                        NOT89AType("NOT89ACountrycode") = "UK"
''                        NOT89AType("NOT89AAmount") = Sheet1.Range("OSIncomeNotified89A_AmountUK").Value
''                        NOT89A.add NOT89AType
''                        Set NOT89AType = Nothing
''                        Set NOT89AType = CreateObject("Scripting.Dictionary")
''                    End If
''                    If Trim(Sheet1.Range("OSIncomeNotified89A_AmountCan").Value) <> "" Then
''                        NOT89AType("NOT89ACountrycode") = "CA"
''                        NOT89AType("NOT89AAmount") = Sheet1.Range("OSIncomeNotified89A_AmountCan").Value
''                        NOT89A.add NOT89AType
''                        Set NOT89AType = Nothing
''                        Set NOT89AType = CreateObject("Scripting.Dictionary")
''                    End If
'
''                    OthersIncDtlsOth.add "NOT89A", NOT89A
''                    NOT89AInc.add "DateRange", DateRange
''                    OthersIncDtlsOth.add "NOT89AInc", NOT89AInc
''==========================================Schema V0.2.2 on 20-01-2026
'                    OthersIncDtlsOthSrc.add OthersIncDtlsOth
'                End If
            OthersInc.add "OthersIncDtlsOthSrc", OthersIncDtlsOthSrc
        
            jsonDictionary.add "OthersInc", OthersInc
       
                
                
    If Trim(Sheet1.Range("IncD.LessDeduction57").Value) <> "" Then
        jsonDictionary("DeductionUs57iia") = DeductionUs57iia
    Else
        jsonDictionary("DeductionUs57iia") = 0
    End If
                
                     
    'Gross Total Income
    If GrossTotIncome <> "" Then
        jsonDictionary("GrossTotIncome") = GrossTotIncome
     Else
        jsonDictionary("GrossTotIncome") = 0
    End If
    
   'GrossTotIncomeIncLTCG112A added by Konda on 07-02-2025
     If GrossTotIncomeIncLTCG112A <> "" Then
        jsonDictionary("GrossTotIncomeIncLTCG112A") = GrossTotIncomeIncLTCG112A
     Else
        jsonDictionary("GrossTotIncomeIncLTCG112A") = 0
    End If
    
    If Investment = "Yes" Then
     Investment = "Y"
    ElseIf Investment = "No" Then
    Investment = "N"
    End If
    
'    If Investment <> "(Select)" And Investment <> "" Then
'    writeXML "<ITRForm:DepPayInvClmUndDednVIA>" & Investment & "</ITRForm:DepPayInvClmUndDednVIA>"
'    End If
    
UpdateProgressBar
    If Trim(Sheet1.Range("IncD.Section80C").Value) <> "" Then
        UsrDeductUndChapVIA("Section80C") = Section80C
     Else
        UsrDeductUndChapVIA("Section80C") = 0
    End If
    
UpdateProgressBar
    If Trim(Sheet1.Range("IncD.Section80CCC").Value) <> "" Then
        UsrDeductUndChapVIA("Section80CCC") = Section80CCC
     Else
        UsrDeductUndChapVIA("Section80CCC") = 0
    End If
UpdateProgressBar

'New Schema updated by Konda as AY-2026-27 on 26-12-2025

If Trim(Sheet1.Range("IncD.Section80CCC").Value) <> "" And Trim(Sheet1.Range("IncD.Section80CCC").Value) > 0 Then
            For i = 1 To UBound(Type_80CCC)
             'Malli updated as per V0.9.1 changes  17/04/2026
'                If Type_80CCC(i) <> "" Then
'                    PensionContributionFund("TypeofIdentifier") = Type_80CCC(i)
'                End If
                
                If Type_80CCC(i) <> "" Then
                  If UCase(Type_80CCC(i)) = UCase("PRAN") Then
                      PensionContributionFund("TypeofIdentifier") = CStr("PRAN")
                  ElseIf UCase(Type_80CCC(i)) = UCase("Other than PRAN") Then
                      PensionContributionFund("TypeofIdentifier") = CStr("OTHPRAN")
                  End If
                Else
                      PensionContributionFund("TypeofIdentifier") = ""
                End If
                
                
                '-------------------------

                If Name_80CCC(i) <> "" Then
                    PensionContributionFund("NameofIdentifier") = CStr(Name_80CCC(i))
                Else
                    PensionContributionFund("NameofIdentifier") = ""
                End If
                
                If Amount_80CCC(i) <> "" Then
'27/07/2026
                'PensionContributionFund("Amount") = CInt(Amount_80CCC(i))
                 PensionContributionFund("Amount") = CDbl(Amount_80CCC(i))
                '----------------
                    Else
                PensionContributionFund("Amount") = ""
                End If
                
                PensionContribution80CCC.add PensionContributionFund
                Set PensionContributionFund = Nothing
                Set PensionContributionFund = CreateObject("Scripting.Dictionary")
            
            Next
        
    UsrDeductUndChapVIA.add "PensionContribution80CCC", PensionContribution80CCC
    
'   Else
'         UsrDeductUndChapVIA.add "PensionContribution80CCC", PensionContribution80CCC
    End If
    
    If Trim(Sheet1.Range("IncD.Section80CCD_SE").Value) <> "" Then
        UsrDeductUndChapVIA("Section80CCDEmployeeOrSE") = Section80CCDEmployeeOrSE
     Else
        UsrDeductUndChapVIA("Section80CCDEmployeeOrSE") = 0
    End If
                     
UpdateProgressBar


'Malli commented as per V0.9.1 version changes
'New Schema updated by Konda as AY-2026-27 on 26-12-2025

'If Trim(Sheet1.Range("IncD.Section80CCD_SE").Value) <> "" And Trim(Sheet1.Range("IncD.Section80CCD_SE").Value) > 0 Then
'
'
'            For i = 1 To UBound(Type_1_80CCD)
'
'                If Type_1_80CCD(i) <> "" Then
'                    PensionContributionFund("TypeofIdentifier") = Type_1_80CCD(i)
'                End If
'
'                If Name_1_80CCD(i) <> "" Then
'                    PensionContributionFund("NameofIdentifier") = Name_1_80CCD(i)
'                End If
'                If Amount_1_80CCD(i) <> "" Then
'                    PensionContributionFund("Amount") = Amount_1_80CCD(i)
'                End If
'
'
'                PensionContribution80CCD1.add PensionContributionFund
'                Set PensionContributionFund = Nothing
'                Set PensionContributionFund = CreateObject("Scripting.Dictionary")
'
'            Next
'
'    UsrDeductUndChapVIA.add "PensionContribution80CCD1", PensionContribution80CCD1
'
''   Else
''        UsrDeductUndChapVIA.add "PensionContribution80CCD1", PensionContribution80CCD1
'    End If
'
  '------------------------------------------------------------

    If Trim(Sheet1.Range("IncD.Section80CCD1B_SE").Value) <> "" Then
        UsrDeductUndChapVIA("Section80CCD1B") = Section80CCD1B
     Else
        UsrDeductUndChapVIA("Section80CCD1B") = 0
    End If
    
'Malli commented as per V0.9.1 version changes
'New Schema updated by Konda as AY-2026-27 on 26-12-2025
'If Trim(Sheet1.Range("IncD.Section80CCD1B_SE").Value) <> "" And Trim(Sheet1.Range("IncD.Section80CCD1B_SE").Value) > 0 Then
'
'
'            For i = 1 To UBound(Type_1b_80CCD)
'
'                If Type_1b_80CCD(i) <> "" Then
'                    PensionContributionFund("TypeofIdentifier") = Type_1b_80CCD(i)
'                End If
'
'                If Name_1b_80CCD(i) <> "" Then
'                    PensionContributionFund("NameofIdentifier") = Name_1b_80CCD(i)
'                End If
'                If Amount_1b_80CCD(i) <> "" Then
'                    PensionContributionFund("Amount") = Amount_1b_80CCD(i)
'                End If
'
'
'                PensionContribution80CCD1B.add PensionContributionFund
'                Set PensionContributionFund = Nothing
'                Set PensionContributionFund = CreateObject("Scripting.Dictionary")
'
'            Next
'
'    UsrDeductUndChapVIA.add "PensionContribution80CCD1B", PensionContribution80CCD1B
'
''    Else
''        UsrDeductUndChapVIA.add "PensionContribution80CCD1B", PensionContribution80CCD1B
'    End If
'---------------------------------
    If Trim(Sheet1.Range("IncD.Section80CCD").Value) <> "" Then
        UsrDeductUndChapVIA("Section80CCDEmployer") = Section80CCDEmployer
     Else
        UsrDeductUndChapVIA("Section80CCDEmployer") = 0
    End If
    
    
    'Malli----------AY_2025_26_21/04/2025
'As per AY-2026-27 New Schema PRANNum commented by Konda on 26-12-2025
'    Dim PRANNum_G As Variant
'
'    PRANNum_G = Sheet1.Range("sheet1.PRAN").Value
'
'    If Sheet1.Range("sheet1.PRAN").Locked = False And Trim(PRANNum_G) <> "" Then
'        UsrDeductUndChapVIA("PRANNum") = UCase(Trim(PRANNum_G))
'    End If

'Malli updated as pr V0.9.1 changes 17/04/2026
If Sheet1.Range("IncD.Section80CCD_SE") > 0 Or Sheet1.Range("IncD.Section80CCD1B_SE") > 0 Then
 Dim mincellsPRANNum As Variant
 
 Dim PRANDtls, PRANDtls_PRANNum, PRANNum_Gen, prn
 
 Set PRANDtls_PRANNum = CreateObject("scripting.dictionary")
 Set PRANDtls = New Collection
 
 mincellsPRANNum = Sheet1.Range("pran_new").Rows.count
 
 
 For prn = 1 To mincellsPRANNum
 
     PRANNum_Gen = Sheet1.Range("pran_new").Cells(prn, 1).Value
     If PRANNum_Gen <> "" Then
         PRANDtls_PRANNum("PRANNum") = CStr(PRANNum_Gen)
    
     
                    PRANDtls.add PRANDtls_PRANNum
                    Set PRANDtls_PRANNum = Nothing
                    Set PRANDtls_PRANNum = CreateObject("Scripting.Dictionary")
     End If
 Next
 Debug.Print PRANDtls.count
 If PRANDtls.count > 0 Then
     UsrDeductUndChapVIA.add "PRANDtls", PRANDtls
 End If
 
 End If

'    '--------------------------------------------------
'    '--------------------------------------------------
    
    
    If Trim(Sheet1.Range("IncD.Section80DValue").Value) <> "" Then
        UsrDeductUndChapVIA("Section80D") = Section80D
     Else
        UsrDeductUndChapVIA("Section80D") = 0
    End If

UpdateProgressBar
                     
    'If SELECT80DD <> "" Then
     '   UsrDeductUndChapVIA("Section80DDUsrType") = SELECT80DD
    'End If
    If Trim(Sheet1.Range("IncD.Section80DD").Value) <> "" Then
        UsrDeductUndChapVIA("Section80DD") = Section80DD
     Else
        UsrDeductUndChapVIA("Section80DD") = 0
    End If
    
    If SELECT80DDS <> "" Then
        UsrDeductUndChapVIA("Section80DDBUsrType") = SELECT80DDS
    End If
    
    'Malli----AY_2025_26--21/04/2025
    Dim NameOfSpecDisease80DDB_G, NameOfSpecDisease80DDB_G_enum As Variant
    
    NameOfSpecDisease80DDB_G = Sheet1.Range("Sheet1.Specified_Disease").Value
    
    If Sheet1.Range("Sheet1.Specified_Disease").Locked = False And Trim(NameOfSpecDisease80DDB_G) <> "" Then
        
        
        If NameOfSpecDisease80DDB_G = "(a) Dementia" Then
        NameOfSpecDisease80DDB_G_enum = "a"
        
        ElseIf NameOfSpecDisease80DDB_G = "(b) Dystonia Musculorum Deformans" Then
        NameOfSpecDisease80DDB_G_enum = "b"
        
        ElseIf NameOfSpecDisease80DDB_G = "(c) Motor Neuron Disease" Then
        NameOfSpecDisease80DDB_G_enum = "c"
        
        ElseIf NameOfSpecDisease80DDB_G = "(d) Ataxia" Then
        NameOfSpecDisease80DDB_G_enum = "d"
        
        ElseIf NameOfSpecDisease80DDB_G = "(e) Chorea" Then
        NameOfSpecDisease80DDB_G_enum = "e"
        
        ElseIf NameOfSpecDisease80DDB_G = "(f) Hemiballismus" Then
        NameOfSpecDisease80DDB_G_enum = "f"
        
        ElseIf NameOfSpecDisease80DDB_G = "(g) Aphasia" Then
        NameOfSpecDisease80DDB_G_enum = "g"
        
        ElseIf NameOfSpecDisease80DDB_G = "(h) Parkinsons Disease" Then
        NameOfSpecDisease80DDB_G_enum = "h"
        
        ElseIf NameOfSpecDisease80DDB_G = "(i) Malignant Cancers" Then
        NameOfSpecDisease80DDB_G_enum = "i"
        
        ElseIf NameOfSpecDisease80DDB_G = "(j) Full Blown Acquired Immuno-Deficiency Syndrome (AIDS)" Then
        NameOfSpecDisease80DDB_G_enum = "j"
        
        ElseIf NameOfSpecDisease80DDB_G = "(k) Chronic Renal failure" Then
        NameOfSpecDisease80DDB_G_enum = "k"
        
        ElseIf NameOfSpecDisease80DDB_G = "(l) Hematological disorders" Then
        NameOfSpecDisease80DDB_G_enum = "l"
        
        ElseIf NameOfSpecDisease80DDB_G = "(m) Hemophilia" Then
        NameOfSpecDisease80DDB_G_enum = "m"
        
        ElseIf NameOfSpecDisease80DDB_G = "(n) Thalassaemia" Then
        NameOfSpecDisease80DDB_G_enum = "n"
        
        Else
         NameOfSpecDisease80DDB_G_enum = ""
        End If
        
        If NameOfSpecDisease80DDB_G_enum <> "" Then
        'UsrDeductUndChapVIA("NameOfSpecDisease80DDB") = UCase(Trim(NameOfSpecDisease80DDB_G_enum))
        UsrDeductUndChapVIA("NameOfSpecDisease80DDB") = (Trim(NameOfSpecDisease80DDB_G_enum))
        End If
        
        
        
    End If
    
    '--------------------------------
    
    
    If Trim(Sheet1.Range("IncD.Section80DDB").Value) <> "" Then
        UsrDeductUndChapVIA("Section80DDB") = Section80DDB
     Else
        UsrDeductUndChapVIA("Section80DDB") = 0
    End If
UpdateProgressBar
                     
    If Trim(Sheet1.Range("IncD.Section80E").Value) <> "" Then
        UsrDeductUndChapVIA("Section80E") = Section80E
     Else
        UsrDeductUndChapVIA("Section80E") = 0
    End If
UpdateProgressBar
                     
    If Trim(Sheet1.Range("IncD.Section80EE").Value) <> "" Then
        UsrDeductUndChapVIA("Section80EE") = Section80EE
    Else
        UsrDeductUndChapVIA("Section80EE") = 0
    End If
           
    If Trim(Sheet1.Range("IncD.Section80EEA").Value) <> "" Then
        UsrDeductUndChapVIA("Section80EEA") = Section80EEA
    Else
        UsrDeductUndChapVIA("Section80EEA") = 0
    End If
    
    If Trim(Sheet1.Range("IncD.Section80EEB").Value) <> "" Then
        UsrDeductUndChapVIA("Section80EEB") = Section80EEB
    Else
        UsrDeductUndChapVIA("Section80EEB") = 0
    End If
                     
    If Trim(Sheet1.Range("IncD.Section80G").Value) <> "" Then
        UsrDeductUndChapVIA("Section80G") = Section80G
     Else
        UsrDeductUndChapVIA("Section80G") = 0
    End If
                    
    If Trim(Sheet1.Range("IncD.Section80GG").Value) <> "" Then
        UsrDeductUndChapVIA("Section80GG") = Section80GG
     Else
        UsrDeductUndChapVIA("Section80GG") = 0
    End If
                    
       'Malli----------AY_2025_26_21/04/2025
              Dim Form10BAAckNum_G As Variant
              
         Form10BAAckNum_G = Sheet1.Range("Sheet1.AckNum").Value
    If Sheet1.Range("Sheet1.AckNum").Locked = False And Trim(Form10BAAckNum_G) <> "" Then
        UsrDeductUndChapVIA("Form10BAAckNum") = UCase(Trim(Form10BAAckNum_G))
    End If
        '------------------------------------
        
        
    If Trim(Sheet1.Range("IncD.Section80GGA").Value) <> "" Then
        UsrDeductUndChapVIA("Section80GGA") = Section80GGA
     Else
        UsrDeductUndChapVIA("Section80GGA") = 0
    End If
                     
UpdateProgressBar
                     
    If Trim(Sheet1.Range("IncD.Section80GGC").Value) <> "" Then
        UsrDeductUndChapVIA("Section80GGC") = Section80GGC
     Else
        UsrDeductUndChapVIA("Section80GGC") = 0
    End If
                     
     'Generate - 06.04.2023.103.ID.01
    If Trim(Sheet1.Range("IncD.AnyOtherDeductions").Value) <> "" Then
            UsrDeductUndChapVIA("AnyOthSec80CCH") = Sheet1.Range("IncD.AnyOtherDeductions").Value
    End If
    'End Change
                     
    ' If SELECT80U <> "" Then
    '    UsrDeductUndChapVIA("Section80UUsrType") = SELECT80U
    'End If
                     
                     
    If Trim(Sheet1.Range("IncD.Section80U").Value) <> "" Then
        UsrDeductUndChapVIA("Section80U") = Section80U
     Else
        UsrDeductUndChapVIA("Section80U") = 0
    End If
  
    If Trim(Sheet1.Range("IncD.Section80TTA").Value) = "" Then
        UsrDeductUndChapVIA("Section80TTA") = 0
     Else
        UsrDeductUndChapVIA("Section80TTA") = Section80TTA
    End If
    
    If Trim(Sheet1.Range("IncD.Section80TTB").Value) = "" Then
        UsrDeductUndChapVIA("Section80TTB") = 0
     Else
        UsrDeductUndChapVIA("Section80TTB") = Section80TTB
    End If
                                                                                                   
    If Trim(Sheet1.Range("IncD.TotalChapVIADeductions_Input").Value) = "" Then
        UsrDeductUndChapVIA("TotalChapVIADeductions") = 0
     Else
        UsrDeductUndChapVIA("TotalChapVIADeductions") = TotalChapVIADeductions
    End If
    jsonDictionary.add "UsrDeductUndChapVIA", UsrDeductUndChapVIA
    
                    If Trim(Sheet1.Range("IncD.Section80C_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80C") = Sheet1.Range("IncD.Section80C_Calc").Value
                     Else
                        DeductUndChapVIA("Section80C") = 0
                    End If
                    If Trim(Sheet1.Range("IncD.Section80CCC_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80CCC") = Sheet1.Range("IncD.Section80CCC_Calc").Value
                     Else
                        DeductUndChapVIA("Section80CCC") = 0
                    End If
                     
                     
                UpdateProgressBar
                    If Trim(Sheet1.Range("IncD.Section80CCD_Calc_SE").Value) <> "" Then
                        DeductUndChapVIA("Section80CCDEmployeeOrSE") = Sheet1.Range("IncD.Section80CCD_Calc_SE").Value
                     Else
                        DeductUndChapVIA("Section80CCDEmployeeOrSE") = 0
                    End If
                     
                UpdateProgressBar
                    If Trim(Sheet1.Range("IncD.Section80CCD1B_Calc_SE").Value) <> "" Then
                        DeductUndChapVIA("Section80CCD1B") = Sheet1.Range("IncD.Section80CCD1B_Calc_SE").Value
                    Else
                        DeductUndChapVIA("Section80CCD1B") = 0
                    End If
                    
                    If Trim(Sheet1.Range("IncD.Section80CCD_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80CCDEmployer") = Round(Sheet1.Range("IncD.Section80CCD_Calc").Value, 0)
                     Else
                        DeductUndChapVIA("Section80CCDEmployer") = 0
                    End If
                     
                     
                    If Trim(Sheet1.Range("IncD.Section80DValue_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80D") = Sheet1.Range("IncD.Section80DValue_Calc").Value
                     Else
                        DeductUndChapVIA("Section80D") = 0
                    End If
        
                UpdateProgressBar
                     
                    If Trim(Sheet1.Range("IncD.Section80DD_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80DD") = Sheet1.Range("IncD.Section80DD_Calc").Value
                     Else
                        DeductUndChapVIA("Section80DD") = 0
                    End If
                    If Trim(Sheet1.Range("IncD.Section80DDB_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80DDB") = Sheet1.Range("IncD.Section80DDB_Calc").Value
                     Else
                        DeductUndChapVIA("Section80DDB") = 0
                    End If
                    If Trim(Sheet1.Range("IncD.Section80E_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80E") = Sheet1.Range("IncD.Section80E_Calc").Value
                     Else
                        DeductUndChapVIA("Section80E") = 0
                    End If
                     
                UpdateProgressBar
                     
                     If Trim(Sheet1.Range("IncD.Section80EE_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80EE") = Sheet1.Range("IncD.Section80EE_Calc").Value
                     Else
                        DeductUndChapVIA("Section80EE") = 0
                     End If
                     
                     If Trim(Sheet1.Range("IncD.Section80EEA_Calc").Value) <> "" Then
                    DeductUndChapVIA("Section80EEA") = Sheet1.Range("IncD.Section80EEA_Calc").Value
                     Else
                    DeductUndChapVIA("Section80EEA") = 0
                     End If
                     
                    If Trim(Sheet1.Range("IncD.Section80EEB_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80EEB") = Sheet1.Range("IncD.Section80EEB_Calc").Value
                     Else
                        DeductUndChapVIA("Section80EEB") = 0
                    End If
                     
                    If Trim(Sheet1.Range("IncD.Section80G_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80G") = Sheet1.Range("IncD.Section80G_Calc").Value
                     Else
                        DeductUndChapVIA("Section80G") = 0
                    End If
                      
                    If Trim(Sheet1.Range("IncD.Section80GG_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80GG") = Sheet1.Range("IncD.Section80GG_Calc").Value
                     Else
                        DeductUndChapVIA("Section80GG") = 0
                    End If
                     
                    If Trim(Sheet1.Range("IncD.Section80GGA_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80GGA") = Sheet1.Range("IncD.Section80GGA_Calc").Value
                     Else
                        DeductUndChapVIA("Section80GGA") = 0
                    End If
                     
                    If Trim(Sheet1.Range("IncD.Section80GGC_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80GGC") = Sheet1.Range("IncD.Section80GGC_Calc").Value
                     Else
                        DeductUndChapVIA("Section80GGC") = 0
                    End If
                     
                    If Trim(Sheet1.Range("IncD.Section80U_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80U") = Sheet1.Range("IncD.Section80U_Calc").Value
                     Else
                        DeductUndChapVIA("Section80U") = 0
                    End If
                     
                    If Trim(Sheet1.Range("IncD.Section80TTA_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80TTA") = Sheet1.Range("IncD.Section80TTA_Calc").Value
                     Else
                        DeductUndChapVIA("Section80TTA") = 0
                    End If
                     
                    If Trim(Sheet1.Range("IncD.Section80TTB_Calc").Value) <> "" Then
                        DeductUndChapVIA("Section80TTB") = Sheet1.Range("IncD.Section80TTB_Calc").Value
                     Else
                        DeductUndChapVIA("Section80TTB") = 0
                    End If
                    
                    'Generate - 06.04.2023.103.ID.01
                    If Sheet1.Range("IncD.AnyOtherDeductions_Calc").Value <> "" Then
                        DeductUndChapVIA("AnyOthSec80CCH") = Sheet1.Range("IncD.AnyOtherDeductions_Calc").Value
                    End If
                    'End Change
                     
                    If Trim(Sheet1.Range("IncD.TotalChapVIADeductions").Value) <> "" Then
                        DeductUndChapVIA("TotalChapVIADeductions") = Sheet1.Range("IncD.TotalChapVIADeductions").Value
                     Else
                        DeductUndChapVIA("TotalChapVIADeductions") = 0
                    End If
                    
                    
                    
                jsonDictionary.add "DeductUndChapVIA", DeductUndChapVIA
                 
                'Newly updated by Bindu as per DE V3 on 4th Feb 2025
                'If Trim(Sheet1.Range("IncD.TotalIncome").Value) <> "" Then
                If Trim(Sheet1.Range("IncD.TotalIncome_New").Value) <> "" Then
                    jsonDictionary("TotalIncome") = TotalIncome
                 Else
                    jsonDictionary("TotalIncome") = 0
                End If
                 
                If (Not IsEmpty(Others_NOI)) Then
                If (UBound(Others_NOI) > 0) Then
                
                'Commented by Konda on 10-03-2026---V0.5
'                For i = 1 To UBound(Others_NOI)
'                 If (Others_NOI(i) = "Defense Medical Disability Pension") Then
'                            Others_NOI(i) = "DMDP"
'
'                 ElseIf Mid(Others_NOI(i), 1, 4) = "Agri" Then
'                 Others_NOI(i) = "AGRI"
'
'                 ElseIf Mid(Others_NOI(i), 1, 12) = "Sec 10(10BC)" Then
'                 Others_NOI(i) = "10(10BC)"
'                 ElseIf Mid(Others_NOI(i), 1, 11) = "Sec 10(10D)" Then
'                 Others_NOI(i) = "10(10D)"
'
'
'                 ElseIf Mid(Others_NOI(i), 1, 10) = "Sec 10(11)" Then
'                 Others_NOI(i) = "10(11)"
'
'                 ElseIf Mid(Others_NOI(i), 1, 10) = "Sec 10(12)" Then
'                 Others_NOI(i) = "10(12)"
'
'                 ElseIf Mid(Others_NOI(i), 1, 10) = "Sec 10(13)" Then
'                 Others_NOI(i) = "10(13)"
'
'                 ElseIf Mid(Others_NOI(i), 1, 10) = "Sec 10(16)" Then
'                 Others_NOI(i) = "10(16)"
'
'                  ElseIf Mid(Others_NOI(i), 1, 10) = "Sec 10(17)" Then
'                 Others_NOI(i) = "10(17)"
'
'
'                 ElseIf Mid(Others_NOI(i), 1, 11) = "Sec 10(17A)" Then
'                 Others_NOI(i) = "10(17A)"
'
'                 ElseIf Mid(Others_NOI(i), 1, 10) = "Sec 10(18)" Then
'                 Others_NOI(i) = "10(18)"
'
'
'                 ElseIf Mid(Others_NOI(i), 1, 10) = "Sec 10(19)" Then
'                 Others_NOI(i) = "10(19)"
'                 ElseIf Mid(Others_NOI(i), 1, 10) = "Sec 10(26)" Then
'                 Others_NOI(i) = "10(26)"
'                 ElseIf Mid(Others_NOI(i), 1, 13) = "Sec 10(26AAA)" Then
'                 Others_NOI(i) = "10(26AAA)"
'                 ElseIf Mid(Others_NOI(i), 1, 10) = "Sec 10(34)" Then
'                 Others_NOI(i) = "10(34)"
'                 ElseIf Others_NOI(i) = "Any exempt income including LTCG on which tax is not payable" Then
'                 Others_NOI(i) = "CG1L"
'                 ElseIf Mid(Others_NOI(i), 1, 9) = "Any Other" Then
'                 Others_NOI(i) = "OTH"
'
'                 ElseIf Mid(Others_NOI(i), 1, 11) = "Sec 10(12C)" Then
'                 Others_NOI(i) = "10(12C)"
'
'
''                 ElseIf Mid(Others_NOI(i), 1, 4) = "LTCG" Then
''                 Others_NOI(i) = "LTCG"
'                 End If
'
'                Next

'Konda updated on 10-03-2026--V0.5

'Category
            For i = 1 To UBound(Others_NOI)
            
                 If (Others_NOI(i) = "Agricultural & related incomes") Then
                        Others_NOI(i) = "AGRI"
                    ElseIf (Others_NOI(i) = "Compensation/other sums received by government or other approved entities") Then
                        Others_NOI(i) = "GOVC"
                    ElseIf (Others_NOI(i) = "Income from specified Investments") Then
                        Others_NOI(i) = "ISI"
                    ElseIf (Others_NOI(i) = "Specified sums received by armed forces personnel") Then
                        Others_NOI(i) = "SSRA"
                    ElseIf (Others_NOI(i) = "Sums received by Senior Citizens/Minors ") Then
                        Others_NOI(i) = "SRSC"
                    ElseIf (Others_NOI(i) = "Sums received by specified Category of Taxpayers") Then
                        Others_NOI(i) = "SRST"
                    ElseIf (Others_NOI(i) = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana") Then
                        Others_NOI(i) = "SRPC"
                    ElseIf (Others_NOI(i) = "Other Incomes") Then
                        Others_NOI(i) = "OTH"
                    End If
             
                Next
 'SubCategory
          For i = 1 To UBound(Others_NOI1)

                    If (Others_NOI1(i) = "10(1)-Agricultural income(Less than or equal to 5000)") Then
                        Others_NOI1(i) = "10(1)"
                    ElseIf (Others_NOI1(i) = "10(30)-subsidy received from or through the Tea Board") Then
                        Others_NOI1(i) = "10(30)"
                    ElseIf (Others_NOI1(i) = "10(31)-Rubber/Coffee/Tea development accounts/funds") Then
                        Others_NOI1(i) = "10(31)"
                    ElseIf (Others_NOI1(i) = "10(10BB)-payments made under the Bhopal Gas Leak Disaster") Then
                        Others_NOI1(i) = "10(10BB)"
                        'Ankita_23/03/2026=========
                    ElseIf (Others_NOI1(i) = "10(10BC)-Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster") Then
                        Others_NOI1(i) = "10(10BC)"
                    ElseIf (Others_NOI1(i) = "10(17A)-Award instituted by Government") Then
                        Others_NOI1(i) = "10(17A)"
                    ElseIf (Others_NOI1(i) = "10(12AB)-any sum received as lump sum amount as per clause (vi) of paragraph 2 of the notification number FX-1/3/2024-PR") Then
                        Others_NOI1(i) = "10(12AB)"
                    ElseIf (Others_NOI1(i) = "10(15)-Interest on specified securities/investments") Then
                        Others_NOI1(i) = "10(15)"
                        'Ankita_removed as per V0.6
'                    ElseIf (Others_NOI1(i) = "10(23EA)-Contributions received from recognised stock exchanges") Then
'                        Others_NOI1(i) = "10(23EA)"
                    ElseIf (Others_NOI1(i) = "10(23FBB)-income referred to in section 115UB, accruing or arising to, or received by, a unit holder of an investment fund") Then
                        Others_NOI1(i) = "10(23FBB)"
                    ElseIf (Others_NOI1(i) = "10(23FD)Unit holder income from Business Trust (certain parts)") Then
                        Others_NOI1(i) = "10(23FD)"
                    ElseIf (Others_NOI1(i) = "10(35)-Income from specified Mutual Funds ") Then
                        Others_NOI1(i) = "10(35)"
                    ElseIf (Others_NOI1(i) = "10(35A)-distributed income referred to in section 115TA received from a securitisation trust") Then
                        Others_NOI1(i) = "10(35A)"
                    ElseIf (Others_NOI1(i) = "10(12C)-Agniveer Corpus Fund income") Then
                        Others_NOI1(i) = "10(12C)"
'                   ElseIf (Others_NOI1(i) = "10(18)-Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award") Then
                    ElseIf (Others_NOI1(i) = "10(18)-Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award") Then
                        Others_NOI1(i) = "10(18)"
                    ElseIf (Others_NOI1(i) = "10(19)-Armed Forces Family pension in case of death during operational duty") Then
                        Others_NOI1(i) = "10(19)"
                    ElseIf (Others_NOI1(i) = "10(23AA)-Sum received by any person on behalf of any Fund established by the armed forces") Then
                        Others_NOI1(i) = "10(23AA)"
                    ElseIf (Others_NOI1(i) = "Defense Medical Disability Pension") Then
                        Others_NOI1(i) = "DMD"
                    ElseIf (Others_NOI1(i) = "10(32)-Minor child’s income—small exemption") Then
                        Others_NOI1(i) = "10(32)"
                    ElseIf (Others_NOI1(i) = "10(43)-Reverse mortgage—payments to senior citizens") Then
                        Others_NOI1(i) = "10(43)"
                    ElseIf (Others_NOI1(i) = "10(19A)-Annual value of one palace in occupation of ex-ruler") Then
                        Others_NOI1(i) = "10(19A)"
                    ElseIf (Others_NOI1(i) = "10(26)-Any income as referred to in section 10(26)") Then
                        Others_NOI1(i) = "10(26)"
                    ElseIf (Others_NOI1(i) = "10(26AAA)-Any income as referred to in section 10(26AAA)") Then
                        Others_NOI1(i) = "10(26AAA)"
                    ElseIf (Others_NOI1(i) = "10(10D)-Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(10D)") Then
                        Others_NOI1(i) = "10(10D)"
                    ElseIf (Others_NOI1(i) = "10(11)-Statutory Provident Fund received") Then
                        Others_NOI1(i) = "10(11)"
                    ElseIf (Others_NOI1(i) = "10(11A)-Sum received from an account opened under the Sukanya Samriddhi Yojana") Then
                        Others_NOI1(i) = "10(11A)"
                    ElseIf (Others_NOI1(i) = "10(12)-Recognized Provident Fund received") Then
                        Others_NOI1(i) = "10(12)"
                    ElseIf (Others_NOI1(i) = "10(12A)-Any payment from the National Pension System Trust to an assessee") Then
                        Others_NOI1(i) = "10(12A)"
                    ElseIf (Others_NOI1(i) = "10(12AA)-any payment from the National Pension System Trust ") Then
                        Others_NOI1(i) = "10(12AA)"
                    ElseIf (Others_NOI1(i) = "10(12B)-Any payment from the National Pension System Trust to an Central Govt. Employee") Then
                        Others_NOI1(i) = "10(12B)"
                    ElseIf (Others_NOI1(i) = "10(12BA)-partial withdrawal made from the National Pension System") Then
                        Others_NOI1(i) = "10(12BA)"
                    ElseIf (Others_NOI1(i) = "10(13)-Approved superannuation fund received") Then
                        Others_NOI1(i) = "10(13)"
                    ElseIf (Others_NOI1(i) = "10(25)-Sum received by trustees on behalf of approved superannuation, gratuity, or pension funds") Then
                        Others_NOI1(i) = "10(25)"
'                    ElseIf (Others_NOI1(i) = "10(25A)-any income under Employees' State Insurance Fund") Then
'                        Others_NOI1(i) = "10(25A)"
                    ElseIf (Others_NOI1(i) = "10(44)-Income received by any person for, or on behalf of, the New Pension System Trust") Then
                        Others_NOI1(i) = "10(44)"
                    ElseIf (Others_NOI1(i) = "10(2)-Member’s share from HUF") Then
                        Others_NOI1(i) = "10(2)"
                    ElseIf (Others_NOI1(i) = "10(16)-Scholarships for education") Then
                        Others_NOI1(i) = "10(16)"
                    
                    'Malli_AY_2026_27 17/06/2026
                    ElseIf (Others_NOI1(i) = "Income exempt as per CBDT Circular") Then
                        Others_NOI1(i) = "Incmexmptcircular"
                        
                    ElseIf (Others_NOI1(i) = "Income exempt as per CBDT Notification") Then
                        Others_NOI1(i) = "Incmexmptnotification"
                        
                    ElseIf (Others_NOI1(i) = "Receipts not in the nature of income") Then
                        Others_NOI1(i) = "Receiptnotincme"
                        
                    '---------------------------
                    
                    End If

                Next
                
                
'=====================================
                    For i = 1 To UBound(Others_NOI)
        
        'Konda updated on 10-03-2026--V0.5
'                            If Others_NOI(i) <> "" Then
'                                ExemptIncAgriOthUs10Dtl("NatureDesc") = Others_NOI(i)
'                            End If
                            
'                            If Others_NOI1(i) <> "" And Others_NOI1(i) <> "Not Applicable" Then
'                                ExemptIncAgriOthUs10Dtl("OthNatOfInc") = UCase(Others_NOI1(i))
'                            End If

                            If Others_NOI(i) <> "" Then
                                ExemptIncAgriOthUs10Dtl("Category") = Others_NOI(i)
                            End If
                            
                            If Others_NOI1(i) <> "" Then
                                ExemptIncAgriOthUs10Dtl("SubCategory") = Others_NOI1(i)
                            End If
 '================================================================
                            
                            'Malli_AY_2026_27
                            If UCase(Others_NOI1(i)) = UCase("Incmexmptcircular") Or UCase(Others_NOI1(i)) = UCase("Incmexmptnotification") Or UCase(Others_NOI1(i)) = UCase("Receiptnotincme") Then
                                If Others_NOI22(i) <> "" Then
                                   ExemptIncAgriOthUs10Dtl("Description") = Trim(Others_NOI22(i))
                                End If
                            End If
                            '----------------
 
 
                            If Others_Amt(i) <> "" Then
                                ExemptIncAgriOthUs10Dtl("OthAmount") = Others_Amt(i)
                            End If
                        ExemptIncAgriOthUs10Dtls.add ExemptIncAgriOthUs10Dtl
                        Set ExemptIncAgriOthUs10Dtl = Nothing
                        Set ExemptIncAgriOthUs10Dtl = CreateObject("Scripting.Dictionary")
                    Next
    ExemptIncAgriOthUs10.add "ExemptIncAgriOthUs10Dtls", ExemptIncAgriOthUs10Dtls
    If Trim(Sheet1.Range("ExemptIncomeTotal").Value) <> "" Then
        ExemptIncAgriOthUs10("ExemptIncAgriOthUs10Total") = ExemptIncAgriOthUs10Total
     Else
        ExemptIncAgriOthUs10("ExemptIncAgriOthUs10Total") = 0
    End If
    jsonDictionary.add "ExemptIncAgriOthUs10", ExemptIncAgriOthUs10
    End If
End If
Set IncomeDeductions = jsonDictionary
End Function
Function Verification() As Object

subProcCaption = "Verification"
noOfProcessSub = 5
Dim jsonDictionary, Declaration
Dim AssesseeVerPAN, AssesseeVerName, FatherName, Capacity, Place As Variant

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set Declaration = CreateObject("Scripting.Dictionary")

AssesseeVerName = UCase(Sheet3.Range("Ver.AssesseeVerName").Value)
FatherName = UCase(Sheet3.Range("Ver.FatherName").Value)
AssesseeVerPAN = UCase(Sheet3.Range("Ver.PAN").Value)
Capacity = Mid(Sheet3.Range("Ver.capacity").Value, 1, 1)
Place = UCase(Sheet3.Range("Ver.Place").Value)
          
    Declaration("AssesseeVerName") = AssesseeVerName
UpdateProgressBar
    Declaration("FatherName") = FatherName
UpdateProgressBar
    Declaration("AssesseeVerPAN") = AssesseeVerPAN
    jsonDictionary.add "Declaration", Declaration
    jsonDictionary("Capacity") = Capacity
UpdateProgressBar
    jsonDictionary("Place") = Place

Set Verification = jsonDictionary
End Function
Function TaxReturnPreparer() As Object
subProcCaption = "Tax Computation"
noOfProcessSub = 10
Dim jsonDictionary
Dim IdentificationNoOfTRP, ReImbFrmGov, NameOfTRP As Variant

Set jsonDictionary = CreateObject("Scripting.Dictionary")

IdentificationNoOfTRP = UCase(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value)
NameOfTRP = UCase(Sheet3.Range("Sheet2.NameOfTRP").Value)
ReImbFrmGov = UCase(Sheet3.Range("Sheet2.ReImbFrmGov").Value)

    If Trim(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value) <> "" Then
        
        If Trim(Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value) <> "" Then
            jsonDictionary("IdentificationNoOfTRP") = IdentificationNoOfTRP
        End If

        If Trim(Sheet3.Range("Sheet2.NameOfTRP").Value) <> "" Then
            jsonDictionary("NameOfTRP") = NameOfTRP
        End If
             
        If Trim(Sheet3.Range("Sheet2.ReImbFrmGov").Value) <> "" Then
            jsonDictionary("ReImbFrmGov") = UVCase(CDbl(ReImbFrmGov))
        Else
            jsonDictionary("ReImbFrmGov") = 0
        End If
        
        Set TaxReturnPreparer = jsonDictionary
    End If
UpdateProgressBar

End Function
Function TaxComputation() As Object
subProcCaption = "Tax Computation"
noOfProcessSub = 10
Dim i As Long
Dim jsonDictionary, IntrstPay
Dim EducationCess, GrossTaxLiability, section89, Section89A, NetTaxLiability, TotalIntrstPay, IntrstPayUs234A, IntrstPayUs234B, IntrstPayUs234C, LateFilingFee234F, TotTaxPlusIntrstPay As Variant
Dim Rebate87A, TaxPayableOnRebate As Long
'Konda updated 0n 27-02-2026--V0.4
Dim FeeFurnish234I

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set IntrstPay = CreateObject("Scripting.Dictionary")

Rebate87A = UCase(Sheet1.Range("IncD.Rebate87A").Value)
TaxPayableOnRebate = UCase(Sheet1.Range("IncD.TaxPayableOnRebate").Value)
EducationCess = Sheet1.Range("IncD.EducationCess").Value
GrossTaxLiability = Sheet1.Range("IncD.GrossTaxLiability").Value
section89 = Sheet1.Range("IncD.Section89").Value
'Section89A = Sheet1.Range("IncD.Section89A").Value
NetTaxLiability = Round(Sheet1.Range("IncD.NetTaxLiability").Value, 0)
TotalIntrstPay = Round(Sheet1.Range("IncD.TotalIntrstPay").Value, 0)
IntrstPayUs234A = UVCase(Round(Sheet1.Range("IncD.IntrstPayUs234A").Value, 0))
IntrstPayUs234B = UVCase(Round(Sheet1.Range("IncD.IntrstPayUs234B").Value, 0))
IntrstPayUs234C = UVCase(Round(Sheet1.Range("IncD.IntrstPayUs234C").Value, 0))
LateFilingFee234F = UVCase(Round(Sheet1.Range("IncD.IntrstPayUs234F").Value, 0))
'Konda updaed on 27-02-2026--V0.4
FeeFurnish234I = UVCase(Round(Sheet1.Range("IncD.Section234I").Value, 0))
'======================
TotTaxPlusIntrstPay = Round(Sheet1.Range("IncD.TotTaxPlusIntrstPay").Value, 0)

    If TotalTaxPayable <> "" Then
      jsonDictionary("TotalTaxPayable") = TotalTaxPayable
    Else
     jsonDictionary("TotalTaxPayable") = 0
    End If
               
    UpdateProgressBar
    
    If Sheet1.Range("IncD.Rebate87A").Value <> "" Then
        jsonDictionary("Rebate87A") = UVCase(CDbl(Rebate87A))
     Else
        jsonDictionary("Rebate87A") = 0
    End If
                 
    UpdateProgressBar
                 
    If Sheet1.Range("IncD.TaxPayableOnRebate").Value <> "" Then
        jsonDictionary("TaxPayableOnRebate") = TaxPayableOnRebate
     Else
        jsonDictionary("TaxPayableOnRebate") = 0
    End If
            
            
UpdateProgressBar
            
    If Trim(Sheet1.Range("IncD.EducationCess").Value) <> "" Then
        jsonDictionary("EducationCess") = EducationCess
     Else
        jsonDictionary("EducationCess") = 0
    End If
             
UpdateProgressBar
             
    If Trim(Sheet1.Range("IncD.GrossTaxLiability").Value) <> "" Then
        jsonDictionary("GrossTaxLiability") = GrossTaxLiability
     Else
        jsonDictionary("GrossTaxLiability") = 0
    End If
             
UpdateProgressBar
             
    If Trim(Sheet1.Range("IncD.Section89").Value) <> "" Then
        jsonDictionary("Section89") = section89
     Else
        jsonDictionary("Section89") = 0
    End If
'    If Trim(Sheet1.Range("IncD.Section89A").Value) <> "" Then
'        jsonDictionary("Section89A") = Section89A
'     Else
'        jsonDictionary("Section89A") = 0
'    End If
    If Trim(Sheet1.Range("IncD.NetTaxLiability").Value) <> "" Then
        jsonDictionary("NetTaxLiability") = NetTaxLiability
     Else
        jsonDictionary("NetTaxLiability") = 0
    End If
                 
UpdateProgressBar
                 
    If Trim(Sheet1.Range("IncD.TotalIntrstPay").Value) <> "" Then
        jsonDictionary("TotalIntrstPay") = TotalIntrstPay
     Else
        jsonDictionary("TotalIntrstPay") = 0
    End If
UpdateProgressBar
             
    IntrstPay("IntrstPayUs234A") = IntrstPayUs234A
    IntrstPay("IntrstPayUs234B") = IntrstPayUs234B
    IntrstPay("IntrstPayUs234C") = IntrstPayUs234C
    IntrstPay("LateFilingFee234F") = LateFilingFee234F
'Konda updated 27-02-2026--V0.4
    IntrstPay("FeeFurnish234I") = FeeFurnish234I
'=========
    jsonDictionary.add "IntrstPay", IntrstPay
            
             
    If Trim(Sheet1.Range("IncD.TotTaxPlusIntrstPay").Value) <> "" Then
        jsonDictionary("TotTaxPlusIntrstPay") = TotTaxPlusIntrstPay
     Else
        jsonDictionary("TotTaxPlusIntrstPay") = 0
    End If
UpdateProgressBar
Set TaxComputation = jsonDictionary
End Function
Function LTCG112A_New() As Object
'Konda added AY 2025-26 on 04-02-2025
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Dim TotSaleCnsdrn, TotCstAcqisn, LongCap112A As Variant
 
'Dim LTCG112A
'Set TotSaleCnsdrn = CreateObject("Scripting.Dictionary")
'Set TotCstAcqisn = CreateObject("Scripting.Dictionary")
'
'Set LTCG112A_New = CreateObject("Scripting.Dictionary")
 
 
If Trim(Sheet1.Range("IncD.Sale_LTCG").Value) <> "" Then
                jsonDictionary("TotSaleCnsdrn") = Sheet1.Range("IncD.Sale_LTCG").Value
            Else
                jsonDictionary("TotSaleCnsdrn") = 0
            End If
            If Trim(Sheet1.Range("IncD.Cost_LTCG").Value) <> "" Then
                jsonDictionary("TotCstAcqisn") = Sheet1.Range("IncD.Cost_LTCG").Value
            Else
                jsonDictionary("TotCstAcqisn") = 0
            End If
            If Trim(Sheet1.Range("IncD.CG_LTCG").Value) <> "" Then
                jsonDictionary("LongCap112A") = Sheet1.Range("IncD.CG_LTCG").Value
            Else
                jsonDictionary("LongCap112A") = 0
            End If
 
'jsonDictionary.add "LTCG112A", LTCG112A_New
Set LTCG112A_New = jsonDictionary
 
End Function
Function TDSonSalaries() As Object
Dim jsonDictionary, TDSonSal, EmployerOrDeductorOrCollectDetl
Dim TotalTDSonSalaries, TDSonSalary, IncChrgSal, TotalTDSSal As Variant
Set TDSonSalary = New Collection

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set TDSonSal = CreateObject("Scripting.Dictionary")
Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")

TotalTDSonSalaries = Sheet2.Range("TDSal.Sum").Value
If Not IsEmpty(TAN_TDS) Then
If UBound(TAN_TDS) > 0 Then
Dim i As Long
subProcCaption = "TDS1"
noOfProcessSub = UBound(TAN_TDS)
 
  For i = 1 To UBound(TAN_TDS)
    If TAN_TDS(i) <> "" Then
        EmployerOrDeductorOrCollectDetl("TAN") = UCase(TAN_TDS(i))
    End If
 
    If TsEmpName(i) <> "" Then
        EmployerOrDeductorOrCollectDetl("EmployerOrDeductorOrCollecterName") = UCase(TsEmpName(i))
    End If
    TDSonSal.add "EmployerOrDeductorOrCollectDetl", EmployerOrDeductorOrCollectDetl
    Set EmployerOrDeductorOrCollectDetl = Nothing
    Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")
    
    If Tot_chrg(i) <> "" Then
        TDSonSal("IncChrgSal") = UVCase(CDbl(Tot_chrg(i)))
    Else
        TDSonSal("IncChrgSal") = 0
    End If
    If Amount4_I(i) <> "" Then
        TDSonSal("TotalTDSSal") = UVCase(CDbl(Amount4_I(i)))
    Else
        TDSonSal("TotalTDSSal") = 0
    End If
    TDSonSalary.add TDSonSal
    Set TDSonSal = Nothing
    Set TDSonSal = CreateObject("Scripting.Dictionary")
UpdateProgressBar
    Next
    jsonDictionary.add "TDSonSalary", TDSonSalary
 End If
End If
jsonDictionary("TotalTDSonSalaries") = TotalTDSonSalaries

Set TDSonSalaries = jsonDictionary
End Function
Function TDSonOthThanSals() As Object
Dim jsonDictionary
Dim TotalTDSonOthThanSals, TDSonOthThanSal As Variant
TotalTDSonOthThanSals = Sheet2.Range("TDSoth.Sum").Value

Set jsonDictionary = CreateObject("Scripting.Dictionary")

    If Not IsEmpty(Tan2_TDS) Then
    If UBound(Tan2_TDS) > 0 Then
        subProcCaption = "TDS2"
        noOfProcessSub = UBound(Tan2_TDS)
        Dim i As Variant
        
        Dim TDSonOthThan, EmployerOrDeductorOrCollectDetl
        Set TDSonOthThanSal = New Collection
        Set TDSonOthThan = CreateObject("Scripting.Dictionary")
        Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")

        
        For i = 1 To UBound(Tan2_TDS)
        
        If Tan2_TDS(i) <> "" Then
            EmployerOrDeductorOrCollectDetl("TAN") = UCase(Tan2_TDS(i))
        End If
                
        If DeductName2_TDS(i) <> "" Then
            EmployerOrDeductorOrCollectDetl("EmployerOrDeductorOrCollecterName") = UCase(DeductName2_TDS(i))
        End If
        TDSonOthThan.add "EmployerOrDeductorOrCollectDetl", EmployerOrDeductorOrCollectDetl
        Set EmployerOrDeductorOrCollectDetl = Nothing
        Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")

        If TaxDeducted_TDS(i) <> "" Then
            TDSonOthThan("AmtForTaxDeduct") = UVCase(CDbl(TaxDeducted_TDS(i)))
        End If
'Konda Commeted as AY-2025-26 V0.4 on 12-02-2025
        
' 'Konda Added on 04-02-2025
'     If SectionTDS_TDS2_1(i) <> "" Then
'
'        If SectionTDS_TDS2_1(i) = "192A-TDS on PF withdrawal" Then
'               TDSonOthThan("SecTDSDeducted") = "192A"
'            ElseIf SectionTDS_TDS2_1(i) = "194 -Dividends" Then
'               TDSonOthThan("SecTDSDeducted") = "194"
'            ElseIf SectionTDS_TDS2_1(i) = "193 -Interest on Securities" Then
'               TDSonOthThan("SecTDSDeducted") = "193"
''            ElseIf SectionTDS_TDS2_1(i) = "192-Salary-Payment to Government employees other than Indian Government employees" Then
''               TDSonOthThan("SecTDSDeducted") = "92A"
''            ElseIf SectionTDS_TDS2_1(i) = "192-Salary-Payment to employees other than Government employees" Then
''               TDSonOthThan("SecTDSDeducted") = "92B"
''            ElseIf SectionTDS_TDS2_1(i) = "192-Salary-Payment to Indian Government employees" Then
''               TDSonOthThan("SecTDSDeducted") = "92C"
'            ElseIf SectionTDS_TDS2_1(i) = "194A-Interest other than 'Interest on securities'" Then
'               TDSonOthThan("SecTDSDeducted") = "94A"
'            ElseIf SectionTDS_TDS2_1(i) = "194DA-Payment in respect of life insurance policy" Then
'               TDSonOthThan("SecTDSDeducted") = "4DA"
'            ElseIf SectionTDS_TDS2_1(i) = "194EE-Payments in respect of deposits under National Savings" Then
'               TDSonOthThan("SecTDSDeducted") = "4EE"
'            ElseIf SectionTDS_TDS2_1(i) = "194IB-Payment of rent by certain individuals or Hindu undivided" Then
'               TDSonOthThan("SecTDSDeducted") = "4IB"
'            ElseIf SectionTDS_TDS2_1(i) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
'               TDSonOthThan("SecTDSDeducted") = "94K"
'            ElseIf SectionTDS_TDS2_1(i) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
'               TDSonOthThan("SecTDSDeducted") = "4BA1"
'            ElseIf SectionTDS_TDS2_1(i) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
'               TDSonOthThan("SecTDSDeducted") = "4BA2"
'            ElseIf SectionTDS_TDS2_1(i) = "194LBB-Income in respect of units of investment fund" Then
'               TDSonOthThan("SecTDSDeducted") = "LBB"
'            ElseIf SectionTDS_TDS2_1(i) = "194P-Deduction of tax in case of specified senior citizen" Then
'               TDSonOthThan("SecTDSDeducted") = "94P"
'        End If
'
'        End If
'
''end konda
'Konda Commeted as AY-2025-26 V0.4 on 12-02-2025
'================================================================================================================================
    'Konda---------AY-2025-26 V0.6.2 on 22-04-2025
              Dim TDSSection_TDS2
            If Section_TDS2(i) <> "" Then
                If Section_TDS2(i) = "192-Salary-Payment to Government employees other than Indian Government employees" Then
                TDSSection_TDS2 = "92A"
                
                ElseIf Section_TDS2(i) = "192-Salary-Payment to employees other than Government employees" Then
                TDSSection_TDS2 = "92B"
                
                ElseIf Section_TDS2(i) = "192-Salary-Payment to Indian Government employees" Then
                TDSSection_TDS2 = "92C"
                
                ElseIf Section_TDS2(i) = "192A-TDS on PF withdrawal" Then
                TDSSection_TDS2 = "192A"
                
                ElseIf Section_TDS2(i) = "193-Interest on Securities" Then
                TDSSection_TDS2 = "193"
                
                ElseIf Section_TDS2(i) = "194-Dividends" Then
                TDSSection_TDS2 = "194"
                
                ElseIf Section_TDS2(i) = "194A-Interest other than 'Interest on securities'" Then
                TDSSection_TDS2 = "94A"
                
                ElseIf Section_TDS2(i) = "194B-Winning from lottery or crossword puzzle" Then
                TDSSection_TDS2 = "94B"
                
                ElseIf Section_TDS2(i) = "194BA-Winnings from online games" Then
                TDSSection_TDS2 = "94BA"
                
                ElseIf Section_TDS2(i) = "194BB-Winning from horse race" Then
                TDSSection_TDS2 = "4BB"
                
                ElseIf Section_TDS2(i) = "194C-Payments to contractors and sub-contractors" Then
                TDSSection_TDS2 = "94C"
                
                ElseIf Section_TDS2(i) = "194D-Insurance commission" Then
                TDSSection_TDS2 = "94D"
                
                ElseIf Section_TDS2(i) = "194DA-Payment in respect of life insurance policy" Then
                TDSSection_TDS2 = "4DA"
                
                ElseIf Section_TDS2(i) = "194E-Payments to non-resident sportsmen or sports associations" Then
                TDSSection_TDS2 = "94E"
                
                ElseIf Section_TDS2(i) = "194EE-Payments in respect of deposits under National Savings" Then
                TDSSection_TDS2 = "4EE"
                
                ElseIf Section_TDS2(i) = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India" Then
                TDSSection_TDS2 = "4F"
                
                ElseIf Section_TDS2(i) = "194G-Commission, price, etc. on sale of lottery tickets" Then
                TDSSection_TDS2 = "4G"
                
                ElseIf Section_TDS2(i) = "194H-Commission or brokerage" Then
                TDSSection_TDS2 = "4H"
                
                ElseIf Section_TDS2(i) = "194I(a)-Rent on hiring of plant and machinery" Then
                TDSSection_TDS2 = "4-IA"
                
                ElseIf Section_TDS2(i) = "194I(b)-Rent on other than plant and machinery" Then
                TDSSection_TDS2 = "4-IB"
                
                ElseIf Section_TDS2(i) = "194IA-TDS on Sale of immovable property" Then
                TDSSection_TDS2 = "4IA"
                
                ElseIf Section_TDS2(i) = "194IB-Payment of rent by certain individuals or Hindu undivided" Then
                TDSSection_TDS2 = "4IB"
                
                ElseIf Section_TDS2(i) = "194IC-Payment under specified agreement" Then
                TDSSection_TDS2 = "4IC"
                
                ElseIf Section_TDS2(i) = "194J(a)-Fees for technical services" Then
                TDSSection_TDS2 = "94J-A"
                
                ElseIf Section_TDS2(i) = "194J(b)-Fees for professional  services or royalty etc" Then
                TDSSection_TDS2 = "94J-B"
                
                ElseIf Section_TDS2(i) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
                TDSSection_TDS2 = "94K"
                
                ElseIf Section_TDS2(i) = "194LA-Payment of compensation on acquisition of certain immovable" Then
                TDSSection_TDS2 = "4LA"
                
                ElseIf Section_TDS2(i) = "194LB-Income by way of Interest from Infrastructure Debt fund" Then
                TDSSection_TDS2 = "4LB"
                
                ElseIf Section_TDS2(i) = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC" Then
                TDSSection_TDS2 = "4LC1"
                
                ElseIf Section_TDS2(i) = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC" Then
                TDSSection_TDS2 = "4LC2"
                
                ElseIf Section_TDS2(i) = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC" Then
                TDSSection_TDS2 = "4LC3"
                
                ElseIf Section_TDS2(i) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
                TDSSection_TDS2 = "4BA1"
                
                ElseIf Section_TDS2(i) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
                TDSSection_TDS2 = "4BA2"
                
                ElseIf Section_TDS2(i) = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR" Then
                TDSSection_TDS2 = "LBA1"
                
                ElseIf Section_TDS2(i) = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR" Then
                TDSSection_TDS2 = "LBA2"
                
                ElseIf Section_TDS2(i) = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR" Then
                TDSSection_TDS2 = "LBA3"
                
                ElseIf Section_TDS2(i) = "194LBB-Income in respect of units of investment fund" Then
                TDSSection_TDS2 = "LBB"
                
                ElseIf Section_TDS2(i) = "194R-Benefits or perquisites of business or profession" Then
                TDSSection_TDS2 = "94R"
                
                ElseIf Section_TDS2(i) = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons" Then
                TDSSection_TDS2 = "94S"
                
                ElseIf Section_TDS2(i) = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released" Then
                TDSSection_TDS2 = "94B-P"
                
                ElseIf Section_TDS2(i) = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released" Then
                TDSSection_TDS2 = "94R-P"
                
                ElseIf Section_TDS2(i) = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released" Then
                TDSSection_TDS2 = "94S-P"
                
                ElseIf Section_TDS2(i) = "194LBC-Income in respect of investment in securitization trust" Then
                TDSSection_TDS2 = "LBC"
                
                ElseIf Section_TDS2(i) = "194LD-TDS on interest on bonds / government securities" Then
                TDSSection_TDS2 = "4LD"
                
                ElseIf Section_TDS2(i) = "194M-Payment of certain sums by certain individuals or HUF" Then
                TDSSection_TDS2 = "94M"
                
                ElseIf Section_TDS2(i) = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso" Then
                TDSSection_TDS2 = "94N"
                
                ElseIf Section_TDS2(i) = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties" Then
                TDSSection_TDS2 = "94N-F"
                
                ElseIf Section_TDS2(i) = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso" Then
                TDSSection_TDS2 = "94N-C"
                
                ElseIf Section_TDS2(i) = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies" Then
                TDSSection_TDS2 = "94N-FT"
                
                ElseIf Section_TDS2(i) = "194O-Payment of certain sums by e-commerce operator to e-commerce participant." Then
                TDSSection_TDS2 = "94O"
                
                ElseIf Section_TDS2(i) = "194P-Deduction of tax in case of specified senior citizen" Then
                TDSSection_TDS2 = "94P"
                
                ElseIf Section_TDS2(i) = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods" Then
                TDSSection_TDS2 = "94Q"
                
                ElseIf Section_TDS2(i) = "195-Other sums payable to a non-resident" Then
                TDSSection_TDS2 = "195"
                
                ElseIf Section_TDS2(i) = "196A-Income in respect of units of non-residents" Then
                TDSSection_TDS2 = "96A"
                
                ElseIf Section_TDS2(i) = "196B-Payments in respect of units to an offshore fund" Then
                TDSSection_TDS2 = "96B"
                
                ElseIf Section_TDS2(i) = "196C-Income from foreign currency bonds or shares of Indian" Then
                TDSSection_TDS2 = "96C"
                
                ElseIf Section_TDS2(i) = "196D-Income of foreign institutional investors from securities" Then
                TDSSection_TDS2 = "96D"
                
                ElseIf Section_TDS2(i) = "196D(1A)-Income of specified fund from securities" Then
                TDSSection_TDS2 = "96DA"
                
                ElseIf Section_TDS2(i) = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released" Then
                TDSSection_TDS2 = "94BA-P"
                
                Else
                TDSSection_TDS2 = ""
                
                 
                End If
                If TDSSection_TDS2 <> "" Then
                   TDSonOthThan("TDSSection") = UCase(TDSSection_TDS2)
                End If
     End If
                
                
                
                '-------------------

        'TAX DETAILS-E2
'        If Year2_TDS(i) <> "" Then
'            TDSonOthThan("DeductedYr") = UCase(Year2_TDS(i))
'        End If
        
        'TAX DETAILS-C2 2024-25 Bindu
        If Year2_TDS(i) <> "" Then
            If Year2_TDS(i) = "2008-09" Then
               TDSonOthThan("DeductedYr") = "2008"
            ElseIf Year2_TDS(i) = "2009-10" Then
               TDSonOthThan("DeductedYr") = "2009"
            ElseIf Year2_TDS(i) = "2010-11" Then
               TDSonOthThan("DeductedYr") = "2010"
            ElseIf Year2_TDS(i) = "2011-12" Then
               TDSonOthThan("DeductedYr") = "2011"
            ElseIf Year2_TDS(i) = "2012-13" Then
               TDSonOthThan("DeductedYr") = "2012"
            ElseIf Year2_TDS(i) = "2013-14" Then
               TDSonOthThan("DeductedYr") = "2013"
            ElseIf Year2_TDS(i) = "2014-15" Then
               TDSonOthThan("DeductedYr") = "2014"
            ElseIf Year2_TDS(i) = "2015-16" Then
               TDSonOthThan("DeductedYr") = "2015"
            ElseIf Year2_TDS(i) = "2016-17" Then
               TDSonOthThan("DeductedYr") = "2016"
            ElseIf Year2_TDS(i) = "2017-18" Then
               TDSonOthThan("DeductedYr") = "2017"
            ElseIf Year2_TDS(i) = "2018-19" Then
               TDSonOthThan("DeductedYr") = "2018"
            ElseIf Year2_TDS(i) = "2019-20" Then
               TDSonOthThan("DeductedYr") = "2019"
            ElseIf Year2_TDS(i) = "2020-21" Then
               TDSonOthThan("DeductedYr") = "2020"
            ElseIf Year2_TDS(i) = "2021-22" Then
               TDSonOthThan("DeductedYr") = "2021"
            ElseIf Year2_TDS(i) = "2022-23" Then
               TDSonOthThan("DeductedYr") = "2022"
            ElseIf Year2_TDS(i) = "2023-24" Then
               TDSonOthThan("DeductedYr") = "2023"
            'added by Chetan C M
            ElseIf Year2_TDS(i) = "2024-25" Then
               TDSonOthThan("DeductedYr") = "2024"
'            New Schema updated by Konda as AY-2026-27 on 26-12-2025
            ElseIf Year2_TDS(i) = "2025-26" Then
               TDSonOthThan("DeductedYr") = "2025"
            End If
            
        End If
        
        
        
                 
        If Range("TDSoth.TotTDSOnAmtPaid").Cells.item(i).Value <> "" Then
            TDSonOthThan("TotTDSOnAmtPaid") = UVCase(CDbl(Range("TDSoth.TotTDSOnAmtPaid").item(i).Value))
        Else
            TDSonOthThan("TotTDSOnAmtPaid") = 0
        End If

        If Range("TDSoth.6income").Cells.item(i).Value <> "" Then
            TDSonOthThan("ClaimOutOfTotTDSOnAmtPaid") = UVCase(CDbl(Range("TDSoth.6income").item(i).Value))
        Else
            TDSonOthThan("ClaimOutOfTotTDSOnAmtPaid") = 0
        End If
        TDSonOthThanSal.add TDSonOthThan
        Set TDSonOthThan = Nothing
        Set TDSonOthThan = CreateObject("Scripting.Dictionary")
UpdateProgressBar
Next
End If
Else
End If
jsonDictionary.add "TDSonOthThanSal", TDSonOthThanSal
jsonDictionary("TotalTDSonOthThanSals") = TotalTDSonOthThanSals
Set TDSonOthThanSals = jsonDictionary
End Function
Function ScheduleTDS3Dtls() As Object
Dim jsonDictionary
Dim TotalTDS3Details, TDS3Details As Variant
TotalTDS3Details = Sheet2.Range("TDS26QB.Sum").Value
Set jsonDictionary = CreateObject("Scripting.Dictionary")

    If Not IsEmpty(PAN_TDS) Then
    If UBound(PAN_TDS) > 0 Then
        subProcCaption = "TDS3"
        noOfProcessSub = UBound(PAN_TDS)
        Dim i As Variant
        
        Dim TDS3
        Set TDS3Details = New Collection
        
        Set TDS3 = CreateObject("Scripting.Dictionary")
               
            For i = 1 To UBound(PAN_TDS)
                If PAN_TDS(i) <> "" Then
                    TDS3("PANofTenant") = UCase(PAN_TDS(i))
                End If
                
                If Tenant_Aadhar_TDS(i) <> "" Then
                    TDS3("AadhaarofTenant") = UCase(Tenant_Aadhar_TDS(i))
                End If
            
                If DeductName2_TDS3(i) <> "" Then
                    TDS3("NameOfTenant") = UCase(DeductName2_TDS3(i))
                End If
    'Konda--------------
                Dim TDSSection_TDS3
            If Section_TDS3(i) <> "" Then
                If Section_TDS3(i) = "192-Salary-Payment to Government employees other than Indian Government employees" Then
                TDSSection_TDS3 = "92A"
                
                ElseIf Section_TDS3(i) = "192-Salary-Payment to employees other than Government employees" Then
                TDSSection_TDS3 = "92B"
                
                ElseIf Section_TDS3(i) = "192-Salary-Payment to Indian Government employees" Then
                TDSSection_TDS3 = "92C"
                
                ElseIf Section_TDS3(i) = "192A-TDS on PF withdrawal" Then
                TDSSection_TDS3 = "192A"
                
                ElseIf Section_TDS3(i) = "193-Interest on Securities" Then
                TDSSection_TDS3 = "193"
                
                ElseIf Section_TDS3(i) = "194-Dividends" Then
                TDSSection_TDS3 = "194"
                
                ElseIf Section_TDS3(i) = "194A-Interest other than 'Interest on securities'" Then
                TDSSection_TDS3 = "94A"
                
                ElseIf Section_TDS3(i) = "194B-Winning from lottery or crossword puzzle" Then
                TDSSection_TDS3 = "94B"
                
                ElseIf Section_TDS3(i) = "194BA-Winnings from online games" Then
                TDSSection_TDS3 = "94BA"
                
                ElseIf Section_TDS3(i) = "194BB-Winning from horse race" Then
                TDSSection_TDS3 = "4BB"
                
                ElseIf Section_TDS3(i) = "194C-Payments to contractors and sub-contractors" Then
                TDSSection_TDS3 = "94C"
                
                ElseIf Section_TDS3(i) = "194D-Insurance commission" Then
                TDSSection_TDS3 = "94D"
                
                ElseIf Section_TDS3(i) = "194DA-Payment in respect of life insurance policy" Then
                TDSSection_TDS3 = "4DA"
                
                ElseIf Section_TDS3(i) = "194E-Payments to non-resident sportsmen or sports associations" Then
                TDSSection_TDS3 = "94E"
                
                ElseIf Section_TDS3(i) = "194EE-Payments in respect of deposits under National Savings" Then
                TDSSection_TDS3 = "4EE"
                
                ElseIf Section_TDS3(i) = "194F-Payments on account of repurchase of units by Mutual Fund or Unit Trust of India" Then
                TDSSection_TDS3 = "4F"
                
                ElseIf Section_TDS3(i) = "194G-Commission, price, etc. on sale of lottery tickets" Then
                TDSSection_TDS3 = "4G"
                
                ElseIf Section_TDS3(i) = "194H-Commission or brokerage" Then
                TDSSection_TDS3 = "4H"
                
                ElseIf Section_TDS3(i) = "194I(a)-Rent on hiring of plant and machinery" Then
                TDSSection_TDS3 = "4-IA"
                
                ElseIf Section_TDS3(i) = "194I(b)-Rent on other than plant and machinery" Then
                TDSSection_TDS3 = "4-IB"
                
                ElseIf Section_TDS3(i) = "194IA-TDS on Sale of immovable property" Then
                TDSSection_TDS3 = "4IA"
                
                ElseIf Section_TDS3(i) = "194IB-Payment of rent by certain individuals or Hindu undivided" Then
                TDSSection_TDS3 = "4IB"
                
                ElseIf Section_TDS3(i) = "194IC-Payment under specified agreement" Then
                TDSSection_TDS3 = "4IC"
                
                ElseIf Section_TDS3(i) = "194J(a)-Fees for technical services" Then
                TDSSection_TDS3 = "94J-A"
                
                ElseIf Section_TDS3(i) = "194J(b)-Fees for professional  services or royalty etc" Then
                TDSSection_TDS3 = "94J-B"
                
                ElseIf Section_TDS3(i) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
                TDSSection_TDS3 = "94K"
                
                ElseIf Section_TDS3(i) = "194LA-Payment of compensation on acquisition of certain immovable" Then
                TDSSection_TDS3 = "4LA"
                
                ElseIf Section_TDS3(i) = "194LB-Income by way of Interest from Infrastructure Debt fund" Then
                TDSSection_TDS3 = "4LB"
                
                ElseIf Section_TDS3(i) = "194LC-194LC (2)(i) and (ia) Income under clause (i) and (ia) of sub-section (2) of section 194LC" Then
                TDSSection_TDS3 = "4LC1"
                
                ElseIf Section_TDS3(i) = "194LC-194LC (2)(ib) Income under clause (ib) of sub-section (2) of section 194LC" Then
                TDSSection_TDS3 = "4LC2"
                
                ElseIf Section_TDS3(i) = "194LC-194LC (2)(ic) Income under clause (ic) of sub-section (2) of section 194LC" Then
                TDSSection_TDS3 = "4LC3"
                
                ElseIf Section_TDS3(i) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
                TDSSection_TDS3 = "4BA1"
                
                ElseIf Section_TDS3(i) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
                TDSSection_TDS3 = "4BA2"
                
                ElseIf Section_TDS3(i) = "194LBA(a)-194LBA(a) income referred to in section 10(23FC)(a) from units of a business trust-NR" Then
                TDSSection_TDS3 = "LBA1"
                
                ElseIf Section_TDS3(i) = "194LBA(b)-194LBA(b) Income referred to in section 10(23FC)(b) from units of a business trust-NR" Then
                TDSSection_TDS3 = "LBA2"
                
                ElseIf Section_TDS3(i) = "194LBA(c)-194LBA(c) Income referred to in section 10(23FCA) from units of a business trust-NR" Then
                TDSSection_TDS3 = "LBA3"
                
                ElseIf Section_TDS3(i) = "194LBB-Income in respect of units of investment fund" Then
                TDSSection_TDS3 = "LBB"
                
                ElseIf Section_TDS3(i) = "194R-Benefits or perquisites of business or profession" Then
                TDSSection_TDS3 = "94R"
                
                ElseIf Section_TDS3(i) = "194S-Payment of consideration for transfer of virtual digital asset by persons other than specified persons" Then
                TDSSection_TDS3 = "94S"
                
                ElseIf Section_TDS3(i) = "Proviso to section 194B-Winnings from lotteries and crossword puzzles where consideration is made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such winnings are released" Then
                TDSSection_TDS3 = "94B-P"
                
                ElseIf Section_TDS3(i) = "First Proviso to sub-section(1) of section 194R-Benefits or perquisites of business or profession where such benefit is provided in kind or where part in cash is not sufficient to meet tax liability and tax required to be deducted is paid before such benefit is released" Then
                TDSSection_TDS3 = "94R-P"
                
                ElseIf Section_TDS3(i) = "Proviso to sub- section(1) of section 194S-Payment for transfer of virtual digital asset where payment is in kind or in exchange of another virtual digital asset and tax required to be deducted is paid before such payment is released" Then
                TDSSection_TDS3 = "94S-P"
                
                ElseIf Section_TDS3(i) = "194LBC-Income in respect of investment in securitization trust" Then
                TDSSection_TDS3 = "LBC"
                
                ElseIf Section_TDS3(i) = "194LD-TDS on interest on bonds / government securities" Then
                TDSSection_TDS3 = "4LD"
                
                ElseIf Section_TDS3(i) = "194M-Payment of certain sums by certain individuals or HUF" Then
                TDSSection_TDS3 = "94M"
                
                ElseIf Section_TDS3(i) = "194N-Payment of certain amounts in cash other than cases covered by first proviso or third proviso" Then
                TDSSection_TDS3 = "94N"
                
                ElseIf Section_TDS3(i) = "194N -First Proviso Payment of certain amounts in cash to non-filers except in case of co-operativesocieties" Then
                TDSSection_TDS3 = "94N-F"
                
                ElseIf Section_TDS3(i) = "194N -Third Proviso Payment of certain amounts in cash to co-operative societies not covered by first proviso" Then
                TDSSection_TDS3 = "94N-C"
                
                ElseIf Section_TDS3(i) = "194N-First Proviso read with Third Proviso Payment of certain amount in cash to non-filers being co-operative societies" Then
                TDSSection_TDS3 = "94N-FT"
                
                ElseIf Section_TDS3(i) = "194O-Payment of certain sums by e-commerce operator to e-commerce participant." Then
                TDSSection_TDS3 = "94O"
                
                ElseIf Section_TDS3(i) = "194P-Deduction of tax in case of specified senior citizen" Then
                TDSSection_TDS3 = "94P"
                
                ElseIf Section_TDS3(i) = "194Q-Deduction of tax at source on payment of certain sum for purchase of goods" Then
                TDSSection_TDS3 = "94Q"
                
                ElseIf Section_TDS3(i) = "195-Other sums payable to a non-resident" Then
                TDSSection_TDS3 = "195"
                
                ElseIf Section_TDS3(i) = "196A-Income in respect of units of non-residents" Then
                TDSSection_TDS3 = "96A"
                
                ElseIf Section_TDS3(i) = "196B-Payments in respect of units to an offshore fund" Then
                TDSSection_TDS3 = "96B"
                
                ElseIf Section_TDS3(i) = "196C-Income from foreign currency bonds or shares of Indian" Then
                TDSSection_TDS3 = "96C"
                
                ElseIf Section_TDS3(i) = "196D-Income of foreign institutional investors from securities" Then
                TDSSection_TDS3 = "96D"
                
                ElseIf Section_TDS3(i) = "196D(1A)-Income of specified fund from securities" Then
                TDSSection_TDS3 = "96DA"
                
                ElseIf Section_TDS3(i) = "194BA(2)-Sub-section (2) of section 194BA Net Winnings from online games where the net winnings are made in kind or cash is not sufficient to meet the tax liability and tax has been paid before such net winnings are released" Then
                TDSSection_TDS3 = "94BA-P"
                
                Else
                TDSSection_TDS3 = ""
                
                 
                End If
                If TDSSection_TDS3 <> "" Then
                   TDS3("TDSSection") = UCase(TDSSection_TDS3)
                End If
     End If
                

      '------------------------------
                If TaxDeducted_TDS3(i) <> "" Then
                    TDS3("GrsRcptToTaxDeduct") = UVCase(CDbl(TaxDeducted_TDS3(i)))
                End If
'Konda Commeted as AY-2025-26 V0.4 on 12-02-2025-------------------------------
' 'Konda Added on 04-02-2025
'     If SectionTDS_TDS2_2(i) <> "" Then
'
'        If SectionTDS_TDS2_2(i) = "192A-TDS on PF withdrawal" Then
'               TDS3("SecTDSDeducted") = "192A"
'            ElseIf SectionTDS_TDS2_2(i) = "194 -Dividends" Then
'               TDS3("SecTDSDeducted") = "194"
'            ElseIf SectionTDS_TDS2_2(i) = "193 -Interest on Securities" Then
'               TDS3("SecTDSDeducted") = "193"
''            ElseIf SectionTDS_TDS2_2(i) = "192-Salary-Payment to Government employees other than Indian Government employees" Then
''               TDS3("SecTDSDeducted") = "92A"
''            ElseIf SectionTDS_TDS2_2(i) = "192-Salary-Payment to employees other than Government employees" Then
''               TDS3("SecTDSDeducted") = "92B"
''            ElseIf SectionTDS_TDS2_2(i) = "192-Salary-Payment to Indian Government employees" Then
''               TDS3("SecTDSDeducted") = "92C"
'            ElseIf SectionTDS_TDS2_2(i) = "194A-Interest other than 'Interest on securities'" Then
'               TDS3("SecTDSDeducted") = "94A"
'            ElseIf SectionTDS_TDS2_2(i) = "194DA-Payment in respect of life insurance policy" Then
'               TDS3("SecTDSDeducted") = "4DA"
'            ElseIf SectionTDS_TDS2_2(i) = "194EE-Payments in respect of deposits under National Savings" Then
'               TDS3("SecTDSDeducted") = "4EE"
'            ElseIf SectionTDS_TDS2_2(i) = "194IB-Payment of rent by certain individuals or Hindu undivided" Then
'               TDS3("SecTDSDeducted") = "4IB"
'            ElseIf SectionTDS_TDS2_2(i) = "194K-Income payable to a resident assessee in respect of units of a specified mutual fund or of the units of the Unit Trust of India" Then
'               TDS3("SecTDSDeducted") = "94K"
'            ElseIf SectionTDS_TDS2_2(i) = "194LBA(a)-Certain income in the form of interest from units of a business trust to a resident unit holder" Then
'               TDS3("SecTDSDeducted") = "4BA1"
'            ElseIf SectionTDS_TDS2_2(i) = "194LBA(b)-Certain income in the form of dividend from units of a business trust to a resident unit holder" Then
'               TDS3("SecTDSDeducted") = "4BA2"
'            ElseIf SectionTDS_TDS2_2(i) = "194LBB-Income in respect of units of investment fund" Then
'               TDS3("SecTDSDeducted") = "LBB"
'            ElseIf SectionTDS_TDS2_2(i) = "194P-Deduction of tax in case of specified senior citizen" Then
'               TDS3("SecTDSDeducted") = "94P"
'        End If
'
'        End If
' 'end konda
'End Commeted as AY-2025-26 V0.4 on 12-02-2025-------------------------------------
                'TAX DETAILS-E3
'                If Year2_TDS3(i) <> "" Then
'                    TDS3("DeductedYr") = UCase(Year2_TDS3(i))
'                End If
                'TAX DETAILS-C3
                
                
                If Year2_TDS3(i) <> "" Then
                    If Year2_TDS3(i) = "2017-18" Then
                       TDS3("DeductedYr") = "2017"
                    ElseIf Year2_TDS3(i) = "2018-19" Then
                       TDS3("DeductedYr") = "2018"
                    ElseIf Year2_TDS3(i) = "2019-20" Then
                       TDS3("DeductedYr") = "2019"
                    ElseIf Year2_TDS3(i) = "2020-21" Then
                       TDS3("DeductedYr") = "2020"
                    ElseIf Year2_TDS3(i) = "2021-22" Then
                       TDS3("DeductedYr") = "2021"
                    ElseIf Year2_TDS3(i) = "2022-23" Then
                       TDS3("DeductedYr") = "2022"
                    ElseIf Year2_TDS3(i) = "2023-24" Then
                       TDS3("DeductedYr") = "2023"
                    'added by Chetan C M
                    ElseIf Year2_TDS3(i) = "2024-25" Then
                       TDS3("DeductedYr") = "2024"
'                    New Schema updated by Konda as AY-2026-27 on 26-12-2025
                    ElseIf Year2_TDS3(i) = "2025-26" Then
                       TDS3("DeductedYr") = "2025"
                    End If
                End If
                
                
                
            
                If Range("TDS26QB.TotTDSOnAmtPaid").Cells.item(i).Value <> "" Then
                    TDS3("TDSDeducted") = UVCase(CDbl(Range("TDS26QB.TotTDSOnAmtPaid").item(i).Value))
                End If
                
                If Range("TDS26QB.6income").Cells.item(i).Value <> "" Then
                    TDS3("TDSClaimed") = UVCase(CDbl(Range("TDS26QB.6income").item(i).Value))
                End If
                UpdateProgressBar
               TDS3Details.add TDS3
               Set TDS3 = Nothing
               Set TDS3 = CreateObject("Scripting.Dictionary")

            Next
            
 End If
 Else
 End If
 jsonDictionary.add "TDS3Details", TDS3Details
            If Not Sheet2.Range("TDS26QB.Sum").Value = "" Then
              jsonDictionary("TotalTDS3Details") = TotalTDS3Details
            Else
              jsonDictionary("TotalTDS3Details") = 0
            End If
 Set ScheduleTDS3Dtls = jsonDictionary
 End Function
 Function TaxPayments() As Object
ScheduleTCS
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")
 Dim TotalTaxPayments As Variant
 Dim TaxPayment As Object
 TotalTaxPayments = Sheet2.Range("TaxP.Sum").Value
 
If Not IsEmpty(BSR_TDS) Then
    If UBound(BSR_TDS) > 0 Then
        subProcCaption = "IT"
        noOfProcessSub = UBound(BSR_TDS)
        Dim i As Variant
               
        Dim Tax
        Set TaxPayment = New Collection
        Set Tax = CreateObject("Scripting.Dictionary")
        
            For i = 1 To UBound(BSR_TDS)
                If BSR_TDS(i) <> "" Then
                    Tax("BSRCode") = UCase(BSR_TDS(i))
                End If
                If DateCredit_TDS(i) <> "" Then
                    Tax("DateDep") = Mid(UCase(DateCredit_TDS(i)), 7, 4) & "-" & Mid(UCase(DateCredit_TDS(i)), 4, 2) & "-" & Mid(UCase(DateCredit_TDS(i)), 1, 2)
                End If
                If SerialNum_TDS(i) <> "" Then
                    Tax("SrlNoOfChaln") = UVCase(CDbl(SerialNum_TDS(i)))
                End If
                If TaxPaid3_TDS(i) <> "" Then
                    Tax("Amt") = UVCase(CDbl(TaxPaid3_TDS(i)))
                Else
                    Tax("Amt") = 0
                End If
                TaxPayment.add Tax
                Set Tax = Nothing
                Set Tax = CreateObject("Scripting.Dictionary")

                UpdateProgressBar
            Next
            
End If
Else
End If
jsonDictionary.add "TaxPayment", TaxPayment
jsonDictionary("TotalTaxPayments") = TotalTaxPayments
Set TaxPayments = jsonDictionary
End Function

'XML Generation Function for 80G
        
Function Schedule80G() As Object
Dim i As Long
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim TotalDonationsUs80GCash, TotalDonationsUs80GOtherMode, TotalDonationsUs80G, TotalEligibleDonationsUs80G As Variant

Dim DoneeWithPan100 As Object
Dim DoneeWithPan50 As Object
Dim DoneeWithPan100Appr As Object
Dim DoneeWithPan50Appr As Object

Dim DoneeWith100, DoneeWith50, DoneeWith100Appr, DoneeWith50Appr

Set DoneeWith100 = CreateObject("Scripting.Dictionary")
Set DoneeWith50 = CreateObject("Scripting.Dictionary")
Set DoneeWith100Appr = CreateObject("Scripting.Dictionary")
Set DoneeWith50Appr = CreateObject("Scripting.Dictionary")

Dim AddressDetail100, AddressDetail50, AddressDetail00Appr, AddressDetai50Appr

Set AddressDetail100 = CreateObject("Scripting.Dictionary")
Set AddressDetail50 = CreateObject("Scripting.Dictionary")
Set AddressDetail00Appr = CreateObject("Scripting.Dictionary")
Set AddressDetai50Appr = CreateObject("Scripting.Dictionary")

Dim Don100Percent, Don50PercentNoApprReqd, Don100PercentApprReqd, Don50PercentApprReqd
Dim TotDon100PercentCash, TotDon100PercentOtherMode, TotDon100Percent, TotEligibleDon100Percent As Variant
Dim TotDon50PercentNoApprReqdCash, TotDon50PercentNoApprReqdOtherMode, TotDon50PercentNoApprReqd, TotEligibleDon50Percent As Variant
Dim TotDon100PercentApprReqdCash, TotDon100PercentApprReqdOtherMode, TotDon100PercentApprReqd, TotEligibleDon100PercentApprReqd As Variant
Dim TotDon50PercentApprReqdCash, TotDon50PercentApprReqdOtherMode, TotDon50PercentApprReqd, TotEligibleDon50PercentApprReqd As Variant

Set Don100Percent = CreateObject("Scripting.Dictionary")
Set Don50PercentNoApprReqd = CreateObject("Scripting.Dictionary")
Set Don100PercentApprReqd = CreateObject("Scripting.Dictionary")
Set Don50PercentApprReqd = CreateObject("Scripting.Dictionary")

Set DoneeWithPan100 = New Collection
Set DoneeWithPan50 = New Collection
Set DoneeWithPan100Appr = New Collection
Set DoneeWithPan50Appr = New Collection

TotalDonationsUs80GCash = UCase(Range("Per5080G.TotalDonationsUs80G").Value)
TotalDonationsUs80GOtherMode = UCase(Range("Per5080G.TotalDonationsOtherUs80G").Value)
TotalDonationsUs80G = UCase(Range("Per5080G.TotalDonationsUs80GTotal").Value)
TotalEligibleDonationsUs80G = UCase(Range("Per5080G.TotalEligibleDonationsUs80G").Value)

TotDon100PercentCash = UCase(Range("Per10080G.TotDon100Percent").Value)
TotDon100PercentOtherMode = UCase(Range("Per10080G.TotDonOther100Percent").Value)
TotDon100Percent = UCase(Range("Per10080G.TotDon100PercentTotal").Value)
TotEligibleDon100Percent = UCase(Range("Per10080G.TotElig100Percent").Value)

TotDon50PercentNoApprReqdCash = UCase(Range("PerNO5080G.TotDon100Percent").Value)
TotDon50PercentNoApprReqdOtherMode = UCase(Range("PerNO5080G.TotDonOther100Percent").Value)
TotDon50PercentNoApprReqd = UCase(Range("PerNO5080G.TotDon100PercentTotal").Value)
TotEligibleDon50Percent = UCase(Range("PerNO5080G.TotElig100Percent").Value)

TotDon100PercentApprReqdCash = UCase(Range("PerYES10080G.TotDon100Percent").Value)
TotDon100PercentApprReqdOtherMode = UCase(Range("PerYES10080G.TotDonOther100Percent").Value)
TotDon100PercentApprReqd = UCase(Range("PerYES10080G.TotDon100PercentTotal").Value)
TotEligibleDon100PercentApprReqd = UCase(Range("PerYES10080G.TotElig100Percent").Value)

TotDon50PercentApprReqdCash = UCase(Range("Per5080G.TotDon100Percent").Value)
TotDon50PercentApprReqdOtherMode = UCase(Range("Per5080G.TotDonother100Percent").Value)
TotDon50PercentApprReqd = UCase(Range("Per5080G.TotDon100PercentTotal").Value)
TotEligibleDon50PercentApprReqd = UCase(Range("Per5080G.TotElig100Percent").Value)
 
If Not IsEmpty(Name1_80GA) Then
If (UBound(Name1_80GA) > 0) Then
 subProcCaption = "80GA"
 noOfProcessSub = UBound(Name1_80GA)
    For i = 1 To UBound(Name1_80GA)
        If Name1_80GA(i) <> "" Then
            DoneeWith100("DoneeWithPanName") = UCase(Name1_80GA(i))
        End If
        If Pan1_80GA(i) <> "" Then
            DoneeWith100("DoneePAN") = UCase(Pan1_80GA(i))
        End If
        If Addr1_80GA(i) <> "" Then
            AddressDetail100("AddrDetail") = UCase(Addr1_80GA(i))
        End If
        If City1_80GA(i) <> "" Then
            AddressDetail100("CityOrTownOrDistrict") = UCase(City1_80GA(i))
        End If
        If State1_80GA(i) <> "" Then
            AddressDetail100("StateCode") = Mid((State1_80GA(i)), 1, 2)
        End If
        If PinCd1_80GA(i) <> "" Then
            AddressDetail100("PinCode") = UVCase(CDbl(PinCd1_80GA(i)))
        End If
        DoneeWith100.add "AddressDetail", AddressDetail100
        Set AddressDetail100 = Nothing
        Set AddressDetail100 = CreateObject("Scripting.Dictionary")

        
        If DonationAmt1_80GA(i) <> "" Then
            DoneeWith100("DonationAmtCash") = UVCase(CDbl(DonationAmt1_80GA(i)))
        Else
            DoneeWith100("DonationAmtCash") = 0
        End If
        
        If DonationAmt1_80GA1(i) <> "" Then
            DoneeWith100("DonationAmtOtherMode") = UVCase(CDbl(DonationAmt1_80GA1(i)))
        Else
            DoneeWith100("DonationAmtOtherMode") = 0
        End If
'Konda updated on 10-03-2026--V0.5
        If Transaction1_80GA(i) <> "" Then
            DoneeWith100("TransactionRefNum") = UCase(Transaction1_80GA(i))
        End If
        
        If IFSC_80GA(i) <> "" Then
            DoneeWith100("IFSCCode") = UCase(IFSC_80GA(i))
        End If
        
'===================
        
        If Trim(Range("Per10080G.DonationAmtTotal").Cells.item(i).Value) <> "" Then
            DoneeWith100("DonationAmt") = UVCase(CDbl(Range("Per10080G.DonationAmtTotal").Cells.item(i).Value))
        Else
            DoneeWith100("DonationAmt") = 0
        End If
        
        
        If Trim(Range("Per10080G.EligibleAmt").Cells.item(i).Value) <> "" Then
            DoneeWith100("EligibleDonationAmt") = UVCase(CDbl(Range("Per10080G.EligibleAmt").Cells.item(i).Value))
        Else
            DoneeWith100("EligibleDonationAmt") = 0
        End If
            DoneeWithPan100.add DoneeWith100
            Set DoneeWith100 = Nothing
            Set DoneeWith100 = CreateObject("Scripting.Dictionary")
        UpdateProgressBar
        Next
        Don100Percent.add "DoneeWithPan", DoneeWithPan100

        If Trim(Range("Per10080G.TotDon100Percent").Value) <> "" Then
            Don100Percent("TotDon100PercentCash") = UVCase(CDbl(TotDon100PercentCash))
        Else
            Don100Percent("TotDon100PercentCash") = 0
        End If
        
        If Trim(Range("Per10080G.TotDonOther100Percent").Value) <> "" Then
            Don100Percent("TotDon100PercentOtherMode") = UVCase(CDbl(TotDon100PercentOtherMode))
        Else
            Don100Percent("TotDon100PercentOtherMode") = 0
        End If
        
        If Trim(Range("Per10080G.TotDon100PercentTotal").Value) <> "" Then
            Don100Percent("TotDon100Percent") = UVCase(CDbl(TotDon100Percent))
        Else
            Don100Percent("TotDon100Percent") = 0
        End If
        
        If Trim(Range("Per10080G.TotElig100Percent").Value) <> "" Then
            Don100Percent("TotEligibleDon100Percent") = UVCase(CDbl(TotEligibleDon100Percent))
        Else
            Don100Percent("TotEligibleDon100Percent") = 0
        End If
        jsonDictionary.add "Don100Percent", Don100Percent
        End If
        End If
             
        If Not IsEmpty(Name2_80GB) Then
        If (UBound(Name2_80GB) > 0) Then
            subProcCaption = "80GB"
            noOfProcessSub = UBound(Name2_80GB)
        
            For i = 1 To UBound(Name2_80GB)
                If Name2_80GB(i) <> "" Then
                    DoneeWith50("DoneeWithPanName") = UCase(Name2_80GB(i))
                End If
                If Pan_80GB(i) <> "" Then
                    DoneeWith50("DoneePAN") = UCase(Pan_80GB(i))
                End If
                If Addr280GB(i) <> "" Then
                    AddressDetail50("AddrDetail") = UCase(Addr280GB(i))
                End If
                If City280GB(i) <> "" Then
                    AddressDetail50("CityOrTownOrDistrict") = UCase(City280GB(i))
                End If
                If State280GB(i) <> "" Then
                    AddressDetail50("StateCode") = Mid((State280GB(i)), 1, 2)
                End If
                If PinCode_80GB(i) <> "" Then
                    AddressDetail50("PinCode") = UVCase(CDbl(PinCode_80GB(i)))
                End If
                DoneeWith50.add "AddressDetail", AddressDetail50
                Set AddressDetail50 = Nothing
                Set AddressDetail50 = CreateObject("Scripting.Dictionary")
    
                If DonationAmt_80GB(i) <> "" Then
                    DoneeWith50("DonationAmtCash") = UVCase(CDbl(DonationAmt_80GB(i)))
                Else
                    DoneeWith50("DonationAmtCash") = 0
                End If
                
                If DonationAmt_80GB1(i) <> "" Then
                    DoneeWith50("DonationAmtOtherMode") = UVCase(CDbl(DonationAmt_80GB1(i)))
                Else
                    DoneeWith50("DonationAmtOtherMode") = 0
                End If
                
        'Konda updated on 10-03-2026--V0.5
                If Transaction1_80GB(i) <> "" Then
                    DoneeWith50("TransactionRefNum") = UCase(Transaction1_80GB(i))
                End If
                
                If IFSC_80GB(i) <> "" Then
                    DoneeWith50("IFSCCode") = UCase(IFSC_80GB(i))
                End If
        
'===================
                If Trim(Range("PerNO5080G.DonationAmtTotal").Cells.item(i).Value) <> "" Then
                    DoneeWith50("DonationAmt") = UVCase(CDbl(Range("PerNO5080G.DonationAmtTotal").Cells.item(i).Value))
                Else
                    DoneeWith50("DonationAmt") = 0
                End If
        
                If Trim(Range("PerNO5080G.EligibleAmt").Cells.item(i).Value) <> "" Then
                    DoneeWith50("EligibleDonationAmt") = Round(UVCase(CDbl(Range("PerNO5080G.EligibleAmt").Cells.item(i).Value)))
                Else
                    DoneeWith50("EligibleDonationAmt") = 0
                End If
                DoneeWithPan50.add DoneeWith50
                Set DoneeWith50 = Nothing
                Set DoneeWith50 = CreateObject("Scripting.Dictionary")
    
                UpdateProgressBar
            Next
            Don50PercentNoApprReqd.add "DoneeWithPan", DoneeWithPan50
            
            If Trim(Range("PerNO5080G.TotDon100Percent").Value) <> "" Then
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqdCash") = UVCase(CDbl(TotDon50PercentNoApprReqdCash))
            Else
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqdCash") = 0
            End If
                             
            If Trim(Range("PerNO5080G.TotDonOther100Percent").Value) <> "" Then
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqdOtherMode") = UVCase(CDbl(TotDon50PercentNoApprReqdOtherMode))
            Else
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqdOtherMode") = 0
            End If
                                  
            If Trim(Range("PerNO5080G.TotDon100PercentTotal").Value) <> "" Then
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqd") = UVCase(CDbl(TotDon50PercentNoApprReqd))
            Else
                Don50PercentNoApprReqd("TotDon50PercentNoApprReqd") = 0
            End If
                          
                     
            If Trim(Range("PerNO5080G.TotElig100Percent").Value) <> "" Then
                Don50PercentNoApprReqd("TotEligibleDon50Percent") = UVCase(CDbl(TotEligibleDon50Percent))
            Else
                Don50PercentNoApprReqd("TotEligibleDon50Percent") = 0
            End If
            
            jsonDictionary.add "Don50PercentNoApprReqd", Don50PercentNoApprReqd
            End If
            End If
            


            If Not IsEmpty(Name_80GC) Then
            If (UBound(Name_80GC) > 0) Then
                subProcCaption = "80GC"
                noOfProcessSub = UBound(Name_80GC)
            
                For i = 1 To UBound(Name_80GC)
                    If Name_80GC(i) <> "" Then
                        DoneeWith100Appr("DoneeWithPanName") = UCase(Name_80GC(i))
                    End If
                    If Pan_80GC(i) <> "" Then
                        DoneeWith100Appr("DoneePAN") = UCase(Pan_80GC(i))
                    End If
                   
                    If Addr_80GC(i) <> "" Then
                        AddressDetail00Appr("AddrDetail") = UCase(Addr_80GC(i))
                    End If
                    If City_80GC(i) <> "" Then
                        AddressDetail00Appr("CityOrTownOrDistrict") = UCase(City_80GC(i))
                    End If
                    If State_80GC(i) <> "" Then
                        AddressDetail00Appr("StateCode") = Mid((State_80GC(i)), 1, 2)
                    End If
                    If PinCode_80GC(i) <> "" Then
                        AddressDetail00Appr("PinCode") = UVCase(CDbl(PinCode_80GC(i)))
                    End If
                    DoneeWith100Appr.add "AddressDetail", AddressDetail00Appr
                    Set AddressDetail00Appr = Nothing
                    Set AddressDetail00Appr = CreateObject("Scripting.Dictionary")
    
                    If DonationAmt_80GC(i) <> "" Then
                        DoneeWith100Appr("DonationAmtCash") = UVCase(CDbl(DonationAmt_80GC(i)))
                    Else
                        DoneeWith100Appr("DonationAmtCash") = 0
                    End If
                    
                    If DonationAmt_80GC1(i) <> "" Then
                        DoneeWith100Appr("DonationAmtOtherMode") = UVCase(CDbl(DonationAmt_80GC1(i)))
                    Else
                        DoneeWith100Appr("DonationAmtOtherMode") = 0
                    End If
         'Konda updated on 10-03-2026--V0.5
                    If Transaction1_80GC(i) <> "" Then
                        DoneeWith100Appr("TransactionRefNum") = UCase(Transaction1_80GC(i))
                    End If
                    
                    If IFSC_80GC(i) <> "" Then
                        DoneeWith100Appr("IFSCCode") = UCase(IFSC_80GC(i))
                    End If
'===================
                    
                    
                    If Trim(Range("PerYES10080G.DonationAmtTotal").Cells.item(i).Value) <> "" Then
                        DoneeWith100Appr("DonationAmt") = UVCase(CDbl(Range("PerYES10080G.DonationAmtTotal").Cells.item(i).Value))
                    Else
                        DoneeWith100Appr("DonationAmt") = 0
                    End If
                    
                    If Trim(Range("PerYES10080G.EligibleAmt").Cells.item(i).Value) <> "" Then
                        DoneeWith100Appr("EligibleDonationAmt") = UVCase(CDbl(Range("PerYES10080G.EligibleAmt").Cells.item(i).Value))
                    Else
                        DoneeWith100Appr("EligibleDonationAmt") = 0
                    End If
                    DoneeWithPan100Appr.add DoneeWith100Appr
                    Set DoneeWith100Appr = Nothing
                    Set DoneeWith100Appr = CreateObject("Scripting.Dictionary")
    
                UpdateProgressBar
                Next
                Don100PercentApprReqd.add "DoneeWithPan", DoneeWithPan100Appr
            
                If Trim(Range("PerYES10080G.TotDon100Percent").Value) <> "" Then
                    Don100PercentApprReqd("TotDon100PercentApprReqdCash") = UVCase(CDbl(TotDon100PercentApprReqdCash))
                Else
                    Don100PercentApprReqd("TotDon100PercentApprReqdCash") = 0
                End If
                
                If Trim(Range("PerYES10080G.TotDonOther100Percent").Value) <> "" Then
                    Don100PercentApprReqd("TotDon100PercentApprReqdOtherMode") = UVCase(CDbl(TotDon100PercentApprReqdOtherMode))
                Else
                    Don100PercentApprReqd("TotDon100PercentApprReqdOtherMode") = 0
                End If
                
                If Trim(Range("PerYES10080G.TotDon100PercentTotal").Value) <> "" Then
                    Don100PercentApprReqd("TotDon100PercentApprReqd") = UVCase(CDbl(TotDon100PercentApprReqd))
                Else
                    Don100PercentApprReqd("TotDon100PercentApprReqd") = 0
                End If
                
                
                If Trim(Range("PerYES10080G.TotElig100Percent").Value) <> "" Then
                    Don100PercentApprReqd("TotEligibleDon100PercentApprReqd") = UVCase(CDbl(TotEligibleDon100PercentApprReqd))
                Else
                    Don100PercentApprReqd("TotEligibleDon100PercentApprReqd") = 0
                End If
                
                jsonDictionary.add "Don100PercentApprReqd", Don100PercentApprReqd
                End If
                End If

                If Not IsEmpty(Name_80GD) Then
                If (UBound(Name_80GD) > 0) Then
                    subProcCaption = "80GD"
                    noOfProcessSub = UBound(Name_80GD)
                    
                    For i = 1 To UBound(Name_80GD)
                    
                        If Name_80GD(i) <> "" Then
                            DoneeWith50Appr("DoneeWithPanName") = UCase(Name_80GD(i))
                        End If
                        If Pan_80GD(i) <> "" Then
                            DoneeWith50Appr("DoneePAN") = UCase(Pan_80GD(i))
                        End If
                        If Addr_80GD(i) <> "" Then
                            AddressDetai50Appr("AddrDetail") = UCase(Addr_80GD(i))
                        End If
                        If City_80GD(i) <> "" Then
                            AddressDetai50Appr("CityOrTownOrDistrict") = UCase(City_80GD(i))
                        End If
                        If State_80GD(i) <> "" Then
                            AddressDetai50Appr("StateCode") = Mid((State_80GD(i)), 1, 2)
                        End If
                        If PinCode_80GD(i) <> "" Then
                            AddressDetai50Appr("PinCode") = UVCase(CDbl(PinCode_80GD(i)))
                        End If
                        DoneeWith50Appr.add "AddressDetail", AddressDetai50Appr
                        Set AddressDetai50Appr = Nothing
                        Set AddressDetai50Appr = CreateObject("Scripting.Dictionary")
    
                        If DonationAmt_80GD(i) <> "" Then
                            DoneeWith50Appr("DonationAmtCash") = UVCase(CDbl(DonationAmt_80GD(i)))
                        Else
                            DoneeWith50Appr("DonationAmtCash") = 0
                        End If
                        
                        If DonationAmt_80GD1(i) <> "" Then
                            DoneeWith50Appr("DonationAmtOtherMode") = UVCase(CDbl(DonationAmt_80GD1(i)))
                        Else
                            DoneeWith50Appr("DonationAmtOtherMode") = 0
                        End If
            'Konda updated on 10-03-2026--V0.5
                        If Transaction1_80GD(i) <> "" Then
                            DoneeWith50Appr("TransactionRefNum") = UCase(Transaction1_80GD(i))
                        End If
                        
                        If IFSC_80GD(i) <> "" Then
                            DoneeWith50Appr("IFSCCode") = UCase(IFSC_80GD(i))
                        End If
'===================
                        
                        
                        If Trim(Range("Per5080G.DonationAmtTotal").Cells.item(i).Value) <> "" Then
                            DoneeWith50Appr("DonationAmt") = UVCase(CDbl(Range("Per5080G.DonationAmtTotal").Cells.item(i).Value))
                        Else
                            DoneeWith50Appr("DonationAmt") = 0
                        End If
                        
                        If Trim(Range("Per5080G.EligibleAmt").Cells.item(i).Value) <> "" Then
                            DoneeWith50Appr("EligibleDonationAmt") = Round(UVCase(CDbl(Range("Per5080G.EligibleAmt").Cells.item(i).Value)))
                        Else
                            DoneeWith50Appr("EligibleDonationAmt") = 0
                        End If
'Change.25.01.2023.102.80GB
                        If Trim(Range("Per5080G.ArnNbr").Cells.item(i).Value) <> "" Then
                            DoneeWith50Appr("ArnNbr") = UCase(Range("Per5080G.ArnNbr").Cells.item(i).Value)
                        
                        End If
                        
'End Change

                        
                        DoneeWithPan50Appr.add DoneeWith50Appr
                        Set DoneeWith50Appr = Nothing
                        Set DoneeWith50Appr = CreateObject("Scripting.Dictionary")
        
                    UpdateProgressBar
                    Next
                    Don50PercentApprReqd.add "DoneeWithPan", DoneeWithPan50Appr
                
                If Trim(Range("Per5080G.TotDon100Percent").Value) <> "" Then
                    Don50PercentApprReqd("TotDon50PercentApprReqdCash") = UVCase(CDbl(Range("Per5080G.TotDon100Percent").Value))
                Else
                    Don50PercentApprReqd("TotDon50PercentApprReqdCash") = 0
                End If
                
                If Trim(Range("Per5080G.TotDonother100Percent").Value) <> "" Then
                    Don50PercentApprReqd("TotDon50PercentApprReqdOtherMode") = UVCase(CDbl(Range("Per5080G.TotDonother100Percent").Value))
                Else
                    Don50PercentApprReqd("TotDon50PercentApprReqdOtherMode") = 0
                End If
                
                If Trim(Range("Per5080G.TotDon100PercentTotal").Value) <> "" Then
                    Don50PercentApprReqd("TotDon50PercentApprReqd") = UVCase(CDbl(Range("Per5080G.TotDon100PercentTotal").Value))
                Else
                    Don50PercentApprReqd("TotDon50PercentApprReqd") = 0
                End If
                
                If Trim(Range("Per5080G.TotElig100Percent").Value) <> "" Then
                    Don50PercentApprReqd("TotEligibleDon50PercentApprReqd") = UVCase(CDbl(Range("Per5080G.TotElig100Percent").Value))
                Else
                    Don50PercentApprReqd("TotEligibleDon50PercentApprReqd") = 0
                End If
                
                jsonDictionary.add "Don50PercentApprReqd", Don50PercentApprReqd
                End If
                End If
        If Trim(Range("Per5080G.TotalDonationsUs80G").Value) <> "" Then
            jsonDictionary("TotalDonationsUs80GCash") = UVCase(CDbl(TotalDonationsUs80GCash))
        Else
            jsonDictionary("TotalDonationsUs80GCash") = 0
        End If
        
        If Trim(Range("Per5080G.TotalDonationsOtherUs80G").Value) <> "" Then
            jsonDictionary("TotalDonationsUs80GOtherMode") = UVCase(CDbl(TotalDonationsUs80GOtherMode))
        Else
            jsonDictionary("TotalDonationsUs80GOtherMode") = 0
        End If
        
        If Trim(Range("Per5080G.TotalDonationsUs80GTotal").Value) <> "" Then
            jsonDictionary("TotalDonationsUs80G") = UVCase(CDbl(TotalDonationsUs80G))
        Else
            jsonDictionary("TotalDonationsUs80G") = 0
        End If
        
        If Trim(Range("Per5080G.TotalEligibleDonationsUs80G").Value) <> "" Then
            jsonDictionary("TotalEligibleDonationsUs80G") = UVCase(CDbl(TotalEligibleDonationsUs80G))
        Else
            jsonDictionary("TotalEligibleDonationsUs80G") = 0
        End If
Set Schedule80G = jsonDictionary
End Function
'XML Generation Function for 80GGA
Function Schedule80GGA() As Object
Dim i As Long
Dim jsonDictionary, DonationDtlsSciRsrchRural, AddressDetail
Dim DonationDtlsSciRsrchRuralDev, TotalDonationAmtCash80GGA, TotalDonationAmtOtherMode80GGA, TotalDonationsUs80GGA, TotalEligibleDonationAmt80GGA As Variant
Set DonationDtlsSciRsrchRuralDev = New Collection
Dim dateDonation

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set DonationDtlsSciRsrchRural = CreateObject("Scripting.Dictionary")
Set AddressDetail = CreateObject("Scripting.Dictionary")


TotalDonationAmtCash80GGA = UCase(Range("Total_DonationInCash_80GGA").Value)
TotalDonationAmtOtherMode80GGA = UCase(Range("Total_DonationInOtherMode_80GGA").Value)
TotalDonationsUs80GGA = UCase(Range("Total_Donation_80GGA").Value)
TotalEligibleDonationAmt80GGA = UCase(Range("Total_Donation_Eligible_80GGA").Value)
subProcCaption = "80GGA"
    If Not IsEmpty(RelevantClauseClaimed_80GGA) Then
    If (UBound(RelevantClauseClaimed_80GGA) > 0) Then
    
    For i = 1 To UBound(RelevantClauseClaimed_80GGA)
 
        If Mid(RelevantClauseClaimed_80GGA(i), 1, InStrRev(RelevantClauseClaimed_80GGA(i), ")")) = "80GGA(2)(a)" Then
           RelevantClauseClaimed_80GGA(i) = "80GGA2a"
        ElseIf Mid(RelevantClauseClaimed_80GGA(i), 1, InStrRev(RelevantClauseClaimed_80GGA(i), ")")) = "80GGA(2)(aa)" Then
           RelevantClauseClaimed_80GGA(i) = "80GGA2aa"
        ElseIf Mid(RelevantClauseClaimed_80GGA(i), 1, InStrRev(RelevantClauseClaimed_80GGA(i), ")")) = "80GGA(2)(b)" Then
           RelevantClauseClaimed_80GGA(i) = "80GGA2b"
        ElseIf Mid(RelevantClauseClaimed_80GGA(i), 1, InStrRev(RelevantClauseClaimed_80GGA(i), ")")) = "80GGA(2)(bb)" Then
           RelevantClauseClaimed_80GGA(i) = "80GGA2bb"
        ElseIf Mid(RelevantClauseClaimed_80GGA(i), 1, InStrRev(RelevantClauseClaimed_80GGA(i), ")")) = "80GGA(2)(c)" Then
           RelevantClauseClaimed_80GGA(i) = "80GGA2c"
        ElseIf Mid(RelevantClauseClaimed_80GGA(i), 1, InStrRev(RelevantClauseClaimed_80GGA(i), ")")) = "80GGA(2)(cc)" Then
           RelevantClauseClaimed_80GGA(i) = "80GGA2cc"
        ElseIf Mid(RelevantClauseClaimed_80GGA(i), 1, InStrRev(RelevantClauseClaimed_80GGA(i), ")")) = "80GGA(2)(d)" Then
           RelevantClauseClaimed_80GGA(i) = "80GGA2d"
        ElseIf Mid(RelevantClauseClaimed_80GGA(i), 1, InStrRev(RelevantClauseClaimed_80GGA(i), ")")) = "80GGA(2)(e)" Then
           RelevantClauseClaimed_80GGA(i) = "80GGA2e"
        End If

    Next
    
    
      For i = 1 To UBound(RelevantClauseClaimed_80GGA)

            If RelevantClauseClaimed_80GGA(i) <> "" Then
                DonationDtlsSciRsrchRural("RelevantClauseUndrDedClaimed") = RelevantClauseClaimed_80GGA(i)
            End If
            
            If Name_of_Donee_80GGA(i) <> "" Then
                DonationDtlsSciRsrchRural("NameOfDonee") = UCase(Name_of_Donee_80GGA(i))
            End If
            
            If Address_80GGA(i) <> "" Then
                AddressDetail("AddrDetail") = UCase(Address_80GGA(i))
            End If
            
            If City_Town_District_80GGA(i) <> "" Then
                AddressDetail("CityOrTownOrDistrict") = UCase(City_Town_District_80GGA(i))
            End If
            
            If State_Code_80GGA(i) <> "" Then
                AddressDetail("StateCode") = Mid((State_Code_80GGA(i)), 1, 2)
            End If
            
            If Pincode_80GGA(i) <> "" Then
                AddressDetail("PinCode") = UVCase(CDbl(Pincode_80GGA(i)))
            End If
            DonationDtlsSciRsrchRural.add "AddressDetail", AddressDetail
            Set AddressDetail = Nothing
            Set AddressDetail = CreateObject("Scripting.Dictionary")
    
            If PAN_of_donee_80GGA(i) <> "" Then
                DonationDtlsSciRsrchRural("DoneePAN") = UCase(PAN_of_donee_80GGA(i))
            End If
            
            If Date_Donation(i) <> "" Then
                dateDonation = Mid((Date_Donation(i)), 7, 4) & "-" & Mid((Date_Donation(i)), 4, 2) & "-" & Mid((Date_Donation(i)), 1, 2)
'                DonationDtlsSciRsrchRural("DateOfDonationCash") = dateDonation
                 'DonationDtlsSciRsrchRural("DonationAmtCashDate") = dateDonation
            End If
            
            If Donation_cash_80GGA(i) <> "" Then
                DonationDtlsSciRsrchRural("DonationAmtCash") = UVCase(CDbl(Donation_cash_80GGA(i)))
            Else
                DonationDtlsSciRsrchRural("DonationAmtCash") = 0
            End If
            
            
            If Donation_other_80GGA(i) <> "" Then
                DonationDtlsSciRsrchRural("DonationAmtOtherMode") = UVCase(CDbl(Donation_other_80GGA(i)))
            Else
                DonationDtlsSciRsrchRural("DonationAmtOtherMode") = 0
            End If
            
            
            If Trim(Range("Donation_total_80GGA").Cells.item(i).Value) <> "" Then
                DonationDtlsSciRsrchRural("DonationAmt") = UVCase(CDbl(Range("Donation_total_80GGA").Cells.item(i).Value))
            Else
                DonationDtlsSciRsrchRural("DonationAmt") = 0
            End If
            
            
            If Trim(Range("Donation_Eligible_80GGA").Cells.item(i).Value) <> "" Then
                DonationDtlsSciRsrchRural("EligibleDonationAmt") = UVCase(CDbl(Range("Donation_Eligible_80GGA").Cells.item(i).Value))
            Else
                DonationDtlsSciRsrchRural("EligibleDonationAmt") = 0
            End If
            DonationDtlsSciRsrchRuralDev.add DonationDtlsSciRsrchRural
            Set DonationDtlsSciRsrchRural = Nothing
            Set DonationDtlsSciRsrchRural = CreateObject("Scripting.Dictionary")
            UpdateProgressBar
            Next
        End If
    End If
    jsonDictionary.add "DonationDtlsSciRsrchRuralDev", DonationDtlsSciRsrchRuralDev
            
            If Trim(Range("Total_DonationInCash_80GGA").Value) > 0 Then
                jsonDictionary("TotalDonationAmtCash80GGA") = UVCase(CDbl(TotalDonationAmtCash80GGA))
            Else
                jsonDictionary("TotalDonationAmtCash80GGA") = 0
            End If
            
            If Trim(Range("Total_DonationInOtherMode_80GGA").Value) > 0 Then
                jsonDictionary("TotalDonationAmtOtherMode80GGA") = UVCase(CDbl(TotalDonationAmtOtherMode80GGA))
            Else
                jsonDictionary("TotalDonationAmtOtherMode80GGA") = 0
            End If
            
            If Trim(Range("Total_Donation_80GGA").Value) > 0 Then
                jsonDictionary("TotalDonationsUs80GGA") = UVCase(CDbl(TotalDonationsUs80GGA))
            Else
                jsonDictionary("TotalDonationsUs80GGA") = 0
            End If
  
            If Trim(Range("Total_Donation_Eligible_80GGA").Value) > 0 Then
                jsonDictionary("TotalEligibleDonationAmt80GGA") = UVCase(CDbl(TotalEligibleDonationAmt80GGA))
            Else
                jsonDictionary("TotalEligibleDonationAmt80GGA") = 0
            End If
Set Schedule80GGA = jsonDictionary
End Function
Function ScheduleTCS() As Object
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")
Dim TotalSchTCS, TCS As Variant
TotalSchTCS = Sheet11.Range("TCS.Sum").Value

If Not IsEmpty(TAN_TCS) Then
    If (UBound(TAN_TCS) > 0) Then
        subProcCaption = "TCS"
        noOfProcessSub = UBound(TAN_TCS)
        Dim i As Variant
        
        Dim TCSDict, EmployerOrDeductorOrCollectDetl
        Set TCS = New Collection
        
        Set TCSDict = CreateObject("Scripting.Dictionary")
        Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")

        
        
            For i = 1 To UBound(TAN_TCS)
                    If TAN_TCS(i) <> "" Then
                        EmployerOrDeductorOrCollectDetl("TAN") = TAN_TCS(i)
                    End If
                
                    If Not IsEmpty(EmployerOrDeductorOrCollecterName_TCS) And UBound(EmployerOrDeductorOrCollecterName_TCS) > 0 Then
                        EmployerOrDeductorOrCollectDetl("EmployerOrDeductorOrCollecterName") = EmployerOrDeductorOrCollecterName_TCS(i)
                    End If
                    TCSDict.add "EmployerOrDeductorOrCollectDetl", EmployerOrDeductorOrCollectDetl
                    Set EmployerOrDeductorOrCollectDetl = Nothing
                    Set EmployerOrDeductorOrCollectDetl = CreateObject("Scripting.Dictionary")
        
                    If Not IsEmpty(AmtTaxCollected_TCS) And UBound(AmtTaxCollected_TCS) > 0 Then
                        TCSDict("AmtTaxCollected") = AmtTaxCollected_TCS(i)
                    End If
                    
                    
                   ' TAX DETAILS-E4
                    'If Not IsEmpty(Year2_TCS) And UBound(Year2_TCS) > 0 Then
                     '   TCSDict("CollectedYr") = UCase(Year2_TCS(i))
                    'End If
                    
                    'TAX DETAILS-C4
                        If Year2_TCS(i) <> "" Then
                            If Year2_TCS(i) = "2008-09" Then
                               TCSDict("CollectedYr") = "2008"
                            ElseIf Year2_TCS(i) = "2009-10" Then
                               TCSDict("CollectedYr") = "2009"
                            ElseIf Year2_TCS(i) = "2010-11" Then
                               TCSDict("CollectedYr") = "2010"
                            ElseIf Year2_TCS(i) = "2011-12" Then
                               TCSDict("CollectedYr") = "2011"
                            ElseIf Year2_TCS(i) = "2012-13" Then
                               TCSDict("CollectedYr") = "2012"
                            ElseIf Year2_TCS(i) = "2013-14" Then
                               TCSDict("CollectedYr") = "2013"
                            ElseIf Year2_TCS(i) = "2014-15" Then
                               TCSDict("CollectedYr") = "2014"
                            ElseIf Year2_TCS(i) = "2015-16" Then
                               TCSDict("CollectedYr") = "2015"
                            ElseIf Year2_TCS(i) = "2016-17" Then
                               TCSDict("CollectedYr") = "2016"
                            ElseIf Year2_TCS(i) = "2017-18" Then
                               TCSDict("CollectedYr") = "2017"
                            ElseIf Year2_TCS(i) = "2018-19" Then
                               TCSDict("CollectedYr") = "2018"
                            ElseIf Year2_TCS(i) = "2019-20" Then
                               TCSDict("CollectedYr") = "2019"
                            ElseIf Year2_TCS(i) = "2020-21" Then
                               TCSDict("CollectedYr") = "2020"
                            ElseIf Year2_TCS(i) = "2021-22" Then
                               TCSDict("CollectedYr") = "2021"
                            ElseIf Year2_TCS(i) = "2022-23" Then
                               TCSDict("CollectedYr") = "2022"
                            ElseIf Year2_TCS(i) = "2023-24" Then
                               TCSDict("CollectedYr") = "2023"
                            'added by Chetan C M
                            ElseIf Year2_TCS(i) = "2024-25" Then
                               TCSDict("CollectedYr") = "2024"
                            'New Schema updated by Konda as AY-2026-27 on 26-12-2025
                            ElseIf Year2_TCS(i) = "2025-26" Then
                               TCSDict("CollectedYr") = "2025"
                            End If
                        End If
                    
                    If Not IsEmpty(BroughtFwdTCSAmt_TCS) And UBound(BroughtFwdTCSAmt_TCS) > 0 Then
                        TCSDict("TotalTCS") = BroughtFwdTCSAmt_TCS(i)
                    End If
                
                    If Not IsEmpty(AmtClaimedOnOwnHands_TCS) And UBound(AmtClaimedOnOwnHands_TCS) > 0 Then
                    TCSDict("AmtTCSClaimedThisYear") = AmtClaimedOnOwnHands_TCS(i)
                    End If
                    TCS.add TCSDict
                    Set TCSDict = Nothing
                    Set TCSDict = CreateObject("Scripting.Dictionary")
                UpdateProgressBar
            Next
                
End If
End If
jsonDictionary.add "TCS", TCS
                jsonDictionary("TotalSchTCS") = TotalSchTCS
Set ScheduleTCS = jsonDictionary
End Function
'XML Generation Function For 80D
'Function Schedule80D() As Object
'subProcCaption = "Schedule 80D"
'noOfProcessSub = 10
'Dim jsonDictionary, Sec80DSelfFamSrCtznHealth
'Dim SeniorCitizenFlag, parentsSeniorCitizenFlag, SelfAndFamily, HealthInsPremSlfFam, PrevHlthChckUpSlfFam, SelfAndFamilySeniorCitizen As String
'Dim HlthInsPremSlfFamSrCtzn, PrevHlthChckUpSlfFamSrCtzn, MedicalExpSlfFamSrCtzn, Parents, HlthInsPremParents, PrevHlthChckUpParents As Long
'Dim ParentsSeniorCitizen, HlthInsPremParentsSrCtzn, PrevHlthChckUpParentsSrCtzn, MedicalExpParentsSrCtzn, EligibleAmountOfDedn As Long
'
'Set jsonDictionary = CreateObject("Scripting.Dictionary")
'Set Sec80DSelfFamSrCtznHealth = CreateObject("Scripting.Dictionary")
'
'If FamilyMember = "Yes" Then
'    FamilyMember = "Y"
'ElseIf FamilyMember = "No" Then
'    FamilyMember = "N"
'ElseIf FamilyMember = "Not claiming for Self/ Family" Then
'    FamilyMember = "S"
'End If
'
'If SeniorCitizen = "Yes" Then
'    SeniorCitizen = "Y"
'ElseIf SeniorCitizen = "No" Then
'    SeniorCitizen = "N"
'ElseIf SeniorCitizen = "Not claiming for Parents" Then
'    SeniorCitizen = "P"
'End If
'
'SeniorCitizenFlag = UCase(FamilyMember)
'parentsSeniorCitizenFlag = UCase(SeniorCitizen)
'SelfAndFamily = Sheet9.Range("Self_And_Family_80D").Value
'HealthInsPremSlfFam = Sheet9.Range("Health_Insurance_80D").Value
'PrevHlthChckUpSlfFam = Sheet9.Range("Preventive_Health_80D").Value
'SelfAndFamilySeniorCitizen = Sheet9.Range("Senior_Citizen_80D").Value
'HlthInsPremSlfFamSrCtzn = Sheet9.Range("Health_InsuranceSC_80D").Value
'PrevHlthChckUpSlfFamSrCtzn = Sheet9.Range("Preventive_Health_SC_80D").Value
'MedicalExpSlfFamSrCtzn = Sheet9.Range("Medical_Expenditure_SC_80D").Value
'Parents = Sheet9.Range("Parents_80D").Value
'HlthInsPremParents = Sheet9.Range("Health_Insurance2_80D").Value
'PrevHlthChckUpParents = Sheet9.Range("Preventive_Health2_80D").Value
'ParentsSeniorCitizen = Sheet9.Range("Parents_SC_80D").Value
'HlthInsPremParentsSrCtzn = Sheet9.Range("Health_Insurance3_80D").Value
'PrevHlthChckUpParentsSrCtzn = Sheet9.Range("Preventive_Health3_80D").Value
'MedicalExpParentsSrCtzn = Sheet9.Range("Medical_Expenditure2_80D").Value
'EligibleAmountOfDedn = Sheet9.Range("Eligible_Amount_80D").Value
'
'
'    If FamilyMember <> "" And FamilyMember <> "(Select)" Then
'        Sec80DSelfFamSrCtznHealth("SeniorCitizenFlag") = SeniorCitizenFlag
'    End If
'
'    If Sheet9.Range("Self_And_Family_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("SelfAndFamily") = UVCase(CDbl(SelfAndFamily))
'    Else
'        Sec80DSelfFamSrCtznHealth("SelfAndFamily") = 0
'    End If
'
'    If Sheet9.Range("Health_Insurance_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("HealthInsPremSlfFam") = UVCase(CDbl(HealthInsPremSlfFam))
'    Else
'        Sec80DSelfFamSrCtznHealth("HealthInsPremSlfFam") = 0
'    End If
'
'    If Sheet9.Range("Preventive_Health_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFam") = UVCase(CDbl(PrevHlthChckUpSlfFam))
'    Else
'        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFam") = 0
'    End If
'
'    If Sheet9.Range("Senior_Citizen_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("SelfAndFamilySeniorCitizen") = UVCase(CDbl(SelfAndFamilySeniorCitizen))
'    Else
'        Sec80DSelfFamSrCtznHealth("SelfAndFamilySeniorCitizen") = 0
'    End If
'
'    If Sheet9.Range("Health_InsuranceSC_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("HlthInsPremSlfFamSrCtzn") = UVCase(CDbl(HlthInsPremSlfFamSrCtzn))
'    Else
'        Sec80DSelfFamSrCtznHealth("HlthInsPremSlfFamSrCtzn") = 0
'    End If
'
'    If Sheet9.Range("Preventive_Health_SC_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFamSrCtzn") = UVCase(CDbl(PrevHlthChckUpSlfFamSrCtzn))
'    Else
'        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFamSrCtzn") = 0
'    End If
'
'    If Sheet9.Range("Medical_Expenditure_SC_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("MedicalExpSlfFamSrCtzn") = UVCase(CDbl(MedicalExpSlfFamSrCtzn))
'    Else
'        Sec80DSelfFamSrCtznHealth("MedicalExpSlfFamSrCtzn") = 0
'    End If
'
'    If SeniorCitizen <> "" And SeniorCitizen <> "(Select)" Then
'        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizenFlag") = parentsSeniorCitizenFlag
'    End If
'
'    If Sheet9.Range("Parents_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("Parents") = UVCase(CDbl(Parents))
'    Else
'        Sec80DSelfFamSrCtznHealth("Parents") = 0
'    End If
'
'    If Sheet9.Range("Health_Insurance2_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("HlthInsPremParents") = UVCase(CDbl(HlthInsPremParents))
'    Else
'        Sec80DSelfFamSrCtznHealth("HlthInsPremParents") = 0
'    End If
'
'    If Sheet9.Range("Preventive_Health2_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParents") = UVCase(CDbl(PrevHlthChckUpParents))
'    Else
'        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParents") = 0
'    End If
'
'    If Sheet9.Range("Parents_SC_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizen") = UVCase(CDbl(ParentsSeniorCitizen))
'    Else
'        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizen") = 0
'    End If
'
'    If Sheet9.Range("Health_Insurance3_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("HlthInsPremParentsSrCtzn") = UVCase(CDbl(HlthInsPremParentsSrCtzn))
'    Else
'        Sec80DSelfFamSrCtznHealth("HlthInsPremParentsSrCtzn") = 0
'    End If
'
'    If Sheet9.Range("Preventive_Health3_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParentsSrCtzn") = UVCase(CDbl(PrevHlthChckUpParentsSrCtzn))
'    Else
'        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParentsSrCtzn") = 0
'    End If
'
'    If Sheet9.Range("Medical_Expenditure2_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("MedicalExpParentsSrCtzn") = UVCase(CDbl(MedicalExpParentsSrCtzn))
'    Else
'        Sec80DSelfFamSrCtznHealth("MedicalExpParentsSrCtzn") = 0
'    End If
'
'    If Sheet9.Range("Eligible_Amount_80D").Value <> "" Then
'        Sec80DSelfFamSrCtznHealth("EligibleAmountOfDedn") = UVCase(CDbl(EligibleAmountOfDedn))
'    Else
'        Sec80DSelfFamSrCtznHealth("EligibleAmountOfDedn") = 0
'    End If
'    jsonDictionary.add "Sec80DSelfFamSrCtznHealth", Sec80DSelfFamSrCtznHealth
'Set Schedule80D = jsonDictionary
'End Function
Function Schedule80D()
On Error Resume Next
subProcCaption = "Schedule 80D"
noOfProcessSub = 10
If Mid(Sheet1.Range("sheet1.Status").Value, 1, 1) <> "F" Then

Dim jsonDictionary, Sec80DSelfFamSrCtznHealth

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set Sec80DSelfFamSrCtznHealth = CreateObject("Scripting.Dictionary")

FamilyMember = Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value
    
    If FamilyMember = "Yes" Then
     FamilyMember = "Y"
     ElseIf FamilyMember = "No" Then
     FamilyMember = "N"
'    ElseIf FamilyMember = "Not Claiming for Self/Family" Then
     ElseIf FamilyMember = "Not claiming for Self/ Family" Then 'Konda_09/05/2025

     FamilyMember = "S"
     End If


    If FamilyMember <> "" And FamilyMember <> "(Select)" Then
        Sec80DSelfFamSrCtznHealth("SeniorCitizenFlag") = UCase(FamilyMember)
    Else
        Sec80DSelfFamSrCtznHealth("SeniorCitizenFlag") = ""
    End If
    
    If Sheet9.Range("Self_And_Family_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("SelfAndFamily") = Sheet9.Range("Self_And_Family_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("SelfAndFamily") = 0
    End If
    'Malli-------------------------16/04/2025
    
    If Sheet9.Range("Health_Insurance_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("HealthInsPremSlfFam") = Sheet9.Range("Health_Insurance_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("HealthInsPremSlfFam") = 0
    End If
    
    '_________________________________
    
Dim i As Long
Dim mIntCells80DA As Variant
Dim Sch80DInsDtls, Sch80DInsDtls_G, InsDtls80Da, Sec80DSelfFamHIDtls

'Set InsDtls80Da = CreateObject("Scripting.Dictionary")
Set Sec80DSelfFamHIDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = New Collection
'Set InsDtls80Da = New Collection

'Sec80DSelfFamHIDtls("HealthInsPremSlfFam") = UVCase(Sheet9.Range("Health_Insurance_80D").Value)

If Sheet9.Range("TotAmtA1.80D").Value > 0 Then

  
mIntCells80DA = Sheet9.Range("NameInsurerA1.80D").Rows.count
 
For i = 1 To mIntCells80DA

If Sheet9.Range("NameInsurerA1.80D").Cells(i, 1).Value <> "" Then

'Dim InsurerName, PolicyNo, ReceiptNo, HealthInsAmt,TotalPayments
Dim InsurerName_80DA, PolicyNo_80DA, ReceiptNo_80DA, HealthInsAmt_80DA, TotalPayments



InsurerName_80DA = Sheet9.Range("NameInsurerA1.80D").Cells(i, 1).Value
PolicyNo_80DA = Sheet9.Range("PolicyNumA1.80D").Cells(i, 1).Value
'ReceiptNo_80DA = Sheet9.Range("ReceiptNumA1.80D").Cells(i, 1).Value
HealthInsAmt_80DA = Sheet9.Range("AmtA1.80D").Cells(i, 1).Value

If InsurerName_80DA <> "" Then
Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80DA)
End If

If PolicyNo_80DA <> "" Then
Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80DA)
End If

'If ReceiptNo_80DA <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80DA)
'End If

If HealthInsAmt_80DA <> "" Then
Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80DA)
End If

'End If

                Sch80DInsDtls.add Sch80DInsDtls_G
                Set Sch80DInsDtls_G = Nothing
                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
End If
Next
'End If
                
                Sec80DSelfFamHIDtls.add "Sch80DInsDtls", Sch80DInsDtls
                Set Sch80DInsDtls = Nothing
                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")


Sec80DSelfFamHIDtls("TotalPayments") = UVCase(Sheet9.Range("TotAmtA1.80D").Value)

                Sec80DSelfFamSrCtznHealth.add "Sec80DSelfFamHIDtls", Sec80DSelfFamHIDtls
                Set Sec80DSelfFamHIDtls = Nothing
                Set Sec80DSelfFamHIDtls = CreateObject("Scripting.Dictionary")
End If
'---------------------------------

    If Sheet9.Range("Preventive_Health_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFam") = Sheet9.Range("Preventive_Health_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFam") = 0
    End If
    
    If Sheet9.Range("Senior_Citizen_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("SelfAndFamilySeniorCitizen") = Sheet9.Range("Senior_Citizen_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("SelfAndFamilySeniorCitizen") = 0
    End If
    
    '----------------------------
    If Sheet9.Range("Health_InsuranceSC_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("HlthInsPremSlfFamSrCtzn") = Sheet9.Range("Health_InsuranceSC_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("HlthInsPremSlfFamSrCtzn") = 0
    End If
    '_____________________________
Dim mIntCells80Db As Variant

Dim InsDtls80Db, Sec80DSelfFamSrCtznHIDtls
Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = New Collection
'Set InsDtls80Db = CreateObject("Scripting.Dictionary")
Set Sec80DSelfFamSrCtznHIDtls = CreateObject("Scripting.Dictionary")
 
' Sec80DSelfFamSrCtznHIDtls("HlthInsPremSlfFamSrCtzn") = UVCase(Sheet9.Range("Health_InsuranceSC_80D").Value)

If Sheet9.Range("TotAmtB1.80D").Value > 0 Then
 
 mIntCells80Db = Sheet9.Range("NameInsurerB1.80D").Rows.count
 
For i = 1 To mIntCells80Db

If Sheet9.Range("NameInsurerB1.80D").Cells(i, 1).Value <> "" Then

 
Dim InsurerName_80Db, PolicyNo_80Db, ReceiptNo_80Db, HealthInsAmt_80Db



InsurerName_80Db = Sheet9.Range("NameInsurerB1.80D").Cells(i, 1).Value
PolicyNo_80Db = Sheet9.Range("PolicyNumB1.80D").Cells(i, 1).Value
'ReceiptNo_80Db = Sheet9.Range("ReceiptNumB1.80D").Cells(i, 1).Value
HealthInsAmt_80Db = Sheet9.Range("AmtB1.80D").Cells(i, 1).Value

If InsurerName_80Db <> "" Then
Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80Db)
End If

If PolicyNo_80Db <> "" Then
Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80Db)
End If

'If ReceiptNo_80Db <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80Db)
'End If

If HealthInsAmt_80Db <> "" Then
Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80Db)
End If

'End If

                Sch80DInsDtls.add Sch80DInsDtls_G
                Set Sch80DInsDtls_G = Nothing
                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
End If
Next
'End If
                Sec80DSelfFamSrCtznHIDtls.add "Sch80DInsDtls", Sch80DInsDtls
                Set Sch80DInsDtls = Nothing
                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")


Sec80DSelfFamSrCtznHIDtls("TotalPayments") = UVCase(Sheet9.Range("TotAmtB1.80D").Value)

                Sec80DSelfFamSrCtznHealth.add "Sec80DSelfFamSrCtznHIDtls", Sec80DSelfFamSrCtznHIDtls
                Set Sec80DSelfFamSrCtznHIDtls = Nothing
                Set Sec80DSelfFamSrCtznHIDtls = CreateObject("Scripting.Dictionary")


End If
    
    '-----------------------------
    If Sheet9.Range("Preventive_Health_SC_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFamSrCtzn") = Sheet9.Range("Preventive_Health_SC_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpSlfFamSrCtzn") = 0
    End If
    
    If Sheet9.Range("Medical_Expenditure_SC_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("MedicalExpSlfFamSrCtzn") = Sheet9.Range("Medical_Expenditure_SC_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("MedicalExpSlfFamSrCtzn") = 0
    End If

    
If Mid(Sheet1.Range("sheet1.Status").Value, 1, 1) <> "H" Then
Dim SeniorCitizen As Variant
SeniorCitizen = Sheet9.Range("DropDown_ValueOf_SC_80D").Value
    
    If SeniorCitizen = "Yes" Then
     SeniorCitizen = "Y"
     ElseIf SeniorCitizen = "No" Then
     SeniorCitizen = "N"
     ElseIf SeniorCitizen = "Not claiming for Parents" Then
     SeniorCitizen = "P"
    End If

    If SeniorCitizen <> "" And SeniorCitizen <> "(Select)" Then
        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizenFlag") = UCase(SeniorCitizen)
    Else
        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizenFlag") = ""
    End If
    
    If Sheet9.Range("Parents_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("Parents") = Sheet9.Range("Parents_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("Parents") = 0
    End If
    '------------------------------------------
    If Sheet9.Range("Health_Insurance2_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("HlthInsPremParents") = Sheet9.Range("Health_Insurance2_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("HlthInsPremParents") = 0
    End If
    '___________________________________________
    Dim mIntCells80Dc As Variant
Dim InsDtls80Dc, Sec80DParentsHIDtls
Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = New Collection
'Set InsDtls80Dc = CreateObject("Scripting.Dictionary")
Set Sec80DParentsHIDtls = CreateObject("Scripting.Dictionary")
 
 'Sec80DParentsHIDtls("HlthInsPremParents") = UVCase(Sheet9.Range("Health_Insurance2_80D").Value)

If Sheet9.Range("TotAmtA2.80D").Value > 0 Then
 
 mIntCells80Dc = Sheet9.Range("NameInsurerA2.80D").Rows.count
 
For i = 1 To mIntCells80Dc

If Sheet9.Range("NameInsurerA2.80D").Cells(i, 1).Value <> "" Then

 
Dim InsurerName_80Dc, PolicyNo_80Dc, ReceiptNo_80Dc, HealthInsAmt_80Dc


InsurerName_80Dc = Sheet9.Range("NameInsurerA2.80D").Cells(i, 1).Value
PolicyNo_80Dc = Sheet9.Range("PolicyNumA2.80D").Cells(i, 1).Value
'ReceiptNo_80Dc = Sheet9.Range("ReceiptNumA2.80D").Cells(i, 1).Value
HealthInsAmt_80Dc = Sheet9.Range("AmtA2.80D").Cells(i, 1).Value

If InsurerName_80Dc <> "" Then
Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80Dc)
End If

If PolicyNo_80Dc <> "" Then
Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80Dc)
End If

'If ReceiptNo_80Dc <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80Dc)
'End If

If HealthInsAmt_80Dc <> "" Then
Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80Dc)
End If

'End If

                Sch80DInsDtls.add Sch80DInsDtls_G
                Set Sch80DInsDtls_G = Nothing
                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
End If
Next
'End If

                Sec80DParentsHIDtls.add "Sch80DInsDtls", Sch80DInsDtls
                Set Sch80DInsDtls = Nothing
                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")


Sec80DParentsHIDtls("TotalPayments") = UVCase(Sheet9.Range("TotAmtA2.80D").Value)

                Sec80DSelfFamSrCtznHealth.add "Sec80DParentsHIDtls", Sec80DParentsHIDtls
                Set Sec80DParentsHIDtls = Nothing
                Set Sec80DParentsHIDtls = CreateObject("Scripting.Dictionary")


End If


    '-------------------------------------------
    If Sheet9.Range("Preventive_Health2_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParents") = Sheet9.Range("Preventive_Health2_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParents") = 0
    End If
    
    If Sheet9.Range("Parents_SC_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizen") = Sheet9.Range("Parents_SC_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("ParentsSeniorCitizen") = 0
    End If
    '---------------------------------------------
    If Sheet9.Range("Health_Insurance3_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("HlthInsPremParentsSrCtzn") = Sheet9.Range("Health_Insurance3_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("HlthInsPremParentsSrCtzn") = 0
    End If
    '________________________________________________
    Dim mIntCells80Dd As Variant
Dim InsDtls80Dd, Sec80DParentsSrCtznHIDtls
Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")
Set Sch80DInsDtls = New Collection
'Set InsDtls80Dd = CreateObject("Scripting.Dictionary")
Set Sec80DParentsSrCtznHIDtls = CreateObject("Scripting.Dictionary")
 

'Sec80DParentsSrCtznHIDtls("HlthInsPremParentsSrCtzn") = UVCase(Sheet9.Range("Health_Insurance3_80D").Value)

If Sheet9.Range("TotAmtB2.80D").Value > 0 Then
 
 mIntCells80Dd = Sheet9.Range("NameInsurerB2.80D").Rows.count
 
For i = 1 To mIntCells80Dd

If Sheet9.Range("NameInsurerB2.80D").Cells(i, 1).Value <> "" Then

 
Dim InsurerName_80Dd, PolicyNo_80Dd, ReceiptNo_80Dd, HealthInsAmt_80Dd



InsurerName_80Dd = Sheet9.Range("NameInsurerB2.80D").Cells(i, 1).Value
PolicyNo_80Dd = Sheet9.Range("PolicyNumB2.80D").Cells(i, 1).Value
'ReceiptNo_80Dd = Sheet9.Range("ReceiptNumB2.80D").Cells(i, 1).Value
HealthInsAmt_80Dd = Sheet9.Range("AmtB2.80D").Cells(i, 1).Value

If InsurerName_80Dd <> "" Then
Sch80DInsDtls_G("InsurerName") = UCase(InsurerName_80Dd)
End If

If PolicyNo_80Dd <> "" Then
Sch80DInsDtls_G("PolicyNo") = UCase(PolicyNo_80Dd)
End If

'If ReceiptNo_80Dd <> "" Then
'Sch80DInsDtls_G("ReceiptNo") = UCase(ReceiptNo_80Dd)
'End If

If HealthInsAmt_80Dd <> "" Then
Sch80DInsDtls_G("HealthInsAmt") = UVCase(HealthInsAmt_80Dd)
End If

'End If

                Sch80DInsDtls.add Sch80DInsDtls_G
                Set Sch80DInsDtls_G = Nothing
                Set Sch80DInsDtls_G = CreateObject("Scripting.Dictionary")
End If
Next
'End If

                Sec80DParentsSrCtznHIDtls.add "Sch80DInsDtls", Sch80DInsDtls
                Set Sch80DInsDtls = Nothing
                Set Sch80DInsDtls = CreateObject("Scripting.Dictionary")


Sec80DParentsSrCtznHIDtls("TotalPayments") = UVCase(Sheet9.Range("TotAmtB2.80D").Value)

                Sec80DSelfFamSrCtznHealth.add "Sec80DParentsSrCtznHIDtls", Sec80DParentsSrCtznHIDtls
                Set Sec80DParentsSrCtznHIDtls = Nothing
                Set Sec80DParentsSrCtznHIDtls = CreateObject("Scripting.Dictionary")


End If

    '------------------------------------------------
    If Sheet9.Range("Preventive_Health3_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParentsSrCtzn") = Sheet9.Range("Preventive_Health3_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("PrevHlthChckUpParentsSrCtzn") = 0
    End If
    
    If Sheet9.Range("Medical_Expenditure2_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("MedicalExpParentsSrCtzn") = Sheet9.Range("Medical_Expenditure2_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("MedicalExpParentsSrCtzn") = 0
    End If
End If
    
    If Sheet9.Range("Eligible_Amount_80D").Value <> "" Then
        Sec80DSelfFamSrCtznHealth("EligibleAmountOfDedn") = Sheet9.Range("Eligible_Amount_80D").Value
    Else
        Sec80DSelfFamSrCtznHealth("EligibleAmountOfDedn") = 0
    End If

    jsonDictionary.add "Sec80DSelfFamSrCtznHealth", Sec80DSelfFamSrCtznHealth

End If

Set Schedule80D = jsonDictionary
End Function
Function TaxPaid() As Object
subProcCaption = "Tax Paid"
noOfProcessSub = 10
Dim i As Long
Dim jsonDictionary, TaxesPaid
Dim AdvanceTax, TDS, TCS, SelfAssessmentTax, TotalTaxesPaid, BalTaxPayable As Variant

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set TaxesPaid = CreateObject("Scripting.Dictionary")

AdvanceTax = Sheet3.Range("IncD.AdvanceTax").Value
TDS = Sheet3.Range("IncD.TDS").Value
TCS = Sheet3.Range("IncD.TCS").Value
SelfAssessmentTax = Sheet3.Range("IncD.SelfAssessmentTax").Value
TotalTaxesPaid = Sheet3.Range("IncD.TotalTaxesPaid").Value
BalTaxPayable = Sheet3.Range("IncD.BalTaxPayable").Value

        If Trim(Sheet3.Range("IncD.AdvanceTax").Value) <> "" Then
            TaxesPaid("AdvanceTax") = AdvanceTax
        Else
            TaxesPaid("AdvanceTax") = 0
        End If
        
        If Trim(Sheet3.Range("IncD.TDS").Value) <> "" Then
            TaxesPaid("TDS") = TDS
        Else
            TaxesPaid("TDS") = 0
        End If
    
        If Trim(Sheet3.Range("IncD.TCS").Value) <> "" Then
            TaxesPaid("TCS") = TCS
        Else
            TaxesPaid("TCS") = 0
        End If
    
        If Trim(Sheet3.Range("IncD.SelfAssessmentTax").Value) <> "" Then
            TaxesPaid("SelfAssessmentTax") = SelfAssessmentTax
        Else
            TaxesPaid("SelfAssessmentTax") = 0
        End If
        
        If Trim(Sheet3.Range("IncD.TotalTaxesPaid").Value) <> "" Then
            TaxesPaid("TotalTaxesPaid") = TotalTaxesPaid
        Else
            TaxesPaid("TotalTaxesPaid") = 0
        End If
        jsonDictionary.add "TaxesPaid", TaxesPaid
                        
        If Trim(Sheet3.Range("IncD.BalTaxPayable").Value) <> "" Then
           jsonDictionary("BalTaxPayable") = BalTaxPayable
        Else
           jsonDictionary("BalTaxPayable") = 0
        End If
            
Set TaxPaid = jsonDictionary
End Function
Function Refund() As Object

subProcCaption = ""
noOfProcessSub = 10
Dim i As Long
Dim jsonDictionary, BankAccountDtls, useForRefund, AddtnlBank
Dim RefundDue, AddtnlBankDetails, IFSCCode, BankName, BankAccountNo, BankAccnType_BA As Variant

''INC_C43 2024-25 Bindu
'Dim BankAccnType_BA As Variant

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set BankAccountDtls = CreateObject("Scripting.Dictionary")
Set useForRefund = CreateObject("Scripting.Dictionary")
Set AddtnlBank = CreateObject("Scripting.Dictionary")

Set AddtnlBankDetails = New Collection
Set IFSCCode = New Collection

RefundDue = Sheet3.Range("IncD.RefundDue").Value

    jsonDictionary("RefundDue") = RefundDue
                
    If (Not IsEmpty(BankIFSC_BA)) Then
     If (UBound(BankIFSC_BA) > 0) Then
        For i = 1 To UBound(BankIFSC_BA)
        
            AddtnlBank("IFSCCode") = UCase(BankIFSC_BA(i))
            AddtnlBank("BankName") = BankName_BA(i)
            If UCase(BankAccntnum_BA(i)) <> "" Then
                AddtnlBank("BankAccountNo") = UCase(BankAccntnum_BA(i))
            Else
                AddtnlBank("BankAccountNo") = 0
            End If
            
            
'            'INC_C43 2024-25 Bindu
'            BankAccnType_BA = Sheet3.Range("SchBA.AcntType").Cells.item(i).Value
'            If UCase(BankAccnType_BA) = UCase("Savings Account") Then
'               AddtnlBank("AccountType") = "SAV"
'            ElseIf UCase(BankAccnType_BA) = UCase("Current Account") Then
'                AddtnlBank("AccountType") = "CUR"
'            ElseIf UCase(BankAccnType_BA) = UCase("Cash Credit Account") Then
'                AddtnlBank("AccountType") = "CC"
'            ElseIf UCase(BankAccnType_BA) = UCase("Over draft account") Then
'                AddtnlBank("AccountType") = "OD"
'            'ElseIf UCase(BankAccnType_BA) = UCase("Non Resident Account") Then
'            'Non Resident Account,Other
'            ElseIf UCase(Mid(BankAccnType_BA, 1, 3)) = UCase(Mid("Non Resident Account", 1, 3)) Then
'                AddtnlBank("AccountType") = "NR"
'             'ElseIf UCase(BankAccnType_BA) = UCase("Other") Then
'            ElseIf UCase(Mid(BankAccnType_BA, 1, 3)) = UCase(Mid("Other", 1, 3)) Then
'                AddtnlBank("AccountType") = "OTH"
'            End If
            
           'INC_C43 2024-25  Malli
            BankAccnType_BA = Sheet3.Range("SchBA.AcntType").Cells.item(i).Value
            If UCase(BankAccnType_BA) = UCase("Savings Account") Then
               AddtnlBank("AccountType") = "SB"
            ElseIf UCase(BankAccnType_BA) = UCase("Current Account") Then
                AddtnlBank("AccountType") = "CA"
            ElseIf UCase(BankAccnType_BA) = UCase("Cash Credit Account") Then
                AddtnlBank("AccountType") = "CC"
            ElseIf UCase(BankAccnType_BA) = UCase("Over draft account") Then
                AddtnlBank("AccountType") = "OD"
            'ElseIf UCase(BankAccnType_BA) = UCase("Non Resident Account") Then
            'Non Resident Account,Other
            ElseIf UCase(Mid(BankAccnType_BA, 1, 3)) = UCase(Mid("Non Resident Account", 1, 3)) Then
                AddtnlBank("AccountType") = "NRO"
             'ElseIf UCase(BankAccnType_BA) = UCase("Other") Then
            ElseIf UCase(Mid(BankAccnType_BA, 1, 3)) = UCase(Mid("Other", 1, 3)) Then
                AddtnlBank("AccountType") = "OTH"
            End If
            
            
            
            
'Malli--------------------------
'Konda--------Umcometed
       If tempxml(i) = "" Then
          tempxml(i) = "false"
       ElseIf tempxml(i) = True Then
          tempxml(i) = "true"
       ElseIf tempxml(i) = False Then
          tempxml(i) = "false"
       End If

    If tempxml(i) <> "" Then
        AddtnlBank("UseForRefund") = tempxml(i)
    End If
    
 'End-----------------
'Malli---------------------------

    AddtnlBankDetails.add AddtnlBank
    Set AddtnlBank = Nothing
    Set AddtnlBank = CreateObject("Scripting.Dictionary")
        
    Next
    End If
    End If

    BankAccountDtls.add "AddtnlBankDetails", AddtnlBankDetails
    jsonDictionary.add "BankAccountDtls", BankAccountDtls
    
UpdateProgressBar
Set Refund = jsonDictionary
End Function

''
' Convert object (Dictionary/Collection/Array) to JSON
'
' @method ToJson
' @param {Variant} JsonValue (Dictionary, Collection, or Array)
' @param {Integer|String} Whitespace "Pretty" print json with given number of spaces per indentation (Integer) or given string
' @return {String}
''
Public Function ToJson(ByVal JsonValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal json_CurrentIndentation As Long = 0) As String
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long
    Dim json_Index As Long
    Dim json_LBound As Long
    Dim json_UBound As Long
    Dim json_IsFirstItem As Boolean
    Dim json_Index2D As Long
    Dim json_LBound2D As Long
    Dim json_UBound2D As Long
    Dim json_IsFirstItem2D As Boolean
    Dim json_Key As Variant
    Dim json_Value As Variant
    Dim json_DateStr As String
    Dim json_Converted As String
    Dim json_SkipItem As Boolean
    Dim json_PrettyPrint As Boolean
    Dim json_Indentation As String
    Dim json_InnerIndentation As String

    json_LBound = -1
    json_UBound = -1
    json_IsFirstItem = True
    json_LBound2D = -1
    json_UBound2D = -1
    json_IsFirstItem2D = True
    json_PrettyPrint = Not IsMissing(Whitespace)

    Select Case VBA.VarType(JsonValue)
    Case VBA.vbNull
        ToJson = "null"
    Case VBA.vbDate
        ' Date
        json_DateStr = ConvertToIso(VBA.CDate(JsonValue))

        ToJson = """" & json_DateStr & """"
    Case VBA.vbString
        ' String (or large number encoded as string)
        If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
            ToJson = JsonValue
        Else
            ToJson = """" & json_Encode(JsonValue) & """"
        End If
    Case VBA.vbBoolean
        If JsonValue Then
            ToJson = "true"
        Else
            ToJson = "false"
        End If
    Case VBA.vbArray To VBA.vbArray + VBA.vbByte
        If json_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
                json_InnerIndentation = VBA.String$(json_CurrentIndentation + 2, Whitespace)
            Else
                json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
                json_InnerIndentation = VBA.Space$((json_CurrentIndentation + 2) * Whitespace)
            End If
        End If

        ' Array
        json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength

        On Error Resume Next

        json_LBound = LBound(JsonValue, 1)
        json_UBound = UBound(JsonValue, 1)
        json_LBound2D = LBound(JsonValue, 2)
        json_UBound2D = UBound(JsonValue, 2)

        If json_LBound >= 0 And json_UBound >= 0 Then
            For json_Index = json_LBound To json_UBound
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    ' Append comma to previous line
                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                End If

                If json_LBound2D >= 0 And json_UBound2D >= 0 Then
                    ' 2D Array
                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                    End If
                    json_BufferAppend json_Buffer, json_Indentation & "[", json_BufferPosition, json_BufferLength

                    For json_Index2D = json_LBound2D To json_UBound2D
                        If json_IsFirstItem2D Then
                            json_IsFirstItem2D = False
                        Else
                            json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                        End If

                        json_Converted = ToJson(JsonValue(json_Index, json_Index2D), Whitespace, json_CurrentIndentation + 2)

                        ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                        If json_Converted = "" Then
                            ' (nest to only check if converted = "")
                            If json_IsUndefined(JsonValue(json_Index, json_Index2D)) Then
                                json_Converted = "null"
                            End If
                        End If

                        If json_PrettyPrint Then
                            json_Converted = vbNewLine & json_InnerIndentation & json_Converted
                        End If

                        json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                    Next json_Index2D

                    If json_PrettyPrint Then
                        json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength
                    End If

                    json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
                    json_IsFirstItem2D = True
                Else
                    ' 1D Array
                    json_Converted = ToJson(JsonValue(json_Index), Whitespace, json_CurrentIndentation + 1)

                    ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                    If json_Converted = "" Then
                        ' (nest to only check if converted = "")
                        If json_IsUndefined(JsonValue(json_Index)) Then
                            json_Converted = "null"
                        End If
                    End If

                    If json_PrettyPrint Then
                        json_Converted = vbNewLine & json_Indentation & json_Converted
                    End If

                    json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                End If
            Next json_Index
        End If

        On Error GoTo 0

        If json_PrettyPrint Then
            json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
            Else
                json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
            End If
        End If

        json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength

        ToJson = json_BufferToString(json_Buffer, json_BufferPosition)

    ' Dictionary or Collection
    Case VBA.vbObject
        If json_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                json_Indentation = VBA.String$(json_CurrentIndentation + 1, Whitespace)
            Else
                json_Indentation = VBA.Space$((json_CurrentIndentation + 1) * Whitespace)
            End If
        End If

        ' Dictionary
        If VBA.TypeName(JsonValue) = "Dictionary" Then
            json_BufferAppend json_Buffer, "{", json_BufferPosition, json_BufferLength
            For Each json_Key In JsonValue.Keys
                ' For Objects, undefined (Empty/Nothing) is not added to object
                json_Converted = ToJson(JsonValue(json_Key), Whitespace, json_CurrentIndentation + 1)
                If json_Converted = "" Then
                    json_SkipItem = json_IsUndefined(JsonValue(json_Key))
                Else
                    json_SkipItem = False
                End If

                If Not json_SkipItem Then
                    If json_IsFirstItem Then
                        json_IsFirstItem = False
                    Else
                        json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                    End If

                    If json_PrettyPrint Then
                        json_Converted = vbNewLine & json_Indentation & """" & json_Key & """: " & json_Converted
                    Else
                        json_Converted = """" & json_Key & """:" & json_Converted
                    End If

                    json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
                End If
            Next json_Key

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_BufferAppend json_Buffer, json_Indentation & "}", json_BufferPosition, json_BufferLength

        ' Collection
        ElseIf VBA.TypeName(JsonValue) = "Collection" Then
            json_BufferAppend json_Buffer, "[", json_BufferPosition, json_BufferLength
            For Each json_Value In JsonValue
                If json_IsFirstItem Then
                    json_IsFirstItem = False
                Else
                    json_BufferAppend json_Buffer, ",", json_BufferPosition, json_BufferLength
                End If

                json_Converted = ToJson(json_Value, Whitespace, json_CurrentIndentation + 1)

                ' For Arrays/Collections, undefined (Empty/Nothing) is treated as null
                If json_Converted = "" Then
                    ' (nest to only check if converted = "")
                    If json_IsUndefined(json_Value) Then
                        json_Converted = "null"
                    End If
                End If

                If json_PrettyPrint Then
                    json_Converted = vbNewLine & json_Indentation & json_Converted
                End If

                json_BufferAppend json_Buffer, json_Converted, json_BufferPosition, json_BufferLength
            Next json_Value

            If json_PrettyPrint Then
                json_BufferAppend json_Buffer, vbNewLine, json_BufferPosition, json_BufferLength

                If VBA.VarType(Whitespace) = VBA.vbString Then
                    json_Indentation = VBA.String$(json_CurrentIndentation, Whitespace)
                Else
                    json_Indentation = VBA.Space$(json_CurrentIndentation * Whitespace)
                End If
            End If

            json_BufferAppend json_Buffer, json_Indentation & "]", json_BufferPosition, json_BufferLength
        End If

        ToJson = json_BufferToString(json_Buffer, json_BufferPosition)
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' Number (use decimals for numbers)
        ToJson = VBA.Replace(JsonValue, ",", ".")
    Case Else
        ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
        ' Use VBA's built-in to-string
        On Error Resume Next
        ToJson = JsonValue
        On Error GoTo 0
    End Select
End Function
Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See json_ParseNumber)

    Dim json_Length As Long
    Dim json_CharIndex As Long
    json_Length = VBA.Len(json_String)

    ' Length with be at least 16 characters and assume will be less than 100 characters
    If json_Length >= 30 And json_Length <= 100 Then
        Dim json_CharCode As String

        json_StringIsLargeNumber = True

        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_CharIndex
    End If
End Function
Private Function json_Encode(ByVal json_Text As Variant) As String
    ' Reference: http://www.ietf.org/rfc/rfc4627.txt
    ' Escape: ", \, /, backspace, form feed, line feed, carriage return, tab
    Dim json_Index As Long
    Dim json_Char As String
    Dim json_AscCode As Long
    Dim json_Buffer As String
    Dim json_BufferPosition As Long
    Dim json_BufferLength As Long

    For json_Index = 1 To VBA.Len(json_Text)
        json_Char = VBA.Mid$(json_Text, json_Index, 1)
        json_AscCode = VBA.AscW(json_Char)

        ' When AscW returns a negative number, it returns the twos complement form of that number.
        ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
        ' https://support.microsoft.com/en-us/kb/272138
        If json_AscCode < 0 Then
            json_AscCode = json_AscCode + 65536
        End If

        ' From spec, ", \, and control characters must be escaped (solidus is optional)

        Select Case json_AscCode
        Case 34
            ' " -> 34 -> \"
            json_Char = "\"""
        Case 92
            ' \ -> 92 -> \\
            json_Char = "\\"
        Case 47
            ' / -> 47 -> \/ (optional)
            If JsonOptions.EscapeSolidus Then
                json_Char = "\/"
            End If
        Case 8
            ' backspace -> 8 -> \b
            json_Char = "\b"
        Case 12
            ' form feed -> 12 -> \f
            json_Char = "\f"
        Case 10
            ' line feed -> 10 -> \n
            json_Char = "\n"
        Case 13
            ' carriage return -> 13 -> \r
            json_Char = "\r"
        Case 9
            ' tab -> 9 -> \t
            json_Char = "\t"
        Case 0 To 31, 127 To 65535
            ' Non-ascii characters -> convert to 4-digit hex
            json_Char = "\u" & VBA.Right$("0000" & VBA.Hex$(json_AscCode), 4)
        End Select

        json_BufferAppend json_Buffer, json_Char, json_BufferPosition, json_BufferLength
    Next json_Index

    json_Encode = json_BufferToString(json_Buffer, json_BufferPosition)
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
Private Function json_IsUndefined(ByVal json_Value As Variant) As Boolean
    ' Empty / Nothing -> undefined
    Select Case VBA.VarType(json_Value)
    Case VBA.vbEmpty
        json_IsUndefined = True
    Case VBA.vbObject
        Select Case VBA.TypeName(json_Value)
        Case "Empty", "Nothing"
            json_IsUndefined = True
        End Select
    End Select
End Function
Private Function json_BufferToString(ByRef json_Buffer As String, ByVal json_BufferPosition As Long) As String
    If json_BufferPosition > 0 Then
        json_BufferToString = VBA.Left$(json_Buffer, json_BufferPosition)
    End If
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo utc_ErrorHandling

    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")

    Exit Function

utc_ErrorHandling:
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

Function UVCase(valV1 As Variant) As Variant
    If (valV1 = "") Then
        valV1 = 0
    End If
        UVCase = valV1
End Function

Function Schedule80GGC() As Object
Dim i As Long
Dim jsonDictionary, Schedule80GGC_1, AddressDetail
Dim Schedule80GGCDetails, TotalDonationAmtCash80GGC, TotalDonationAmtOtherMode80GGC, TotalDonationsUs80GGC, TotalEligibleDonationAmt80GGC As Variant
Set Schedule80GGCDetails = New Collection
Dim dateDonation
Dim dob As Variant

Set jsonDictionary = CreateObject("Scripting.Dictionary")
Set Schedule80GGC_1 = CreateObject("Scripting.Dictionary")
'Set AddressDetail = CreateObject("Scripting.Dictionary")
'Set Schedule80GGCDetails = CreateObject("Scripting.Dictionary")

TotalDonationAmtCash80GGC = UCase(Range("Total_DonationInCash_80GGC").Value)
TotalDonationAmtOtherMode80GGC = UCase(Range("Total_DonationInOtherMode_80GGC").Value)
TotalDonationsUs80GGC = UCase(Range("Total_Donation_80GGC").Value)
TotalEligibleDonationAmt80GGC = UCase(Range("Total_Donation_Eligible_80GGC").Value)


subProcCaption = "80GGC"
    If Not IsEmpty(DateofDonation_80GGC) Then
    If (UBound(DateofDonation_80GGC) > 0) Then
    
      For i = 1 To UBound(DateofDonation_80GGC)

            dob = Mid(DateofDonation_80GGC(i), 7, 4) & "-" & Mid(DateofDonation_80GGC(i), 4, 2) & "-" & Mid(DateofDonation_80GGC(i), 1, 2)

            If DateofDonation_80GGC(i) <> "" Then
                Schedule80GGC_1("DonationDate") = dob
            End If
            
            If Donation_cash_80GGC(i) <> "" Then
                Schedule80GGC_1("DonationAmtCash") = UVCase(CDbl(Donation_cash_80GGC(i)))
            Else
                Schedule80GGC_1("DonationAmtCash") = 0
            End If
            
            If Donation_other_80GGC(i) <> "" Then
                Schedule80GGC_1("DonationAmtOtherMode") = UVCase(CDbl(Donation_other_80GGC(i)))
            Else
                Schedule80GGC_1("DonationAmtOtherMode") = 0
            End If

            
'            Dim natureoftrans
'
'            If UCase(NatureofTransaction_80GGC(i)) = UCase("UPI transfer") Then
'               natureoftrans = "UPI"
'            ElseIf UCase(NatureofTransaction_80GGC(i)) = UCase("Cheque issued") Then
'               natureoftrans = "Cheque"
'            ElseIf UCase(NatureofTransaction_80GGC(i)) = UCase("IMPS") Then
'               natureoftrans = "IMPS"
'            ElseIf UCase(NatureofTransaction_80GGC(i)) = UCase("NEFT") Then
'               natureoftrans = "NEFT"
'            ElseIf UCase(NatureofTransaction_80GGC(i)) = UCase("RTGS") Then
'               natureoftrans = "RTGS"
'            End If
'
'
'            If NatureofTransaction_80GGC(i) <> "" Then
'                Schedule80GGC_1("NatureOfTrans") = natureoftrans
'            End If
            
            
            
            If BankIFSC_80GGC(i) <> "" Then
                Schedule80GGC_1("IFSCCode") = UCase(BankIFSC_80GGC(i))
            End If
            
'            If BankName_80GGC(i) <> "" Then
'                Schedule80GGC_1("BankName") = UCase(BankName_80GGC(i))
'            End If
            
'            If BankAccntnum_80GGC(i) <> "" Then
'                Schedule80GGC_1("BankAccountNo") = UVCase(CDbl(BankAccntnum_80GGC(i)))
'            End If
            
             If ChequeNumber_80GGC(i) <> "" Then
                Schedule80GGC_1("TransactionRefNum") = Trim(ChequeNumber_80GGC(i))
             End If
            
            
            If Trim(Range("TotalDonation_80GGC").Cells.item(i).Value) <> "" Then
                Schedule80GGC_1("DonationAmt") = UVCase(CDbl(Range("TotalDonation_80GGC").Cells.item(i).Value))
            Else
                Schedule80GGC_1("DonationAmt") = 0
            End If
            
            If Trim(Range("EligibleAmountofDonation_80GGC").Cells.item(i).Value) <> "" Then
                Schedule80GGC_1("EligibleDonationAmt") = UVCase(CDbl(Range("EligibleAmountofDonation_80GGC").Cells.item(i).Value))
            Else
                Schedule80GGC_1("EligibleDonationAmt") = 0
            End If
'Konda updated on 17-03-2026--V0.6
            If Trim(Range("Name_80GGC").Cells.item(i).Value) <> "" Then
                Schedule80GGC_1("PoliticalPartyName") = UCase(Range("Name_80GGC").Cells.item(i).Value)
            End If
            
            If Trim(Range("PAN_80GGC").Cells.item(i).Value) <> "" Then
                Schedule80GGC_1("PoliticalPartyPAN") = UCase(Range("PAN_80GGC").Cells.item(i).Value)
            End If
'====================
        
            Schedule80GGCDetails.add Schedule80GGC_1
            Set Schedule80GGC_1 = Nothing
            Set Schedule80GGC_1 = CreateObject("Scripting.Dictionary")
    
            UpdateProgressBar
            Next
        End If
    End If
    jsonDictionary.add "Schedule80GGCDetails", Schedule80GGCDetails
            
            If Trim(Range("Total_DonationInCash_80GGC").Value) > 0 Then
                jsonDictionary("TotalDonationAmtCash80GGC") = UVCase(CDbl(TotalDonationAmtCash80GGC))
            Else
                jsonDictionary("TotalDonationAmtCash80GGC") = 0
            End If
            
            If Trim(Range("Total_DonationInOtherMode_80GGC").Value) > 0 Then
                jsonDictionary("TotalDonationAmtOtherMode80GGC") = UVCase(CDbl(TotalDonationAmtOtherMode80GGC))
            Else
                jsonDictionary("TotalDonationAmtOtherMode80GGC") = 0
            End If
            
            If Trim(Range("Total_Donation_80GGC").Value) > 0 Then
                jsonDictionary("TotalDonationsUs80GGC") = UVCase(CDbl(TotalDonationsUs80GGC))
            Else
                jsonDictionary("TotalDonationsUs80GGC") = 0
            End If
  
            If Trim(Range("Total_Donation_Eligible_80GGC").Value) > 0 Then
                jsonDictionary("TotalEligibleDonationAmt80GGC") = UVCase(CDbl(TotalEligibleDonationAmt80GGC))
            Else
                jsonDictionary("TotalEligibleDonationAmt80GGC") = 0
            End If
Set Schedule80GGC = jsonDictionary
End Function

Function Schedule80U_1() As Object
subProcCaption = "Schedule 80U"
noOfProcessSub = 10
Dim jsonDictionary, Schedule80U_D

 Dim NatureOfDisability, DeductionAmount, Form10IAFilingDate, Form10IAAckNum, UDIDNum As Variant
 
 Dim NatureOfDisability_1, DeductionAmount_1, Form10IAFilingDate_1, Form10IAAckNum_1, UDIDNum_1 As Variant
 
 Set jsonDictionary = CreateObject("Scripting.Dictionary")
' Set Schedule80U = CreateObject("Scripting.Dictionary")
' Set Schedule80U_D = CreateObject("Scripting.Dictionary")
 
' MsgBox Nature_of_disability_80U
 
 NatureOfDisability_1 = Sheet14.Range("Naturedisability_80U").Cells(1, 1).Value
 'MsgBox (NatureOfDisability_1)
 DeductionAmount_1 = Sheet14.Range("Amtdeduction_80U").Value
' Form10IAFilingDate_1 = Sheet14.Range("DatefilingFm10IA_80U").Value
 Form10IAAckNum_1 = Sheet14.Range("AckNoFm10IAfiled_80U").Value
 UDIDNum_1 = Sheet14.Range("UDIDNum_80U").Value
 
 
 'Malli---------19/04/2025
 'TypeOfDisability ,FormAckNum11A
 Dim TypeOfDisability_1, FormAckNum11A_1, TypeOfDisability_1_80U
 TypeOfDisability_1 = Sheet14.Range("Disability_80U").Value
' FormAckNum11A_1 = Sheet14.Range("AcknowledgeNum11A2_80U").Value
 
 
 '------------------------

Dim Ndisability80U As Variant
Ndisability80U = Mid(NatureOfDisability_1, 1, 1)

If IsNumeric(Ndisability80U) Then

If Ndisability80U = "1" Or Ndisability80U = "2" Then
jsonDictionary("NatureOfDisability") = Ndisability80U
'Else: jsonDictionary("NatureOfDisability") = ""
End If
 
 'Malli-----------19/04/2025
 If TypeOfDisability_1 <> "" And TypeOfDisability_1 <> "(Select)" Then
'Konda_06/05/2025_Commented and updated as per New Schema_v0.7 and V0.7.1
'Konda---------as per Clarification tracker-14/05/2025
'If TypeOfDisability_1 = "(i) autism, cerebral palsy, or multiple disabilities and" Then
If TypeOfDisability_1 = "(i) autism, cerebral palsy, or multiple disabilities" Then
    TypeOfDisability_1_80U = "1"

 ElseIf TypeOfDisability_1 = "(ii) others" Then
    TypeOfDisability_1_80U = "2"
 
' If TypeOfDisability_1 = "(i) blindness" Then
' TypeOfDisability_1_80U = "1"
'
' ElseIf TypeOfDisability_1 = "(ii) low vision" Then
' TypeOfDisability_1_80U = "2"
'
' ElseIf TypeOfDisability_1 = "(iii) leprosy-cured" Then
' TypeOfDisability_1_80U = "3"
'
' ElseIf TypeOfDisability_1 = "(iv) hearing impairment" Then
' TypeOfDisability_1_80U = "4"
'
' ElseIf TypeOfDisability_1 = "(v) locomotor disability" Then
' TypeOfDisability_1_80U = "5"
'
' ElseIf TypeOfDisability_1 = "(vi) mental retardation" Then
' TypeOfDisability_1_80U = "6"
'
' ElseIf TypeOfDisability_1 = "(vii) mental illness" Then
' TypeOfDisability_1_80U = "7"
'
' ElseIf TypeOfDisability_1 = "(viii) autism" Then
' TypeOfDisability_1_80U = "8"
'
' ElseIf TypeOfDisability_1 = "(ix) cerebral palsy" Then
' TypeOfDisability_1_80U = "9"
'
' ElseIf TypeOfDisability_1 = "(x) multiple disability" Then
' TypeOfDisability_1_80U = "10"
 
 Else
 TypeOfDisability_1_80U = ""
 End If
 
 End If
 
 If TypeOfDisability_1_80U <> "" Then
 jsonDictionary("TypeOfDisability") = UCase(TypeOfDisability_1_80U)
 End If
 '--------------------------
 
If Sheet14.Range("Amtdeduction_80U").Value <> "" And Sheet14.Range("Amtdeduction_80U").Value <> "0" Then
jsonDictionary("DeductionAmount") = UVCase(CDbl(DeductionAmount_1))
Else
jsonDictionary("DeductionAmount") = 0
End If

'If Sheet14.Range("DatefilingFm10IA_80U").Value <> "" Then
'Dim dob As Variant
'dob = Mid(Form10IAFilingDate_1, 7, 4) & "-" & Mid(Form10IAFilingDate_1, 4, 2) & "-" & Mid(Form10IAFilingDate_1, 1, 2)
'jsonDictionary("Form10IAFilingDate") = dob
'End If

If Sheet14.Range("AckNoFm10IAfiled_80U").Value <> "" Then
'jsonDictionary("Form10IAAckNum") = UVCase(Form10IAAckNum_1)
jsonDictionary("Form10IAAckNum") = UCase(Form10IAAckNum_1)
End If

'If Sheet14.Range("UDIDNum_80U").Value <> "" Then
'Schedule80U_D(UDIDNum) = UVCase(UDIDNum_1)
'End If
'Konda_06/05/2025_Commented and updated as per New Schema_v0.7 and V0.7.1
'Malli------------19/04/2025
'If FormAckNum11A_1 <> "" Then
'jsonDictionary("FormAckNum11A") = UCase(FormAckNum11A_1)
'End If
'---------------------------

If Sheet14.Range("UDIDNum_80U").Value <> "" Then
'jsonDictionary("UDIDNum") = UVCase(UDIDNum_1)
jsonDictionary("UDIDNum") = UCase(UDIDNum_1)
End If

'jsonDictionary.add "Schedule80U", Schedule80U_D

Set Schedule80U_1 = jsonDictionary

End If
'jsonDictionary.add Schedule80U_D
'Set Schedule80U = jsonDictionary


End Function
Function Schedule80DD_1() As Object
subProcCaption = "Schedule 80DD"
noOfProcessSub = 10
Dim jsonDictionary, Schedule80DD_D

Dim NatureOfDisability, DeductionAmount, DependentType, DependentPan, DependentAadhaar, Form10IAFilingDate, Form10IAAckNum, UDIDNum As Variant
Dim NatureOfDisability_1, DeductionAmount_1, DependentType_1, DependentPan_1, DependentAadhaar_1, Form10IAFilingDate_1, Form10IAAckNum_1, UDIDNum_1 As Variant
'Dim Type_dependent_80DD_1, Aadhaar_dependent_80DD_1, PAN_dependent_80DD_1, UDID_Num_80DD_1, AckNo_ofForm10IAfiled_80DD_1, Date_filingFm10IA_80DD_1, Amount_of_deduction_80DD_1, Nature_of_disability_80DD_1          As Variant


Set jsonDictionary = CreateObject("Scripting.Dictionary")
'Set Schedule80DD_D = CreateObject("Scripting.Dictionary")

NatureOfDisability_1 = Sheet14.Range("Naturedisability_80DD").Cells(1, 1).Value
DeductionAmount_1 = Sheet14.Range("Amtdeduction_80DD").Value
DependentType_1 = Sheet14.Range("Typedependent_80DD").Value
DependentPan_1 = Sheet14.Range("PANdependent_80DD").Value
DependentAadhaar_1 = Sheet14.Range("Aadhaardependent_80DD").Value
'DependentPan_1 = Sheet14.Range("PAN_dependent_80DD").Value
'Form10IAFilingDate_1 = Sheet14.Range("DatefilingFm10IA_80DD").Value
Form10IAAckNum_1 = Sheet14.Range("AckNoFm10IAfiled_80DD").Value
UDIDNum_1 = Sheet14.Range("UDIDNum_80DD").Value
 
 
 'Malli---------19/04/2025
 Dim TypeOfDisability_1, TypeOfDisability_1_80DD, FormAckNum11A_1
 
 TypeOfDisability_1 = Sheet14.Range("Disability_80DD").Value
' FormAckNum11A_1 = Sheet14.Range("AcknowledgeNum11A2_80DD").Value
 
 '------------------------
 
Dim Ndisability80DD As Variant  '---1
Ndisability80DD = Mid(NatureOfDisability_1, 1, 1)

If IsNumeric(Ndisability80DD) Then

If Ndisability80DD = "1" Or Ndisability80DD = "2" Then
jsonDictionary("NatureOfDisability") = Ndisability80DD
'Else: jsonDictionary("NatureOfDisability") = ""
End If


'Malli--------19/04/2025
If TypeOfDisability_1 <> "" And TypeOfDisability_1 <> "(Select)" Then

'Konda_06/05/2025_Commented and updated as per New Schema_v0.7 and V0.7.1
'Konda---------as per Clarification tracker-14/05/2025
'If TypeOfDisability_1 = "(i) autism, cerebral palsy, or multiple disabilities and" Then
If TypeOfDisability_1 = "(i) autism, cerebral palsy, or multiple disabilities" Then
TypeOfDisability_1_80DD = "1"

ElseIf TypeOfDisability_1 = "(ii) others" Then
TypeOfDisability_1_80DD = "2"

'---------------------------------
'If TypeOfDisability_1 = "(i) blindness" Then
'TypeOfDisability_1_80DD = "1"
'
'ElseIf TypeOfDisability_1 = "(ii) low vision" Then
'TypeOfDisability_1_80DD = "2"
'
'ElseIf TypeOfDisability_1 = "(iii) leprosy-cured" Then
'TypeOfDisability_1_80DD = "3"
'
'ElseIf TypeOfDisability_1 = "(iv) hearing impairment" Then
'TypeOfDisability_1_80DD = "4"
'
'ElseIf TypeOfDisability_1 = "(v) locomotor disability" Then
'TypeOfDisability_1_80DD = "5"
'
'ElseIf TypeOfDisability_1 = "(vi) mental retardation" Then
'TypeOfDisability_1_80DD = "6"
'
'ElseIf TypeOfDisability_1 = "(vii) mental illness" Then
'TypeOfDisability_1_80DD = "7"
'
'ElseIf TypeOfDisability_1 = "(viii) autism" Then
'TypeOfDisability_1_80DD = "8"
'
'ElseIf TypeOfDisability_1 = "(ix) cerebral palsy" Then
'TypeOfDisability_1_80DD = "9"
'
'ElseIf TypeOfDisability_1 = "(x) multiple disability" Then
'TypeOfDisability_1_80DD = "10"
Else
TypeOfDisability_1_80DD = ""
End If

End If

'------2
If TypeOfDisability_1_80DD <> "" Then
jsonDictionary("TypeOfDisability") = UCase(TypeOfDisability_1_80DD)
End If

'-----------------------

'------3
If Sheet14.Range("Amtdeduction_80DD").Value <> "" And Sheet14.Range("Amtdeduction_80DD").Value <> "0" Then
jsonDictionary("DeductionAmount") = UVCase(CDbl(DeductionAmount_1))
Else
jsonDictionary("DeductionAmount") = 0
End If
'-------4
Dim Tdept_typ80DD
Tdept_typ80DD = Mid(DependentType_1, 1, 1)
If Tdept_typ80DD <> "" And Tdept_typ80DD <> "(" Then
jsonDictionary("DependentType") = Tdept_typ80DD
'ElseIf Tdept_typ80DD <> "" Or Tdept_typ80DD <> "(" Then
'jsonDictionary("DependentType") = ""
End If
'--------5
If Sheet14.Range("PANdependent_80DD").Value <> "" Then
'jsonDictionary("DependentPan") = UVCase(DependentPan_1)
jsonDictionary("DependentPan") = UCase(DependentPan_1)
End If
'--------6
If Sheet14.Range("Aadhaardependent_80DD").Value <> "" Then
'jsonDictionary("DependentAadhaar") = UVCase(DependentAadhaar_1)
jsonDictionary("DependentAadhaar") = UCase(DependentAadhaar_1)
End If
'------7
'If Sheet14.Range("DatefilingFm10IA_80DD").Value <> "" Then
'Dim dob As Variant
'dob = Mid(Form10IAFilingDate_1, 7, 4) & "-" & Mid(Form10IAFilingDate_1, 4, 2) & "-" & Mid(Form10IAFilingDate_1, 1, 2)
'jsonDictionary("Form10IAFilingDate") = dob
'End If
'------8
If Sheet14.Range("AckNoFm10IAfiled_80DD").Value <> "" Then
'jsonDictionary("Form10IAAckNum") = UVCase(Form10IAAckNum_1)
jsonDictionary("Form10IAAckNum") = UCase(Form10IAAckNum_1)
End If

'------9
'Konda updated as new schem for V0.7 and V0.7.1 on 07-05-2025
'Malli------19/04/2025
'If FormAckNum11A_1 <> "" Then
'jsonDictionary("FormAckNum11A") = UCase(FormAckNum11A_1)
'End If
'---------


'---------10
If Sheet14.Range("UDIDNum_80DD").Value <> "" Then
'jsonDictionary("UDIDNum") = UVCase(UDIDNum_1)
jsonDictionary("UDIDNum") = UCase(UDIDNum_1)
End If

'jsonDictionary.add "", Schedule80DD_D
Set Schedule80DD_1 = jsonDictionary

End If

End Function

'Malli------------09/04/2025
'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80E() As Object
Dim i As Long
Dim mIntCells80E As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80EDtls_G, Schedule80EDtls
Set Schedule80EDtls_G = CreateObject("Scripting.Dictionary")

Set Schedule80EDtls = CreateObject("Scripting.Dictionary")
Set Schedule80EDtls = New Collection

Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
subProcCaption = "80E"

  
  If Sheet17.Range("TotAmt.80E").Value > 0 Then
     
     mIntCells80E = Sheet17.Range("LoanfrmBankOrInstitute.80E").count
  
  For i = 1 To mIntCells80E
  
  If (Sheet17.Range("LoanfrmBankOrInstitute.80E").Cells(i, 1).Value <> "" And Sheet17.Range("LoanfrmBankOrInstitute.80E").Cells(i, 1).Value <> "(Select)") Then
   
   LoanTknFrom_G = Sheet17.Range("LoanfrmBankOrInstitute.80E").Cells(i, 1).Value
'   IFSCCode_G = Sheet17.Range("IFSC.80E").Cells(i, 1).Value  'Ankita_05/05/2025_Commented as per DESheet_v0.7
   BankOrInstnName_G = Sheet17.Range("bankName.80E").Cells(i, 1).Value
'   PAN_G = Sheet17.Range("PAN.80E").Cells(i, 1).Value   'Ankita_05/05/2025_Commented as per DESheet_v0.7
   LoanAccNoOfBankOrInstnRefNo_G = Sheet17.Range("loanAccNum.80E").Cells(i, 1).Value
   DateofLoan_G = Sheet17.Range("loanDate.80E").Cells(i, 1).Value
   TotalLoanAmt_G = Sheet17.Range("loanAmt.80E").Cells(i, 1).Value
   LoanOutstndngAmt_G = Sheet17.Range("loanOutstanding.80E").Cells(i, 1).Value
   InterestPaid_G = Sheet17.Range("Intrst.80E").Cells(i, 1).Value
   
    
   
   If LoanTknFrom_G <> "" Then
            If LoanTknFrom_G = "Bank" Then
            LoanTknFrom_G = "B"
     'Konda updated on 27-04-2025 SIT-92116
            'ElseIf LoanTknFrom_G = "Instituition" Then
            ElseIf LoanTknFrom_G = "Institution" Then
        '-----------
            LoanTknFrom_G = "I"
            End If
        Schedule80EDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        Schedule80EDtls_G("LoanTknFrom") = ""
   End If
'Konda updated as new schem for V0.7 and V0.7.1 on 07-05-2025
'   If IFSCCode_G <> "" Then
'   Schedule80EDtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
'   End If
'
'   If PAN_G <> "" Then
'   Schedule80EDtls_G("PAN") = UCase(Trim(PAN_G))
'   End If
   
   If BankOrInstnName_G <> "" Then
   Schedule80EDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   Schedule80EDtls_G("BankOrInstnName") = ""
   End If
   
   
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   Schedule80EDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   Schedule80EDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   Schedule80EDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   Schedule80EDtls_G("DateofLoan") = ""
   End If
   
    
   Schedule80EDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   Schedule80EDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
   'Schedule80EDtls_G("InterestPaid") = UVCase(InterestPaid_G)
    Schedule80EDtls_G("Interest80E") = UVCase(InterestPaid_G)
   
  'End If
  
                Schedule80EDtls.add Schedule80EDtls_G
                Set Schedule80EDtls_G = Nothing
                Set Schedule80EDtls_G = CreateObject("Scripting.Dictionary")
  End If
  Next
                jsonDictionary.add "Schedule80EDtls", Schedule80EDtls
                Set Schedule80EDtls = Nothing
                Set Schedule80EDtls = CreateObject("Scripting.Dictionary")

  'End If
        
TotalInterestPaid_G = Sheet17.Range("TotAmt.80E").Value
'jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
jsonDictionary("TotalInterest80E") = UVCase(TotalInterestPaid_G)

Set Schedule80E = jsonDictionary
 End If
 
End Function

'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80EE() As Object
Dim i As Long
Dim mIntCells80EE As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80EEDtls_G, Schedule80EEDtls
Set Schedule80EEDtls_G = CreateObject("Scripting.Dictionary")
 

Set Schedule80EEDtls = CreateObject("Scripting.Dictionary")
Set Schedule80EEDtls = New Collection

Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
subProcCaption = "80EE"


'Malli------23/04/2025
'Ankita_05/05/2025_Commented as per DESheet_v0.7
'If Sheet17.Range("ResidentialHP.80E").Value > 0 Then
'--------------------------------------------
'  'ResHPVal
'  jsonDictionary("ResHPVal") = UVCase(Sheet17.Range("ResidentialHP.80E").Value)
  
  If Sheet17.Range("TotAmt.80EE").Value > 0 Then
     
     mIntCells80EE = Sheet17.Range("LoanfrmBankOrInstitute.80EE").count
  
  For i = 1 To mIntCells80EE
  
  If (Sheet17.Range("LoanfrmBankOrInstitute.80EE").Cells(i, 1).Value <> "" And Sheet17.Range("LoanfrmBankOrInstitute.80EE").Cells(i, 1).Value <> "(Select)") Then
  
   LoanTknFrom_G = Sheet17.Range("LoanfrmBankOrInstitute.80EE").Cells(i, 1).Value
'   IFSCCode_G = Sheet17.Range("IFSC.80EE").Cells(i, 1).Value  'Ankita_05/05/2025_Commented as per DESheet_v0.7
   BankOrInstnName_G = Sheet17.Range("bankName.80EE").Cells(i, 1).Value
'   PAN_G = Sheet17.Range("PAN.80EE").Cells(i, 1).Value  'Ankita_05/05/2025_Commented as per DESheet_v0.7
   LoanAccNoOfBankOrInstnRefNo_G = Sheet17.Range("loanAccNum.80EE").Cells(i, 1).Value
   DateofLoan_G = Sheet17.Range("loanDate.80EE").Cells(i, 1).Value
   TotalLoanAmt_G = Sheet17.Range("loanAmt.80EE").Cells(i, 1).Value
   LoanOutstndngAmt_G = Sheet17.Range("loanOutstanding.80EE").Cells(i, 1).Value
   InterestPaid_G = Sheet17.Range("Intrst.80EE").Cells(i, 1).Value
   
  
   
   If LoanTknFrom_G <> "" Then
            If LoanTknFrom_G = "Bank" Then
            LoanTknFrom_G = "B"
    'Konda updated on 27-04-2025 SIT-92116
            'ElseIf LoanTknFrom_G = "Instituition" Then
            ElseIf LoanTknFrom_G = "Institution" Then
    'end-----------------
        
            LoanTknFrom_G = "I"
            End If
        Schedule80EEDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        Schedule80EEDtls_G("LoanTknFrom") = ""
   End If
'Konda updated as new schem for V0.7 and V0.7.1 on 07-05-2025
'   If IFSCCode_G <> "" Then
'   Schedule80EEDtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
'   End If
'
'   If PAN_G <> "" Then
'   Schedule80EEDtls_G("PAN") = UCase(Trim(PAN_G))
'   End If
   
   If BankOrInstnName_G <> "" Then
   Schedule80EEDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   Schedule80EEDtls_G("BankOrInstnName") = ""
   End If
   
   
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   Schedule80EEDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   Schedule80EEDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   Schedule80EEDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   Schedule80EEDtls_G("DateofLoan") = ""
   End If
   
    
   Schedule80EEDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   Schedule80EEDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
   'Schedule80EEDtls_G("InterestPaid") = UVCase(InterestPaid_G)
    Schedule80EEDtls_G("Interest80EE") = UVCase(InterestPaid_G)
   
  'End If
  
                Schedule80EEDtls.add Schedule80EEDtls_G
                Set Schedule80EEDtls_G = Nothing
                Set Schedule80EEDtls_G = CreateObject("Scripting.Dictionary")
  End If
  Next
                jsonDictionary.add "Schedule80EEDtls", Schedule80EEDtls
                Set Schedule80EEDtls = Nothing
                Set Schedule80EEDtls = CreateObject("Scripting.Dictionary")

  'End If       'Konda updated as new schem for V0.7 and V0.7.1 on 07-05-2025
        
TotalInterestPaid_G = Sheet17.Range("TotAmt.80EE").Value
'jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
jsonDictionary("TotalInterest80EE") = UVCase(TotalInterestPaid_G)

Set Schedule80EE = jsonDictionary
End If

End Function
'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80EEA() As Object
Dim i As Long
Dim mIntCells80EEA As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80EEADtls_G, Schedule80EEADtls
Set Schedule80EEADtls_G = CreateObject("Scripting.Dictionary")
 

Set Schedule80EEADtls = CreateObject("Scripting.Dictionary")
Set Schedule80EEADtls = New Collection

Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
subProcCaption = "80EEA"
   
   'Malli----------23/04/2025
  If Sheet17.Range("Stampduty.80EEA").Value > 0 Then
  '-----------------------------
   
   'PropStmpDtyVal
   jsonDictionary("PropStmpDtyVal") = UVCase(Sheet17.Range("Stampduty.80EEA").Value)
   
  If Sheet17.Range("TotAmt.80EEA").Value > 0 Then
     
     mIntCells80EEA = Sheet17.Range("LoanfrmBankOrInstitute.80EEA").count
  
  For i = 1 To mIntCells80EEA
  
  If (Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Cells(i, 1).Value <> "" And Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Cells(i, 1).Value <> "(Select)") Then
  
   
   
   LoanTknFrom_G = Sheet17.Range("LoanfrmBankOrInstitute.80EEA").Cells(i, 1).Value
'   IFSCCode_G = Sheet17.Range("IFSC.80EEA").Cells(i, 1).Value 'Ankita_05/05/2025_Commented as per DESheet_v0.7
   BankOrInstnName_G = Sheet17.Range("bankName.80EEA").Cells(i, 1).Value
'   PAN_G = Sheet17.Range("PAN.80EEA").Cells(i, 1).Value      'Ankita_05/05/2025_Commented as per DESheet_v0.7
   LoanAccNoOfBankOrInstnRefNo_G = Sheet17.Range("loanAccNum.80EEA").Cells(i, 1).Value
   DateofLoan_G = Sheet17.Range("loanDate.80EEA").Cells(i, 1).Value
   TotalLoanAmt_G = Sheet17.Range("loanAmt.80EEA").Cells(i, 1).Value
   LoanOutstndngAmt_G = Sheet17.Range("loanOutstanding.80EEA").Cells(i, 1).Value
   InterestPaid_G = Sheet17.Range("Intrst.80EEA").Cells(i, 1).Value
   
  
   
   If LoanTknFrom_G <> "" Then
            If LoanTknFrom_G = "Bank" Then
            LoanTknFrom_G = "B"
            
        'Konda updated on 27-04-2025 SIT-92116
            'ElseIf LoanTknFrom_G = "Instituition" Then
            ElseIf LoanTknFrom_G = "Institution" Then
        '-----------
            LoanTknFrom_G = "I"
            End If
        Schedule80EEADtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        Schedule80EEADtls_G("LoanTknFrom") = ""
   End If
'Konda updated as new schem for V0.7 and V0.7.1 on 07-05-2025
'   If IFSCCode_G <> "" Then
'   Schedule80EEADtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
'   End If
'
'    If PAN_G <> "" Then
'   Schedule80EEADtls_G("PAN") = UCase(Trim(PAN_G))
'   End If
   
   If BankOrInstnName_G <> "" Then
   Schedule80EEADtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   Schedule80EEADtls_G("BankOrInstnName") = ""
   End If
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   Schedule80EEADtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   Schedule80EEADtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   Schedule80EEADtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   Schedule80EEADtls_G("DateofLoan") = ""
   End If
   
    
   Schedule80EEADtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   Schedule80EEADtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
   'Schedule80EEADtls_G("InterestPaid") = UVCase(InterestPaid_G)
   Schedule80EEADtls_G("Interest80EEA") = UVCase(InterestPaid_G)
  'End If
  
                Schedule80EEADtls.add Schedule80EEADtls_G
                Set Schedule80EEADtls_G = Nothing
                Set Schedule80EEADtls_G = CreateObject("Scripting.Dictionary")
  End If
  Next
                jsonDictionary.add "Schedule80EEADtls", Schedule80EEADtls
                Set Schedule80EEADtls = Nothing
                Set Schedule80EEADtls = CreateObject("Scripting.Dictionary")
  End If
        

TotalInterestPaid_G = Sheet17.Range("TotAmt.80EEA").Value
'jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
jsonDictionary("TotalInterest80EEA") = UVCase(TotalInterestPaid_G)

Set Schedule80EEA = jsonDictionary
End If

End Function
'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80EEB() As Object
Dim i As Long
Dim mIntCells80EEB As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80EEBDtls_G, Schedule80EEBDtls
Set Schedule80EEBDtls_G = CreateObject("Scripting.Dictionary")
'Set Schedule80EDtls_G = New Collection

Set Schedule80EEBDtls = CreateObject("Scripting.Dictionary")
Set Schedule80EEBDtls = New Collection

Dim VehicleRegNo_G, VehicleValue_G, LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
subProcCaption = "80EEB"

  
  If Sheet17.Range("TotAmt.80EEB").Value > 0 Then
     
     mIntCells80EEB = Sheet17.Range("LoanfrmBankOrInstitute.80EEB").count
  
  For i = 1 To mIntCells80EEB
  
  If (Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Cells(i, 1).Value <> "" And Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Cells(i, 1).Value <> "(Select)") Then
  
   
   
   LoanTknFrom_G = Sheet17.Range("LoanfrmBankOrInstitute.80EEB").Cells(i, 1).Value
'   IFSCCode_G = Sheet17.Range("IFSC.80EEB").Cells(i, 1).Value     'Ankita_05/05/2025_Commented as per DESheet_v0.7
   BankOrInstnName_G = Sheet17.Range("bankName.80EEB").Cells(i, 1).Value
'   PAN_G = Sheet17.Range("PAN.80EEB").Cells(i, 1).Value       'Ankita_05/05/2025_Commented as per DESheet_v0.7
   LoanAccNoOfBankOrInstnRefNo_G = Sheet17.Range("loanAccNum.80EEB").Cells(i, 1).Value
   DateofLoan_G = Sheet17.Range("loanDate.80EEB").Cells(i, 1).Value
   TotalLoanAmt_G = Sheet17.Range("loanAmt.80EEB").Cells(i, 1).Value
   LoanOutstndngAmt_G = Sheet17.Range("loanOutstanding.80EEB").Cells(i, 1).Value
'   VehicleValue_G = Sheet17.Range("Vehicle_value.80EEB").Cells(i, 1).Value  'Ankita_05/05/2025_Commented as per DESheet_v0.7
   VehicleRegNo_G = Sheet17.Range("VehicleRegNum.80EEB").Cells(i, 1).Value
   InterestPaid_G = Sheet17.Range("Intrst.80EEB").Cells(i, 1).Value
   
  
   
   If LoanTknFrom_G <> "" Then
            If LoanTknFrom_G = "Bank" Then
            LoanTknFrom_G = "B"
            
        'Konda updated on 27-04-2025 SIT-92116
            'ElseIf LoanTknFrom_G = "Instituition" Then
            ElseIf LoanTknFrom_G = "Institution" Then
        '-----------
            LoanTknFrom_G = "I"
            End If
        Schedule80EEBDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
   Else
        Schedule80EEBDtls_G("LoanTknFrom") = ""
   End If
'Konda updated as new schem for V0.7 and V0.7.1 on 07-05-2025
'   If IFSCCode_G <> "" Then
'   Schedule80EEBDtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
'   End If
   
'   If PAN_G <> "" Then
'   Schedule80EEBDtls_G("PAN") = UCase(Trim(PAN_G))
'   End If
   
   If BankOrInstnName_G <> "" Then
   Schedule80EEBDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
   Else
   Schedule80EEBDtls_G("BankOrInstnName") = ""
   End If
   
   
   
   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
   Schedule80EEBDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
   Else
   Schedule80EEBDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
   End If
   
   If DateofLoan_G <> "" Then
   Schedule80EEBDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
   Else
   Schedule80EEBDtls_G("DateofLoan") = ""
   End If
   
    
   Schedule80EEBDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
   Schedule80EEBDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
'Konda updated as new schem for V0.7 and V0.7.1 on 07-05-2025
   'Schedule80EEBDtls_G("VehicleValue") = UVCase(VehicleValue_G)
   '-----------------------------
   
   If VehicleRegNo_G <> "" Then
   Schedule80EEBDtls_G("VehicleRegNo") = UCase(Trim(VehicleRegNo_G))
   Else
   Schedule80EEBDtls_G("VehicleRegNo") = ""
   End If
   
   'Schedule80EEBDtls_G("InterestPaid") = UVCase(InterestPaid_G)
    Schedule80EEBDtls_G("Interest80EEB") = UVCase(InterestPaid_G)
  'End If
  
                Schedule80EEBDtls.add Schedule80EEBDtls_G
                Set Schedule80EEBDtls_G = Nothing
                Set Schedule80EEBDtls_G = CreateObject("Scripting.Dictionary")
  
  End If
  Next
  
               jsonDictionary.add "Schedule80EEBDtls", Schedule80EEBDtls
               Set Schedule80EEBDtls = Nothing
               Set Schedule80EEBDtls = CreateObject("Scripting.Dictionary")
 ' End If
           

TotalInterestPaid_G = Sheet17.Range("TotAmt.80EEB").Value
'jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
jsonDictionary("TotalInterest80EEB") = UVCase(TotalInterestPaid_G)

Set Schedule80EEB = jsonDictionary
End If
End Function


'Malli_80EE_AY_2025_26-------09/04/2025
Function Schedule80C() As Object
Dim i As Long
Dim mIntCells80C As Variant
Dim jsonDictionary
Set jsonDictionary = CreateObject("Scripting.Dictionary")

Dim Schedule80CDtls_G, Schedule80CDtls
Set Schedule80CDtls_G = CreateObject("Scripting.Dictionary")
 
Set Schedule80CDtls = CreateObject("Scripting.Dictionary")
Set Schedule80CDtls = New Collection

Dim NatureOfPayment_G, IdentificationNo_G, Amount_G, TotalAmt_G
subProcCaption = "80C"

  
If Sheet15.Range("TotAmount.80C").Value > 0 Then
'Konda updated as new schem for V0.7 and V0.7.1
     'mIntCells80C = Sheet15.Range("NaturePayment.80C").count
     mIntCells80C = Sheet15.Range("Amount.80C").count
  
   For i = 1 To mIntCells80C
   'If (Sheet15.Range("NaturePayment.80C").Cells(i, 1).Value <> "" And Sheet15.Range("NaturePayment.80C").Cells(i, 1).Value <> "(Select)") Then

'Konda updated as new schem for V0.7 and V0.7.1 on 07-05-2025
'    If Sheet15.Range("NaturePayment.80C").Cells(i, 1).Value <> "" Then

  If Sheet15.Range("Amount.80C").Cells(i, 1).Value <> "" Then
  
'   NatureOfPayment_G = Sheet15.Range("NaturePayment.80C").Cells(i, 1).Value
   IdentificationNo_G = Sheet15.Range("Identification_Number.80C").Cells(i, 1).Value
   Amount_G = Sheet15.Range("Amount.80C").Cells(i, 1).Value
   
'Konda updated as new schem for V0.7 and V0.7.1 on 07-05-2025
'   If NatureOfPayment_G <> "" Then
'    Schedule80CDtls_G("NatureOfPayment") = UCase(Trim(NatureOfPayment_G))
'   Else
'   Schedule80CDtls_G("NatureOfPayment") = ""
'   End If
    
     Schedule80CDtls_G("Amount") = UVCase(Amount_G)
    
   If IdentificationNo_G <> "" Then
    Schedule80CDtls_G("IdentificationNo") = UCase(Trim(IdentificationNo_G))
   Else
   Schedule80CDtls_G("IdentificationNo") = ""
   End If
   
   
    
   'End If
                Schedule80CDtls.add Schedule80CDtls_G
                Set Schedule80CDtls_G = Nothing
                Set Schedule80CDtls_G = CreateObject("Scripting.Dictionary")
   
   End If
   Next
   
              jsonDictionary.add "Schedule80CDtls", Schedule80CDtls
              Set Schedule80CDtls = Nothing
              Set Schedule80CDtls = CreateObject("Scripting.Dictionary")
   
   
'End If
              
TotalAmt_G = Sheet15.Range("TotAmount.80C").Value
jsonDictionary("TotalAmt") = UVCase(TotalAmt_G)

Set Schedule80C = jsonDictionary
End If
End Function

'Malli_80EE_AY_2025_26-------09/04/2025
'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Function Schedule80CCC() As Object
'Dim i As Long
'Dim mIntCells80CCC As Variant
'Dim jsonDictionary
'Set jsonDictionary = CreateObject("Scripting.Dictionary")
'
'Dim Schedule80CCCDtls_G, Schedule80CCCDtls
'Set Schedule80CCCDtls_G = CreateObject("Scripting.Dictionary")
'
'Set Schedule80CCCDtls = CreateObject("Scripting.Dictionary")
'Set Schedule80CCCDtls = New Collection
'
'Dim InsurerName_G, PolicyDocNo_G, Amount_G, TotalAmt_G
'subProcCaption = "80CCC"
'
'
''If Sheet15.Range("TotAmount.80CCC").Value > 0 Then
'If Sheet15.Range("TotAmount.80CC").Value > 0 Then
'
'     mIntCells80CCC = Sheet15.Range("Name_of_insurer.80CCC").count
'     'mIntCells80CCC = Sheet17.Range("Name_of_insurer").count
'   For i = 1 To mIntCells80CCC
'   'If (Sheet15.Range("Name_of_insurer.80CCC").Cells(i, 1).Value <> "" And Sheet15.Range("Name_of_insurer.80CCC").Cells(i, 1).Value <> "(Select)") Then
'   If Sheet15.Range("Name_of_insurer.80CCC").Cells(i, 1).Value <> "" Then
'   'If (Sheet17.Range("Name_of_insurer").Cells(i, 1).Value <> "" And Sheet17.Range("Name_of_insurer").Cells(i, 1).Value <> "(Select)") Then
'
'   InsurerName_G = Sheet15.Range("Name_of_insurer.80CCC").Cells(i, 1).Value
'   'InsurerName_G = Sheet17.Range("Name_of_insurer").Cells(i, 1).Value
'
'   PolicyDocNo_G = Sheet15.Range("Policy_Document_Number.80CCC").Cells(i, 1).Value
'   Amount_G = Sheet15.Range("Amount.80CCC").Cells(i, 1).Value
'
'
'
'   If InsurerName_G <> "" Then
'    Schedule80CCCDtls_G("InsurerName") = UCase(Trim(InsurerName_G))
'   Else
'   Schedule80CCCDtls_G("InsurerName") = ""
'   End If
'
'   If PolicyDocNo_G <> "" Then
'    Schedule80CCCDtls_G("PolicyDocNo") = UCase(Trim(PolicyDocNo_G))
'   Else
'   Schedule80CCCDtls_G("PolicyDocNo") = ""
'   End If
'
'      Schedule80CCCDtls_G("Amount") = UVCase(Amount_G)
'
'   'End If
'                Schedule80CCCDtls.add Schedule80CCCDtls_G
'                Set Schedule80CCCDtls_G = Nothing
'                Set Schedule80CCCDtls_G = CreateObject("Scripting.Dictionary")
'   End If
'   Next
'
'              jsonDictionary.add "Schedule80CCCDtls", Schedule80CCCDtls
'              Set Schedule80CCCDtls = Nothing
'              Set Schedule80CCCDtls = CreateObject("Scripting.Dictionary")
'
'
''End If
'
''TotalAmt_G = Sheet15.Range("TotAmount.80CCC").Value
'TotalAmt_G = Sheet15.Range("TotAmount.80CC").Value
'
'jsonDictionary("TotalAmt") = UVCase(TotalAmt_G)
'
'Set Schedule80CCC = jsonDictionary
'End If
'End Function
'Malli_80EE_AY_2025_26-------09/04/2025


'Konda_29/01/2026=========
'Function Schedule24B() As Object
'Dim i As Long
'Dim mIntCells24B As Variant
'Dim jsonDictionary
'Set jsonDictionary = CreateObject("Scripting.Dictionary")
'
'
'
'Dim ScheduleUs24BDtls_G, ScheduleUs24BDtls
'
'Set ScheduleUs24BDtls_G = CreateObject("Scripting.Dictionary")
'
'Set ScheduleUs24BDtls = CreateObject("Scripting.Dictionary")
'Set ScheduleUs24BDtls = New Collection
'
'subProcCaption = "24B"
'
'Dim LoanTknFrom_G, IFSCCode_G, BankOrInstnName_G, PAN_G, LoanAccNoOfBankOrInstnRefNo_G, DateofLoan_G, TotalLoanAmt_G, LoanOutstndngAmt_G, InterestPaid_G, TotalInterestPaid_G
'
'If Sheet16.Range("TotAmt.24b").Value > 0 Then
' mIntCells24B = Sheet16.Range("LoanfrmBankOrInstitute.24b").count
'
'  For i = 1 To mIntCells24B
'   If (Sheet16.Range("LoanfrmBankOrInstitute.24b").Cells(i, 1).Value <> "" And Sheet16.Range("LoanfrmBankOrInstitute.24b").Cells(i, 1).Value <> "(Select)") Then
'
'   LoanTknFrom_G = Sheet16.Range("LoanfrmBankOrInstitute.24b").Cells(i, 1).Value
''   IFSCCode_G = Sheet16.Range("IFSC.24b").Cells(i, 1).Value  'Ankita_05/05/2025_Commented as per DESheet_v0.7
'   BankOrInstnName_G = Sheet16.Range("bankName.24b").Cells(i, 1).Value
''   PAN_G = Sheet16.Range("PAN.24b").Cells(i, 1).Value  'Ankita_05/05/2025_Commented as per DESheet_v0.7
'   LoanAccNoOfBankOrInstnRefNo_G = Sheet16.Range("loanAccNum.24b").Cells(i, 1).Value
'   DateofLoan_G = Sheet16.Range("loanDate.24b").Cells(i, 1).Value
'   TotalLoanAmt_G = Sheet16.Range("loanAmt.24b").Cells(i, 1).Value
'   LoanOutstndngAmt_G = Sheet16.Range("loanOutstanding.24b").Cells(i, 1).Value
'   InterestPaid_G = Sheet16.Range("Intrst.24b").Cells(i, 1).Value
'
'
'   If LoanTknFrom_G <> "" Then
'            If LoanTknFrom_G = "Bank" Then
'            LoanTknFrom_G = "B"
'            ElseIf LoanTknFrom_G = "Other than bank" Then
'            LoanTknFrom_G = "I"
'            End If
'        ScheduleUs24BDtls_G("LoanTknFrom") = UCase(LoanTknFrom_G)
'   Else
'        ScheduleUs24BDtls_G("LoanTknFrom") = ""
'   End If
'
'   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'
''   If IFSCCode_G <> "" Then
''   ScheduleUs24BDtls_G("IFSCCode") = UCase(Trim(IFSCCode_G))
''   End If
''
''    If PAN_G <> "" Then
''   ScheduleUs24BDtls_G("PAN") = UCase(Trim(PAN_G))
''   End If
'
'   If BankOrInstnName_G <> "" Then
'   ScheduleUs24BDtls_G("BankOrInstnName") = UCase(Trim(BankOrInstnName_G))
'   Else
'   ScheduleUs24BDtls_G("BankOrInstnName") = ""
'   End If
'
'   If LoanAccNoOfBankOrInstnRefNo_G <> "" Then
'   ScheduleUs24BDtls_G("LoanAccNoOfBankOrInstnRefNo") = UCase(Trim(LoanAccNoOfBankOrInstnRefNo_G))
'   Else
'   ScheduleUs24BDtls_G("LoanAccNoOfBankOrInstnRefNo") = ""
'   End If
'
'   If DateofLoan_G <> "" Then
'   ScheduleUs24BDtls_G("DateofLoan") = Mid(DateofLoan_G, 7, 4) & "-" & Mid(DateofLoan_G, 4, 2) & "-" & Mid(DateofLoan_G, 1, 2)
'   Else
'   ScheduleUs24BDtls_G("DateofLoan") = ""
'   End If
'
'
'   ScheduleUs24BDtls_G("TotalLoanAmt") = UVCase(TotalLoanAmt_G)
'   ScheduleUs24BDtls_G("LoanOutstndngAmt") = UVCase(LoanOutstndngAmt_G)
'  ' ScheduleUs24BDtls_G("InterestPaid") = UVCase(InterestPaid_G)
'    ScheduleUs24BDtls_G("InterestUs24B") = UVCase(InterestPaid_G)
''End If
'       ScheduleUs24BDtls.add ScheduleUs24BDtls_G
'                Set ScheduleUs24BDtls_G = Nothing
'                Set ScheduleUs24BDtls_G = CreateObject("Scripting.Dictionary")
'
'  End If
'  Next
'              jsonDictionary.add "ScheduleUs24BDtls", ScheduleUs24BDtls
'                Set ScheduleUs24BDtls = Nothing
'                Set ScheduleUs24BDtls = CreateObject("Scripting.Dictionary")
''End If
'
'TotalInterestPaid_G = Sheet16.Range("TotAmt.24b").Value
''jsonDictionary("TotalInterestPaid") = UVCase(TotalInterestPaid_G)
'jsonDictionary("TotalInterestUs24B") = UVCase(TotalInterestPaid_G)
'
'
'Set Schedule24B = jsonDictionary
'End If
'
'End Function

'Konda_EA10_13A__AY_2025_26-------06-05-2025
Function ScheduleEA10_13A() As Object
subProcCaption = "ScheduleEA10_13A"
noOfProcessSub = 10
Dim jsonDictionary

Dim Placeofwork_G, ActlHRARecv_G, ActlRentPaid_G, DtlsSalUsSec171_G, BasicSalary_G, DearnessAllwnc_G, ActlRentPaid10Per_G, Sal40Or50Per_G, EligbleExmpAllwncUs13A_G

Set jsonDictionary = CreateObject("Scripting.Dictionary")
   
   Placeofwork_G = Sheet18.Range("Sch10of13A_PlaceofWrk").Cells(1, 1).Value

   ActlHRARecv_G = Sheet18.Range("Sch10of13A_ActlHRArecivedA").Value

   ActlRentPaid_G = Sheet18.Range("Sch10of13A_ActlRentpaid").Value
   DtlsSalUsSec171_G = Sheet18.Range("Sch10of13A_DetlsofSalpersec17of1").Value
   BasicSalary_G = Sheet18.Range("Sch10of13A_BasicSalary").Value
   DearnessAllwnc_G = Sheet18.Range("Sch10of13A_DearAllowance").Value
   ActlRentPaid10Per_G = Sheet18.Range("Sch10of13A_Actlrentpaid10persalaryB").Value
   Sal40Or50Per_G = Sheet18.Range("Sch10of13A_50Por40Pofsalary").Value
   EligbleExmpAllwncUs13A_G = Sheet18.Range("Sch10of13A_ElgiblExmptAllwnce10of13A").Value
  
  
  
Dim Placeofwork_EA As Variant
Placeofwork_EA = Mid(Placeofwork_G, 1, 1)

If IsNumeric(Placeofwork_EA) Then

If Placeofwork_EA = "1" Or Placeofwork_EA = "2" Then
jsonDictionary("Placeofwork") = Placeofwork_EA

End If
   
 
 If ActlHRARecv_G <> "" Then
 jsonDictionary("ActlHRARecv") = UVCase(ActlHRARecv_G)
 End If

 
If ActlRentPaid_G <> "" Then
jsonDictionary("ActlRentPaid") = UVCase(ActlRentPaid_G)

End If

If DtlsSalUsSec171_G <> "" Then
jsonDictionary("DtlsSalUsSec171") = UVCase(DtlsSalUsSec171_G)

End If

If BasicSalary_G <> "" Then
jsonDictionary("BasicSalary") = UVCase(BasicSalary_G)

End If

If DearnessAllwnc_G <> "" Then
'jsonDictionary("DearnessAllwnc") = UVCase(DearnessAllwnc_G)
'Konda--------08_05_2025
jsonDictionary("DearnessAllwnc") = DearnessAllwnc_G
'-------------------
End If

If ActlRentPaid10Per_G <> "" Then
jsonDictionary("ActlRentPaid10Per") = UVCase(ActlRentPaid10Per_G)

End If

If Sal40Or50Per_G <> "" Then
jsonDictionary("Sal40Or50Per") = UVCase(Sal40Or50Per_G)

End If

If EligbleExmpAllwncUs13A_G <> "" Then
jsonDictionary("EligbleExmpAllwncUs13A") = UVCase(EligbleExmpAllwncUs13A_G)

End If


Set ScheduleEA10_13A = jsonDictionary

End If


End Function















