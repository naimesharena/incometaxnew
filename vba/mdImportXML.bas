Attribute VB_Name = "mdImportXML"
Option Explicit

'Dim dom1, dom As DOMDocument
Dim end_TaxP, rngname_TaxP As Variant
Dim end_TDSoth, rngname_TDSoth As Variant
Dim end_TDS3oth, rngname_TDS3oth As Variant
Dim end_TCS, rngname_TCS As Variant
Dim sPassword As String
Public choicexml As Variant
Public Javaxml, Excelxml, Common_xml, AssmntYear As String

Dim rowcount, cnt As Variant
Dim xml As Variant

Function OpenXMLFileDialog() As String
On Error GoTo endline
    Dim FileDialogBox As FileDialog
    Dim FileObject As Variant
    Dim filname As String
    Dim intChoice As Long
    
    intChoice = 0
    
    Set FileDialogBox = Application.FileDialog(msoFileDialogFilePicker)
    With FileDialogBox
        .ButtonName = "Select"
        .AllowMultiSelect = False
        .Filters.add "XML Files", "*.xml", 1
        .Title = "Choose XML file to use for Filling TDS/TaxPayment Data in Utility"
        .InitialView = msoFileDialogViewDetails
'        choicexml = Application.InputBox( _
'                            prompt:="Enter 1 for Prefilled XML 2 for Excel XML ", _
'                            Title:="XML", _
'                            Default:=0, _
'                            Type:=1)
'        If Not (choicexml = "1" Or choicexml = "2") Then
'            MsgBox "Invalid Choice"
'            End
'        End If
        intChoice = .Show
        If intChoice <> 0 Then
            For Each FileObject In .SelectedItems
                filname = FileObject
                OpenXMLFileDialog = filname
            Next FileObject
        End If
        On Error GoTo 0
    End With
    
    If intChoice = 0 Then
        Application.EnableEvents = True
        Set FileDialogBox = Nothing
        CloseMsg
    End If
    
    Set FileDialogBox = Nothing
    
endline:

End Function


Sub ImportXML()
'On Error GoTo endline

'Dim XpathOfPI As String
'
'    Dim NodeList As IXMLDOMNodeList
'    Dim Node As IXMLDOMNode
'    Dim InnerNode As IXMLDOMNode
'    Dim rowcount As Long
'    Dim cellcount As Long
'    Dim rowRange As Range
'    Dim cellrange As Range
'
'    'Application.EnableEvents = False
'
'    fmsgbox "Import personal/tax details from downloaded Pre-filled XML or Import from already generated XML of the current assessment year."
'
'    xml = ""
'    AssmntYear = ""
'    Set dom1 = New DOMDocument
'    dom1.Load (OpenXMLFileDialog)
'
'    On Error Resume Next
'
'    XpathOfPI = "/ns3:ITR/ns2:ITR1/Form_ITR1"
'    Set NodeList = dom1.SelectNodes(XpathOfPI)
'
'    For Each Node In NodeList
'     AssmntYear = Node.SelectSingleNode("AssessmentYear").Text
'        If AssmntYear = "2020" Then
'        End If
'    Next Node
'
'    xml = dom1.xml
'
'    xml = Replace(xml, "ITRForm:", "")
'
'    Set dom = New DOMDocument
'    dom.LoadXML (xml)
'    Javaxml = "/ns3:ITR/ns2:ITR1"
'    Excelxml = "/ITRETURN:ITR/ITR1FORM:ITR1"
'    If AssmntYear <> "" Then
'    Common_xml = Javaxml                                         '"/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
'   Else
'   Common_xml = Excelxml
'   End If
'
'
'   Prefillexml
'
'   ProgressBarHide
'
'endline:
'    Application.EnableEvents = True
'
End Sub
Sub SalaryXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
   

    XpathOfPI = Common_xml & "/ITR1_IncomeDeductions " '"/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

  
            Sheet1.Range("IncD.IncomeFromSal").Value = UCase(node.SelectSingleNode("GrossSalary").text)
            Sheet1.Range("IncD.Allowances").Value = UCase(node.SelectSingleNode("Salary").text)
            Sheet1.Range("IncD.Perquisites").Value = UCase(node.SelectSingleNode("PerquisitesValue").text)
            Sheet1.Range("IncD.Profits").Value = UCase(node.SelectSingleNode("ProfitsInSalary").text)
            
            
            Sheet1.Range("Deductions_16").Value = UCase(node.SelectSingleNode("DeductionUs16").text)
            
'            Sheet1.Range("IncD.Deduction16ia").Value = UCase(Node.SelectSingleNode("DeductionUs16ia").Text)
            
            Sheet1.Range("IncD.Deduction16").Value = UCase(node.SelectSingleNode("EntertainmentAlw16ii").text)
            
            Sheet1.Range("IncD.Deduction16ic").Value = UCase(node.SelectSingleNode("ProfessionalTaxUs16iii").text)
            
            Sheet1.Range("IncD.TypeOfHP").Value = UCase(node.SelectSingleNode("TypeOfHP").text)
                        
            If Sheet1.Range("IncD.TypeOfHP").Value = "L" Then
               Sheet1.Range("IncD.TypeOfHP").Value = "Let Out"
             ElseIf Sheet1.Range("IncD.TypeOfHP").Value = "S" Then
            Sheet1.Range("IncD.TypeOfHP").Value = "Self Occupied"
            ElseIf Sheet1.Range("IncD.TypeOfHP").Value = "D" Then
            Sheet1.Range("IncD.TypeOfHP").Value = "Deemed Let Out"
            End If

            
           Sheet1.Range("IncD.GrossRentRecieved").Value = UCase(node.SelectSingleNode("GrossRentReceived").text)
           Sheet1.Range("IncD.TaxPaidLocalAuthorities").Value = UCase(node.SelectSingleNode("TaxPaidlocalAuth").text)
           Sheet1.Range("IncD.InterestBorrowedCapital").Value = UCase(node.SelectSingleNode("InterestPayable").text)
           Sheet1.Range("IncD.Arrears").Value = UCase(node.SelectSingleNode("ArrearsUnrealizedRentRcvd").text)
           
           Sheet1.Range("IncD.LessDeduction57").Value = UCase(node.SelectSingleNode("DeductionUs57iia").text)
           Sheet1.Range("IncD.IncomeFromOS").Value = UCase(node.SelectSingleNode("IncomeOthSrc").text)
          

            
        
           Sheet1.Range("IncD.Section80C").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80C").text)
           Sheet1.Range("IncD.Section80CCC").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCC").text)
           Sheet1.Range("IncD.Section80CCD_SE").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCDEmployeeOrSE").text)
           Sheet1.Range("IncD.Section80CCD1B_SE").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCD1B").text)
           Sheet1.Range("IncD.Section80CCD").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCDEmployer").text)
'          Sheet1.Range("IncD.Section80CCG").Value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCG").Text)
           
           
          ' Sheet1.Range("SELECT80D").Value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/HealthInsurancePremium").Text)
'           Sheet1.Range("SELECT80D").Value = Findtext(UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/HealthInsurancePremium").Text), "Selection80D")
'           Sheet1.Range("IncD.Section80D").Value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/Sec80DHealthInsurancePremiumUsr").Text)
'
'
'           Sheet1.Range("SELECT80DB").Value = Findtext(UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/MedicalExpenditure").Text), "Selection80DB")
'           Sheet1.Range("IncD.Section80DB").Value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/Sec80DMedicalExpenditureUsr").Text)
'
'
'           Sheet1.Range("SELECT80DC").Value = Findtext(UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/PreventiveHealthCheckUp").Text), "Selection80DC")
'           Sheet1.Range("IncD.Section80DC").Value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80DHealthInsPremium/Sec80DPreventiveHealthCheckUpUsr").Text)
'
           Sheet1.Range("IncD.Section80D").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80D").text)
           Sheet1.Range("SELECT80DD").Value = Findtext(UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80DDUsrType").text), "Selection80DD")
           Sheet1.Range("IncD.Section80DD").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80DD").text)
           

           Sheet1.Range("SELECT80DDS").Value = Findtext(UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80DDBUsrType").text), "Selection80DDB")
           Sheet1.Range("IncD.Section80DDB").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80DDB").text)
           
           
           Sheet1.Range("IncD.Section80E").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80E").text)
           Sheet1.Range("IncD.Section80EE").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80EE").text)
           Sheet1.Range("IncD.Section80EEA").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80EEA").text)
           Sheet1.Range("IncD.Section80EEB").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80EEB").text)
           Sheet1.Range("IncD.Section80GG").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80GG").text)
           Sheet1.Range("IncD.Section80GGA").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80GGA").text)
'          Sheet1.Range("IncD.Section80QQB").Value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80QQB").Text)
'          Sheet1.Range("IncD.Section80RRB").Value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80RRB").Text)
           Sheet1.Range("IncD.Section80TTA").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80TTA").text)
           Sheet1.Range("IncD.Section80TTB").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80TTB").text)
           'Sheet1.Range("IncD.Section80CCG").Value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCG").Text)
        
           Sheet1.Range("IncD.Section80GGC").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80GGC").text)
'          Sheet1.Range("IncD.Section80CCG").Value = UCase(Node.SelectSingleNode("UsrDeductUndChapVIA/Section80CCG").Text)
           
           Sheet1.Range("SELECT80U").Value = Findtext(UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80UUsrType").text), "Selection80U")
           Sheet1.Range("IncD.Section80U").Value = UCase(node.SelectSingleNode("UsrDeductUndChapVIA/Section80U").text)
           
            
       
        Next InnerNode
    Next node
End Sub


Sub ReliefXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
    XpathOfPI = Common_xml & "/ITR1_TaxComputation  " '"/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
            
            Sheet1.Range("IncD.Section89").Value = UCase(node.SelectSingleNode("Section89").text)
            
            
    Next node

End Sub



Sub sec38XMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim Capacity As Variant
    
    
    
    XpathOfPI = Common_xml '"/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

            
            Sheet3.Range("EI.Sec10_38").Value = UCase(node.SelectSingleNode("TaxPaid/TaxesPaid/ExcIncSec1038").text)
            Sheet3.Range("EI.Sec10_34").Value = UCase(node.SelectSingleNode("TaxPaid/TaxesPaid/ExcIncSec1034").text)
            Sheet3.Range("ExcempIncome").Value = UCase(node.SelectSingleNode("ns2:TaxExmpIntInc").text)
            
            Capacity = UCase(node.SelectSingleNode("Verification/Capacity").text)
            If Capacity = "S" Then
            Capacity = "Self"
            ElseIf Capacity = "R" Then
            Capacity = "Representative"
            Else
            Capacity = ""
            End If
            
            Sheet3.Range("Ver.capacity").Value = Capacity
            
            
            Sheet3.Range("Ver.Place").Value = UCase(node.SelectSingleNode("Verification/Place").text)
'            Sheet3.Range("Ver.Date").Value = Mid(UCase(Node.SelectSingleNode("Verification/Date").Text), 9, 2) & "/" & Mid(UCase(Node.SelectSingleNode("Verification/Date").Text), 6, 2) & "/" & Mid(UCase(Node.SelectSingleNode("Verification/Date").Text), 1, 4)
            
            
            
           Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value = UCase(node.SelectSingleNode("TaxReturnPreparer/IdentificationNoOfTRP").text)
           Sheet3.Range("Sheet2.NameOfTRP").Value = UCase(node.SelectSingleNode("TaxReturnPreparer/NameOfTRP").text)
           Sheet3.Range("Sheet2.ReImbFrmGov").Value = UCase(node.SelectSingleNode("TaxReturnPreparer/ReImbFrmGov").text)
         
        Next InnerNode
    Next node
   
End Sub



Function ValidateXML()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
     Dim AssmntYear1 As String
     ValidateXML = True

    XpathOfPI = Common_xml & "/Form_ITR1"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
     AssmntYear1 = node.SelectSingleNode("AssessmentYear").text
        If AssmntYear1 = "2025" Then          'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
         Exit Function
         End If
    Next node
     ValidateXML = False
End Function
Sub ITXMLImport()
On Error Resume Next
    Dim XpathOfIT As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim BSRCodeColNo, DateDepColNo, SrlNoChallanColNo, AmtColNo As Long
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long
    
    
    
    XpathOfIT = Common_xml & "/TaxPayments/TaxPayment"  '"/ns2:ITR/TaxPayments/TaxPayment"
    Set Nodelist = dom.SelectNodes(XpathOfIT)

    BSRCodeColNo = Sheet2.Range("TaxP.BSRCode").Column
    DateDepColNo = Sheet2.Range("TaxP.DateDep").Column
    SrlNoChallanColNo = Sheet2.Range("TaxP.SrlNoOfChaln").Column
    AmtColNo = Sheet2.Range("TaxP.Amt").Column
    
    TotalExRow = Range("TaxP.BSRCode").Rows.count
    
    TotalXMLRow = Nodelist.Length
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
        For Each InnerNode In node.ChildNodes
            Sheet2.Cells(rowcount, BSRCodeColNo).Value = UCase(node.SelectSingleNode("BSRCode").text)
            strDate = node.SelectSingleNode("DateDep").text
            YYYY = Mid(strDate, 1, 4)
            MM = Mid(strDate, 6, 2)
            DD = Mid(strDate, 9, 2)
            strDate = DD & "/" & MM & "/" & YYYY
      
            Sheet2.Cells(rowcount, DateDepColNo).Value = strDate
            Sheet2.Cells(rowcount, SrlNoChallanColNo).Value = node.SelectSingleNode("SrlNoOfChaln").text
            Sheet2.Cells(rowcount, AmtColNo).Value = node.SelectSingleNode("Amt").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub



Sub setDiffTblinfo_IT()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet2.Range("TaxP.BSRCode").count
    Set rangecells = Sheet2.Range("TaxP.BSRCode").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
'Change-28.07.2023.101 -> Discussion with Arun Sir
'    DefinedgridNameRange = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.Amt"
    DefinedgridNameRange = "TaxP.BSRCode||TaxP.DateDep||TaxP.SrlNoOfChaln||TaxP.Amt||IT.FormulaOFS||FormulaOfQ||FormulaOfSAT||FormulaOfSAT1||FormulaOfSATNew||IT_DueDate"  'Ayush_DueDate_08/09/2025
End Sub

Sub AddDiffRows_IT(DiffRows As Long)
    setDiffTblinfo_IT
    Sheet2.Activate
    searchLastRow ("TaxP.BSRCode")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub

Sub TCSXMLImport()
On Error Resume Next
    Dim XpathOfTCS As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTCS As Long
    Dim TANColNo, DEDNameColNo, TaxColNo, ClaimColNo, AmntClaimedBySpouseTCS, AmtTaxCollected, CollectedYr As Variant
    Dim TotalExRow, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    
    
    
    
    XpathOfTCS = Common_xml & "/ScheduleTCS/TCS" ' "/ns2:ITR/ScheduleTCS/TCS"
    Set Nodelist = dom.SelectNodes(XpathOfTCS)
    
    TANColNo = Sheet11.Range("TCS.TAN").Column
    DEDNameColNo = Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").Column
    AmtTaxCollected = Sheet11.Range("TCS.AmountCollected").Column
    CollectedYr = Sheet11.Range("TCS.CollectionYear").Column
    TaxColNo = Sheet11.Range("TCS.TotalTCS").Column
    ClaimColNo = Sheet11.Range("TCS.AmtTCSClaimedThisYear").Column
    'AmntClaimedBySpouseTCS = Sheet11.Range("TCS.AmtClaimedBySpouse").Column
    
    
    TotalExRow = Range("TCS.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
    Sheet11.Range("TCS.TAN").ClearContents
    Sheet11.Range("TCS.EmployerOrDeductorOrCollecterName").ClearContents
    Sheet11.Range("TCS.AmountCollected").ClearContents
    Sheet11.Range("TCS.CollectionYear").ClearContents
    Sheet11.Range("TCS.TotalTCS").ClearContents
    Sheet11.Range("TCS.AmtTCSClaimedThisYear").ClearContents
    'Sheet11.Range("TCS.AmtClaimedBySpouse").ClearContents
    
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
        For Each InnerNode In node.ChildNodes
            Sheet11.Cells(rowcount, TANColNo).Value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet11.Cells(rowcount, DEDNameColNo).Value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            Sheet11.Cells(rowcount, AmtTaxCollected).Value = node.SelectSingleNode("AmtTaxCollected").text
            Sheet11.Cells(rowcount, CollectedYr).Value = node.SelectSingleNode("CollectedYr").text
            Sheet11.Cells(rowcount, TaxColNo).Value = node.SelectSingleNode("TotalTCS").text
            Sheet11.Cells(rowcount, ClaimColNo).Value = node.SelectSingleNode("AmtTCSClaimedThisYear").text
            
            'Sheet11.Cells(rowcount, AmntClaimedBySpouseTCS).Value = Node.SelectSingleNode("AmtClaimedBySpouse").Text
        Next InnerNode
        cnt = cnt + 1
    Next node
End If
    RecTCS = cnt
    

End Sub

Function Findtext(myinput As Variant, tblrange As Variant) As Variant
On Error Resume Next
    Dim rng As Range
    Dim found As Boolean
    Dim searchtext As Variant
    found = False
    searchtext = ""
    
    For Each rng In Range(tblrange)
        If InStr(rng.Value, "-") > 0 Then
            If Mid(rng.Value, 1, InStr(rng.Value, "-") - 1) = myinput Then
                found = True
                searchtext = rng.Value
                Exit For
            End If
        End If
    Next rng

    If found Then
         Findtext = searchtext
    Else
        Findtext = myinput
    End If
End Function

Function Findtext1(myinput As Variant, tblrange As Variant) As Variant
On Error Resume Next
    Dim rng As Range
    Dim found As Boolean
    Dim searchtext As Variant
    found = False
    searchtext = ""
    
    For Each rng In Range(tblrange)
        If InStrRev(rng.Value, ")") > 0 Then
            If Mid(rng.Value, 1, InStrRev(rng.Value, ")")) = myinput Then
                found = True
                searchtext = rng.Value
                Exit For
            End If
        End If
    Next rng

    If found Then
         Findtext1 = searchtext
    Else
        Findtext1 = myinput
    End If
End Function


Sub PersonalInfoXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim GenderUpload As String
    Dim sCountry, iCountry As Variant
    Dim sState, iState As Variant
    Dim sEmpCat, iEmpCat As Variant
        
        
        
    
        
    XpathOfPI = Common_xml & "/PersonalInfo" '"/ns2:ITR/PersonalInfo"   '"/ITR/ITR1/PartA_GEN1/OrgFirmInfo"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

            Sheet1.Range("sheet1.FirstName").Value = UCase(node.SelectSingleNode("AssesseeName/FirstName").text)
            Sheet1.Range("sheet1.MiddleName").Value = UCase(node.SelectSingleNode("AssesseeName/MiddleName").text)
            Sheet1.Range("sheet1.SurNameOrOrgName").Value = UCase(node.SelectSingleNode("AssesseeName/SurNameOrOrgName").text)
            Sheet1.Range("sheet1.PAN").Value = UCase(node.SelectSingleNode("PAN").text)
   
            dob = node.SelectSingleNode("DOB").text
            YYYY = Mid(dob, 1, 4)
            MM = Mid(dob, 6, 2)
            DD = Mid(dob, 9, 2)
            dob = DD & "/" & MM & "/" & YYYY
           
            Sheet1.Range("sheet1.DOB").Value = dob
       
        'Address
            Sheet1.Range("sheet1.ResidenceNo").Value = UCase(node.SelectSingleNode("Address/ResidenceNo").text)
            Sheet1.Range("sheet1.ResidenceName").Value = UCase(node.SelectSingleNode("Address/ResidenceName").text)
            Sheet1.Range("sheet1.RoadOrStreet").Value = UCase(node.SelectSingleNode("Address/RoadOrStreet").text)
            Sheet1.Range("sheet1.LocalityOrArea").Value = UCase(node.SelectSingleNode("Address/LocalityOrArea").text)
            Sheet1.Range("sheet1.CityOrTownOrDistrict").Value = UCase(node.SelectSingleNode("Address/CityOrTownOrDistrict").text)
            
            
            iState = UCase(node.SelectSingleNode("Address/StateCode").text)
            sState = Findtext(iState, "StateList")
            Sheet1.Range("sheet1.StateCode1").Value = sState
        
            Application.EnableEvents = False
            
            iCountry = UCase(node.SelectSingleNode("Address/CountryCode").text)
            sCountry = Findtext(iCountry, "CountList")
            Sheet1.Range("sheet1.Country").Value = sCountry
            
            Application.EnableEvents = True
                        
            Sheet1.Range("sheet1.PinCode").Value = UCase(node.SelectSingleNode("Address/PinCode").text)
            Sheet1.Range("sheet1.ZipCode").Value = UCase(node.SelectSingleNode("Address/ZipCode").text)
            Sheet1.Range("sheet1.MobileCountryCode").Value = UCase(node.SelectSingleNode("Address/CountryCodeMobile").text)
            Sheet1.Range("sheet1.Mobileno").Value = UCase(node.SelectSingleNode("Address/MobileNo").text)
            Sheet1.Range("sheet1.EmailAddress").Value = UCase(node.SelectSingleNode("Address/EmailAddress").text)
        
        'Employee Category/AADHAAR/ GENDER/Status
        
            iEmpCat = UCase(node.SelectSingleNode("EmployerCategory").text)
            
              
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
            
            
            'sEmpCat = Findtext(sEmpCat, "EmpCatList")
            Sheet1.Range("sheet1.EmployerCategory1").Value = sEmpCat
            
            Sheet1.Range("Sheet1.Aadhaar").Value = UCase(node.SelectSingleNode("AadhaarCardNo").text)
'           Enhancement
'           Sheet1.Range("Sheet1.AadhaarEnrol").Value = UCase(node.SelectSingleNode("AadhaarEnrolmentId").text)
            Sheet1.Range("IncD.IncomeFromSal").Value = UCase(node.SelectSingleNode("ITR1_IncomeDeductions/Salary").text)
            Sheet1.Range("IncD.IncomeFromOS").Value = UCase(node.SelectSingleNode("ITR1_IncomeDeductions/IncomeOthSrc").text)
            
        Next InnerNode
    Next node
   
End Sub
Sub FilingInfoXMLImport()
On Error Resume Next
    Dim XpathOfPI As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim GenderUpload As String
    Dim DateofOriginalfile, DateofOriginalfile1, NoticeDateussec, Filingtype As Variant
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
    Dim iProvisoFlag, sProvisoFlag As Variant
    Dim iDepositAmountFlag, sDepositAmountFlag As Variant
    Dim iAggrigateAmountFlag, sAggrigateAmountFlag As Variant
    Dim iAggrigateAmountFlag1, sAggrigateAmountFlag1 As Variant

    
        
   
        
        
    XpathOfPI = Common_xml & "/FilingStatus"
    Set Nodelist = dom.SelectNodes(XpathOfPI)

    For Each node In Nodelist
        'For Each InnerNode In Node.ChildNodes
        ' For Each InnerNode In Node.ChildNodes
        'ReturnFile Sec/ Port 5A
        
            iReturnFile = UCase(node.SelectSingleNode("ReturnFileSec").text)
            'sReturnFile = Findtext(iReturnFile & " ", "ReturnSecList")
            'Sheet1.Range("sheet1.ReturnFileSec").Value = sReturnFile
       
            
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
'Konda updated on 23-12-2025
           'Range("sheet1.ReturnFileSec").Value = "139(9A) - After condonation of delay u/s 119(2)(b)"
           '119(2)(b)- After condonation of delay
           Range("sheet1.ReturnFileSec").Value = "119(2)(b)- After condonation of delay"
        End If


       iProvisoFlag = UCase(node.SelectSingleNode("SeventhProvisio139").text)
        If iProvisoFlag = "Y" Then
        sProvisoFlag = "Yes"
        ElseIf iProvisoFlag = "N" Then
        sProvisoFlag = "No"
        End If
            
        Sheet1.Range("sheet1.SeventhProvisoFlag").Value = sProvisoFlag
        
        iDepositAmountFlag = UCase(node.SelectSingleNode("DepAmtAggAmtExcd1CrPrYrFlg").text)
        If iDepositAmountFlag = "Y" Then
        sDepositAmountFlag = "Yes"
        ElseIf iDepositAmountFlag = "N" Then
        sDepositAmountFlag = "No"
        End If
'Change-22.11.2022.102.16F
'        Sheet1.Range("Sheet1.DepositAmountFlag").Value = sDepositAmountFlag
'        Sheet1.Range("Sheet1.DepositAmount").Value = UCase(Node.SelectSingleNode("AmtSeventhProvisio139i").text)
'---end

        iAggrigateAmountFlag = UCase(node.SelectSingleNode("IncrExpAggAmt2LkTrvFrgnCntryFlg").text)
        If iAggrigateAmountFlag = "Y" Then
        sAggrigateAmountFlag = "Yes"
        ElseIf iAggrigateAmountFlag = "N" Then
        sAggrigateAmountFlag = "No"
        End If
            
        Sheet1.Range("Sheet1.AggrigateAmountFlag").Value = sAggrigateAmountFlag
        Sheet1.Range("Sheet1.AggrigateAmount").Value = UCase(node.SelectSingleNode("AmtSeventhProvisio139ii").text)
        
        iAggrigateAmountFlag1 = UCase(node.SelectSingleNode("IncrExpAggAmt1LkElctrctyPrYrFlg").text)
        If iAggrigateAmountFlag1 = "Y" Then
        sAggrigateAmountFlag1 = "Yes"
        ElseIf iAggrigateAmountFlag1 = "N" Then
        sAggrigateAmountFlag1 = "No"
        End If
            
        Sheet1.Range("Sheet1.AggrigateAmountFlag1").Value = sAggrigateAmountFlag
        Sheet1.Range("Sheet1.AggrigateAmount1").Value = UCase(node.SelectSingleNode("AmtSeventhProvisio139iii").text)
       
       
     
            
'            iPort5A = UCase(Node.SelectSingleNode("PortugeseCC5A").Text)
'            sPort5A = IIf(iPort5A = "Y", "Yes", IIf(iPort5A = "N", "No", iPort5A))
'            Sheet1.Range("sheet1.PortugeseCC5A").Value = sPort5A
'
'        Sheet1.Range("Portugese").Value = Node.SelectSingleNode("PANOfSpouse").Text
       ReturnFileSec = Sheet1.Range("sheet1.ReturnFileSec1")
    ReturnFileSec = Mid(ReturnFileSec, 1, 2)
    If ReturnFileSec = "17" Then
Sheet1.Range("Sheet1.ReceiptNo").Value = node.SelectSingleNode("ReceiptNo").text

DateofOriginalfile = node.SelectSingleNode("OrigRetFiledDate").text
    
        
        Sheet1.Range("Sheet1.OrigRetFiledDate").Value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
     Else
        Sheet1.Range("Sheet1.ReceiptNo").Value = node.SelectSingleNode("ReceiptNo").text
        
        DateofOriginalfile = node.SelectSingleNode("OrigRetFiledDate").text
        
        Sheet1.Range("Sheet1.OrigRetFiledDate").Value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
        
        Sheet1.Range("sheet1.NoticeNo").Value = node.SelectSingleNode("NoticeNo").text
        
        NoticeDateussec = node.SelectSingleNode("NoticeDateUnderSec").text
        
        Sheet1.Range("sheet1.NoticeDate").Value = Mid(NoticeDateussec, 9, 2) & "/" & Mid(NoticeDateussec, 6, 2) & "/" & Mid(NoticeDateussec, 1, 4)
     End If
     
'
          If ReturnFileSec = "18" Then
     Sheet1.Range("Sheet1.ReceiptNo").Value = node.SelectSingleNode("ReceiptNo").text

DateofOriginalfile = node.SelectSingleNode("OrigRetFiledDate").text
    
        
        Sheet1.Range("Sheet1.OrigRetFiledDate").Value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
     Else
        Sheet1.Range("Sheet1.ReceiptNo").Value = node.SelectSingleNode("ReceiptNo").text
        
        DateofOriginalfile = node.SelectSingleNode("OrigRetFiledDate").text
        
        Sheet1.Range("Sheet1.OrigRetFiledDate").Value = Mid(DateofOriginalfile, 9, 2) & "/" & Mid(DateofOriginalfile, 6, 2) & "/" & Mid(DateofOriginalfile, 1, 4)
        
        Sheet1.Range("sheet1.NoticeNo").Value = node.SelectSingleNode("NoticeNo").text
        
        NoticeDateussec = node.SelectSingleNode("NoticeDateUnderSec").text
        
        Sheet1.Range("sheet1.NoticeDate").Value = Mid(NoticeDateussec, 9, 2) & "/" & Mid(NoticeDateussec, 6, 2) & "/" & Mid(NoticeDateussec, 1, 4)
     End If


     
        'Next InnerNode
    Next node
   

End Sub
Sub VeriInfoXMLImport()
On Error Resume Next
    Dim XpathOfVerification As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim dob As String
    Dim Status1 As String
    Dim YYYY As String
    Dim MM  As String
    Dim DD As String
    Dim GenderUpload As String
    Dim sReturnFile, iReturnFile As Variant
    Dim sPort5A, iPort5A As Variant
        
   
        
        
    XpathOfVerification = Common_xml & "/Verification"
    Set Nodelist = dom.SelectNodes(XpathOfVerification)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes

            Sheet3.Range("Ver.FatherName").Value = UCase(node.SelectSingleNode("Declaration/FatherName").text)
                                                
        Next InnerNode
    Next node
    
End Sub
Sub Sch80DXMLImport()
On Error Resume Next
    Dim XpathOf80D As String
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode

    XpathOf80D = Common_xml & "/Schedule80D/Sec80DSelfFamSrCtznHealth"
    Set Nodelist = dom.SelectNodes(XpathOf80D)

    For Each node In Nodelist
        For Each InnerNode In node.ChildNodes
   
            Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = UCase(node.SelectSingleNode("SeniorCitizenFlag").text)
            If Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Y" Then
            Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "Yes"
            ElseIf Sheet1.Range("DropDown_ValueOf_FamilyM_80D").Value = "N" Then
            Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "No"
            ElseIf Sheet1.Range("DropDown_ValueOf_FamilyM_80D").Value = "S" Then
            Sheet9.Range("DropDown_ValueOf_FamilyM_80D").Value = "No Claiming for Self/Family"
            End If
                                            
            
            Sheet9.Range("Health_Insurance_80D").Value = UCase(node.SelectSingleNode("HealthInsPremSlfFam").text)
            Sheet9.Range("Preventive_Health_80D").Value = UCase(node.SelectSingleNode("PrevHlthChckUpSlfFam").text)
            Sheet9.Range("Health_InsuranceSC_80D").Value = UCase(node.SelectSingleNode("HlthInsPremSlfFamSrCtzn").text)
            Sheet9.Range("Preventive_Health_SC_80D").Value = UCase(node.SelectSingleNode("PrevHlthChckUpSlfFamSrCtzn").text)
            Sheet9.Range("Medical_Expenditure_SC_80D").Value = UCase(node.SelectSingleNode("MedicalExpSlfFamSrCtzn").text)
            
            Sheet9.Range("DropDown_ValueOf_SC_80D").Value = UCase(node.SelectSingleNode("ParentsSeniorCitizenFlag").text)
            If Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Y" Then
            Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Yes"
            ElseIf Sheet1.Range("DropDown_ValueOf_SC_80D").Value = "N" Then
            Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "No"
            ElseIf Sheet1.Range("DropDown_ValueOf_SC_80D").Value = "P" Then
            Sheet9.Range("DropDown_ValueOf_SC_80D").Value = "Not claiming for Parents"
            End If
            
            
            Sheet9.Range("Health_Insurance2_80D").Value = UCase(node.SelectSingleNode("HlthInsPremParents").text)
            Sheet9.Range("Preventive_Health2_80D").Value = UCase(node.SelectSingleNode("PrevHlthChckUpParents").text)
            Sheet9.Range("Health_Insurance3_80D").Value = UCase(node.SelectSingleNode("HlthInsPremParentsSrCtzn").text)
            Sheet9.Range("Preventive_Health3_80D").Value = UCase(node.SelectSingleNode("PrevHlthChckUpParentsSrCtzn").text)
            Sheet9.Range("Medical_Expenditure2_80D").Value = UCase(node.SelectSingleNode("MedicalExpParentsSrCtzn").text)
                                             
            
                                            
        Next InnerNode
    Next node
    
End Sub


Sub RefundInfoXMLImport()
On Error Resume Next
    Dim XpathOfRefund As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim IFSC, BankName, ACCNO, CheckBox, CheckBox1 As Variant
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long




    XpathOfRefund = Common_xml & "/Refund/BankAccountDtls/AddtnlBankDetails"  '"ns2:ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfRefund)

    IFSC = Range("SchBA.IFSC").Column
    BankName = Range("SchBA.BankName").Column
    ACCNO = Range("SchBA.AcntNo").Column
    CheckBox = Range("tempxml").Column
    
    TotalExRow = Range("SchBA.IFSC").Rows.count
    
    TotalXMLRow = Nodelist.Length
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
        For Each InnerNode In node.ChildNodes
        

            
            Sheet3.Cells(rowcount, IFSC).Value = UCase(node.SelectSingleNode("IFSCCode").text)
            Sheet3.Cells(rowcount, BankName).Value = UCase(node.SelectSingleNode("BankName").text)
            Sheet3.Cells(rowcount, ACCNO).Value = UCase(node.SelectSingleNode("BankAccountNo").text)
            Sheet3.Cells(rowcount, CheckBox).Value = UCase(node.SelectSingleNode("UseForRefund").text)
            
            If CheckBox = True Then
            CheckBox = "true"
            ElseIf CheckBox = False Then
            CheckBox = "false"
            ElseIf CheckBox = "" Then
            CheckBox = "false"
            End If
            
LinkCheckBoxes
                   
            
           
           
               
            
            
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
    XpathOfRefund = Common_xml & "/Refund/BankAccountDtls/PriBankDetails"
    Set Nodelist = dom.SelectNodes(XpathOfRefund)
End Sub

Sub TDSonSalaryXMLImport()
On Error Resume Next
    Dim XpathOfTDS As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TANNoEmployer, TDSNameOfEmployer, TDSIncomeCharge, TDSTotalTax As Long
    Dim strDate As String
    Dim YYYY, MM, DD As String
    Dim TotalExRow As Long



    XpathOfTDS = Common_xml & "/TDSonSalaries/TDSonSalary"   '"ns2:ITR/TDSonSalaries"
    Set Nodelist = dom.SelectNodes(XpathOfTDS)

    TANNoEmployer = Range("TDSal.TAN").Column
    TDSNameOfEmployer = Range("TDSal.EmployerOrDeductorOrCollecterName").Column
    TDSIncomeCharge = Range("TDSal.IncChrgSalary").Column
    TDSTotalTax = Range("TDSal.TotalTDSSalary").Column
    
    TotalExRow = Range("TDSal.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
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
        For Each InnerNode In node.ChildNodes
        UpdateProgressBar
            Sheet2.Cells(rowcount, TANNoEmployer).Value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet2.Cells(rowcount, TDSNameOfEmployer).Value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            Sheet2.Cells(rowcount, TDSIncomeCharge).Value = UCase(node.SelectSingleNode("IncChrgSal").text)
            Sheet2.Cells(rowcount, TDSTotalTax).Value = UCase(node.SelectSingleNode("TotalTDSSal").text)
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
End Sub




Sub TDSOthXMLImport()
On Error Resume Next
    Dim XpathOfTDS2 As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim TANColNo, DEDNameColNo, DEDAmountDeducted, UTNColNo, FYColNo, TaxColNo, ClaimColNo, BroughtFwdTDSAmt As Variant
    
    
    
   
    
    XpathOfTDS2 = Common_xml & "/TDSonOthThanSals/TDSonOthThanSal"  '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOfTDS2)
    
    TANColNo = Sheet2.Range("TDSoth.TAN").Column
    DEDNameColNo = Sheet2.Range("TDSoth.EmployerOrDeductorOrCollecterName").Column
    DEDAmountDeducted = Sheet2.Range("TDSoth.AmountDeducted").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    FYColNo = Sheet2.Range("TDSoth.DeductedYear").Column
    BroughtFwdTDSAmt = Sheet2.Range("TDSoth.TotTDSOnAmtPaid").Column
    TaxColNo = Sheet2.Range("TDSoth.6income").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
    
    TotalExRow = Range("TDSoth.TAN").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet2.Range("TDSoth.TAN").ClearContents
        Sheet2.Range("TDSoth.EmployerOrDeductorOrCollecterName").ClearContents
        Sheet2.Range("TDSoth.AmountDeducted").ClearContents
        Sheet2.Range("TDSoth.DeductedYear").ClearContents
        Sheet2.Range("TDSoth.TotTDSOnAmtPaid").ClearContents
        Sheet2.Range("TDSoth.6income").ClearContents
        'Sheet2.Range("TDSoth.AmtClaimedBySpouse").ClearContents
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
        For Each InnerNode In node.ChildNodes
            Sheet2.Cells(rowcount, TANColNo).Value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/TAN").text)
            Sheet2.Cells(rowcount, DEDNameColNo).Value = UCase(node.SelectSingleNode("EmployerOrDeductorOrCollectDetl/EmployerOrDeductorOrCollecterName").text)
            Sheet2.Cells(rowcount, DEDAmountDeducted).Value = UCase(node.SelectSingleNode("AmtForTaxDeduct").text)
            'Sheet2.Cells(rowcount, UTNColNo).Value = UCase(Node.SelectSingleNode("UniqueTDSCerNo").Text)
            
           Sheet2.Cells(rowcount, FYColNo).Value = UCase(node.SelectSingleNode("DeductedYr").text)
            
            Sheet2.Cells(rowcount, BroughtFwdTDSAmt).Value = UCase(node.SelectSingleNode("TotTDSOnAmtPaid").text)
            Sheet2.Cells(rowcount, TaxColNo).Value = node.SelectSingleNode("ClaimOutOfTotTDSOnAmtPaid").text
            'Sheet2.Cells(rowcount, ClaimColNo).Value = Node.SelectSingleNode("AmtClaimedBySpouse").Text
         
        Next InnerNode
        cnt = cnt + 1
    Next node
End If
    RecTDS1 = cnt
   
End Sub

Sub AddDiffRows_TDSoth(DiffRows As Long)
    setDiffTblinfo_TDSoth
    Sheet2.Activate
    searchLastRow ("TDSoth.TAN")
    insertRowUnderSectionWithFormula (DiffRows)
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TDSoth)
End Sub


Sub setDiffTblinfo_TDSoth()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet2.Range("TDSoth.TAN").count
    Set rangecells = Sheet2.Range("TDSoth.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDSoth = ccount
    DefinedgridNameRange = "TDSoth.TAN||TDSoth.EmployerOrDeductorOrCollecterName||TDsOthr.SectionTDS||TDSoth.AmountDeducted||TDSoth.DeductedYear||TDSoth.TotTDSOnAmtPaid||TDSoth.6income"
End Sub


'New 26QC

Sub TDSOthXMLImport1()
On Error Resume Next
    Dim XpathOfTDS21 As String
    Dim TotalXMLRow1 As Long
    Dim TotalDiffRow1 As Long
    Dim RecTDS11, RecTDS21, cnt1 As Long
    Dim nodeList1 As IXMLDOMNodeList
    Dim Node1 As IXMLDOMNode
    Dim InnerNode1 As IXMLDOMNode
    Dim TotalExRow1 As Long
    Dim TANColNo1, TDSAADhaar, DEDNameColNo1, DEDAmountDeducted1, UTNColNo1, FYColNo1, TaxColNo1, ClaimColNo1, BroughtFwdTDSAmt1 As Variant
    
    
    
    XpathOfTDS21 = Common_xml & "/ScheduleTDS3Dtls/TDS3Details"  '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set nodeList1 = dom.SelectNodes(XpathOfTDS21)
    
    TANColNo1 = Sheet2.Range("TDS26QB.PAN").Column
    TDSAADhaar = Sheet2.Range("TDS26QB.Aadhar_Number").Column
    DEDNameColNo1 = Sheet2.Range("TDS26QB.EmployerOrDeductorName").Column
    DEDAmountDeducted1 = Sheet2.Range("TDS26QB.AmountDeducted").Column
    'UTNColNo = Sheet2.Range("TDSoth.UTN").Column
    FYColNo1 = Sheet2.Range("TDS26QB.DeductedYear").Column
    BroughtFwdTDSAmt1 = Sheet2.Range("TDS26QB.TotTDSOnAmtPaid").Column
    TaxColNo1 = Sheet2.Range("TDS26QB.6income").Column
    'ClaimColNo = Sheet2.Range("TDSoth.AmtClaimedBySpouse").Column
    
    TotalExRow1 = Range("TDS26QB.PAN").Rows.count
    
    TotalXMLRow1 = nodeList1.Length
    TotalDiffRow1 = TotalXMLRow1 - TotalExRow1
    
    If (TotalXMLRow1 > 0) Then
        Sheet2.Range("TDS26QB.PAN").ClearContents
        Sheet2.Range("TDS26QB.Aadhar_Number").ClearContents
        Sheet2.Range("TDS26QB.EmployerOrDeductorName").ClearContents
        Sheet2.Range("TDS26QB.AmountDeducted").ClearContents
        Sheet2.Range("TDS26QB.DeductedYear").ClearContents
        Sheet2.Range("TDS26QB.TotTDSOnAmtPaid").ClearContents
        Sheet2.Range("TDS26QB.6income").ClearContents
        'Sheet2.Range("TDSoth.AmtClaimedBySpouse").ClearContents
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
        For Each InnerNode1 In Node1.ChildNodes
            Sheet2.Cells(rowcount, TANColNo1).Value = UCase(Node1.SelectSingleNode("PANofTenant").text)
             Sheet2.Cells(rowcount, TDSAADhaar).Value = UCase(Node1.SelectSingleNode("AadhaarofTenant").text)
            Sheet2.Cells(rowcount, DEDNameColNo1).Value = UCase(Node1.SelectSingleNode("NameOfTenant").text)
            Sheet2.Cells(rowcount, DEDAmountDeducted1).Value = UCase(Node1.SelectSingleNode("GrsRcptToTaxDeduct").text)
            'Sheet2.Cells(rowcount, UTNColNo).Value = UCase(Node.SelectSingleNode("UniqueTDSCerNo").Text)
            Sheet2.Cells(rowcount, FYColNo1).Value = UCase(Node1.SelectSingleNode("DeductedYr").text)
            Sheet2.Cells(rowcount, BroughtFwdTDSAmt1).Value = UCase(Node1.SelectSingleNode("TDSDeducted").text)
            Sheet2.Cells(rowcount, TaxColNo1).Value = Node1.SelectSingleNode("TDSClaimed").text
           
            'Sheet2.Cells(rowcount, ClaimColNo).Value = Node.SelectSingleNode("AmtClaimedBySpouse").Text
        Next InnerNode1
        cnt1 = cnt1 + 1
    Next Node1
End If
    RecTDS11 = cnt1
    
End Sub






Sub AddDiffRows_TDSoth1(DiffRows As Long)
    setDiffTblinfo_TDSoth1
    Sheet2.Activate
    searchLastRow ("TDS26QB.PAN")
    insertRowUnderSectionWithFormula (DiffRows)
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TDSoth)
End Sub
Sub setDiffTblinfo_TDSoth1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet2.Range("TDS26QB.PAN").count
    Set rangecells = Sheet2.Range("TDSoth.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDSoth = ccount
    'DefinedgridNameRange = "TDS26QB.PAN||TDS26QB.EmployerOrDeductorName||TDsOthr.SectionTDS_ii||TDS26QB.AmountDeducted||TDS26QB.DeductedYear||TDS26QB.TotTDSOnAmtPaid||TDS26QB.6income"
     DefinedgridNameRange = "TDS26QB.PAN||TDS26QB.EmployerOrDeductorName||TDsOthr2.SectionTDSDeducted||TDsOthr.SectionTDS_ii||TDS26QB.AmountDeducted||TDS26QB.DeductedYear||TDS26QB.TotTDSOnAmtPaid||TDS26QB.6income"
End Sub


Function searchLastRow(ByVal gridRangeName As String) As String
On Error Resume Next
 strCurrActiveCellRange = Replace(ActiveSheet.Range(gridRangeName).AddressLocal, "$", "")
 strNewActiveCellRange = Mid(strCurrActiveCellRange, InStr(1, strCurrActiveCellRange, ":") + 1, Len(strCurrActiveCellRange))
 ActiveSheet.Range(strNewActiveCellRange).Select
 searchLastRow = strCurrActiveCellRange
End Function

Function getRowNo(RangeAddress As String) As Long
    Dim ColStart As Long
    Dim ColEnd As Long
    
    ColStart = InStr(1, RangeAddress, "$")
    ColEnd = InStr(ColStart + 1, RangeAddress, "$")
    ColStart = ColEnd
    ColEnd = InStr(ColStart + 1, RangeAddress, "$")
    
    getRowNo = Mid(Mid(RangeAddress, ColStart + 1, ColEnd - ColStart - 1), 1, Len(Mid(RangeAddress, ColStart + 1, ColEnd - ColStart - 1)) - 1)
End Function

Sub AddDiffRows_TDS1(DiffRows As Long)
    setDiffTblinfo_TDS1
    Sheet2.Activate
    searchLastRow ("TDSal.TAN")
    insertRowUnderSectionWithFormula (DiffRows)
    'InsertDiffRowsAndFillFormulas (DiffRows)
    'Call ExendRangeNameToTable(DiffRows, rngname_TaxP)
End Sub

Sub AddDiffRows_BANK(DiffRows As Long)
    setTableInfo_BANK
    Sheet3.Activate
    searchLastRow ("SchBA.IFSC")
    insertRowUnderSectionWithFormula (DiffRows)
    
    'LinkCheckBoxes
End Sub

Sub setTableInfo_BANK()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.IFSC").count
    Set rangecells = Range("SchBA.IFSC").Cells
    For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).Value = "") Or (rangecells.item(mIntCtr).Value = "(Select)")) Then
               ccount = ccount + 1
           End If
    Next
    'DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.AcntType||SchBA.CheckBox"
    'Added by Shrutika(24/04/2025)
     DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.AcntType||SchBA.CheckBox||tempxml"
End Sub

Sub setDiffTblinfo_TDS1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Range("TDSal.TAN").count
    Set rangecells = Range("TDSal.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    DefinedgridNameRange = "TDSal.TAN||TDSal.EmployerOrDeductorOrCollecterName||TDSal.IncChrgSalary||TDSal.TotalTDSSalary"
End Sub
 
Sub AddDiffRows_TCS(DiffRows As Long)
    setTblinfo_TCS
    Sheet11.Activate
    searchLastRow ("TCS.TAN")
    EfilingCommon.insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub ExemptXMLImport()
On Error Resume Next
    Dim XpathOfExempt As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim NatureColNo, DescriptionColNo, AmtColNo As Long
    
    Dim TotalExRow As Long
   
    XpathOfExempt = Common_xml & "/ITR1_IncomeDeductions/ExemptIncAgriOthUs10/ExemptIncAgriOthUs10Dtls"
    Set Nodelist = dom.SelectNodes(XpathOfExempt)

    NatureColNo = Sheet1.Range("Others.NOI").Column
    DescriptionColNo = Sheet1.Range("Nature_Others").Column
    AmtColNo = Sheet1.Range("Others.Amount").Column
    
    TotalExRow = Range("Others.NOI").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet1.Range("Others.NOI").ClearContents
        Sheet1.Range("Nature_Others").ClearContents
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
        For Each InnerNode In node.ChildNodes
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "OTH" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(34)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(34) (Exempted Dividend Income)"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(26AAA)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(26AAA) Any income as referred to in section 10(26AAA)"
            End If
            
             If UCase(node.SelectSingleNode("NatureDesc").text) = "10(26)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(26) Any income as referred to in section 10(26)"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(19)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(19) Armed Forces Family pension in case of death during operational duty"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(10D)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10D) Any sum received under a life insurance policy, including the sum allocated by way of bonus on such policy except sum as mentioned in sub-clause (a) to (d) of Sec.10(1)"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(11)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(11) Statutory Provident Fund received"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(12)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(12) Recognized Provident Fund received"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(13)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(13) Approved superannuation fund received"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(16)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(16) Scholarships granted to meet the cost of education"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(17)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(17)-Allowance MP/MLA/MLC" 'Ankita
            End If
            
             If UCase(node.SelectSingleNode("NatureDesc").text) = "10(17A)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(17A) Award instituted by Government"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(18)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(18)-Pension received by winner of  ""Param Vir Chakra"" or ""Maha Vir Chakra"" or ""Vir Chakra"" or such other gallantry award"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "10(10BC)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10BC) Any amount from the Central/State Govt./local authority by way of compensation on account of any disaster"
            End If
            
            If UCase(node.SelectSingleNode("NatureDesc").text) = "DMDP" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Defense Medical Disability Pension"
            End If
        
            If UCase(node.SelectSingleNode("NatureDesc").text) = "AGRI" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Agriculture Income (less than equal to Rs.5000)"
            End If
            
            Sheet1.Cells(rowcount, DescriptionColNo).Value = node.SelectSingleNode("OthNatOfInc").text
            Sheet1.Cells(rowcount, AmtColNo).Value = node.SelectSingleNode("OthAmount").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub
Sub ExemptXMLImport1()
On Error Resume Next
    Dim XpathOfExempt As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim NatureColNo, DescriptionColNo, AmtColNo As Long
    
    Dim TotalExRow As Long
   
    XpathOfExempt = Common_xml & "/ITR1_IncomeDeductions/OthersInc/OthersIncDtlsOthSrc"  '"/ns2:ITR/TaxPayments/TaxPayment"
    Set Nodelist = dom.SelectNodes(XpathOfExempt)

    NatureColNo = Sheet1.Range("Others.NOI_2").Column
    DescriptionColNo = Sheet1.Range("Nature_Others_2").Column
    AmtColNo = Sheet1.Range("Others.Amount_2").Column
    
    TotalExRow = Range("Others.NOI_2").Rows.count
    
    TotalXMLRow = Nodelist.Length
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
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "SAV" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Savings Bank Account"
            End If
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "IFD" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Deposit (Bank/Post Office/Cooperative Society)"
            End If
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "TAX" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Interest from Income Tax Refund"
            End If
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "FAP" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Family pension"
            End If
            
            
            If UCase(node.SelectSingleNode("OthSrcNatureDesc").text) = "OTH" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
            End If
            
            Sheet1.Cells(rowcount, DescriptionColNo).Value = node.SelectSingleNode("OthSrcOthNatOfInc").text
            Sheet1.Cells(rowcount, AmtColNo).Value = node.SelectSingleNode("OthSrcOthAmount").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub
Sub ExemptXMLImport2()
On Error Resume Next
    Dim XpathOfExempt As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1 As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim NatureColNo, DescriptionColNo, AmtColNo As Long
    
    Dim TotalExRow As Long
   
    XpathOfExempt = Common_xml & "/ITR1_IncomeDeductions/AllwncExemptUs10/AllwncExemptUs10Dtls"  '"/ns2:ITR/TaxPayments/TaxPayment"
    Set Nodelist = dom.SelectNodes(XpathOfExempt)

    NatureColNo = Sheet1.Range("Others.NOI_1").Column
    DescriptionColNo = Sheet1.Range("Nature_Others_1").Column
    AmtColNo = Sheet1.Range("Others.Amount_1").Column
    
    TotalExRow = Range("Others.NOI_1").Rows.count
    
    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow
    
    If (TotalXMLRow > 0) Then
        Sheet1.Range("Others.NOI_1").ClearContents
        Sheet1.Range("Nature_Others_1").ClearContents
        Sheet1.Range("Others.Amount_1").ClearContents
        
    End If
    
    If (TotalDiffRow > 0) Then
     AddDiffRows_Exempt2 (TotalDiffRow)
    End If
 
    rowcount = getRowNo(Sheet1.Range("Others.NOI_1").name)
    rowcount = rowcount - 1
    cnt = 0
    
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            
            Sheet1.Cells(rowcount, NatureColNo).Value = Findtext("Sec " & (node.SelectSingleNode("SalNatureDesc").text), "PART_Nature_1")
            
            
            If UCase(node.SelectSingleNode("SalNatureDesc").text) = "10(10B)(i)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette"
            End If
            
            
            If UCase(node.SelectSingleNode("SalNatureDesc").text) = "10(10B)(ii)" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government"
            End If
            
            If UCase(node.SelectSingleNode("SalNatureDesc").text) = "OTH" Then
            Sheet1.Cells(rowcount, NatureColNo).Value = "Any Other"
            End If
            
            Sheet1.Cells(rowcount, DescriptionColNo).Value = node.SelectSingleNode("SalOthNatOfInc").text
            Sheet1.Cells(rowcount, AmtColNo).Value = node.SelectSingleNode("SalOthAmount").text
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub

Sub setDiffTblinfo_Exempt()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Others.NOI").count
    Set rangecells = Sheet1.Range("Others.NOI").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    'Malli_AY_2026_27 17/08/2026
    'DefinedgridNameRange = "Others.NOI||Nature_Others||Others.Amount"
    DefinedgridNameRange = "Others.NOI||Nature_Others||Sheet5.DescEI||Others.Amount"
    '----------------------------
End Sub
Sub AddDiffRows_Exempt(DiffRows As Long)
    setDiffTblinfo_Exempt
    Sheet1.Activate
    searchLastRow ("Others.NOI")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub AddDiffRows_Exempt1(DiffRows As Long)
    setDiffTblinfo_Exempt1
    Sheet1.Activate
    searchLastRow ("Others.NOI_2")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub setDiffTblinfo_Exempt1()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Others.NOI_2").count
    Set rangecells = Sheet1.Range("Others.NOI_2").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    DefinedgridNameRange = "Others.NOI_2||Nature_Others_2||Others.Amount_2||Others.Amount_2_1||Others.Amount_2_2||Others.Amount_2_3||Others.Amount_2_4||Others.Amount_2_5"
 End Sub
 Sub AddDiffRows_Exempt2(DiffRows As Long)
    setDiffTblinfo_Exempt2
    Sheet1.Activate
    searchLastRow ("Others.NOI_1")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub setDiffTblinfo_Exempt2()
   Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet1.Range("Others.NOI_1").count
    Set rangecells = Sheet1.Range("Others.NOI_1").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
'Konda updated 0n 12-03-2026---V0.5
'    DefinedgridNameRange = "Others.NOI_1||Nature_Others_1||Others.Amount_1"
     DefinedgridNameRange = "Others.NOI_1||Others.Amount_1"
 End Sub
Sub XMLImport_80G_AP()
On Error Resume Next
    Dim XpathOf80G_A As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
    
    
  
    
    XpathOf80G_A = Common_xml & "/Schedule80G/Don100Percent/DoneeWithPan"  '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOf80G_A)
    
    NameColNo = Sheet4.Range("Per10080G.DoneeName").Column
    AddressColNo = Sheet4.Range("Per10080G.AddrDetail").Column
    CityColNo = Sheet4.Range("Per10080G.CityOrTownOrDistrict").Column
    StateCodeColNo = Sheet4.Range("Per10080G.StateCode").Column
    PincodeColNo = Sheet4.Range("Per10080G.PinCode").Column
    PanofDoneeColNo = Sheet4.Range("Per10080G.DoneePAN").Column
    AmountColNo = Sheet4.Range("Per10080G.DonationAmt").Column
    DonationColNo = Sheet4.Range("Per10080G.DonationAmtOther").Column
    
    TotalExRow = Range("Per10080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.Length
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
        For Each InnerNode In node.ChildNodes
            Sheet4.Cells(rowcount, NameColNo).Value = UCase(node.SelectSingleNode("DoneeWithPanName").text)
            Sheet4.Cells(rowcount, AddressColNo).Value = UCase(node.SelectSingleNode("AddressDetail/AddrDetail").text)
            Sheet4.Cells(rowcount, CityColNo).Value = UCase(node.SelectSingleNode("AddressDetail/CityOrTownOrDistrict").text)
            Dim iState As Variant
            iState = UCase(node.SelectSingleNode("AddressDetail/StateCode").text)
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            
            
            If iState = "99" Then
            iState = ""
            End If
            
            Sheet4.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet4.Cells(rowcount, PincodeColNo).Value = UCase(node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet4.Cells(rowcount, PanofDoneeColNo).Value = node.SelectSingleNode("DoneePAN").text
            Sheet4.Cells(rowcount, AmountColNo).Value = node.SelectSingleNode("DonationAmtCash").text
            Sheet4.Cells(rowcount, DonationColNo).Value = node.SelectSingleNode("DonationAmtOtherMode").text
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub
Sub setDiffTblinfo_80G_A()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet4.Range("Per10080G.DoneeName").count
    Set rangecells = Sheet4.Range("Per10080G.DoneeName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
'Konda updated on 10-03-2026--V0.5
'    DefinedgridNameRange = "Per10080G.DoneeName||Per10080G.AddrDetail||Per10080G.CityOrTownOrDistrict||Per10080G.StateCode||Per10080G.PinCode||Per10080G.DoneePAN||Per10080G.DonationAmt||Per10080G.DonationAmtOther||Per10080G.DonationAmtTotal||Per10080G.EligibleAmt"
    DefinedgridNameRange = "Per10080G.DoneeName||Per10080G.AddrDetail||Per10080G.CityOrTownOrDistrict||Per10080G.StateCode||Per10080G.PinCode||Per10080G.DoneePAN||Per10080G.DonationAmt||Per10080G.DonationAmtOther||Per10080G.Traref||Per10080G.IFSC||Per10080G.DonationAmtTotal||Per10080G.EligibleAmt"
'========================
End Sub

Sub AddDiffRows_80G_A(DiffRows As Long)
    setDiffTblinfo_80G_A
    Sheet4.Activate
    searchLastRow ("Per10080G.DoneeName")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub XMLImport_80G_BP()
On Error Resume Next
    Dim XpathOf80G_B As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
    
    
  
    XpathOf80G_B = Common_xml & "/Schedule80G/Don50PercentNoApprReqd/DoneeWithPan" '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOf80G_B)
    
    NameColNo = Sheet4.Range("PerNO5080G.DoneeName").Column
    AddressColNo = Sheet4.Range("PerNO5080G.AddrDetail").Column
    CityColNo = Sheet4.Range("PerNO5080G.CityOrTownOrDistrict").Column
    StateCodeColNo = Sheet4.Range("PerNO5080G.StateCode").Column
    PincodeColNo = Sheet4.Range("PerNO5080G.PinCode").Column
    PanofDoneeColNo = Sheet4.Range("PerNO5080G.DoneePAN").Column
    AmountColNo = Sheet4.Range("PerNO5080G.DonationAmt").Column
    DonationColNo = Sheet4.Range("PerNO5080G.DonationAmtOther").Column
    
    TotalExRow = Range("PerNO5080G.DoneeName").Rows.count
    
    TotalXMLRow = Nodelist.Length
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
        For Each InnerNode In node.ChildNodes
            Sheet4.Cells(rowcount, NameColNo).Value = UCase(node.SelectSingleNode("DoneeWithPanName").text)
            Sheet4.Cells(rowcount, AddressColNo).Value = UCase(node.SelectSingleNode("AddressDetail/AddrDetail").text)
            Sheet4.Cells(rowcount, CityColNo).Value = UCase(node.SelectSingleNode("AddressDetail/CityOrTownOrDistrict").text)
            Dim iState As Variant
            iState = UCase(node.SelectSingleNode("AddressDetail/StateCode").text)
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet4.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet4.Cells(rowcount, PincodeColNo).Value = UCase(node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet4.Cells(rowcount, PanofDoneeColNo).Value = node.SelectSingleNode("DoneePAN").text
            Sheet4.Cells(rowcount, AmountColNo).Value = node.SelectSingleNode("DonationAmtCash").text
            Sheet4.Cells(rowcount, DonationColNo).Value = node.SelectSingleNode("DonationAmtOtherMode").text
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub
Sub setDiffTblinfo_80G_B()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet4.Range("PerNO5080G.DoneeName").count
    Set rangecells = Sheet4.Range("PerNO5080G.DoneeName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
'Konda updated on 10-03-2026---V0.5
'    DefinedgridNameRange = "PerNO5080G.DoneeName||PerNO5080G.AddrDetail||PerNO5080G.CityOrTownOrDistrict||PerNO5080G.StateCode||PerNO5080G.PinCode||PerNO5080G.DoneePAN||PerNO5080G.DonationAmt||PerNO5080G.DonationAmtOther||PerNO5080G.DonationAmtTotal||PerNO5080G.EligibleAmt"
    DefinedgridNameRange = "PerNO5080G.DoneeName||PerNO5080G.AddrDetail||PerNO5080G.CityOrTownOrDistrict||PerNO5080G.StateCode||PerNO5080G.PinCode||PerNO5080G.DoneePAN||PerNO5080G.DonationAmt||PerNO5080G.DonationAmtOther||PerNO5080G.Traref||PerNO5080G.IFSC||PerNO5080G.DonationAmtTotal||PerNO5080G.EligibleAmt"
'======================
End Sub

Sub AddDiffRows_80G_B(DiffRows As Long)
    setDiffTblinfo_80G_B
    Sheet4.Activate
    searchLastRow ("PerNO5080G.DoneeName")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub

Sub XMLImport_80G_CP()
On Error Resume Next
    Dim XpathOf80G_C As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
    
    XpathOf80G_C = Common_xml & "/Schedule80G/Don100PercentApprReqd/DoneeWithPan"  '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOf80G_C)
    
    NameColNo = Sheet4.Range("PerYES10080G.DoneeWithPanName").Column
    AddressColNo = Sheet4.Range("PerYES10080G.AddrDetail").Column
    CityColNo = Sheet4.Range("PerYES10080G.CityOrTownOrDistrict").Column
    StateCodeColNo = Sheet4.Range("PerYES10080G.StateCode").Column
    PincodeColNo = Sheet4.Range("PerYES10080G.PinCode").Column
    PanofDoneeColNo = Sheet4.Range("PerYES10080G.DoneePAN").Column
    AmountColNo = Sheet4.Range("PerYES10080G.DonationAmt").Column
    DonationColNo = Sheet4.Range("PerYES10080G.DonationAmtOther").Column
    
    TotalExRow = Range("PerYES10080G.DoneeWithPanName").Rows.count
    
    TotalXMLRow = Nodelist.Length
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
     AddDiffRows_80G_C (TotalDiffRow)
    End If
    
    
    rowcount = getRowNo(Sheet4.Range("PerYES10080G.DoneeWithPanName").name)
    rowcount = rowcount - 1
    cnt = 0
    For Each node In Nodelist
        rowcount = rowcount + 1
        For Each InnerNode In node.ChildNodes
            Sheet4.Cells(rowcount, NameColNo).Value = UCase(node.SelectSingleNode("DoneeWithPanName").text)
            Sheet4.Cells(rowcount, AddressColNo).Value = UCase(node.SelectSingleNode("AddressDetail/AddrDetail").text)
            Sheet4.Cells(rowcount, CityColNo).Value = UCase(node.SelectSingleNode("AddressDetail/CityOrTownOrDistrict").text)
            Dim iState As Variant
            iState = UCase(node.SelectSingleNode("AddressDetail/StateCode").text)
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            
            Sheet4.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet4.Cells(rowcount, PincodeColNo).Value = UCase(node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet4.Cells(rowcount, PanofDoneeColNo).Value = node.SelectSingleNode("DoneePAN").text
            Sheet4.Cells(rowcount, AmountColNo).Value = node.SelectSingleNode("DonationAmtCash").text
            Sheet4.Cells(rowcount, DonationColNo).Value = node.SelectSingleNode("DonationAmtOtherMode").text
            
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub

Sub setDiffTblinfo_80G_C()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet4.Range("PerYES10080G.DoneeWithPanName").count
    Set rangecells = Sheet4.Range("PerYES10080G.DoneeWithPanName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
'Konda updated on 10-03-2026--V0.5
'    DefinedgridNameRange = "PerYES10080G.DoneeWithPanName||PerYES10080G.AddrDetail||PerYES10080G.CityOrTownOrDistrict||PerYES10080G.StateCode||PerYES10080G.PinCode||PerYES10080G.DoneePAN||PerYES10080G.DonationAmt||PerYES10080G.DonationAmtOther||PerYES10080G.DonationAmtTotal||PerYES10080G.EligibleAmt"
    DefinedgridNameRange = "PerYES10080G.DoneeWithPanName||PerYES10080G.AddrDetail||PerYES10080G.CityOrTownOrDistrict||PerYES10080G.StateCode||PerYES10080G.PinCode||PerYES10080G.DoneePAN||PerYES10080G.DonationAmt||PerYES10080G.DonationAmtOther||PerYES10080G.Traref||PerYES10080G.IFSC||PerYES10080G.DonationAmtTotal||PerYES10080G.EligibleAmt"
'==========================
End Sub
Sub AddDiffRows_80G_C(DiffRows As Long)
    setDiffTblinfo_80G_C
    Sheet4.Activate
    searchLastRow ("PerYES10080G.DoneeWithPanName")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub

Sub XMLImport_80G_DP()
On Error Resume Next
    Dim XpathOf80G_D As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim RecTDS1, RecTDS2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant
    
   
    
    
    
    XpathOf80G_D = Common_xml & "/Schedule80G/Don50PercentApprReqd/DoneeWithPan"  '"ns2:ITR/TDSonOthThanSals/TDSonOthThanSal"
    Set Nodelist = dom.SelectNodes(XpathOf80G_D)
    
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
        For Each InnerNode In node.ChildNodes
            Sheet4.Cells(rowcount, NameColNo).Value = UCase(node.SelectSingleNode("DoneeWithPanName").text)
            Sheet4.Cells(rowcount, AddressColNo).Value = UCase(node.SelectSingleNode("AddressDetail/AddrDetail").text)
            Sheet4.Cells(rowcount, CityColNo).Value = UCase(node.SelectSingleNode("AddressDetail/CityOrTownOrDistrict").text)
            Dim iState As Variant
            iState = UCase(node.SelectSingleNode("AddressDetail/StateCode").text)
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If
            If iState = "99" Then
            iState = ""
            End If
            Sheet4.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet4.Cells(rowcount, PincodeColNo).Value = UCase(node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet4.Cells(rowcount, PanofDoneeColNo).Value = node.SelectSingleNode("DoneePAN").text
            Sheet4.Cells(rowcount, AmountColNo).Value = node.SelectSingleNode("DonationAmtCash").text
            Sheet4.Cells(rowcount, DonationColNo).Value = node.SelectSingleNode("DonationAmtOtherMode").text
            
        Next InnerNode
        cnt = cnt + 1
    Next node
    RecTDS1 = cnt
    
End Sub

Sub setDiffTblinfo_80G_D()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet4.Range("Per5080G.DoneeWithPanName").count
    Set rangecells = Sheet4.Range("Per5080G.DoneeWithPanName").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    'DefinedgridNameRange = "Per5080G.DoneeWithPanName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.DonationAmt||Per5080G.DonationAmtOther||Per5080G.DonationAmtTotal||Per5080G.EligibleAmt"
''Change.25.01.2023.102.80GC6
'     DefinedgridNameRange = "Per5080G.DoneeWithPanName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.DonationAmt||Per5080G.DonationAmtOther||Per5080G.DonationAmtTotal||Per5080G.EligibleAmt||Per5080G.DonationAmt_temp"
'Commented by Konda 10-03-2026---V0.5
'     DefinedgridNameRange = "Per5080G.DoneeWithPanName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.ArnNbr||Per5080G.DonationAmt||Per5080G.DonationAmtOther||Per5080G.DonationAmtTotal||Per5080G.EligibleAmt||Per5080G.DonationAmt_temp"
'============================
'End Change
'Konda updated on 10-03-2026---V0.5
    DefinedgridNameRange = "Per5080G.DoneeWithPanName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.ArnNbr||Per5080G.DonationAmt||Per5080G.DonationAmtOther||Per5080G.Traref||Per5080G.IFSC||Per5080G.DonationAmtTotal||Per5080G.EligibleAmt||Per5080G.DonationAmt_temp"
'===================
End Sub


Sub AddDiffRows_80G_D(DiffRows As Long)
    setDiffTblinfo_80G_D
    Sheet4.Activate
    searchLastRow ("Per5080G.DoneeWithPanName")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub
Sub XMLImport_80GGA()
On Error Resume Next
    Dim XpathOf80GGA As String
    Dim TotalXMLRow As Long
    Dim TotalDiffRow As Long
    Dim Rec80GGA1, Rec80GGA2, cnt As Long
    Dim Nodelist As IXMLDOMNodeList
    Dim node As IXMLDOMNode
    Dim InnerNode As IXMLDOMNode
    Dim TotalExRow As Long
    Dim RelevantClause, NameColNo, AddressColNo, CityColNo, StateCodeColNo, PincodeColNo, PanofDoneeColNo, AmountColNo, DonationColNo As Variant

    XpathOf80GGA = Common_xml & "/Schedule80GGA/DonationDtlsSciRsrchRuralDev"
    Set Nodelist = dom.SelectNodes(XpathOf80GGA)

    RelevantClause = Sheet12.Range("RelevantClauseClaimed_80GGA").Column
    NameColNo = Sheet12.Range("Name_of_Donee_80GGA").Column
    AddressColNo = Sheet12.Range("Address_80GGA").Column
    CityColNo = Sheet12.Range("City_Town_District_80GGA").Column
    StateCodeColNo = Sheet12.Range("State_Code_80GGA").Column
    PincodeColNo = Sheet12.Range("Pincode_80GGA").Column
    PanofDoneeColNo = Sheet12.Range("PAN_of_donee_80GGA").Column
    AmountColNo = Sheet12.Range("Donation_cash_80GGA").Column
    DonationColNo = Sheet12.Range("Donation_other_80GGA").Column

    TotalExRow = Range("RelevantClauseClaimed_80GGA").Rows.count

    TotalXMLRow = Nodelist.Length
    TotalDiffRow = TotalXMLRow - TotalExRow

    If (TotalXMLRow > 0) Then
        Sheet12.Range("RelevantClauseClaimed_80GGA").ClearContents
        Sheet12.Range("Name_of_Donee_80GGA").ClearContents
        Sheet12.Range("Address_80GGA").ClearContents
        Sheet12.Range("City_Town_District_80GGA").ClearContents
        Sheet12.Range("State_Code_80GGA").ClearContents
        Sheet12.Range("Pincode_80GGA").ClearContents
        Sheet12.Range("PAN_of_donee_80GGA").ClearContents
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
        For Each InnerNode In node.ChildNodes
         
        Temp80GGA = (node.SelectSingleNode("RelevantClauseUndrDedClaimed").text)
        
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
            Sheet12.Cells(rowcount, NameColNo).Value = UCase(node.SelectSingleNode("NameOfDonee").text)
            Sheet12.Cells(rowcount, AddressColNo).Value = UCase(node.SelectSingleNode("AddressDetail/AddrDetail").text)
            Sheet12.Cells(rowcount, CityColNo).Value = UCase(node.SelectSingleNode("AddressDetail/CityOrTownOrDistrict").text)

            Dim iState As Variant
            iState = UCase(node.SelectSingleNode("AddressDetail/StateCode").text)
            If Len(iState) = "1" Then
            iState = "0" & iState
            End If


            If iState = "99" Then
            iState = ""
            End If

            Sheet12.Cells(rowcount, StateCodeColNo).Value = Findtext(iState, "StateList")
            Sheet12.Cells(rowcount, PincodeColNo).Value = UCase(node.SelectSingleNode("AddressDetail/PinCode").text)
            Sheet12.Cells(rowcount, PanofDoneeColNo).Value = node.SelectSingleNode("DoneePAN").text
            Sheet12.Cells(rowcount, AmountColNo).Value = node.SelectSingleNode("DonationAmtCash").text
            Sheet12.Cells(rowcount, DonationColNo).Value = node.SelectSingleNode("DonationAmtOtherMode").text

        Next InnerNode
        cnt = cnt + 1
    Next node
    Rec80GGA1 = cnt
End Sub

Sub setDiffTblinfo_80GGA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long

    ccount = 0
    mIntCells = Sheet12.Range("RelevantClauseClaimed_80GGA").count
    Set rangecells = Sheet12.Range("RelevantClauseClaimed_80GGA").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    DefinedgridNameRange = "RelevantClauseClaimed_80GGA||Name_of_Donee_80GGA||Address_80GGA||City_Town_District_80GGA||State_Code_80GGA||Pincode_80GGA||PAN_of_donee_80GGA||Donation_cash_80GGA||Donation_other_80GGA||Donation_total_80GGA||Donation_Eligible_80GGA"
End Sub
Sub AddDiffRows_80GGA(DiffRows As Long)
    setDiffTblinfo_80GGA
    Sheet12.Activate
    searchLastRow ("RelevantClauseClaimed_80GGA")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub

Sub Prefillexml()

InitProgBar
ProgressFrameCaption = "Importing XML "

noOfProcessMain = 5
noOfProcessSub = 18
UserForm1.Show vbModeless

If Not ValidateXML() Then
   fmsgbox "Invalid XML.Please Retry."
  
   
 Else
   mainProcCaption = "Importing 80G"
    UpdateProgressBar
    ExemptXMLImport1
    XMLImport_80G_AP
    UpdateProgressBar
    XMLImport_80G_DP
    UpdateProgressBar
    XMLImport_80G_CP
    UpdateProgressBar
    XMLImport_80G_BP
    UpdateProgressBar
    XMLImport_80GGA
    UpdateProgressBar
    mainProcCaption = "Importing Exempt Income"
    ExemptXMLImport
    ExemptXMLImport2
    UpdateProgressBar
    sec38XMLImport
    UpdateProgressBar
    mainProcCaption = "Importing Direct Investments"
    UpdateProgressBar
    mainProcCaption = "Importing Salary and Deductions"
    UpdateProgressBar
    SalaryXMLImport
    UpdateProgressBar
    mainProcCaption = "Importing Schedule 80D"
    UpdateProgressBar
    Sch80DXMLImport
    UpdateProgressBar
    mainProcCaption = "Importing TDS"
    UpdateProgressBar
    TDSonSalaryXMLImport
    UpdateProgressBar
    TDSOthXMLImport
    UpdateProgressBar
    TDSOthXMLImport1
    UpdateProgressBar
    ITXMLImport
    mainProcCaption = "Importing TCS"
    UpdateProgressBar
    TCSXMLImport
     mainProcCaption = "Importing Verification Details"
    UpdateProgressBar
    VeriInfoXMLImport
    UpdateProgressBar
    RefundInfoXMLImport
    mainProcCaption = "Importing Personal Information"
    UpdateProgressBar
    PersonalInfoXMLImport
    mainProcCaption = "Importing Filing Status"
    UpdateProgressBar
    FilingInfoXMLImport
    UpdateProgressBar
    ReliefXMLImport
    UpdateProgressBar
    Sheet1.Activate

    fmsgbox "Filled up TDS on Salary, Others , Tax Payments and Personal Information as per the XML File used for Pre Filling." & vbCr & vbLf & "Please Verify the Data Filled"
  
    End If
End Sub


Sub AddDiffRows_80GGC(DiffRows As Long)
    setDiffTblinfo_80GGC
    Sheet13.Activate
    searchLastRow ("DateofDonation_80GGC")
    insertRowUnderSectionWithFormula (DiffRows)
End Sub

Sub setDiffTblinfo_80GGC()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long

    ccount = 0
    mIntCells = Sheet13.Range("DateofDonation_80GGC").count
    Set rangecells = Sheet13.Range("DateofDonation_80GGC").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TaxP = ccount
    'DefinedgridNameRange = "DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||NatureofTransaction_80GGC||Chequeno_80GGC||IFSC_80GGC||"
'    DefinedgridNameRange = "DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||Chequeno_80GGC||IFSC_80GGC||"
     DefinedgridNameRange = "DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||Name_80GGC||PAN_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||Chequeno_80GGC||IFSC_80GGC||"
End Sub



