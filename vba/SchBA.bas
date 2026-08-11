Attribute VB_Name = "SchBA"
Option Explicit

Public Msgbox_BA As String
Public ColCount1_BA, ColCount2_BA, ColCount3_BA, ColCount4_BA, ColCount55_B5, ColCount5_BA, ColCount6_BA, ColCount1_BA_IFSC, ColCount55_BA As Variant


'INC_C43 2024-25 Bindu
Public ColCount55_BA6, ColCount55_BA6_Other As Variant

Public BankName_BA, BankNameCode, BankAddr_BA, BankJointHolder_BA, BankAccntnum_BA, AccountType_BA, BankAccntBalanc_BA, BankIFSC_BA, BankAccountStatus_BA, tempxml, CheckBox As Variant
'INC_C43
Public BankAccnType_BA As Variant

Public IBAN As Variant
Public IBANNameOfBank As Variant
Public IBANCOUNTRY As Variant
Public IBANACCNO As Variant

Public end_IBAN, end_IBANACCNO, end_IBANCOUNTRY, end_IBANNameOfBank As Variant
Public rngname_IBAN, rngname_IBANACCNO, rngname_IBANCOUNTRY, rngname_IBANNameOfBank As Variant

Sub ValidateSheetBA()
Application.ScreenUpdating = False
Dim vbMessgaeCaption As String
vbMessgaeCaption = "ITR 1: AY: 2026-27"             'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
ValidateBA
MsgBox ("Sheet BA is OK "), vbOKOnly, vbMessgaeCaption
Application.ScreenUpdating = True
End Sub

Function CheckIFSC1(Tval As Variant, Tadd As Variant)
On Error Resume Next
Application.EnableEvents = False
 sPassword = EfilingCommon.getmsgstate
   Sheet3.Unprotect Password:=sPassword


If Tval <> "" Then
    
     If Not CheckIFSC(UCase(Tval)) Then
        fmsgbox "Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details in Sheet Taxes Paid and Verification."
           Sheet3.Range(Tadd) = ""
           Sheet3.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
    
    If Not ValidateIFSCList(UCase(Tval)) Then
           fmsgbox "Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details in Sheet Taxes Paid and Verification."
           Sheet3.Range(Tadd) = ""
           Sheet3.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
    Sheet3.Range(Tadd).Select
   
    Sheet3.Range(Tadd).Offset(0, 1) = UCase(Sheet3.Range(Tadd).Offset(0, 9).Value)
End If

If Tval = "" Then
    Sheet3.Range(Tadd).Offset(0, 1) = ""
End If

endfd:
Sheet3.Protect Password:=sPassword
Application.EnableEvents = True

End Function
Sub PrevBA_Click()
Sheet3.Activate
End Sub
Sub NextBA_Click()
Sheet4.Activate
End Sub

Sub AddRows_BA()
Dim vRows As Long
Dim sourceSheet As Worksheet

Set sourceSheet = ThisWorkbook.Sheets("Taxes Paid and Verification")
sourceSheet.Activate
'Malli
'EfilingCommon.DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.AcntType||SchBA.CheckBox||tempxml"
'EfilingCommon.DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.AcntType"
 EfilingCommon.DefinedgridNameRange = "SchBA.IFSC||SchBA.BankName||SchBA.AcntNo||SchBA.AcntType||SchBA.CheckBox||tempxml" 'Added by Ankita on 16/01/2025
 ActiveCellRange = EfilingCommon.searchLastRow("SchBA.IFSC")
 vRows = EfilingCommon.insertRowUnderSectionWithFormulaOne
'LinkCheckBoxes  'Already added in function (insertRowUnderSectionWithFormulaOne)
End Sub

Sub ValidateBA()
Application.ScreenUpdating = False
Dim vbMessgaeCaption As String
vbMessgaeCaption = "Error"
Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("Taxes Paid and Verification")

     If Not ValidateSchBA Then
        sourceSheet.Activate
        'MsgBox (Msgbox_BA), vbOKOnly, vbMessgaeCaption
        fmsgbox (Msgbox_BA)
        Application.ScreenUpdating = True
        CloseMsg
    End If

End Sub

Function ValidateSchBA() As Boolean

ValidateSchBA = True
Msgbox_BA = ""
    setTableInfo_BA_IFSC
    setTableInfo_BA
    setTableInfo_BA3
   ' setTableInfo_BA4  'For AY_2024_25 Changes Removed -Malli comented
    setTableInfo_BA5
    
    'INC_C43 2024-25 Bindu
    setTableInfo_BA6
    
    
    'ColCount1_BA_IFSC = WorksheetFunction.Max(0, ColCount1_BA, ColCount1_BA_IFSC, ColCount4_BA, ColCount55_B5, ColCount55_BA6)
    ColCount1_BA_IFSC = WorksheetFunction.Max(0, ColCount1_BA, ColCount1_BA_IFSC, ColCount4_BA, ColCount55_B5, ColCount55_BA6)
    
        'For AY_2024_25 Changes Removed -Malli comented
        'If Not ValidateCheckBox_BA Then ValidateSchBA = False
        
        If Not ValidateIFSC Then ValidateSchBA = False
       ' If Not ValidateBankName_BA Then ValidateSchBA = False
        If Not ValidateAccntNumber_BA Then ValidateSchBA = False
        
        
        

End Function
'Function ValidateCheckBox_BA() As Boolean
'ValidateCheckBox_BA = True
'Dim rangecells As Range
'Dim rangecells1 As Range
'Set rangecells = Range("SchBA.IFSC").Cells
'Set rangecells1 = Range("tempxml").Cells
'Dim i As Long
'setTableInfo_BA4
''ReDim tempxml(ColCount55_BA)
'ReDim BankIFSC_BA(ColCount55_BA)
''INC_C43 2024-25 Bindu
'Dim rangecells2 As Range
'Set rangecells2 = Range("SchBA.AcntType").Cells
'ReDim BankAccnType_BA(ColCount55_BA)
'
'
'For i = 1 To ColCount55_BA
'    BankIFSC_BA(i) = rangecells.item(i).Value
'
'    If isdropdownblank(BankIFSC_BA(i)) Then
'        Msgbox_BA = Msgbox_BA + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory." & Chr(13)
'        ValidateCheckBox_BA = False
'        Exit Function
'
'    End If
'
'    If Len(BankIFSC_BA(i)) > 11 Then
'        Msgbox_BA = Msgbox_BA + "* IFS Code at Sr.No " & i & " in Sheet Taxes Paid and Verification cannot exceed 11 characters" & Chr(13)
'        ValidateCheckBox_BA = False
'        Exit Function
'    End If
'
'    If Not ValidateIFSCList(UCase(BankIFSC_BA(i))) Then
'        Msgbox_BA = Msgbox_BA + "* Invalid IFS Code at Sr.No " & i & "  in Sheet Taxes Paid and Verification.Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
'        ValidateCheckBox_BA = False
'        Exit Function
'    End If
'
'
'
'    Next
'
'End Function

'Added by Ankita on 21/01/2025
Function ValidateCheckBox_BA() As Boolean
ValidateCheckBox_BA = True
Dim rangecells As Range
Dim rangecells1 As Range
Set rangecells = Range("SchBA.IFSC").Cells
Set rangecells1 = Range("tempxml").Cells
Dim i As Long
setTableInfo_BA4
'ReDim tempxml(ColCount55_BA)
ReDim BankIFSC_BA(ColCount55_BA)
For i = 1 To ColCount55_BA
    BankIFSC_BA(i) = rangecells.item(i).Value
   
If isdropdownblank(BankIFSC_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory." & Chr(13)
        ValidateCheckBox_BA = False
        Exit Function

    End If
    
    If Len(BankIFSC_BA(i)) > 11 Then
        Msgbox_BA = Msgbox_BA + "* IFS Code at Sr.No " & i & " in Sheet Taxes Paid and Verification cannot exceed 11 characters" & Chr(13)
        ValidateCheckBox_BA = False
        Exit Function
    End If
    
    If Not ValidateIFSCList(UCase(BankIFSC_BA(i))) Then
        Msgbox_BA = Msgbox_BA + "* Invalid IFS Code at Sr.No " & i & "  in Sheet Taxes Paid and Verification. Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
        ValidateCheckBox_BA = False
        Exit Function
    End If
    
    

    Next
 
End Function

Function ValidateIFSC() As Boolean
ValidateIFSC = True
'setTableInfo_BA_IFSC
Dim rangecells As Range
Dim rangecells1 As Range
Dim rangecells2 As Range
Dim rangecells3 As Range
Dim rangecells4 As Range

Set rangecells1 = Range("SchBA.BankName").Cells
Set rangecells2 = Range("SchBA.AcntNo").Cells
Set rangecells3 = Range("SchBA.IFSC").Cells
'Set rangecells4 = Range("tempxml").Cells
Set rangecells4 = Range("tempxml").Cells  'Added by Ankita on 16/01/2025

'INC_C43 2024-25 Bindu
Dim rangecells5 As Range
Set rangecells5 = Range("SchBA.AcntType").Cells


Dim i As Long
ReDim BankIFSC_BA(ColCount1_BA_IFSC)
ReDim BankName_BA(ColCount1_BA_IFSC)
ReDim BankAccntnum_BA(ColCount1_BA_IFSC)
'ReDim tempxml(ColCount1_BA_IFSC)
ReDim tempxml(ColCount1_BA_IFSC) 'Added by Ankita on 16/01/2025

'INC_C43
ReDim BankAccnType_BA(ColCount1_BA_IFSC)

   'INC_E43
   
'    If Not ColCount55_B5 > 0 Then
'        Msgbox_BA = Msgbox_BA + "* Please select atleast one saving/current/cash credit/overdraft/non resident bank acount  in which you prefer to get your refund" & Chr(13)
'        ValidateIFSC = False
'        Exit Function
'    End If
 'Malli---------------------------
'    'INC_C43 2024-25 Bindu
'     'If ColCount1_BA_IFSC > 0 Then
'        If ColCount55_BA6_Other > 0 Or Not ColCount55_B5 > 0 Then
'          Msgbox_BA = Msgbox_BA + "*Please select atleast one saving/current/cash credit/overdraft/non resident bank acount  in which you prefer to get your refund" & Chr(13)
'          ValidateIFSC = False
'          Exit Function
'        End If
'''    'End If

'------------------------------------

'Added by Ankita 0n 16/01/2025
ReDim tempxml(ColCount1_BA_IFSC)
    If Not ColCount55_B5 > 0 Then
        Msgbox_BA = Msgbox_BA + "* ""Please select atleast one account in which you prefer to get your refund""" & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If


'Added by Ankita 0n 16/01/2025
    If ColCount55_BA6_Other > 0 Or Not ColCount55_B5 > 0 Then
    'Changed on 20/03/2025 by Ankita
'     If Not ColCount55_BA6_Other > 0 Then
        '  Msgbox_BA = Msgbox_BA + "*""Please select atleast one saving/current/cash credit/overdraft/non resident bank account in which you prefer to get your refund""" & Chr(13)
      'Added by Shrutika(24/04/2025)NewDev
        Msgbox_BA = Msgbox_BA + "* "" Bank account with type “Others” cannot be selected for refund credit as its not eligible for refund.""" & Chr(13)
          ValidateIFSC = False
          Exit Function
    End If


'Newly added by Bindu

If ColCount1_BA_IFSC > 0 Then

For i = 1 To ColCount1_BA_IFSC
    
    BankName_BA(i) = rangecells1.item(i).Value
    BankAccntnum_BA(i) = rangecells2.item(i).Value
    BankIFSC_BA(i) = rangecells3.item(i).Value
    tempxml(i) = rangecells4.item(i).Value  'Uncommented by Ankita on 16/01/2025
    
    'INC_C43 2024-25 Bindu
    BankAccnType_BA(i) = rangecells5.item(i).Value
        
    
    
'    If BankName_BA(i) = "" Then
'        'Msgbox_BA = Msgbox_BA + "* Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
'        'AY_2024_25 Change "Malli
'        Msgbox_BA = Msgbox_BA + "* Please enter the Bank name at Sr.No " & i & " in Sheet Taxes Paid and Verification " & Chr(13)
'        ValidateIFSC = False
'        Exit Function
'    End If
    
    
    
    

    If Len(BankName_BA(i)) > 125 Then
        Msgbox_BA = Msgbox_BA + "* Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification cannot exceed 125 characters" & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If
    
    If isdropdownblank(BankIFSC_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* IFS Code of the Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory." & Chr(13)
        ValidateIFSC = False
        Exit Function

    End If
    

    'Added by Ankita on 21/01/2025
     
    If isdropdownblank(BankName_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* Please enter the Bank name at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateIFSC = False
        Exit Function

    End If
    
    If Len(BankIFSC_BA(i)) > 11 Then
        Msgbox_BA = Msgbox_BA + "* IFS Code at Sr.No " & i & " in Sheet Taxes Paid and Verification cannot exceed 11 characters" & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If
    
    If Not ValidateIFSCList(UCase(BankIFSC_BA(i))) Then
        Msgbox_BA = Msgbox_BA + "* Invalid IFS Code at Sr.No " & i & " in Sheet Taxes Paid and Verification.Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If
    
    'Added by ankita on 21/01/2025
        If BankName_BA(i) = "" Then
        'Msgbox_BA = Msgbox_BA + "* Bank Name at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        'AY_2024_25 Change "Malli
        Msgbox_BA = Msgbox_BA + "* Please enter the Bank name at Sr.No " & i & " in Sheet Taxes Paid and Verification " & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If
    
    
    UpdateProgressBar
    
    If Not chkCompulsory(BankAccntnum_BA(i)) Then
        'Msgbox_BA = Msgbox_BA + "* Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
       'AY_2024_25 Change "Malli
        Msgbox_BA = Msgbox_BA + "* Please enter the Account Number at Sr.No " & i & " in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If
    
    
    
    If Not ValidateBankAccountNumber_BA(BankAccntnum_BA(i), i) Then
        ValidateIFSC = False
        Exit Function
    End If
    
    'INC_C43 2024-25 Bindu
    If isdropdownblank(BankAccnType_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* Please select type of account at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory." & Chr(13)
        ValidateIFSC = False
        Exit Function
    End If
    
Next

Else
 
       Msgbox_BA = Msgbox_BA + "* Please Enter Bank Account details in Sheet Taxes Paid and Verification" & Chr(13)
       ValidateIFSC = False
       Exit Function

End If

End Function
'Function ValidateBankName_BA() As Boolean
'ValidateBankName_BA = True
'subProcCaption = "Validating BA"
'noOfProcessSub = ColCount1_BA_IFSC
'Dim rangecells As Range
'
'Set rangecells = Range("SchBA.BankName").Cells
'
'Dim i As Long
'ReDim BankName_BA(ColCount1_BA_IFSC)
'For i = 1 To ColCount1_BA_IFSC
'

'
'Next
'End Function
Function BankCode() As Boolean
BankCode = True
Dim rangecells As Range

Set rangecells = Range("BankNameWithCode").Cells
Dim i As Long
ReDim BankNameCode(ColCount1_BA_IFSC)
For i = 1 To ColCount1_BA_IFSC
    
    BankNameCode(i) = rangecells.item(i).Value
    
    If IsNumeric(BankName_BA(i)) Then
        BankCode = False
        Exit Function
    End If
    
    Next
End Function

Function ValidateNameofHolders_BA() As Boolean
ValidateNameofHolders_BA = True

Dim rangecells As Range
Dim i, UTNintCount As Long
Set rangecells = Range("SchBA.JointHoderName").Cells
UTNintCount = Range("SchBA.JointHoderName").count

ReDim BankJointHolder_BA(UTNintCount)
For i = 1 To UTNintCount
BankJointHolder_BA(i) = rangecells.item(i).Value
If Not rangecells.item(i).Value = "" Then
End If

If BankJointHolder_BA(i) <> "" Then
If Len(BankJointHolder_BA(i)) > 125 Then
        Msgbox_BA = Msgbox_BA + "* Name of the Joint Account Holder at Sr.No " & i & " in Sheet Taxes Paid and Verification cannot exceed 125 characters" & Chr(13)
        ValidateNameofHolders_BA = False
        Exit Function
    End If
End If
    Next
End Function

Function ValidateAccntStatus_BA() As Boolean
ValidateAccntStatus_BA = True
Dim rangecells As Range
Set rangecells = Range("SchBA.AcntStatus").Cells
Dim i As Long
ReDim BankAccountStatus_BA(ColCount1_BA_IFSC)

For i = 1 To ColCount1_BA_IFSC
    BankAccountStatus_BA(i) = rangecells.item(i).Value
      
    If BankAccountStatus_BA(i) = "(Select)" Then
        Msgbox_BA = Msgbox_BA + "* Account Status at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntStatus_BA = False
        Exit Function
    End If
    
    Next
End Function

Function AccountType() As Boolean
AccountType = True
Dim rangecells As Range
Set rangecells = Range("SchBA.AccountType").Cells
Dim i As Long
ReDim AccountType_BA(ColCount1_BA_IFSC)

For i = 1 To ColCount1_BA_IFSC
    AccountType_BA(i) = rangecells.item(i).Value
   
   
    If Not chkCompulsory(AccountType_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* Type of Account  at Sr.No " & i + 1 & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        AccountType = False
        Exit Function
    End If
          
    If (AccountType_BA(i)) = "(Select)" Then
        Msgbox_BA = Msgbox_BA + "* Type of Account  at Sr.No " & i + 1 & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        AccountType = False
        Exit Function
    End If
    
    Next
End Function

Function ValidateAccntNumber_BA() As Boolean
ValidateAccntNumber_BA = True
Dim rangecells As Range
Set rangecells = Range("SchBA.AcntNo").Cells
Dim i As Long
ReDim BankAccntnum_BA(ColCount1_BA_IFSC)
For i = 1 To ColCount1_BA_IFSC
    BankAccntnum_BA(i) = rangecells.item(i).Value
    
    
    
    Next
End Function
Function ValidateAccntBalance_BA() As Boolean
ValidateAccntBalance_BA = True
Dim rangecells As Range
Set rangecells = Range("SchBA.AccntBalance").Cells
Dim i As Long
ReDim BankAccntBalanc_BA(ColCount1_BA_IFSC)
For i = 1 To ColCount1_BA_IFSC
    BankAccntBalanc_BA(i) = rangecells.item(i).Value
      
    If Not chkCompulsory(BankAccntBalanc_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* Account Balance at Sr.No " & i & " in Sheet Taxes Paid and Verification is mandatory" & Chr(13)
        ValidateAccntBalance_BA = False
        Exit Function
    End If
    
    If Not IsNumeric(BankAccntBalanc_BA(i)) Then
        Msgbox_BA = Msgbox_BA + "* Account Balance at Sr.No " & i & " Should be a numeric value" & Chr(13)
        ValidateAccntBalance_BA = False
        Exit Function
    End If
    
    If ((BankAccntBalanc_BA(i) > 99999999999999#) Or (BankAccntBalanc_BA(i) < -99999999999999#)) Then
        Msgbox_BA = Msgbox_BA + "* Account Balance at Sr.No " & i & " cannot exceed 14 digits" & Chr(13)
        ValidateAccntBalance_BA = False
        Exit Function
    End If
    
    Next
End Function

Sub setTableInfo_BA_IFSC()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.IFSC").count
    Set rangecells = Range("SchBA.IFSC").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount1_BA_IFSC = ccount
End Sub

Sub setTableInfo_BA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.BankName").count
    Set rangecells = Range("SchBA.BankName").Cells
    For mIntCtr = 1 To mIntCells
            If Not ((rangecells.item(mIntCtr).Value = "") Or (rangecells.item(mIntCtr).Value = "(Select)")) Then
               ccount = ccount + 1
           End If
    Next
    ColCount1_BA = ccount
End Sub

Sub setTableInfo_BA3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.AcntNo").count
    Set rangecells = Range("SchBA.AcntNo").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
               ccount = ccount + 1
           End If
    Next
    ColCount4_BA = ccount
End Sub
Sub setTableInfo_BA5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long

    'INC_C43 2024-25 Bindu
    Dim ccount1 As Long
    Dim rangecells1 As Range
    Dim mIntCells1 As Long
    mIntCells1 = Range("SchBA.AcntType").count
    Set rangecells1 = Range("SchBA.AcntType").Cells
    ccount1 = 0
    ccount = 0
    mIntCells = Range("tempxml").count
    Set rangecells = Range("tempxml").Cells
    For mIntCtr = 1 To mIntCells
            If rangecells.item(mIntCtr).Value = True Then
               ccount = ccount + 1
'               If Trim(rangecells1.item(mIntCtr).Value) = Trim("Other") Then
'                     ccount1 = ccount1 + 1
'               End If
               'Changed on 20/03/2025 by Ankita
               'Changed on 02/04/2025 by Ankita
'               If Trim(rangecells1.item(mIntCtr).Value) <> Trim("Other") Then
                If Trim(rangecells1.item(mIntCtr).Value) = Trim("Other") Then
                     ccount1 = ccount1 + 1
               End If
           End If
    Next
    ColCount55_B5 = ccount
    ColCount55_BA6_Other = ccount1
End Sub
Sub setTableInfo_BA6() 'INC_C43 2024-25 Bindu
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
   
    ccount = 0
    
    mIntCells = Range("SchBA.AcntType").count
    Set rangecells = Range("SchBA.AcntType").Cells
    For mIntCtr = 1 To mIntCells
      If Not ((rangecells.item(mIntCtr).Value = "") Or (rangecells.item(mIntCtr).Value = "(Select)")) Then
        ccount = ccount + 1
        
      End If
    Next
    ColCount55_BA6 = ccount
    
   
End Sub
Sub setTableInfo_BA4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    Dim j As Long
     Dim b As Long
    
    ccount = 0
  
    mIntCells = Sheet3.Range("tempxml").count
    Set rangecells = Sheet3.Range("tempxml").Cells
    For mIntCtr = 1 To mIntCells
      'j = 1
      b = 0
         For j = mIntCtr To mIntCells
         
        If Not (Trim(rangecells.item(mIntCtr).Value) = False Or Trim(rangecells.item(mIntCtr).Value) = "") Or Not (Trim(rangecells.item(j).Value) = False Or Trim(rangecells.item(j).Value) = "") Then
            b = 1
            
          End If
        Next
        If b = 1 Then
        ccount = ccount + 1
        End If
    Next
    ColCount55_BA = ccount
    'rngname_AuditInfo2 = "Audit.Act;Audit.Act_Description;Audit.Sections;Audit.Date;"
End Sub
Function ValidateBankAccountNumber_BA(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_BA = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB() As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
           
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter_Bank(BankAccountNumber) Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_BA = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification " & Chr(13)
            ValidateBankAccountNumber_BA = False
            Exit Function
        End If
    
    End If
    
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0)) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification  " & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If

    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification " & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BA = False
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
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification." & Chr(13)
            ValidateBankAccountNumber_BA = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Bank Account Number at Sr.No " & cc & " in Sheet Taxes Paid and Verification is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_BA = False
        Exit Function
    End If
End Function


Function checkforNonMandatoryCol() As Boolean
checkforNonMandatoryCol = True
    Dim temp As Double
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("SchBA.JointHoderName").count
    Set rangecells = Range("SchBA.JointHoderName").Cells
    
    temp = ColCount1_BA_IFSC
    
    For mIntCtr = (temp + 1) To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                Msgbox_BA = Msgbox_BA + "* Please Enter all the Mandatory fiels at Sr.No " & mIntCtr & " in Sheet Taxes Paid and Verification" & Chr(13)
                checkforNonMandatoryCol = False
                Exit Function
           End If
    Next
End Function

Function ValidateMismatchSchBA()
On Error Resume Next
ValidateMismatchSchBA = True
Msgbox_BA = ""
If (ColCount1_BA_IFSC + 1) <> (Sheet3.Range("NumBankAccountsHeld").Value) Then
    ValidateMismatchSchBA = False
    Msgbox_BA = Msgbox_BA + "* Sheet Taxes Paid and Verification No. of bank accounts held during the P.Y. should be equal to the number of rows filled in Bank Details."
    Sheet3.Activate
    Sheet3.Range("NumBankAccountsHeld").Activate
End If
    
End Function

Function ValidateBankAccountNumber_BARow1(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_BARow1 = True
    Dim numfound As Boolean
    Dim countnum As Long
    Dim myB() As Variant
    Dim i As Long
    Dim zeroCount As Long
    Dim BeforeZero, AfterZero As String
    errmsgVerification = ""
    numfound = False
    countnum = 0
    BeforeZero = ""
    AfterZero = ""
    zeroCount = 1
    
    If Len(BankAccountNumber) > 0 Then
        If Not checkfieldspecialcharacter_Bank(BankAccountNumber) Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet :Taxes Paid and Verification, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_BARow1 = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification " & Chr(13)
            ValidateBankAccountNumber_BARow1 = False
            Exit Function
        End If
    End If
    
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is mandatory in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0)) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification  " & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If
    
    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification " & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
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
            Msgbox_BA = Msgbox_BA & "* Bank Account Number at Sr.No " & cc & " is invalid in Sheet Taxes Paid and Verification." & Chr(13)
            ValidateBankAccountNumber_BARow1 = False
            Exit Function
        End If
    End If

    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Bank Account Number at Sr.No " & cc & " in Sheet Taxes Paid and Verification is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_BARow1 = False
        Exit Function
    End If
End Function



Function ValidateBankName_BA() As Boolean
ValidateBankName_BA = True
Dim rangecells As Range
Set rangecells = Range("SchBA.AcntType").Cells
Dim i As Long
ReDim BankAccnType_BA(ColCount1_BA_IFSC)
For i = 1 To ColCount1_BA_IFSC
    BankAccnType_BA(i) = rangecells.item(i).Value
    

    Next
End Function
