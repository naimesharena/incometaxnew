Attribute VB_Name = "Sch80GGC"
Option Explicit

Public MsgBox_80GGC As String

Public DateofDonation_80GGC As Variant
Public NatureofTransaction_80GGC As Variant
Public Donation_cash_80GGC As Variant
Public Donation_other_80GGC As Variant
Public Donation_total_80GGC As Variant
Public Donation_Eligible_80GGC As Variant
Public ChequeNumber_80GGC As Variant
Public BankAccntnum_80GGC, BankIFSC_80GGC, BankName_80GGC As Variant
Public Total_DonationInCash_80GGC, Total_DonationInOThMode_80GGC, Total_TotalDonation_80GGC As Variant

Public rngname_80GGC As Variant
Public Total_Donation_Eligible_80GGC As Variant

Public end80GGC, end80GGC1, end80GGC2, end80GGC3, end80GGC4, end80GGC5, end80GGC6, end80GGC7, end80GGC8 As Long
Sub ValidateSheet80GGC_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2026-27"                  'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
 Validate80GGC
 'MsgBox "Sheet 80GGC is OK", vbOKOnly, vbMessgaeCaption
 fmsgboxoK "Sheet 80GGC is OK"
End Sub
Sub Validate80GGC()
    If Not Validate80GGC_1 Then
        Sheet13.Activate
        fmsgbox (MsgBox_80GGC)
        CloseMsg
    End If
End Sub

Sub setTableInfo80GGC() '80GGC_C1
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("DateofDonation_80GGC").Cells
    mIntCells = Sheet13.Range("DateofDonation_80GGC").count
    For mIntCtr = 1 To mIntCells
         If Not rangecells.item(mIntCtr).Value = "" Then
          ccount = ccount + 1
        End If
    Next
    end80GGC = ccount
    'rngname_80GGC = "DateofDonation_80GGC;Donationincash_80GGC;Donationinothermode_80GGC;TotalDonation_80GGC;EligibleAmountofDonation_80GGC;NatureofTransaction_80GGC;Chequeno_80GGC;IFSC_80GGC;"
     rngname_80GGC = "DateofDonation_80GGC;Donationincash_80GGC;Donationinothermode_80GGC;TotalDonation_80GGC;EligibleAmountofDonation_80GGC;Chequeno_80GGC;IFSC_80GGC;"


End Sub
Sub setTableInfo80GGC1() '80GGC_C2
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("Donationincash_80GGC").Cells
    mIntCells = Sheet13.Range("Donationincash_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC1 = ccount
End Sub

Sub setTableInfo80GGC2() '80GGC_C3
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("Donationinothermode_80GGC").Cells
    mIntCells = Sheet13.Range("Donationinothermode_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC2 = ccount
End Sub

'Sub setTableInfo80GGC3() '80GGC_C4
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet13.Range("NatureofTransaction_80GGC").Cells
'    mIntCells = Sheet13.Range("NatureofTransaction_80GGC").count
'    For mIntCtr = 1 To mIntCells
'        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80GGC3 = ccount
'End Sub

Sub setTableInfo80GGC4() '80GGC_C5
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("Chequeno_80GGC").Cells
    mIntCells = Sheet13.Range("Chequeno_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        ccount = ccount + 1
        End If
    Next
    end80GGC4 = ccount
End Sub

Sub setTableInfo80GGC5() '80GGC_C6
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet13.Range("IFSC_80GGC").Cells
    mIntCells = Sheet13.Range("IFSC_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC5 = ccount
End Sub


'Ankita_16/03/2026====================
Sub setTableInfo80GGC6()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet11.Range("Name_80GGC").Cells
    mIntCells = Sheet11.Range("Name_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC6 = ccount
End Sub

'Ankita_16/03/2026====================
Sub setTableInfo80GGC7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet11.Range("PAN_80GGC").Cells
    mIntCells = Sheet11.Range("PAN_80GGC").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGC7 = ccount
End Sub


'Sub setTableInfo80GGC6() '80GGC_C7
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet13.Range("NameofDonor_80GGC").Cells
'    mIntCells = Sheet13.Range("NameofDonor_80GGC").count
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80GGC6 = ccount
'End Sub

'Sub setTableInfo80GGC7() '80GGC_C8
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet13.Range("AccountofDonor_80GGC").Cells
'    mIntCells = Sheet13.Range("AccountofDonor_80GGC").count
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80GGC7 = ccount
'End Sub


'Ankita_16/03/2026=======
Sub AddRows80GGC() '80CCG
    Dim vRows  As Long
    Sheets("80GGC").Activate
    'EfilingCommon.DefinedgridNameRange = "||DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||NatureofTransaction_80GGC||Chequeno_80GGC||IFSC_80GGC||"
'    EfilingCommon.DefinedgridNameRange = "||DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||Chequeno_80GGC||IFSC_80GGC||"  'Malli
     EfilingCommon.DefinedgridNameRange = "||DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||Name_80GGC||PAN_80GGC||Chequeno_80GGC||IFSC_80GGC||"  'Malli
     ActiveCellRange = EfilingCommon.searchLastRow("DateofDonation_80GGC")
    'vRows = EfilingCommon.insertRowUnderSectionWithFormula
    EfilingCommon.insertRowUnderSectionWithFormula80GGC HOIflag:=1
End Sub
Function CheckIFSC_80GGC(Tval As Variant, Tadd As Variant) '80GGC_C5
On Error Resume Next
Application.EnableEvents = False
 sPassword = EfilingCommon.getmsgstate
   Sheet13.Unprotect Password:=sPassword

If Tval <> "" Then
     If Not CheckIFSC(UCase(Tval)) Then
        fmsgbox "Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details in Schedule 80GGC Sheet."
           Sheet13.Range(Tadd) = ""
           'Sheet13.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
    
    If Not ValidateIFSCList(UCase(Tval)) Then
           fmsgbox "Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets in Bank Details in Schedule 80GGC Sheet."
           Sheet13.Range(Tadd) = ""
           'Sheet13.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
     
     Dim uiffsc As Variant
    uiffsc = UCase(Sheet13.Range(Tadd).Value)
    Sheet13.Range(Tadd).Value = uiffsc
    'Sheet13.Range(Tadd).Select
   
    'Sheet13.Range(Tadd).Offset(0, 1) = UCase(Sheet13.Range(Tadd).Offset(0, 8).Value)
End If

'If Tval = "" Then
'    Sheet13.Range(Tadd).Offset(0, 1) = ""
'End If

endfd:
Sheet13.Protect Password:=sPassword
Application.EnableEvents = True

End Function

Function Validate80GGC_1() As Boolean
    Validate80GGC_1 = True
    
    MsgBox_80GGC = "Schedule 80GGC : " & Chr(10)
        setTableInfo80GGC
        setTableInfo80GGC1
        setTableInfo80GGC2
        'setTableInfo80GGC3
        setTableInfo80GGC4
        setTableInfo80GGC5
        'setTableInfo80GGC6
        'setTableInfo80GGC7
        
        
    ' end80GGC = WorksheetFunction.Max(0, end80GGC, end80GGC1, end80GGC2, end80GGC3, end80GGC4, end80GGC5)
        end80GGC = WorksheetFunction.Max(0, end80GGC, end80GGC1, end80GGC2, end80GGC4, end80GGC5)

     
        If end80GGC > 0 Then
            If Not ValidateDateofDonation_80GGC Then Validate80GGC_1 = False ' DATE
            If Not ValidateMandatorySh80GGC Then Validate80GGC_1 = False  'Ankita_29/05/2025
            If Not ValidateDonationAmt_80GGC Then Validate80GGC_1 = False 'DONATION IN CASH/OTHER MODE
            'If Not ValidateNatureofTransaction_80GGC Then Validate80GGC_1 = False 'NATURE OF TRANSACTION
            If Not ValidateChequeNumber_80GGC Then Validate80GGC_1 = False 'TRANSACTION REF & NATURE OF TRANSACTION
            'If Not ValidateBank_IFSC_Account Then Validate80GGC_1 = False 'IFSC
            If Not ValidateDonation_total_80GGC Then Validate80GGC_1 = False ' TOTAL DONATION
            If Not ValidateDonation_Eligible_80GGC Then Validate80GGC_1 = False 'ELIGIBLE DONATION
            If Not ValidateTotal_Donation_InCash_80GGC Then Validate80GGC_1 = False 'TOTAL DONATION IN CASH
            If Not ValidateTotal_Donation_OtherMode_80GGC Then Validate80GGC_1 = False 'TOTAL DONATION IN OTHER MODE
            If Not ValidateTotalof_Total_Donation_80GGC Then Validate80GGC_1 = False 'TOTAL - TOATL DONATION
            If Not ValidatePoliticalParty_Name_80GGC Then Validate80GGC_1 = False 'TOTAL - TOATL DONATION   'Ankita_16/03/2026======
            If Not ValidatePoliticalParty_PAN_80GGC Then Validate80GGC_1 = False 'TOTAL - TOATL DONATION    'Ankita_16/03/2026======
        End If
        
End Function

'Ankita_29/05/2025
Function ValidateMandatorySh80GGC() As Boolean
ValidateMandatorySh80GGC = True
Dim i As Long
Dim flag As Boolean
Dim rangecells As Range
flag = True

For i = 1 To Sheet13.Range("DateofDonation_80GGC").Rows.count
    If Sheet13.Range("DateofDonation_80GGC").item(i).Value <> "" Or Sheet13.Range("Donationincash_80GGC").item(i).Value <> "" Or _
        Sheet13.Range("Donationinothermode_80GGC").item(i).Value <> "" Or Sheet13.Range("Chequeno_80GGC").item(i).Value <> "" _
        Or Sheet13.Range("IFSC_80GGC").item(i).Value <> "" Then
        If Sheet13.Range("DateofDonation_80GGC").item(i).Value = "" Then
            flag = False

'Ankita_29/05/2025

                MsgBox_80GGC = MsgBox_80GGC + "* please select "" date of Contribution "" " & Chr(13)
'                 MsgBox_80GGC = MsgBox_80GGC + "* please select ""Date"" at Sr. No." & i & " in Schedule 80GGC "" " & Chr(13)
 
        End If
        If flag = False Then
            ValidateMandatorySh80GGC = False
            Exit Function
        End If
    End If
Next
End Function
 


Function ValidateDateofDonation_80GGC() As Boolean
    ValidateDateofDonation_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet13.Range("DateofDonation_80GGC").Cells
    Dim i As Long
    ReDim DateofDonation_80GGC(end80GGC)
    
    subProcCaption = "Validating 80GGC"
    noOfProcessSub = end80GGC
    
    For i = 1 To end80GGC
        DateofDonation_80GGC(i) = rangecells.item(i).Value
        
        'Ankita_29/05/2025
'        If Not chkCompulsory(DateofDonation_80GGC(i)) Then
'            MsgBox_80GGC = MsgBox_80GGC + "* please select ""Date"" at Sr. No." & i & " in Schedule 80GGC" & Chr(13)
'            ValidateDateofDonation_80GGC = False
'            Exit Function
'        End If
'        UpdateProgressBar
    Next
End Function
Function ValidateTotal_Donation_InCash_80GGC() As Boolean
 ValidateTotal_Donation_InCash_80GGC = True
 Total_DonationInCash_80GGC = Sheet13.Range("Total_DonationInCash_80GGC").Value
 
 
 If Len(Total_DonationInCash_80GGC) > 14 Then
 'Contribution
    'MsgBox_80GGC = MsgBox_80GGC + "* Total amount of Donation in Cash cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    MsgBox_80GGC = MsgBox_80GGC + "* Total amount of Contribution in Cash cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    ValidateTotal_Donation_InCash_80GGC = False
    Exit Function
End If


End Function
Function ValidateDonationAmt_80GGC() As Boolean
ValidateDonationAmt_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet13.Range("Donationincash_80GGC").Cells
    Set rangecells1 = Sheet13.Range("Donationinothermode_80GGC").Cells
    Dim i As Long
    ReDim Donation_cash_80GGC(end80GGC)
    ReDim Donation_other_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        Donation_cash_80GGC(i) = rangecells.item(i).Value
        Donation_other_80GGC(i) = rangecells1.item(i).Value
        
       If rangecells.item(i).Locked = False Then
            If Not chkCompulsory(Donation_cash_80GGC(i)) Then
                'Enter the amount of Contribution either in field "Contribution in cash" or "Contribution in other mode in schedule 80GGC"
                'MsgBox_80GGC = MsgBox_80GGC + "* Enter the amount of donation either in field Donation in cash or Donation in other mode at Sr. No " & i & " is Mandatory in Schedule 80GGC." & Chr(13)
                'MsgBox_80GGC = MsgBox_80GGC + "* Enter the amount of Contribution either in field ""Contribution in cash"" or ""Contribution in other mode"" at Sr. No " & i & " is Mandatory in Schedule 80GGC." & Chr(13)
                 MsgBox_80GGC = MsgBox_80GGC + "* Enter the amount of Contribution either in field ""Contribution in cash"" or ""Contribution in other mode"" at Sr. No " & i & " in Schedule 80GGC." & Chr(13)
                
                ValidateDonationAmt_80GGC = False
                Exit Function
            End If
       End If
       
       If rangecells1.item(i).Locked = False Then
            If Not chkCompulsory(Donation_other_80GGC(i)) Then
               'SIT-66176 'Malli
               'MsgBox_80GGC = MsgBox_80GGC + "* Enter the amount of donation either in field Donation in cash or Donation in other mode at Sr. No " & i & " is Mandatory in Schedule 80GGC." & Chr(13)
                'MsgBox_80GGC = MsgBox_80GGC + "* Enter the amount of donation either in field ""Donation in cash"" or ""Donation in other mode"" at Sr. No " & i & " is Mandatory in Schedule 80GGC." & Chr(13)
                'MsgBox_80GGC = MsgBox_80GGC + "*Enter the amount of Contribution either in field ""Contribution in cash"" or ""Contribution in other mode"" at Sr. No " & i & " is Mandatory in Schedule 80GGC." & Chr(13)
                MsgBox_80GGC = MsgBox_80GGC + "*Enter the amount of Contribution either in field ""Contribution in cash"" or ""Contribution in other mode"" at Sr. No " & i & " in Schedule 80GGC." & Chr(13)
                ValidateDonationAmt_80GGC = False
                Exit Function
            End If
       End If
       
       
       
       
    Next
End Function
'Function ValidateNatureofTransaction_80GGC() As Boolean
'    ValidateNatureofTransaction_80GGC = True
''    setTableInfo80GGC
'    Dim rangecells As Range
'    Set rangecells = Sheet13.Range("NatureofTransaction_80GGC").Cells
'    Dim i As Long
'    ReDim NatureofTransaction_80GGC(end80GGC)
'
'    For i = 1 To end80GGC
'        NatureofTransaction_80GGC(i) = rangecells.item(i).Value
'
'        If rangecells.item(i).Locked = False Then
'
'            If (NatureofTransaction_80GGC(i) = "(Select)") Or (NatureofTransaction_80GGC(i) = "") Then
'                MsgBox_80GGC = MsgBox_80GGC + "* please enter transaction reference number of donation transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13)
'                ValidateNatureofTransaction_80GGC = False
'                Exit Function
'            End If
'        End If
'        UpdateProgressBar
'    Next
'End Function
Function ValidateChequeNumber_80GGC() As Boolean
    ValidateChequeNumber_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Dim rangecells1 As Range
    Dim rangecells2 As Range
    
    Set rangecells2 = Sheet13.Range("Donationinothermode_80GGC").Cells
    ReDim Donation_other_80GGC(end80GGC)
    
    Set rangecells = Sheet13.Range("Chequeno_80GGC").Cells
    Dim i As Long
    ReDim ChequeNumber_80GGC(end80GGC)
    
'    Set rangecells1 = Sheet13.Range("NatureofTransaction_80GGC").Cells Malli
'    ReDim NatureofTransaction_80GGC(end80GGC)
    
    
    Dim rangecells3 As Range
    Set rangecells3 = Sheet13.Range("IFSC_80GGC").Cells
    ReDim BankIFSC_80GGC(end80GGC)
    
    
    For i = 1 To end80GGC
        ChequeNumber_80GGC(i) = rangecells.item(i).Value
        'NatureofTransaction_80GGC(i) = rangecells1.item(i).Value
        BankIFSC_80GGC(i) = rangecells3.item(i).Value
        Donation_other_80GGC(i) = rangecells2.item(i).Value
        
        
        
    If rangecells2.item(i).Locked = False Then
           
          'Malli comented
'         If Donation_other_80GGC(i) > 0 Then
'           If (NatureofTransaction_80GGC(i) = "(Select)") Or (NatureofTransaction_80GGC(i) = "") Then
'                MsgBox_80GGC = MsgBox_80GGC + "* please select Nature of  transaction of donation transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13)
'                ValidateChequeNumber_80GGC = False
'                Exit Function
'             End If
'
'
'          End If
        
          
        If Donation_other_80GGC(i) > 0 Then  'Newly added by Bindu
            If Not chkCompulsory(ChequeNumber_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13)
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
          'Malli comented
'          If Len(ChequeNumber_80GGC(i)) > 0 Or Len(BankIFSC_80GGC(i)) > 1 Then
'             If (NatureofTransaction_80GGC(i) = "(Select)") Or (NatureofTransaction_80GGC(i) = "") Then
'                MsgBox_80GGC = MsgBox_80GGC + "* please select Nature of  transaction of donation transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13)
'                ValidateChequeNumber_80GGC = False
'                Exit Function
'             End If
'          End If
    'Malli comented
             
            
         
             
            If isdropdownblank(BankIFSC_80GGC(i)) Then
                'MsgBox_80GGC = MsgBox_80GGC + "* please enter your bank IFSC from which Donation is made at Sr.No " & i & " in Schedule 80GGC." & Chr(13)
                MsgBox_80GGC = MsgBox_80GGC + "* please enter ""your bank IFSC from which Contribution is made"" at Sr.No " & i & " in Schedule 80GGC." & Chr(13)
                
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
            
            
            If Len(ChequeNumber_80GGC(i)) > 50 Then
                MsgBox_80GGC = MsgBox_80GGC + "* ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr.No " & i & " in Schedule 80GGC cannot exceed 50 characters." & Chr(13)
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
            
            
            'Newly added by Bindu
            If Not checkfieldspecialcharacter_Trans(ChequeNumber_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr.No " & i & " in Schedule 80GGC is invalid, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
            
            
            If Len(BankIFSC_80GGC(i)) > 11 Then
                MsgBox_80GGC = MsgBox_80GGC + "* IFS Code at Sr.No " & i & " in Schedule 80GGC cannot exceed 11 characters" & Chr(13)
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
            
            If Not ValidateIFSCList(UCase(BankIFSC_80GGC(i))) Then
                MsgBox_80GGC = MsgBox_80GGC + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80GGC.Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
                ValidateChequeNumber_80GGC = False
                Exit Function
            End If
           End If
           
     End If
     Sheet13.Activate  'Ankita_29/05/2025_while validating going in other sheet.
     
        UpdateProgressBar
    Next
End Function

Function ValidateBank_IFSC_Account() As Boolean
ValidateBank_IFSC_Account = True
'setTableInfo80GGC
'Dim rangecells As Range
'Dim rangecells1 As Range
'Dim rangecells2 As Range
Dim rangecells3 As Range


'Set rangecells1 = Range("AccountofDonor_80GGC").Cells
'Set rangecells2 = Range("NameofDonor_80GGC").Cells
Set rangecells3 = Range("IFSC_80GGC").Cells


Dim i As Long
ReDim BankIFSC_80GGC(end80GGC)
'ReDim BankName_80GGC(end80GGC)
'ReDim BankAccntnum_80GGC(end80GGC)



For i = 1 To end80GGC
    
    'BankName_80GGC(i) = rangecells2.item(i).Value
    'BankAccntnum_80GGC(i) = rangecells1.item(i).Value
    BankIFSC_80GGC(i) = rangecells3.item(i).Value
    
    
 If rangecells3.item(i).Locked = False Then
 
'    If BankName_80GGC(i) = "" Then
'        MsgBox_80GGC = MsgBox_80GGC + "* please enter your bank name from which Donation is made at Sr.No " & i & " in Schedule 80GGC" & Chr(13)
'        ValidateBank_IFSC_Account = False
'        'Exit Function
'    End If
'
'    If Len(BankName_80GGC(i)) > 125 Then
'        MsgBox_80GGC = MsgBox_80GGC + "* Bank Name at Sr.No " & i & " in Schedule 80GGC cannot exceed 125 characters" & Chr(13)
'        ValidateBank_IFSC_Account = False
'        'Exit Function
'    End If
'
    If isdropdownblank(BankIFSC_80GGC(i)) Then
        MsgBox_80GGC = MsgBox_80GGC + "* please enter ""your bank IFSC from which Contribution is made"" at Sr.No " & i & " in Schedule 80GGC." & Chr(13)
        ValidateBank_IFSC_Account = False
        'Exit Function

    End If
    

    If Len(BankIFSC_80GGC(i)) > 11 Then
        MsgBox_80GGC = MsgBox_80GGC + "* IFS Code at Sr.No " & i & " in Schedule 80GGC cannot exceed 11 characters" & Chr(13)
        ValidateBank_IFSC_Account = False
        'Exit Function
    End If
    
    If Not ValidateIFSCList(UCase(BankIFSC_80GGC(i))) Then
        MsgBox_80GGC = MsgBox_80GGC + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80GGC.Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
        ValidateBank_IFSC_Account = False
        'Exit Function
    End If
    
    UpdateProgressBar
    
'    If Not chkCompulsory(BankAccntnum_80GGC(i)) Then
'        MsgBox_80GGC = MsgBox_80GGC + "* please enter your bank account number from which Donation is made at Sr.No " & i & " in Schedule 80GGC" & Chr(13)
'        ValidateBank_IFSC_Account = False
'        'Exit Function
'    End If
    
    
    
'    If Not ValidateBankAccountNumber_80GGC(BankAccntnum_80GGC(i), i) Then
'        ValidateBank_IFSC_Account = False
'       ' Exit Function
'    End If

  End If

Next
End Function

Function ValidateBankAccountNumber_80GGC(BankAccountNumber As Variant, cc As Long) As Boolean
    ValidateBankAccountNumber_80GGC = True
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
            MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
            ValidateBankAccountNumber_80GGC = False
            Exit Function
        End If

        If BankAccountNumber = 0 Then
            MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC " & Chr(13)
            ValidateBankAccountNumber_80GGC = False
            Exit Function
        End If
    
    End If
    
    If (Len(BankAccountNumber) < 1) Or (Len(BankAccountNumber) > 20) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is mandatory in 80GGC" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If Trim(BankAccountNumber) = "" Or IsEmpty(BankAccountNumber) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is mandatory in 80GGC" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If ((InStr(BankAccountNumber, " ") > 0) Or (InStr(BankAccountNumber, "//") > 0) Or (InStr(BankAccountNumber, "--") > 0)) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC  " & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If (Mid(BankAccountNumber, 1, 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-") Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If

    If Mid(BankAccountNumber, Len(BankAccountNumber), 1) = "/" Or Mid(BankAccountNumber, 1, 1) = "-" Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC " & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If
    
    If Not IsNumeric(Mid(BankAccountNumber, Len(BankAccountNumber), 1)) Then
        MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
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
            MsgBox_80GGC = MsgBox_80GGC & "* Bank Account Number at Sr.No " & cc & " is invalid in 80GGC." & Chr(13)
            ValidateBankAccountNumber_80GGC = False
            Exit Function
        End If
    End If


    If countnum < 1 Then
        errmsgVerification = errmsgVerification & "* Bank Account Number at Sr.No " & cc & " in 80GGC is invalid. Account Number should have a minimum of 1 numeric digit (1-9)" & Chr(13)
        ValidateBankAccountNumber_80GGC = False
        Exit Function
    End If
End Function
Function ValidateDonation_total_80GGC() As Boolean
ValidateDonation_total_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet13.Range("TotalDonation_80GGC").Cells
    Dim i As Long
    ReDim Donation_total_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        Donation_total_80GGC(i) = rangecells.item(i).Value
        If Len(Donation_total_80GGC(i)) > 14 Then
            MsgBox_80GGC = MsgBox_80GGC + "* Total Contribution Amount at Sr. No " & i & " cannot exceed 14 digits in Schedule 80GGC." & Chr(13)
            ValidateDonation_total_80GGC = False
            Exit Function
        End If
    Next
End Function

Function ValidateDonation_Eligible_80GGC() As Boolean
ValidateDonation_Eligible_80GGC = True
'    setTableInfo80GGC
    Dim rangecells As Range
    Set rangecells = Sheet13.Range("EligibleAmountofDonation_80GGC").Cells
    Dim i As Long
    ReDim Donation_Eligible_80GGC(end80GGC)
    
    For i = 1 To end80GGC
        Donation_Eligible_80GGC(i) = rangecells.item(i).Value
        If Len(Donation_Eligible_80GGC(i)) > 14 Then
            'MsgBox_80GGC = MsgBox_80GGC + "* Eligible donation Amount at Sr. No " & i & " cannot exceed 14 digits in Schedule 80GGC." & Chr(13)
            MsgBox_80GGC = MsgBox_80GGC + "* Eligible contribution Amount at Sr. No " & i & " cannot exceed 14 digits in Schedule 80GGC." & Chr(13)
            ValidateDonation_Eligible_80GGC = False
            Exit Function
        End If
    Next
End Function
Function ValidateTotal_Donation_OtherMode_80GGC() As Boolean
 ValidateTotal_Donation_OtherMode_80GGC = True
 Total_DonationInOThMode_80GGC = Sheet13.Range("Total_DonationInOtherMode_80GGC").Value
 
 If Len(Total_DonationInOThMode_80GGC) > 14 Then
    'MsgBox_80GGC = MsgBox_80GGC + "* Total amount of Donation in Other Mode cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    MsgBox_80GGC = MsgBox_80GGC + "* Total amount of Contribution in Other Mode cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    ValidateTotal_Donation_OtherMode_80GGC = False
    Exit Function
End If

End Function
Function ValidateTotalof_Total_Donation_80GGC() As Boolean
 ValidateTotalof_Total_Donation_80GGC = True
 Total_TotalDonation_80GGC = Sheet13.Range("Total_Donation_80GGC").Value
 
 If Len(Total_TotalDonation_80GGC) > 14 Then
    'MsgBox_80GGC = MsgBox_80GGC + "* Total amount Total Donation cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    MsgBox_80GGC = MsgBox_80GGC + "* Total amount Total Contribution cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    ValidateTotalof_Total_Donation_80GGC = False
    Exit Function
End If

End Function
Function ValidateTotal_Donation_Eligible_80GGC() As Boolean
 ValidateTotal_Donation_Eligible_80GGC = True
 Total_Donation_Eligible_80GGC = Sheet13.Range("Total_Donation_Eligible_80GGC").Value
 
 If Len(Total_Donation_Eligible_80GGC) > 14 Then
    'MsgBox_80GGC = MsgBox_80GGC + "* Total Eligible donation amount cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    MsgBox_80GGC = MsgBox_80GGC + "* Total Eligible Contribution amount cannot be greater than 14 digits in Schedule 80GGC." & Chr(13)
    ValidateTotal_Donation_Eligible_80GGC = False
    Exit Function
End If

End Function

Sub Prev80GGC_Click()
Sheet12.Activate '80GGC SHEET
End Sub

'Malli
Sub eventsen()
Application.EnableEvents = True

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
DefinedgridNameRange = "DateofDonation_80GGC||Donationincash_80GGC||Donationinothermode_80GGC||TotalDonation_80GGC||EligibleAmountofDonation_80GGC||Chequeno_80GGC||IFSC_80GGC||"
End Sub
End Sub
Sub Next_80GGCClick()
Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("80U-80DD")
    sourceSheet.Activate
End Sub

'Ankita_16/03/2026===========
Function ValidatePoliticalParty_Name_80GGC() As Boolean
ValidatePoliticalParty_Name_80GGC = True

    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet13.Range("Name_80GGC").Cells
    
    Dim i As Long
    ReDim Political_Name_80GGC(end80GGC)
    For i = 1 To end80GGC
        Political_Name_80GGC(i) = rangecells.item(i).Value
        If rangecells.item(i).Locked = False Then
'            If (Donation_other_80GGC(i) > "0") Then
            If (Donation_other_80GGC(i) > 0 Or Donation_cash_80GGC(i) > 0) Then
            If Not chkCompulsory(Political_Name_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* ""Please enter name of Political Party in schedule 80GGC"" at Sr. No " & i & " " & Chr(13)
                ValidatePoliticalParty_Name_80GGC = False
                Exit Function
            End If
            
            If Not CheckSpecialCharacter4New(Political_Name_80GGC(i)) Then
                MsgBox_80GGC = MsgBox_80GGC + "* Name of political party characters < >  are not allowed in schedule 80GGC" & Chr(13)
                ValidatePoliticalParty_Name_80GGC = False
                Exit Function
            End If
          End If
            
    End If
    Next
End Function

Function CheckSpecialCharacter4New(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArraynew As Variant
    Dim iCharCountNew, iSpecialCharNew As Long
    
    CheckSpecialCharacter4New = True
    specialCharArraynew = Array(">", "<")
    For iCharCountNew = 1 To Len(emailid)
        For iSpecialCharNew = 0 To UBound(specialCharArraynew)
        If Mid(emailid, iCharCountNew, 1) = specialCharArraynew(iSpecialCharNew) Then
            CheckSpecialCharacter4New = False
            Exit Function
        End If
        Next
    Next
End Function

'Ankita_16/03/2026===========
Function ValidatePoliticalParty_PAN_80GGC() As Boolean
    ValidatePoliticalParty_PAN_80GGC = True

    Dim rangecells As Range
    Dim rangecells1 As Range
    
    Set rangecells = Sheet13.Range("PAN_80GGC").Cells
    Dim i As Long
    ReDim Political_PAN_80GGC(end80GGC)
    For i = 1 To end80GGC
        Political_PAN_80GGC(i) = rangecells.item(i).Value
        If rangecells.item(i).Locked = False Then
        
'           If (Donation_other_80GGC(i) > "0") Then
            If (Donation_other_80GGC(i) > 0 Or Donation_cash_80GGC(i) > 0) Then
            
        If Not chkCompulsory(Political_PAN_80GGC(i)) Then
             MsgBox_80GGC = MsgBox_80GGC + "* ""Please enter PAN of Political Party in schedule 80GGC"" at Sr. No " & i & "" & Chr(13)
             ValidatePoliticalParty_PAN_80GGC = False
             Exit Function
        End If
            
        If ((UCase(Political_PAN_80GGC(i)) = UCase(Sheet1.Range("sheet1.PAN").Value)) Or (UCase(Political_PAN_80GGC(i)) = UCase(Sheet3.Range("Ver.PAN").Value))) Then
              MsgBox_80GGC = MsgBox_80GGC + "* ""Political party PAN cannot be same as assesse PAN or verification PAN in schedule 80GGC "" at Sr. No " & i & "" & Chr(13)
              ValidatePoliticalParty_PAN_80GGC = False
                Exit Function
         End If
  
        If Not checkfieldspecialcharacter(Political_PAN_80GGC(i)) Then
             MsgBox_80GGC = MsgBox_80GGC + "* PAN of Political Party at Sr. No  " & i & " cannot Contain Special Characters in Schedule 80GGC." & Chr(13)
             ValidatePoliticalParty_PAN_80GGC = False
             Exit Function
        End If
        
        If Not CheckPAN1(UCase(Political_PAN_80GGC(i))) Then
             MsgBox_80GGC = MsgBox_80GGC + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
             ValidatePoliticalParty_PAN_80GGC = False
             Exit Function
        End If
            
          End If
        End If
        UpdateProgressBar
    Next
End Function

