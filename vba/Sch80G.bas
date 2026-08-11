Attribute VB_Name = "Sch80G"
Option Explicit

Public MsgBox_80GA, MsgBox_80GB, MsgBox_80GC, MsgBox_80GD, MsgBox_80GE, MsgBox_80G As String

Dim errmsg80G As Variant

Public Name1_80GA, Addr1_80GA, City1_80GA, State1_80GA, PinCd1_80GA, Pan1_80GA, DonationAmt1_80GA, DonationAmt1_80GA1, TotDon100Percent_80GA, DonationAmtTotal_80GA, DonationAmtTotOfTotal_80GA, DonationAmtTotOfOtherMode_80GA As Variant
Public Name2_80GB, Addr280GB, City280GB, State280GB, PinCode_80GB, Pan_80GB, DonationAmt_80GB, DonationAmt_80GB1, TotDon100Percent_80GB, DonationAmtTotal_80GB, DonationAmtTotOfTotal_80GB, DonationAmtTotOfOtherMode_80GB As Variant
Public Name_80GC, Addr_80GC, City_80GC, State_80GC, PinCode_80GC, Pan_80GC, DonationAmt_80GC, DonationAmt_80GC1, TotDon100Percent_80GC, DonationAmtTotal_80GC, DonationAmtTotOfTotal_80GC, DonationAmtTotOfOtherMode_80GC As Variant
Public Name_80GD, Addr_80GD, City_80GD, State_80GD, PinCode_80GD, Pan_80GD, ArnNumber_80GD, DonationAmt_80GD, DonationAmt_80GD1, TotDon100Percent_80GD, TotDon100Percent_80G, DonationAmtTotal_80GD, DonationAmtTotOfTotal_80GD, DonationAmtTotOfOtherMode_80GD, DonationTransaction_80GA As Variant
Public TotalDonationsUs80G_E, TotalDonationsUs80G_E2, TotalDonationsUs80G_E3 As Variant
Public Transaction1_80GA, Transaction1_80GB, Transaction1_80GC, Transaction1_80GD, Transaction1_80G As Variant
Public IFSC_80GA, IFSC_80GB, IFSC_80GC, IFSC_80GD As Variant

'Column count for 80G_A
'Ankita_06/03/2026
Public ColCountA, ColCountA_1, ColCountA_2, ColCountA_3, ColCountA_4, ColCountA_5, ColCountA_6, ColCountA_7, ColCountA_8 As Long
Public ColCountB, ColCountB_1, ColCountB_2, ColCountB_3, ColCountB_4, ColCountB_5, ColCountB_6, ColCountB_7, ColCountB_8 As Long
Public ColCountC, ColCountC_1, ColCountC_2, ColCountC_3, ColCountC_4, ColCountC_5, ColCountC_6, ColCountC_7, ColCountC_8 As Long
Public ColCountD, ColCountD_1, ColCountD_2, ColCountD_3, ColCountD_4, ColCountD_5, ColCountD_6, ColCountD_7, ColCountD_8, ColCountDARN As Long
Sub Next_80GClick()
Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("80GGA")
    sourceSheet.Activate
End Sub

Sub Validate80G_All()
Dim vbMessgaeCaption As String
 vbMessgaeCaption = "Error"
 
   Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("80G")
    
If Not Validate80G_A Then
    sourceSheet.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    fmsgbox (MsgBox_80GA)
    CloseMsg
End If

If Not Validate80G_B Then
    sourceSheet.Activate
    'MsgBox (MsgBox_80GB), vbOKOnly, vbMessgaeCaption
     fmsgbox (MsgBox_80GB)
    CloseMsg
End If


If Not Validate80G_C Then
    sourceSheet.Activate
    'MsgBox (MsgBox_80GC), vbOKOnly, vbMessgaeCaption
    fmsgbox (MsgBox_80GC)
    CloseMsg
End If

If Not Validate80G_D Then
    sourceSheet.Activate
    'MsgBox (MsgBox_80GD), vbOKOnly, vbMessgaeCaption
    fmsgbox (MsgBox_80GD)
    CloseMsg
End If

If Not ValidateTotDon100Percent_80G Then
     sourceSheet.Activate
    'MsgBox (MsgBox_80G), vbOKOnly, vbMessgaeCaption
    fmsgbox (MsgBox_80G)
    CloseMsg
End If


'Ankita_03/04/2026
If Not ValidatePAN_Comb90 Then
     sourceSheet.Activate
' If errmsg80G <> "" Then
    fmsgbox (MsgBox_80G)
' End If
End If

If Not ValidatePAN_CashRestriction Then
     sourceSheet.Activate
' If errmsg80G <> "" Then
'    fmsgbox (MsgBox_80G)
' End If
End If
'--------------------



End Sub



'<------Validation for the 1st Grid 80G _A starts------------->

Function Validate80G_A() As Boolean
subProcCaption = "Validating 80GA"
    Validate80G_A = True
    MsgBox_80GA = "Schedule 80G " & Chr(10) & "A." & Chr(10)
        setTableInfo_A
        setTableInfo_A1
        setTableInfo_A2
        setTableInfo_A3
        setTableInfo_A4
        setTableInfo_A5
        setTableInfo_A6
        setTableInfo_A7
        'Ayush_6/3/2025
        setTableInfo_A8
        'Ayush_6/3/2025
        ColCountA = WorksheetFunction.Max(0, ColCountA, ColCountA_1, ColCountA_2, ColCountA_3, ColCountA_4, ColCountA_5, ColCountA_6, ColCountA_7, ColCountA_8)
    
        If Not ValidateNameDonee1_80GA Then Validate80G_A = False
        If Not ValidateAddr1_80GA Then Validate80G_A = False
        If Not ValidateCity1_80GA Then Validate80G_A = False
        If Not ValidateStateCode1_80GA Then Validate80G_A = False
        If Not ValidatePinCode_80GA Then Validate80G_A = False
        If Not ValidatePan1_80GA Then Validate80G_A = False
        If Not ValidateDonationAmt_80GA Then Validate80G_A = False
        If Not ValidateTotDon100Percent_80GA Then Validate80G_A = False
        'Change.28.02.2023.102.IDS.29/30/31A
        If Not ValidateDonationAmtTotal_80GA Then Validate80G_A = False
        If Not ValidateDonationAmtTotOfTotal_80GA Then Validate80G_A = False
        If Not ValidateDonationAmtTotOfOtherMode_80GA Then Validate80G_A = False
        
        'Ankita_6/3/2025
'        If Sheet4.Range("Per10080G.DonationAmtOther").Value > 0 Then
        If Not ValidateTransaction1_80GA Then Validate80G_A = False
        If Not ValidateIFSC_80GA Then Validate80G_A = False
'        End If
        
        'End Change IDS.29/30/31A
End Function
'Change.28.02.2023.102.IDS.29/30/31
Function ValidateDonationAmtTotal_80GA() As Boolean
    ValidateDonationAmtTotal_80GA = True
    'setTableInfo_A
    noOfProcessSub = ColCountA
    Dim rangecells As Range
    Set rangecells = Range("Per10080G.DonationAmtTotal").Cells
    Dim i As Long
    ReDim DonationAmtTotal_80GA(ColCountA)
    For i = 1 To ColCountA
        DonationAmtTotal_80GA(i) = rangecells.item(i).Value
        
        If Len(DonationAmtTotal_80GA(i)) > 14 Then
            MsgBox_80GA = MsgBox_80GA + "* Total donation in Schedule 80G_A cannot be greater than 14 digits at Sr. No " & i & " " & Chr(13)
            
            ValidateDonationAmtTotal_80GA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateDonationAmtTotOfTotal_80GA() As Boolean
 ValidateDonationAmtTotOfTotal_80GA = True
 DonationAmtTotOfTotal_80GA = Range("Per10080G.TotDon100PercentTotal").Value
 
 If Len(DonationAmtTotOfTotal_80GA) > 14 Then
    MsgBox_80GA = MsgBox_80GA + "* Total amount of Total donation  in Schedule 80G_A cannot be greater than 14 digits " & Chr(13)
    ValidateDonationAmtTotOfTotal_80GA = False
Exit Function
End If

End Function
Function ValidateDonationAmtTotOfOtherMode_80GA() As Boolean
 ValidateDonationAmtTotOfOtherMode_80GA = True
 DonationAmtTotOfOtherMode_80GA = Range("Per10080G.TotDonOther100Percent").Value
 
 If Len(DonationAmtTotOfOtherMode_80GA) > 14 Then
    MsgBox_80GA = MsgBox_80GA + "* Total amount of Donation in Other Mode  in Schedule 80G_A cannot be greater than 14 digits " & Chr(13)
    ValidateDonationAmtTotOfOtherMode_80GA = False
Exit Function
End If

End Function
'End Change IDS29/30/31
Function ValidateNameDonee1_80GA() As Boolean
    ValidateNameDonee1_80GA = True
    'setTableInfo_A
    noOfProcessSub = ColCountA
    Dim rangecells As Range
    Set rangecells = Range("Per10080G.DoneeName").Cells
    Dim i As Long
    ReDim Name1_80GA(ColCountA)
    For i = 1 To ColCountA
        Name1_80GA(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Name1_80GA(i)) Then
           ' MsgBox_80GA = MsgBox_80GA + "* Name of the Donee at Sr. No " & i & " in Schedule 80G_A is mandatory" & Chr(13)
           MsgBox_80GA = MsgBox_80GA + "* Please enter name of Donee" & Chr(13)
            ValidateNameDonee1_80GA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateAddr1_80GA() As Boolean
    ValidateAddr1_80GA = True
    'setTableInfo_A
    Dim rangecells As Range
    Set rangecells = Range("Per10080G.AddrDetail").Cells
    Dim i As Long
    ReDim Addr1_80GA(ColCountA)
    For i = 1 To ColCountA
        Addr1_80GA(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Addr1_80GA(i)) Then
           ' MsgBox_80GA = MsgBox_80GA + "* Address of the Donee at Sr. No " & i & " in Schedule 80G_A is mandatory" & Chr(13)
            MsgBox_80GA = MsgBox_80GA + "* Please enter address of Donee" & Chr(13)

            ValidateAddr1_80GA = False
            Exit Function
        End If
    Next
End Function

Function ValidateCity1_80GA() As Boolean
    ValidateCity1_80GA = True
    'setTableInfo_A
    Dim rangecells As Range
    Set rangecells = Range("Per10080G.CityOrTownOrDistrict").Cells
    Dim i As Long
    ReDim City1_80GA(ColCountA)
    For i = 1 To ColCountA
        City1_80GA(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(City1_80GA(i)) Then
            'MsgBox_80GA = MsgBox_80GA + "* City/Town/District of the Donee at Sr. No " & i & " in Schedule 80G_A is mandatory" & Chr(13)
            MsgBox_80GA = MsgBox_80GA + "*Please enter city/town/district of Donee" & Chr(13)
            ValidateCity1_80GA = False
            Exit Function
        End If
        
        
    Next
End Function

Function ValidateStateCode1_80GA() As Boolean
ValidateStateCode1_80GA = True
    '
    'setTableInfo_A
    Dim rangecells As Range
    Set rangecells = Range("Per10080G.StateCode").Cells
    Dim i As Long
    ReDim State1_80GA(ColCountA)
      ReDim StateCode_Per10080G(ColCountA)
    For i = 1 To ColCountA
        State1_80GA(i) = rangecells.item(i).Value
        
        If ((State1_80GA(i) = "(Select)") Or (State1_80GA(i) = "")) Then
          MsgBox_80GA = MsgBox_80GA + "* Selection of State Code at Sr. No " & i & " in Schedule 80G_A is mandatory" & Chr(13)
          ValidateStateCode1_80GA = False
            Exit Function
        End If
        'ramy
        'mmmm
        
        'Malli------------17/09/2024
        If State1_80GA(i) <> "" Then
        Sheets("80G").Activate
        Dim PIN_targetadd, state_targetadd As String
        'state_targetadd = Target.address
         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
        PIN_targetadd = Replace(state_targetadd, "H", "I")
        If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet4.Range(Replace(state_targetadd, "H", "I")).Value = ""

        End If
        '------------------------------
       
    Next
End Function

Function ValidatePinCode_80GA() As Boolean
ValidatePinCode_80GA = True
    'setTableInfo_A
    Dim rangecells As Range
    Set rangecells = Range("Per10080G.PinCode").Cells
    Dim i As Long
    ReDim PinCd1_80GA(ColCountA)
    For i = 1 To ColCountA
        PinCd1_80GA(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(PinCd1_80GA(i)) Then
            MsgBox_80GA = MsgBox_80GA + "* Please enter pin code of Donee" & Chr(13)
            ValidatePinCode_80GA = False
            Exit Function
        End If

        If Not chkNumeric(PinCd1_80GA(i)) Then
            MsgBox_80GA = MsgBox_80GA + "* Pin Code at Sr. No " & i & " in Schedule 80G_A Must be 6 digits Numeric Value" & Chr(13)
            ValidatePinCode_80GA = False
            Exit Function
        End If
        
         If Not checkfieldspecialcharacter(PinCd1_80GA(i)) Then
             MsgBox_80GA = MsgBox_80GA + "* Pin Code at Sr. No  " & i & " in Schedule 80G_A,  Cannot Contain Special Characters" & Chr(13)
             ValidatePinCode_80GA = False
             Exit Function
        End If
    Next
End Function

Function ValidatePan1_80GA() As Boolean
ValidatePan1_80GA = True
    'setTableInfo_A
    Dim rangecells As Range
    'Set rangecells = Range("Per10080G.DoneePAN").Cells
    Set rangecells = Range("Per10080G.DoneePAN").Cells(1, 1)
    Dim i As Long
     
    ReDim Pan1_80GA(ColCountA)
    For i = 1 To ColCountA
   
        Pan1_80GA(i) = rangecells.item(i).Value
        'MsgBox (Pan1_80GA(i))
        If Not chkCompulsory(Pan1_80GA(i)) Then
            MsgBox_80GA = MsgBox_80GA + "* Please enter PAN of Donee" & Chr(13)
            ValidatePan1_80GA = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(Pan1_80GA(i)) Then
             MsgBox_80GA = MsgBox_80GA + "* PAN No at Sr. No  " & i & " in Schedule 80G_A,  Cannot Contain Special Characters" & Chr(13)
             ValidatePan1_80GA = False
             Exit Function
        End If
        
        If Not CheckDoneePAN(UCase(Pan1_80GA(i))) Then
             MsgBox_80GA = MsgBox_80GA + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
             ValidatePan1_80GA = False
             Exit Function
        End If
        If ((UCase(Pan1_80GA(i)) = UCase(Sheet1.Range("sheet1.PAN").Value)) Or (UCase(Pan1_80GA(i)) = UCase(Sheet3.Range("Ver.PAN").Value))) Then
        MsgBox_80GA = MsgBox_80GA + "* PAN of donee at Sr. No  " & i & " is Invalid." & Chr(13) & " Donee PAN cannot be assessee PAN or verification PAN.in Schedule 80G_A" & Chr(13)
        ValidatePan1_80GA = False
        Exit Function
        End If
    Next
End Function
Function ValidateDonationAmt_80GA() As Boolean
ValidateDonationAmt_80GA = True
    'setTableInfo_A
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Range("Per10080G.DonationAmt").Cells
    Set rangecells1 = Range("Per10080G.DonationAmtOther").Cells
    Dim i As Long
    ReDim DonationAmt1_80GA(ColCountA)
    ReDim DonationAmt1_80GA1(ColCountA)
    
    For i = 1 To ColCountA
        DonationAmt1_80GA(i) = rangecells.item(i).Value
         DonationAmt1_80GA1(i) = rangecells1.item(i).Value
         

        
        If Not chkCompulsory(DonationAmt1_80GA(i)) And (DonationAmt1_80GA1(i) = "") Then
            MsgBox_80GA = MsgBox_80GA + "* Enter the amount of donation either in field Donation in cash or Donation in other mode" & Chr(13)
            ValidateDonationAmt_80GA = False
            Exit Function
            

            
        End If
    Next
End Function
'Ankita_6/3/2025
Function ValidateTransaction1_80GA() As Boolean
    ValidateTransaction1_80GA = True
    'setTableInfo_A
    Dim rangecells As Range
    'Set rangecells = Range("Per10080G.DoneePAN").Cells
    Set rangecells = Range("Per10080G.Traref").Cells
    Dim i As Long
     
    ReDim Transaction1_80GA(ColCountA)
    For i = 1 To ColCountA
   
        Transaction1_80GA(i) = rangecells.item(i).Value
    If rangecells.item(i).Locked = False Then
        If Not chkCompulsory(Transaction1_80GA(i)) Then
            MsgBox_80GA = MsgBox_80GA + "*Please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr. No " & i & "" & Chr(13)
            ValidateTransaction1_80GA = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter4(Transaction1_80GA(i)) Then
             MsgBox_80GA = MsgBox_80GA + "* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS at Sr. No " & i & " in Schedule 80G_A, cannot Contain Special Characters" & Chr(13)
             ValidateTransaction1_80GA = False
             Exit Function
        End If
    End If
    Next
End Function
'Ankita_6/3/2025
Function ValidateTransaction1_80GB() As Boolean
    ValidateTransaction1_80GB = True
    Dim rangecells As Range
    Set rangecells = Range("PerNO5080G.Traref").Cells
    Dim i As Long
     
    ReDim Transaction1_80GB(ColCountB)
    For i = 1 To ColCountB
   
        Transaction1_80GB(i) = rangecells.item(i).Value
    If rangecells.item(i).Locked = False Then
        If Not chkCompulsory(Transaction1_80GB(i)) Then
            MsgBox_80GB = MsgBox_80GB + "*Please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr. No " & i & "" & Chr(13)
            ValidateTransaction1_80GB = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter4(Transaction1_80GB(i)) Then
             MsgBox_80GB = MsgBox_80GB + "* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS at Sr. No " & i & " in Schedule 80G_B, cannot Contain Special Characters" & Chr(13)
             ValidateTransaction1_80GB = False
             Exit Function
        End If
     End If
    Next
End Function
'Ankita_6/3/2025
Function ValidateTransaction1_80GC() As Boolean
ValidateTransaction1_80GC = True
    Dim rangecells As Range
    Set rangecells = Range("PerYES10080G.Traref").Cells
    Dim i As Long
     
    ReDim Transaction1_80GC(ColCountC)
    For i = 1 To ColCountC
   
        Transaction1_80GC(i) = rangecells.item(i).Value
    If rangecells.item(i).Locked = False Then
        If Not chkCompulsory(Transaction1_80GC(i)) Then
            MsgBox_80GC = MsgBox_80GC + "*Please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr. No " & i & "" & Chr(13)
            ValidateTransaction1_80GC = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter4(Transaction1_80GC(i)) Then
             MsgBox_80GC = MsgBox_80GC + "* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS at Sr. No " & i & " in Schedule 80G_C, cannot Contain Special Characters" & Chr(13)
             ValidateTransaction1_80GC = False
             Exit Function
        End If
      End If
    Next
End Function
'Ayush_6/3/2025
Function ValidateTransaction1_80GD() As Boolean
ValidateTransaction1_80GD = True
    Dim rangecells As Range
    Set rangecells = Range("Per5080G.Traref").Cells
    Dim i As Long
     
    ReDim Transaction1_80GD(ColCountD)
    For i = 1 To ColCountD
   
        Transaction1_80GD(i) = rangecells.item(i).Value
    
    If rangecells.item(i).Locked = False Then
        If Not chkCompulsory(Transaction1_80GD(i)) Then
            MsgBox_80GD = MsgBox_80GD + "*Please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr. No " & i & "" & Chr(13)
            ValidateTransaction1_80GD = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter4(Transaction1_80GD(i)) Then
             MsgBox_80GD = MsgBox_80GD + "* Transaction Reference number/Cheque number/IMPS/NEFT/RTGS at Sr. No " & i & " in Schedule 80G_D, cannot Contain Special Characters" & Chr(13)
             ValidateTransaction1_80GD = False
             Exit Function
        End If
    End If
    Next
End Function
'Ankita_6/3/2025
Function ValidateIFSC_80GA() As Boolean
ValidateIFSC_80GA = True
    Dim rangecells As Range
    Set rangecells = Range("Per10080G.IFSC").Cells
    Dim i As Long
     
    ReDim IFSC_80GA(ColCountA)
    For i = 1 To ColCountA
   
        IFSC_80GA(i) = rangecells.item(i).Value
    If rangecells.item(i).Locked = False Then
        If Not chkCompulsory(IFSC_80GA(i)) Then
            MsgBox_80GA = MsgBox_80GA + "*Please enter ""your bank IFSC from which Contribution is made"" at Sr. No " & i & "" & Chr(13)
            ValidateIFSC_80GA = False
            Exit Function
        End If
        If Len(IFSC_80GA(i)) > 11 Then
            MsgBox_80GA = MsgBox_80GA + "* IFS Code at Sr.No " & i & " in Schedule 80GA cannot exceed 11 characters" & Chr(13)
            ValidateIFSC_80GA = False
        End If
    End If
        If Not ValidateIFSCList(UCase(IFSC_80GA(i))) Then
            MsgBox_80GA = MsgBox_80GA + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80GA. Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
            ValidateIFSC_80GA = False
        End If
    Next
End Function
'Ankita_6/3/2025
Function ValidateIFSC_80GB() As Boolean
ValidateIFSC_80GB = True
    Dim rangecells As Range
    Set rangecells = Range("PerNO5080G.IFSC").Cells
    Dim i As Long
     
    ReDim IFSC_80GB(ColCountB)
    For i = 1 To ColCountB
   
        IFSC_80GB(i) = rangecells.item(i).Value
    If rangecells.item(i).Locked = False Then
        If Not chkCompulsory(IFSC_80GB(i)) Then
            MsgBox_80GB = MsgBox_80GB + "*Please enter ""your bank IFSC from which Contribution is made"" at Sr. No  " & i & "" & Chr(13)
            ValidateIFSC_80GB = False
            Exit Function
        End If
        If Len(IFSC_80GB(i)) > 11 Then
            MsgBox_80GB = MsgBox_80GB + "* IFS Code at Sr.No " & i & " in Schedule 80GB cannot exceed 11 characters" & Chr(13)
            ValidateIFSC_80GB = False
        End If
    End If
        If Not ValidateIFSCList(UCase(IFSC_80GB(i))) Then
            MsgBox_80GB = MsgBox_80GB + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80GB. Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
            ValidateIFSC_80GB = False
        End If

    Next
End Function
'Ankita_6/3/2025
Function ValidateIFSC_80GC() As Boolean
ValidateIFSC_80GC = True
    Dim rangecells As Range
    Set rangecells = Range("PerYES10080G.IFSC").Cells
    Dim i As Long
     
    ReDim IFSC_80GC(ColCountC)
    For i = 1 To ColCountC
   
        IFSC_80GC(i) = rangecells.item(i).Value
    If rangecells.item(i).Locked = False Then
        If Not chkCompulsory(IFSC_80GC(i)) Then
            MsgBox_80GC = MsgBox_80GC + "*Please enter ""your bank IFSC from which Contribution is made"" at Sr. No  " & i & "" & Chr(13)
            ValidateIFSC_80GC = False
            Exit Function
        End If
        If Len(IFSC_80GC(i)) > 11 Then
            MsgBox_80GC = MsgBox_80GC + "* IFS Code at Sr.No " & i & " in Schedule 80GC cannot exceed 11 characters" & Chr(13)
            ValidateIFSC_80GC = False
        End If
    End If
        If Not ValidateIFSCList(UCase(IFSC_80GC(i))) Then
            MsgBox_80GC = MsgBox_80GC + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80GC. Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
            ValidateIFSC_80GC = False
        End If

    Next
End Function
'Ankita_6/3/2025
Function ValidateIFSC_80GD() As Boolean
ValidateIFSC_80GD = True
    Dim rangecells As Range
    Set rangecells = Range("Per5080G.IFSC").Cells
    Dim i As Long
     
    ReDim IFSC_80GD(ColCountD)
    For i = 1 To ColCountD
   
        IFSC_80GD(i) = rangecells.item(i).Value
    If rangecells.item(i).Locked = False Then
        If Not chkCompulsory(IFSC_80GD(i)) Then
            MsgBox_80GD = MsgBox_80GD + "*Please enter ""your bank IFSC from which Contribution is made"" at Sr. No  " & i & "" & Chr(13)
            ValidateIFSC_80GD = False
            Exit Function
        End If
        If Len(IFSC_80GD(i)) > 11 Then
            MsgBox_80GD = MsgBox_80GD + "* IFS Code at Sr.No " & i & " in Schedule 80GD cannot exceed 11 characters" & Chr(13)
            ValidateIFSC_80GD = False
        End If
    End If
        If Not ValidateIFSCList(UCase(IFSC_80GD(i))) Then
            MsgBox_80GD = MsgBox_80GD + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80GD. Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
            ValidateIFSC_80GD = False
        End If

    Next
End Function

Function ValidateTotDon100Percent_80GA() As Boolean
 ValidateTotDon100Percent_80GA = True
 TotDon100Percent_80GA = Range("Per10080G.TotDon100Percent").Value
 
 If Len(TotDon100Percent_80GA) > 14 Then
    MsgBox_80GA = MsgBox_80GA + "* Total amount of Donation in Cash  in Schedule 80G_A cannot be greater than 14 digits " & Chr(13)
    ValidateTotDon100Percent_80GA = False
Exit Function
End If

End Function


Sub setTableInfo_A()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per10080G.DoneeName").Cells
    mIntCells = Range("Per10080G.DoneeName").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountA = ccount
End Sub

Sub setTableInfo_A1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per10080G.AddrDetail").Cells
    mIntCells = Range("Per10080G.AddrDetail").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountA_1 = ccount
End Sub


Sub setTableInfo_A2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per10080G.CityOrTownOrDistrict").Cells
    mIntCells = Range("Per10080G.CityOrTownOrDistrict").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountA_2 = ccount
End Sub

Sub setTableInfo_A3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per10080G.StateCode").Cells
    mIntCells = Range("Per10080G.StateCode").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value = "") Or (rangecells.item(mIntCtr).Value = "(Select)")) Then
        ccount = ccount + 1
        End If
    Next
    ColCountA_3 = ccount
End Sub


Sub setTableInfo_A4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per10080G.PinCode").Cells
    mIntCells = Range("Per10080G.PinCode").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountA_4 = ccount
End Sub

Sub setTableInfo_A5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per10080G.DoneePAN").Cells
    mIntCells = Range("Per10080G.DoneePAN").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountA_5 = ccount
End Sub

Sub setTableInfo_A6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per10080G.DonationAmt").Cells
    mIntCells = Range("Per10080G.DonationAmt").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountA_6 = ccount
End Sub
Sub setTableInfo_A7()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per10080G.DonationAmtOther").Cells
    mIntCells = Range("Per10080G.DonationAmtOther").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountA_7 = ccount
End Sub
Sub setTableInfo_A8()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per10080G.Traref").Cells
    mIntCells = Range("Per10080G.Traref").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountA_8 = ccount
End Sub


' <----------------End of VAliation For Grid 1 Schedule 80G_A-------------------------->


' <---------------Start of Validation for the Grid 2 Schedule 80G-B--------------------->

Function Validate80G_B() As Boolean
subProcCaption = "Validating 80GB"
    Validate80G_B = True
    MsgBox_80GB = "Schedule 80G " & Chr(10) & "B." & Chr(10)
        setTableInfo_B
        setTableInfo_B1
        setTableInfo_B2
        setTableInfo_B3
        setTableInfo_B4
        setTableInfo_B5
        setTableInfo_B6
        setTableInfo_B7
    
    ColCountB = WorksheetFunction.Max(0, ColCountB, ColCountB_1, ColCountB_2, ColCountB_3, ColCountB_4, ColCountB_5, ColCountB_6, ColCountB_7)
    
            If Not ValidateDoneeName2_80GB Then Validate80G_B = False
            If Not ValidateAddr2_80GB Then Validate80G_B = False
            If Not ValidateCity2_80GB Then Validate80G_B = False
            If Not ValidateState2_80GB Then Validate80G_B = False
            If Not ValidatePinCode2_80GB Then Validate80G_B = False
            If Not ValidatePanDonee2_80GB Then Validate80G_B = False
            If Not ValidateDonationAmt2_80GB Then Validate80G_B = False
            If Not ValidateTotDon100Percent_80GB Then Validate80G_B = False
            'Change.28.02.2023.102.IDS.32/33/34A
            If Not ValidateDonationAmtTotal_80GB Then Validate80G_B = False
            If Not ValidateDonationAmtTotOfTotal_80GB Then Validate80G_B = False
            If Not ValidateDonationAmtTotOfOtherMode_80GB Then Validate80G_B = False
            'End Change IDS.32/33/34A
             If Not ValidateTransaction1_80GB Then Validate80G_B = False
             If Not ValidateIFSC_80GB Then Validate80G_B = False

        
        
'    If ((ColCountB <> ColCountB_1) Or (ColCountB <> ColCountB_2) Or (ColCountB <> ColCountB_3) Or (ColCountB <> ColCountB_4) Or (ColCountB <> ColCountB_5)) Then
'        MsgBox_80GB = MsgBox_80GB + "Enter All mandatory Fields in the Scedule in 80G B " & Chr(13)
'        Validate80G_B = False
'    End If
'
'       If ColCountB_7 > 0 Then
'        If (ColCountB <> ColCountB_7) Then
'         MsgBox_80GB = MsgBox_80GB + "Enter All Mandatory Fields in Schedule 80G_A."
'          Validate80G_B = False
'          Exit Function
'        End If
'        End If
'
'        If ColCountB_6 > 0 Then
'        If (ColCountB <> ColCountB_6) Then
'         MsgBox_80GB = MsgBox_80GB + "Enter All Mandatory Fields in Schedule 80G_A."
'          Validate80G_B = False
'          Exit Function
'        End If
'        End If
    

    
End Function

'Change.28.02.2023.102.IDS.32/33/34
Function ValidateDonationAmtTotal_80GB() As Boolean
    ValidateDonationAmtTotal_80GB = True
    'setTableInfo_A
    noOfProcessSub = ColCountB
    Dim rangecells As Range
    Set rangecells = Range("PerNO5080G.DonationAmtTotal").Cells
    Dim i As Long
    ReDim DonationAmtTotal_80GB(ColCountB)
    For i = 1 To ColCountB
        DonationAmtTotal_80GB(i) = rangecells.item(i).Value
        
        If Len(DonationAmtTotal_80GB(i)) > 14 Then
            MsgBox_80GB = MsgBox_80GB + "* Total donation in Schedule 80G_B cannot be greater than 14 digits at Sr. No " & i & " " & Chr(13)
            
            ValidateDonationAmtTotal_80GB = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
Function ValidateDonationAmtTotOfTotal_80GB() As Boolean
 ValidateDonationAmtTotOfTotal_80GB = True
 DonationAmtTotOfTotal_80GB = Range("PerNO5080G.TotDon100PercentTotal").Value
 
 If Len(DonationAmtTotOfTotal_80GB) > 14 Then
    MsgBox_80GB = MsgBox_80GB + "* Total amount of Total donation  in Schedule 80G_B cannot be greater than 14 digits " & Chr(13)
    ValidateDonationAmtTotOfTotal_80GB = False
Exit Function
End If

End Function
Function ValidateDonationAmtTotOfOtherMode_80GB() As Boolean
 ValidateDonationAmtTotOfOtherMode_80GB = True
 DonationAmtTotOfOtherMode_80GB = Range("PerNO5080G.TotDonOther100Percent").Value
 
 If Len(DonationAmtTotOfOtherMode_80GB) > 14 Then
    MsgBox_80GB = MsgBox_80GB + "* Total amount of Donation in Other Mode  in Schedule 80G_B cannot be greater than 14 digits " & Chr(13)
    ValidateDonationAmtTotOfOtherMode_80GB = False
Exit Function
End If

End Function
'End Change IDS32/33/34
Function ValidateDoneeName2_80GB() As Boolean
    ValidateDoneeName2_80GB = True
'    setTableInfo_B
    noOfProcessSub = ColCountB
    Dim rangecells As Range
    Set rangecells = Range("PerNO5080G.DoneeName").Cells
    Dim i As Long
    ReDim Name2_80GB(ColCountB)
    For i = 1 To ColCountB
        Name2_80GB(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Name2_80GB(i)) Then
           ' MsgBox_80GB = MsgBox_80GB + "* Name of the Donee at Sr.NO " & i & " in Schedule 80G_B is mandatory" & Chr(13)
            MsgBox_80GB = MsgBox_80GB + "* Please enter name of Donee" & Chr(13)

            ValidateDoneeName2_80GB = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateAddr2_80GB() As Boolean
    ValidateAddr2_80GB = True
    Dim rangecells As Range
    Set rangecells = Range("PerNO5080G.AddrDetail").Cells
    Dim i As Long
    ReDim Addr280GB(ColCountB)
    For i = 1 To ColCountB
        Addr280GB(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Addr280GB(i)) Then
            MsgBox_80GB = MsgBox_80GB + "* Please enter address of Donee" & Chr(13)
            ValidateAddr2_80GB = False
            Exit Function
        End If
    Next
End Function

Function ValidateCity2_80GB() As Boolean
    ValidateCity2_80GB = True
    Dim rangecells As Range
    Set rangecells = Range("PerNO5080G.CityOrTownOrDistrict").Cells
    Dim i As Long
    ReDim City280GB(ColCountB)
    For i = 1 To ColCountB
        City280GB(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(City280GB(i)) Then
            MsgBox_80GB = MsgBox_80GB + "* Please enter city/town/district of Donee" & Chr(13)
            ValidateCity2_80GB = False
            Exit Function
        End If
    Next
End Function

Function ValidateState2_80GB() As Boolean
    ValidateState2_80GB = True
    Dim rangecells As Range
    Set rangecells = Range("PerNO5080G.StateCode").Cells
    Dim i As Long
    ReDim State280GB(ColCountB)
    For i = 1 To ColCountB
        State280GB(i) = rangecells.item(i).Value
        
        If ((State280GB(i) = "(Select)") Or (State280GB(i) = "")) Then
            MsgBox_80GB = MsgBox_80GB + "* Selection of State Code at Sr.NO " & i & " in Schedule 80G_B is mandatory" & Chr(13)
            ValidateState2_80GB = False
            Exit Function
        End If
          'Malli------------17/09/2024
        If State280GB(i) <> "" Then
        Sheets("80G").Activate
        Dim PIN_targetadd, state_targetadd As String
        'state_targetadd = Target.address
         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
        PIN_targetadd = Replace(state_targetadd, "H", "I")
        If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet4.Range(Replace(state_targetadd, "H", "I")).Value = ""

        End If
        '------------------------------
        
    Next
End Function

Function ValidatePinCode2_80GB() As Boolean
    ValidatePinCode2_80GB = True
    Dim rangecells As Range
    Set rangecells = Range("PerNO5080G.PinCode").Cells
    Dim i As Long
    ReDim PinCode_80GB(ColCountB)
    For i = 1 To ColCountB
        PinCode_80GB(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(PinCode_80GB(i)) Then
            MsgBox_80GB = MsgBox_80GB + "* Please enter pin code of Donee" & Chr(13)
            ValidatePinCode2_80GB = False
            Exit Function
        End If
                
        If Not chkNumeric(PinCode_80GB(i)) Then
            MsgBox_80GB = MsgBox_80GB + "* Pin Code at Sr. No " & i & " in Schedule 80G_B Must be 6 digits Numeric Value" & Chr(13)
            ValidatePinCode2_80GB = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(PinCode_80GB(i)) Then
             MsgBox_80GB = MsgBox_80GB + "* Pin Code at Sr. No  " & i & " in Schedule 80G_B,  Cannot Contain Special Characters" & Chr(13)
             ValidatePinCode2_80GB = False
             Exit Function
        End If
    Next
End Function

Function ValidatePanDonee2_80GB() As Boolean
    ValidatePanDonee2_80GB = True
    Dim rangecells As Range
    'Set rangecells = Range("PerNO5080G.DoneePAN").Cells
    'Malli Changed
    Set rangecells = Range("PerNO5080G.DoneePAN").Cells(1, 1)
    Dim i As Long
    ReDim Pan_80GB(ColCountB)
    For i = 1 To ColCountB
        Pan_80GB(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Pan_80GB(i)) Then
            MsgBox_80GB = MsgBox_80GB + "* Please enter PAN of Donee" & Chr(13)
            ValidatePanDonee2_80GB = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(Pan_80GB(i)) Then
             MsgBox_80GB = MsgBox_80GB + "* PAN No at Sr. No  " & i & " in Schedule 80G_B,  Cannot Contain Special Characters" & Chr(13)
             ValidatePanDonee2_80GB = False
             Exit Function
        End If
        
        If Not CheckDoneePAN(UCase(Pan_80GB(i))) Then
             MsgBox_80GB = MsgBox_80GB + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
             ValidatePanDonee2_80GB = False
             Exit Function
        End If
        
          If ((UCase(Pan_80GB(i)) = UCase(Sheet1.Range("sheet1.PAN").Value)) Or (UCase(Pan_80GB(i)) = UCase(Sheet3.Range("Ver.PAN").Value))) Then
        MsgBox_80GB = MsgBox_80GB + "* PAN of donee at Sr. No  " & i & " is Invalid." & Chr(13) & " Donee PAN cannot be assessee PAN or verification PAN. in Schedule 80G_B. " & Chr(13)
        ValidatePanDonee2_80GB = False
        Exit Function
        End If
    Next
End Function
Function ValidateDonationAmt2_80GB() As Boolean
    ValidateDonationAmt2_80GB = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Range("PerNO5080G.DonationAmt").Cells
    Set rangecells1 = Range("PerNO5080G.DonationAmtOther").Cells
     
    Dim i As Long
    ReDim DonationAmt_80GB(ColCountB)
    ReDim DonationAmt_80GB1(ColCountB)
    
    For i = 1 To ColCountB
        DonationAmt_80GB(i) = rangecells.item(i).Value
        DonationAmt_80GB1(i) = rangecells1.item(i).Value
        
          If Not chkCompulsory(DonationAmt_80GB(i)) And (DonationAmt_80GB1(i) = "") Then
            MsgBox_80GB = MsgBox_80GB + "Enter the amount of donation either in field Donation in cash or Donation in other mode" & Chr(13)
            ValidateDonationAmt2_80GB = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotDon100Percent_80GB() As Boolean
 ValidateTotDon100Percent_80GB = True
 TotDon100Percent_80GB = Range("PerNO5080G.TotDon100Percent").Value
 
 If Len(TotDon100Percent_80GB) > 14 Then
    MsgBox_80GB = MsgBox_80GB + "* Total amount of Donation in Cash in Schedule 80G_B cannot be greater than 14 digits " & Chr(13)
    ValidateTotDon100Percent_80GB = False
Exit Function
End If

End Function

Sub setTableInfo_B()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerNO5080G.DoneeName").Cells
    mIntCells = Range("PerNO5080G.DoneeName").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountB = ccount
End Sub


Sub setTableInfo_B1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerNO5080G.AddrDetail").Cells
    mIntCells = Range("PerNO5080G.AddrDetail").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountB_1 = ccount
End Sub



Sub setTableInfo_B2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerNO5080G.CityOrTownOrDistrict").Cells
    mIntCells = Range("PerNO5080G.CityOrTownOrDistrict").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountB_2 = ccount
End Sub



Sub setTableInfo_B3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerNO5080G.StateCode").Cells
    mIntCells = Range("PerNO5080G.StateCode").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value = "") Or (rangecells.item(mIntCtr).Value = "(Select)")) Then
        ccount = ccount + 1
        
        End If
        'Malli------------17/09/2024
'        If rangecells.item(mIntCtr).Value <> "" Then
'        Dim PIN_targetadd, state_targetadd As String
'        'state_targetadd = Target.address
'         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
'        PIN_targetadd = Replace(state_targetadd, "H", "I")
'        If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet4.Range(Replace(state_targetadd, "H", "I")).Value = ""
'
'        End If
        '------------------------------
       
    Next
    ColCountB_3 = ccount
End Sub



Sub setTableInfo_B4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerNO5080G.PinCode").Cells
    mIntCells = Range("PerNO5080G.PinCode").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountB_4 = ccount
End Sub



Sub setTableInfo_B5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerNO5080G.DoneePAN").Cells
    mIntCells = Range("PerNO5080G.DoneePAN").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountB_5 = ccount
End Sub



Sub setTableInfo_B6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerNO5080G.DonationAmt").Cells
    mIntCells = Range("PerNO5080G.DonationAmt").count
    For mIntCtr = 1 To mIntCells
         If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountB_6 = ccount
End Sub
Sub setTableInfo_B7()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerNO5080G.DonationAmtOther").Cells
    mIntCells = Range("PerNO5080G.DonationAmtOther").count
    For mIntCtr = 1 To mIntCells
         If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountB_7 = ccount
End Sub

'Ankita_06/03/2026========
Sub setTableInfo_B8()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerNO5080G.Traref").Cells
    mIntCells = Range("PerNO5080G.Traref").count
    For mIntCtr = 1 To mIntCells
         If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountB_8 = ccount
End Sub

' <---------------End of Validations for Grid 2 ------------------------------>

' <---------Start of VAlidations For Grid 3 --------------------------------->


Function Validate80G_C() As Boolean
subProcCaption = "Validating 80GC"
    Validate80G_C = True
    MsgBox_80GC = "Schedule 80G " & Chr(10) & "C." & Chr(10)
        setTableInfo_C
        setTableInfo_C1
        setTableInfo_C2
        setTableInfo_C3
        setTableInfo_C4
        setTableInfo_C5
        setTableInfo_C6
        setTableInfo_C7
    
    ColCountC = WorksheetFunction.Max(0, ColCountC, ColCountC_1, ColCountC_2, ColCountC_3, ColCountC_4, ColCountC_5, ColCountC_6, ColCountC_7)
    
            If Not ValidateDoneeName3_80GC Then Validate80G_C = False
'       If Len(Range("PerYES10080G.DoneeWithPanName").Item(1).Value) > 0 Then
            If Not ValidateAddr3_80GC Then Validate80G_C = False
            If Not ValidateCity3_80GC Then Validate80G_C = False
            If Not ValidateState3_80GC Then Validate80G_C = False
            If Not ValidatePinCode3_80GC Then Validate80G_C = False
            If Not ValidatePanDonee3_80GC Then Validate80G_C = False
            If Not ValidateDonationAmt3_80GC Then Validate80G_C = False
            If Not ValidateTotDon100Percent_80GC Then Validate80G_C = False
            'Change.28.02.2023.102.IDS.35/36/37
            If Not ValidateDonationAmtTotal_80GC Then Validate80G_C = False
            If Not ValidateDonationAmtTotOfTotal_80GC Then Validate80G_C = False
            If Not ValidateDonationAmtTotOfOtherMode_80GC Then Validate80G_C = False
            'End Change IDS.35/36/37
            If Not ValidateTransaction1_80GC Then Validate80G_C = False
            If Not ValidateIFSC_80GC Then Validate80G_C = False

        
'
'    If ((ColCountC <> ColCountC_1) Or (ColCountC <> ColCountC_2) Or (ColCountC <> ColCountC_3) Or (ColCountC <> ColCountC_4) Or (ColCountC <> ColCountC_5)) Then
'        MsgBox_80GC = MsgBox_80GC + "Enter All mandatory Fields in the Scedule in 80G C " & Chr(13)
'        Validate80G_C = False
'    End If
'
'
'     If ColCountC_7 > 0 Then
'        If (ColCountC <> ColCountC_7) Then
'         MsgBox_80GC = MsgBox_80GC + "Enter All Mandatory Fields in Schedule 80G_A."
'          Validate80G_C = False
'          Exit Function
'        End If
'        End If
'
'        If ColCountC_6 > 0 Then
'        If (ColCountC <> ColCountC_6) Then
'         MsgBox_80GC = MsgBox_80GC + "Enter All Mandatory Fields in Schedule 80G_A."
'          Validate80G_C = False
'          Exit Function
'        End If
'        End If
    

    
    
End Function
'Change.28.02.2023.102.IDS.35/36/37
Function ValidateDonationAmtTotal_80GC() As Boolean
    ValidateDonationAmtTotal_80GC = True
    'setTableInfo_A
    noOfProcessSub = ColCountC
    Dim rangecells As Range
    Set rangecells = Range("PerYES10080G.DonationAmtTotal").Cells
    Dim i As Long
    ReDim DonationAmtTotal_80GC(ColCountC)
    For i = 1 To ColCountC
        DonationAmtTotal_80GC(i) = rangecells.item(i).Value
        
        If Len(DonationAmtTotal_80GC(i)) > 14 Then
            MsgBox_80GC = MsgBox_80GC + "* Total donation in Schedule 80G_C cannot be greater than 14 digits at Sr. No " & i & " " & Chr(13)
            
            ValidateDonationAmtTotal_80GC = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
Function ValidateDonationAmtTotOfTotal_80GC() As Boolean
 ValidateDonationAmtTotOfTotal_80GC = True
 DonationAmtTotOfTotal_80GC = Range("PerYES10080G.TotDon100PercentTotal").Value
 
 If Len(DonationAmtTotOfTotal_80GC) > 14 Then
    MsgBox_80GC = MsgBox_80GC + "* Total amount of Total donation  in Schedule 80G_C cannot be greater than 14 digits " & Chr(13)
    ValidateDonationAmtTotOfTotal_80GC = False
Exit Function
End If

End Function
Function ValidateDonationAmtTotOfOtherMode_80GC() As Boolean
 ValidateDonationAmtTotOfOtherMode_80GC = True
 DonationAmtTotOfOtherMode_80GC = Range("PerYES10080G.TotDonOther100Percent").Value
 
 If Len(DonationAmtTotOfOtherMode_80GC) > 14 Then
    MsgBox_80GC = MsgBox_80GC + "* Total amount of Donation in Other Mode  in Schedule 80G_C cannot be greater than 14 digits " & Chr(13)
    ValidateDonationAmtTotOfOtherMode_80GC = False
Exit Function
End If

End Function
'End Change IDS.35/36/37

Function ValidateDoneeName3_80GC() As Boolean
    ValidateDoneeName3_80GC = True
'    setTableInfo_C
    noOfProcessSub = ColCountC
    Dim rangecells As Range
    Set rangecells = Range("PerYES10080G.DoneeWithPanName").Cells
    Dim i As Long
    ReDim Name_80GC(ColCountC)
    For i = 1 To ColCountC
        Name_80GC(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Name_80GC(i)) Then
            'MsgBox_80GC = MsgBox_80GC + "* Name of the Donee at Sr.NO " & i & " in Schedule 80G_C is mandatory" & Chr(13)
            MsgBox_80GC = MsgBox_80GC + "* Please enter name of Donee" & Chr(13)
            ValidateDoneeName3_80GC = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateAddr3_80GC() As Boolean
    ValidateAddr3_80GC = True
    Dim rangecells As Range
    Set rangecells = Range("PerYES10080G.AddrDetail").Cells
    Dim i As Long
    ReDim Addr_80GC(ColCountC)
    For i = 1 To ColCountC
        Addr_80GC(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Addr_80GC(i)) Then
            MsgBox_80GC = MsgBox_80GC + "* Please enter address of Donee" & Chr(13)
            ValidateAddr3_80GC = False
            Exit Function
        End If
    Next
End Function

Function ValidateCity3_80GC() As Boolean
    ValidateCity3_80GC = True
    Dim rangecells As Range
    Set rangecells = Range("PerYES10080G.CityOrTownOrDistrict").Cells
    Dim i As Long
    ReDim City_80GC(ColCountC)
    For i = 1 To ColCountC
        City_80GC(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(City_80GC(i)) Then
            MsgBox_80GC = MsgBox_80GC + "* Please enter city/town/district of Donee" & Chr(13)
            ValidateCity3_80GC = False
            Exit Function
        End If
    Next
End Function

Function ValidateState3_80GC() As Boolean
    ValidateState3_80GC = True
    Dim rangecells As Range
    Set rangecells = Range("PerYES10080G.StateCode").Cells
    Dim i As Long
    ReDim State_80GC(ColCountC)
    For i = 1 To ColCountC
        State_80GC(i) = rangecells.item(i).Value
        
        If ((State_80GC(i) = "(Select)") Or (State_80GC(i) = "")) Then
            MsgBox_80GC = MsgBox_80GC + "* Selection of State Code at Sr.NO " & i & " in Schedule 80G_C is mandatory" & Chr(13)
            ValidateState3_80GC = False
            Exit Function
        End If
        'Malli------------17/09/2024
        If State_80GC(i) <> "" Then
        Sheets("80G").Activate
        Dim PIN_targetadd, state_targetadd As String
        'state_targetadd = Target.address
         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
        PIN_targetadd = Replace(state_targetadd, "H", "I")
        If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet4.Range(Replace(state_targetadd, "H", "I")).Value = ""
        End If
        '------------------------------
       
    Next
End Function

Function ValidatePinCode3_80GC() As Boolean
    ValidatePinCode3_80GC = True
    Dim rangecells As Range
    Set rangecells = Range("PerYES10080G.PinCode").Cells
    Dim i As Long
    ReDim PinCode_80GC(ColCountC)
    For i = 1 To ColCountC
        PinCode_80GC(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(PinCode_80GC(i)) Then
            MsgBox_80GC = MsgBox_80GC + "* Please enter pin code of Donee" & Chr(13)
            ValidatePinCode3_80GC = False
            Exit Function
        End If
        
        If Not chkNumeric(PinCode_80GC(i)) Then
            MsgBox_80GC = MsgBox_80GC + "* Pin Code at Sr. No " & i & " in Schedule 80G_C Must be 6 digits Numeric Value" & Chr(13)
            ValidatePinCode3_80GC = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(PinCode_80GC(i)) Then
             MsgBox_80GC = MsgBox_80GC + "* Pin Code at Sr. No  " & i & " in Schedule 80G_C,  Cannot Contain Special Characters" & Chr(13)
             ValidatePinCode3_80GC = False
             Exit Function
        End If
        
            
        
    Next
End Function

Function ValidatePanDonee3_80GC() As Boolean
    ValidatePanDonee3_80GC = True
    Dim rangecells As Range
    'Set rangecells = Range("PerYES10080G.DoneePAN").Cells
    'Malli Changed
    Set rangecells = Range("PerYES10080G.DoneePAN").Cells(1, 1)
    Dim i As Long
    ReDim Pan_80GC(ColCountC)
    For i = 1 To ColCountC
        Pan_80GC(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Pan_80GC(i)) Then
            MsgBox_80GC = MsgBox_80GC + "* Please enter PAN of Donee" & Chr(13)
            ValidatePanDonee3_80GC = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(Pan_80GC(i)) Then
             MsgBox_80GC = MsgBox_80GC + "* PAN No at Sr. No  " & i & " in Schedule 80G_C,  Cannot Contain Special Characters" & Chr(13)
             ValidatePanDonee3_80GC = False
             Exit Function
        End If
        
        If Not CheckDoneePAN(UCase(Pan_80GC(i))) Then
             MsgBox_80GC = MsgBox_80GC + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
             ValidatePanDonee3_80GC = False
             Exit Function
        End If
         If ((UCase(Pan_80GC(i)) = UCase(Sheet1.Range("sheet1.PAN").Value)) Or (UCase(Pan_80GC(i)) = UCase(Sheet3.Range("Ver.PAN").Value))) Then
        MsgBox_80GC = MsgBox_80GC + "* PAN of donee at Sr. No  " & i & " is Invalid." & Chr(13) & " Donee PAN cannot be assessee PAN or verification PAN. in Schedule 80G_C. " & Chr(13)
        ValidatePanDonee3_80GC = False
        Exit Function
        End If
    Next
End Function
Function ValidateDonationAmt3_80GC() As Boolean
    ValidateDonationAmt3_80GC = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    
    Set rangecells1 = Range("PerYES10080G.DonationAmtOther").Cells
    Set rangecells = Range("PerYES10080G.DonationAmt").Cells
    Dim i As Long
    
    ReDim DonationAmt_80GC(ColCountC)
    ReDim DonationAmt_80GC1(ColCountC)
    
    For i = 1 To ColCountC
    
        DonationAmt_80GC(i) = rangecells.item(i).Value
        DonationAmt_80GC1(i) = rangecells1.item(i).Value
        
        If Not chkCompulsory(DonationAmt_80GC(i)) And (DonationAmt_80GC1(i) = "") Then
            MsgBox_80GC = MsgBox_80GC + "* Enter the amount of donation either in field Donation in cash or Donation in other mode" & Chr(13)
            ValidateDonationAmt3_80GC = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotDon100Percent_80GC() As Boolean
 ValidateTotDon100Percent_80GC = True
 TotDon100Percent_80GC = Range("PerYES10080G.TotDon100Percent").Value
 
 If Len(TotDon100Percent_80GC) > 14 Then
    MsgBox_80GC = MsgBox_80GC + "* Total amount of Donation in Cash in Schedule 80G_C cannot be greater than 14 digits " & Chr(13)
    ValidateTotDon100Percent_80GC = False
Exit Function
End If

End Function
Sub setTableInfo_C()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerYES10080G.DoneeWithPanName").Cells
    mIntCells = Range("PerYES10080G.DoneeWithPanName").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountC = ccount
End Sub


Sub setTableInfo_C1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerYES10080G.AddrDetail").Cells
    mIntCells = Range("PerYES10080G.AddrDetail").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountC_1 = ccount
End Sub



Sub setTableInfo_C2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerYES10080G.CityOrTownOrDistrict").Cells
    mIntCells = Range("PerYES10080G.CityOrTownOrDistrict").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountC_2 = ccount
End Sub



Sub setTableInfo_C3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerYES10080G.StateCode").Cells
    mIntCells = Range("PerYES10080G.StateCode").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value = "") Or (rangecells.item(mIntCtr).Value = "(Select)")) Then
        ccount = ccount + 1
        End If
    Next
    ColCountC_3 = ccount
End Sub



Sub setTableInfo_C4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerYES10080G.PinCode").Cells
    mIntCells = Range("PerYES10080G.PinCode").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountC_4 = ccount
End Sub



Sub setTableInfo_C5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerYES10080G.DoneePAN").Cells
    mIntCells = Range("PerYES10080G.DoneePAN").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountC_5 = ccount
End Sub



Sub setTableInfo_C6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerYES10080G.DonationAmt").Cells
    mIntCells = Range("PerYES10080G.DonationAmt").count
    For mIntCtr = 1 To mIntCells
         If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountC_6 = ccount
End Sub

Sub setTableInfo_C7()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerYES10080G.DonationAmtOther").Cells
    mIntCells = Range("PerYES10080G.DonationAmtOther").count
    For mIntCtr = 1 To mIntCells
         If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountC_7 = ccount
End Sub

Sub setTableInfo_C8()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("PerYES10080G.Traref").Cells
    mIntCells = Range("PerYES10080G.Traref").count
    For mIntCtr = 1 To mIntCells
         If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountC_8 = ccount
End Sub

' <-------------End of Validations for Grid 3 --------------------->

' <-------------Start of VAlidation for Grid 4 ---------------------->



Function Validate80G_D() As Boolean
subProcCaption = "Validating 80GD"
    Validate80G_D = True
    MsgBox_80GD = "Schedule 80G " & Chr(10) & "D." & Chr(10)
        setTableInfo_D
        setTableInfo_D1
        setTableInfo_D2
        setTableInfo_D3
        setTableInfo_D4
        setTableInfo_D5
        setTableInfo_D6
        setTableInfo_D7
'Change.27.01.2023.102.80GARN_A0.2
        setTableInfo_DARN
 'End Change====
    ColCountD = WorksheetFunction.Max(0, ColCountD, ColCountD_1, ColCountD_2, ColCountD_3, ColCountD_4, ColCountD_5, ColCountD_6, ColCountD_7, ColCountDARN)
    
            If Not ValidateDoneeName4_80GD Then Validate80G_D = False
'            If Len(Range("Per5080G.DoneeWithPanName").Item(1).Value) > 0 Then
            If Not ValidateAddr4_80GD Then Validate80G_D = False
            If Not ValidateCity4_80GD Then Validate80G_D = False
            If Not ValidateState4_80GD Then Validate80G_D = False
            If Not ValidatePinCode4_80GD Then Validate80G_D = False
            If Not ValidatePanDonee4_80GD Then Validate80G_D = False
'Change.27.01.2023.102.80GARN_B0.2
            If Not ValidateARNnumber_80GD Then Validate80G_D = False
'End Change====
            If Not ValidateDonationAmt4_80GD Then Validate80G_D = False
            If Not ValidateTotDon100Percent_80GD Then Validate80G_D = False
            
            If Not ValidatePAN_ARN_80GD Then Validate80G_D = False
            'Change.28.02.2023.102.IDS.38/39/40
            If Not ValidateDonationAmtTotal_80GD Then Validate80G_D = False
            If Not ValidateDonationAmtTotOfTotal_80GD Then Validate80G_D = False
            If Not ValidateDonationAmtTotOfOtherMode_80GD Then Validate80G_D = False
            If Not ValidateTotalDonationsUs80G_E Then Validate80G_D = False
            If Not ValidateTotalDonationsUs80G_E2 Then Validate80G_D = False
            If Not ValidateTotalDonationsUs80G_E3 Then Validate80G_D = False
            'End Change IDS.38/39/40
            If Not ValidateTransaction1_80GD Then Validate80G_D = False
            If Not ValidateIFSC_80GD Then Validate80G_D = False

        
'
'    If ((ColCountD <> ColCountD_1) Or (ColCountD <> ColCountD_2) Or (ColCountD <> ColCountD_3) Or (ColCountD <> ColCountD_4) Or (ColCountD <> ColCountD_5)) Then
'        MsgBox_80GD = MsgBox_80GD + "Enter All mandatory Fields in the Scedule in 80G D " & Chr(13)
'        Validate80G_D = False
'    End If
'
'
'     If ColCountD_7 > 0 Then
'        If (ColCountD <> ColCountD_7) Then
'         MsgBox_80GD = MsgBox_80GD + "Enter All Mandatory Fields in Schedule 80G_A."
'          Validate80G_D = False
'          Exit Function
'        End If
'        End If
'
'        If ColCountD_6 > 0 Then
'        If (ColCountD <> ColCountD_6) Then
'         MsgBox_80GD = MsgBox_80GD + "Enter All Mandatory Fields in Schedule 80G_A."
'          Validate80G_D = False
'          Exit Function
'        End If
'        End If
    
    
End Function
'Change.28.02.2023.102.IDS.38/39/40
Function ValidateDonationAmtTotal_80GD() As Boolean
    ValidateDonationAmtTotal_80GD = True
    'setTableInfo_A
    noOfProcessSub = ColCountD
    Dim rangecells As Range
    Set rangecells = Range("Per5080G.DonationAmtTotal").Cells
    Dim i As Long
    ReDim DonationAmtTotal_80GD(ColCountD)
    For i = 1 To ColCountD
        DonationAmtTotal_80GD(i) = rangecells.item(i).Value
        
        If Len(DonationAmtTotal_80GD(i)) > 14 Then
            MsgBox_80GD = MsgBox_80GD + "* Total donation in Schedule 80G_D cannot be greater than 14 digits at Sr. No " & i & " " & Chr(13)
            ValidateDonationAmtTotal_80GD = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function
Function ValidateDonationAmtTotOfTotal_80GD() As Boolean
 ValidateDonationAmtTotOfTotal_80GD = True
 DonationAmtTotOfTotal_80GD = Range("Per5080G.TotDon100PercentTotal").Value
 
 If Len(DonationAmtTotOfTotal_80GD) > 14 Then
    MsgBox_80GD = MsgBox_80GD + "* Total amount of Total donation in Schedule 80G_D cannot be greater than 14 digits " & Chr(13)
    ValidateDonationAmtTotOfTotal_80GD = False
Exit Function
End If

End Function
Function ValidateDonationAmtTotOfOtherMode_80GD() As Boolean
 ValidateDonationAmtTotOfOtherMode_80GD = True
 DonationAmtTotOfOtherMode_80GD = Range("Per5080G.TotDonother100Percent").Value
 
 If Len(DonationAmtTotOfOtherMode_80GD) > 14 Then
    MsgBox_80GD = MsgBox_80GD + "* Total amount of Donation in Other Mode in Schedule 80G_D cannot be greater than 14 digits " & Chr(13)
    ValidateDonationAmtTotOfOtherMode_80GD = False
Exit Function
End If

End Function
'End Change IDS.38/39/40

Function ValidateDoneeName4_80GD() As Boolean
    ValidateDoneeName4_80GD = True
'    setTableInfo_D
    noOfProcessSub = ColCountD
    Dim rangecells As Range
    Set rangecells = Range("Per5080G.DoneeWithPanName").Cells
    Dim i As Long
    ReDim Name_80GD(ColCountD)
    For i = 1 To ColCountD
        Name_80GD(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Name_80GD(i)) Then
           ' MsgBox_80GD = MsgBox_80GD + "* Name of the Donee at Sr.NO " & i & " in Schedule 80G_D is mandatory" & Chr(13)
            MsgBox_80GD = MsgBox_80GD + "* Please enter name of Donee" & Chr(13)

            ValidateDoneeName4_80GD = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateAddr4_80GD() As Boolean
    ValidateAddr4_80GD = True
    Dim rangecells As Range
    Set rangecells = Range("Per5080G.AddrDetail").Cells
    Dim i As Long
    ReDim Addr_80GD(ColCountD)
    For i = 1 To ColCountD
        Addr_80GD(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Addr_80GD(i)) Then
            MsgBox_80GD = MsgBox_80GD + "* Please enter address of Donee" & Chr(13)
            ValidateAddr4_80GD = False
            Exit Function
        End If
    Next
End Function

Function ValidateCity4_80GD() As Boolean
    ValidateCity4_80GD = True
    Dim rangecells As Range
    Set rangecells = Range("Per5080G.CityOrTownOrDistrict").Cells
    Dim i As Long
    ReDim City_80GD(ColCountD)
    For i = 1 To ColCountD
        City_80GD(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(City_80GD(i)) Then
            MsgBox_80GD = MsgBox_80GD + "* Please enter city/town/district of Donee" & Chr(13)
            ValidateCity4_80GD = False
            Exit Function
        End If
    Next
End Function

Function ValidateState4_80GD() As Boolean
    ValidateState4_80GD = True
    Dim rangecells As Range
    Set rangecells = Range("Per5080G.StateCode").Cells
    Dim i As Long
    ReDim State_80GD(ColCountD)
    For i = 1 To ColCountD
        State_80GD(i) = rangecells.item(i).Value
        
        If ((State_80GD(i) = "(Select)") Or (State_80GD(i) = "")) Then
            MsgBox_80GD = MsgBox_80GD + "* Selection of State Code at Sr.NO " & i & " in Schedule 80G_D is mandatory" & Chr(13)
            ValidateState4_80GD = False
            Exit Function
        End If
        'Malli------------17/09/2024
        If State_80GD(i) <> "" Then
        Sheets("80G").Activate
        Dim PIN_targetadd, state_targetadd As String
        'state_targetadd = Target.address
         state_targetadd = Replace(rangecells.item(i).Address, "$", "")
        PIN_targetadd = Replace(state_targetadd, "H", "I")
        If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet4.Range(Replace(state_targetadd, "H", "I")).Value = ""

        End If
        '------------------------------
       
    Next
End Function

Function ValidatePinCode4_80GD() As Boolean
    ValidatePinCode4_80GD = True
    Dim rangecells As Range
    Set rangecells = Range("Per5080G.PinCode").Cells
    Dim i As Long
    ReDim PinCode_80GD(ColCountD)
    For i = 1 To ColCountD
        PinCode_80GD(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(PinCode_80GD(i)) Then
            MsgBox_80GD = MsgBox_80GD + "* Please enter pin code of Donee" & Chr(13)
            ValidatePinCode4_80GD = False
            Exit Function
        End If
        
        If Not chkNumeric(PinCode_80GD(i)) Then
            MsgBox_80GD = MsgBox_80GD + "* Pin Code at Sr. No " & i & " in Schedule 80G_D Must be 6 digits Numeric Value" & Chr(13)
            ValidatePinCode4_80GD = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(PinCode_80GD(i)) Then
             MsgBox_80GD = MsgBox_80GD + "* Pin Code at Sr. No  " & i & " in Schedule 80G_D,  Cannot Contain Special Characters" & Chr(13)
             ValidatePinCode4_80GD = False
             Exit Function
        End If
        
    Next
End Function

Function ValidatePanDonee4_80GD() As Boolean
    ValidatePanDonee4_80GD = True
    Dim rangecells As Range
    'Set rangecells = Range("Per5080G.DoneePAN").Cells
    'Malli changed
    Set rangecells = Range("Per5080G.DoneePAN").Cells(1, 1)
    Dim i As Long
    ReDim Pan_80GD(ColCountD)
    For i = 1 To ColCountD
        Pan_80GD(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Pan_80GD(i)) Then
            MsgBox_80GD = MsgBox_80GD + "* Please enter PAN of Donee" & Chr(13)
            ValidatePanDonee4_80GD = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(Pan_80GD(i)) Then
             MsgBox_80GD = MsgBox_80GD + "* PAN No at Sr. No  " & i & " in Schedule 80G_D,  Cannot Contain Special Characters" & Chr(13)
             ValidatePanDonee4_80GD = False
             Exit Function
        End If
        
        If Not CheckDoneePAN(UCase(Pan_80GD(i))) Then
             MsgBox_80GD = MsgBox_80GD + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
             ValidatePanDonee4_80GD = False
             Exit Function
        End If
             If ((UCase(Pan_80GD(i)) = UCase(Sheet1.Range("sheet1.PAN").Value)) Or (UCase(Pan_80GD(i)) = UCase(Sheet3.Range("Ver.PAN").Value))) Then
         MsgBox_80GD = MsgBox_80GD + "* PAN of donee at Sr. No  " & i & " is Invalid." & Chr(13) & " Donee PAN cannot be assessee PAN or verification PAN.  in Schedule 80G_D. " & Chr(13)
        ValidatePanDonee4_80GD = False
        Exit Function
        End If
    Next
End Function
Function ValidateDonationAmt4_80GD() As Boolean
    ValidateDonationAmt4_80GD = True
    Dim rangecells As Range
    Dim rangecells1 As Range
    
    Set rangecells = Range("Per5080G.DonationAmt").Cells
    Set rangecells1 = Range("Per5080G.DonationAmtOther").Cells
    Dim i As Long
    ReDim DonationAmt_80GD(ColCountD)
    ReDim DonationAmt_80GD1(ColCountD)
    For i = 1 To ColCountD
    
        DonationAmt_80GD(i) = rangecells.item(i).Value
        DonationAmt_80GD1(i) = rangecells1.item(i).Value
        
         If Not chkCompulsory(DonationAmt_80GD(i)) And (DonationAmt_80GD1(i) = "") Then
            MsgBox_80GD = MsgBox_80GD + "* Enter the amount of donation either in field Donation in cash or Donation in other mode" & Chr(13)
            ValidateDonationAmt4_80GD = False
            Exit Function
        End If
        
    Next
End Function

Function ValidateTotDon100Percent_80GD() As Boolean
 ValidateTotDon100Percent_80GD = True
 TotDon100Percent_80GD = Range("Per5080G.TotDon100Percent").Value
 
 If Len(TotDon100Percent_80GD) > 14 Then
    MsgBox_80GD = MsgBox_80GD + "* Total amount of Donation in Cash in Schedule 80G_D cannot be greater than 14 digits " & Chr(13)
    ValidateTotDon100Percent_80GD = False
Exit Function
End If

End Function

Function ValidateTotDon100Percent_80G() As Boolean
ValidateTotDon100Percent_80G = True
 TotDon100Percent_80G = Range("Per5080G.TotalDonationsUs80G").Value
 
 If Len(TotDon100Percent_80G) > 14 Then
    MsgBox_80GD = MsgBox_80GD + "* Total amount of Donation in Cash in Schedule 80G cannot be greater than 14 digits " & Chr(13)
    ValidateTotDon100Percent_80G = False
Exit Function
End If

End Function

Sub setTableInfo_D()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.DoneeWithPanName").Cells
    mIntCells = Range("Per5080G.DoneeWithPanName").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountD = ccount
End Sub


Sub setTableInfo_D1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.AddrDetail").Cells
    mIntCells = Range("Per5080G.AddrDetail").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountD_1 = ccount
End Sub



Sub setTableInfo_D2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.CityOrTownOrDistrict").Cells
    mIntCells = Range("Per5080G.CityOrTownOrDistrict").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountD_2 = ccount
End Sub



Sub setTableInfo_D3()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.StateCode").Cells
    mIntCells = Range("Per5080G.StateCode").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value = "") Or (rangecells.item(mIntCtr).Value = "(Select)")) Then
        ccount = ccount + 1
        End If
    Next
    ColCountD_3 = ccount
End Sub



Sub setTableInfo_D4()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.PinCode").Cells
    mIntCells = Range("Per5080G.PinCode").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountD_4 = ccount
End Sub



Sub setTableInfo_D5()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.DoneePAN").Cells
    mIntCells = Range("Per5080G.DoneePAN").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    ColCountD_5 = ccount
End Sub



Sub setTableInfo_D6()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.DonationAmt").Cells
    mIntCells = Range("Per5080G.DonationAmt").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountD_6 = ccount
End Sub
Sub setTableInfo_D7()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.DonationAmtOther").Cells
    mIntCells = Range("Per5080G.DonationAmtOther").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountD_7 = ccount
End Sub
Sub setTableInfo_D8()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.Traref").Cells
    mIntCells = Range("Per5080G.Traref").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountD_8 = ccount
End Sub

 Sub Print80G_Click()

printWorkSheet
End Sub

 Sub Help80G_Click()
 sPassword = EfilingCommon.getmsgstate
ActiveWorkbook.Unprotect Password:=sPassword
Sheet6.Activate
Sheet6.Visible = xlSheetVisible
ActiveWorkbook.Protect Password:=sPassword
End Sub

 Sub Prev80G_Click()
Sheet9.Activate
End Sub
Sub ValidateSheet80G_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2026-27"     'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
Validate80G_All
'MsgBox "Sheet 80G is OK", vbOKOnly, vbMessgaeCaption
fmsgboxoK "Sheet 80G is OK"

End Sub

 Sub AddRowsSch80G_A_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("Per10080G.DoneeName||Per10080G.AddrDetail||Per10080G.CityOrTownOrDistrict||Per10080G.StateCode||Per10080G.PinCode||Per10080G.DoneePAN||Per10080G.DonationAmt||Per10080G.DonationAmtTotal||Per10080G.EligibleAmt||Per10080G.DonationAmtOther||Per10080G.Traref||Per10080G.IFSC||Data_80G_A||comb_80G_A||Comb_donation_80G_A||Comb_PAN_80G_A")
ActiveCellRange = EfilingCommon.searchLastRow("Per10080G.DoneeName")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

'80G_B.
 Sub AddRows80G_2_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("PerNO5080G.DoneeName||PerNO5080G.AddrDetail||PerNO5080G.CityOrTownOrDistrict||PerNO5080G.StateCode||PerNO5080G.PinCode||PerNO5080G.DoneePAN||PerNO5080G.DonationAmt||PerNO5080G.DonationAmtTotal||PerNO5080G.EligibleAmt||PerNO5080G.DonationAmtOther||PerNO5080G.Traref||PerNO5080G.IFSC||Data_80G_B||comb_80G_B||Comb_donation_80G_B||Comb_PAN_80G_B")
ActiveCellRange = EfilingCommon.searchLastRow("PerNO5080G.DoneeName")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

'80G_C
 Sub AddRowsSch80G_C_Click()
Dim vRows As Long
EfilingCommon.DefinedgridNameRange = ("PerYES10080G.DoneeWithPanName||PerYES10080G.AddrDetail||PerYES10080G.CityOrTownOrDistrict||PerYES10080G.StateCode||PerYES10080G.PinCode||PerYES10080G.DoneePAN||PerYES10080G.DonationAmt||PerYES10080G.DonationAmtTotal||PerYES10080G.EligibleAmt||PerYES10080G.DonationAmtOther||PerYES10080G.Traref||PerYES10080G.IFSC||Data_80G_C||comb_80G_C||Comb_donation_80G_C||Comb_PAN_80G_C")
ActiveCellRange = EfilingCommon.searchLastRow("PerYES10080G.DoneeWithPanName")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub

'80G_D
 Sub AddRowsSch80G_D_Click()
Dim vRows As Long
'EfilingCommon.DefinedgridNameRange = ("Per5080G.DoneeWithPanName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.DonationAmt||Per5080G.DonationAmtTotal||Per5080G.EligibleAmt||Per5080G.DonationAmtOther")
'Change.25.01.2023.102.80GC7
' EfilingCommon.DefinedgridNameRange = ("Per5080G.DoneeWithPanName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.DonationAmt||Per5080G.DonationAmtTotal||Per5080G.EligibleAmt||Per5080G.DonationAmtOther||Per5080G.DonationAmt_temp")
 EfilingCommon.DefinedgridNameRange = ("Per5080G.DoneeWithPanName||Per5080G.AddrDetail||Per5080G.CityOrTownOrDistrict||Per5080G.StateCode||Per5080G.PinCode||Per5080G.DoneePAN||Per5080G.ArnNbr||Per5080G.DonationAmt||Per5080G.DonationAmtTotal||Per5080G.EligibleAmt||Per5080G.DonationAmtOther||Per5080G.DonationAmt_temp||Per5080G.Traref||Per5080G.IFSC||Data_80G_D||comb_80G_D||Comb_donation_80G_D||Comb_PAN_80G_D")
'ENd Change====
ActiveCellRange = EfilingCommon.searchLastRow("Per5080G.DoneeWithPanName")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
'Change.27.01.2023.102.80G0.2V
Function ValidateARNnumber_80GD() As Boolean
    ValidateARNnumber_80GD = True
'    setTableInfo_D
    noOfProcessSub = ColCountD
    Dim rangecells As Range
    Set rangecells = Range("Per5080G.ArnNbr").Cells
    Dim i, j As Long
    Dim PANValue As Variant
    ReDim ArnNumber_80GD(ColCountD)
    For i = 1 To Range("Per5080G.ArnNbr").count
        'ArnNumber_80GD(i) = rangecells.item(i).Value
        PANValue = Range("Per5080G.DoneePAN").Cells(i, 1).Value
        
'        For j = 1 To Range("Per5080G.ArnNbr").count
'            If PANValue <> "" And PANValue = Range("Per5080G.DoneePAN").Cells(j, 1).Value And Range("Per5080G.ArnNbr").Cells(i, 1).Value = "" Then
'                MsgBox_80GD = MsgBox_80GD + "* Please enter ARN (Donation reference Number) at Sr.NO " & i & " in Schedule 80G_D." & Chr(13)
'                ValidateARNnumber_80GD = False
'                Exit Function
'            End If
'        Next
'        If Range("Per5080G.DoneePAN").Cells(i, 1).Value <> "" And Range("Per5080G.ArnNbr").Cells(i, 1).Value = "" Then
'            MsgBox_80GD = MsgBox_80GD + "* Please enter ARN (Donation reference Number) at Sr.NO " & i & " in Schedule 80G_D." & Chr(13)
'            ValidateARNnumber_80GD = False
'            Exit Function
'        End If
        UpdateProgressBar
    Next
End Function

Sub setTableInfo_DARN()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    Set rangecells = Range("Per5080G.ArnNbr").Cells
    mIntCells = Range("Per5080G.ArnNbr").count
    For mIntCtr = 1 To mIntCells
        If Not ((rangecells.item(mIntCtr).Value) = "") Then
        ccount = ccount + 1
        End If
    Next
    ColCountDARN = ccount
End Sub
'End Change====
'Change.28.02.2023.102.IDS.41/42
Function ValidateTotalDonationsUs80G_E() As Boolean
ValidateTotalDonationsUs80G_E = True
 TotalDonationsUs80G_E = Range("Per5080G.TotalDonationsUs80G").Value
 
 If Len(TotalDonationsUs80G_E) > 14 Then
    MsgBox_80GD = MsgBox_80GD + "* Total amount of Donation in Cash in Schedule 80GE cannot be greater than 14 digits " & Chr(13)
    ValidateTotalDonationsUs80G_E = False
Exit Function
End If

End Function

Function ValidateTotalDonationsUs80G_E2() As Boolean
ValidateTotalDonationsUs80G_E2 = True
 TotalDonationsUs80G_E2 = Range("Per5080G.TotalDonationsOtherUs80G").Value
 
 If Len(TotalDonationsUs80G_E2) > 14 Then
    MsgBox_80GD = MsgBox_80GD + "* Total amount of Donation in Other Mode in Schedule 80GE cannot be greater than 14 digits " & Chr(13)
    ValidateTotalDonationsUs80G_E2 = False
Exit Function
End If

End Function

Function ValidateTotalDonationsUs80G_E3() As Boolean
ValidateTotalDonationsUs80G_E3 = True
 TotalDonationsUs80G_E3 = Range("Per5080G.TotalDonationsUs80GTotal").Value
 
 If Len(TotalDonationsUs80G_E3) > 14 Then
    MsgBox_80GD = MsgBox_80GD + "* Amount of Total Donation Schedule 80GE cannot be greater than 14 digits " & Chr(13)
    ValidateTotalDonationsUs80G_E3 = False
Exit Function
End If

End Function

'End Change IDS.41/42


'Ankita_06/03/2026====

'Function ValidateChequeNumber_80G() As Boolean
'    ValidateChequeNumber_80G = True
'    Dim rangecells As Range
'    Dim rangecells1 As Range
'    Dim rangecells2 As Range
'
'    Set rangecells2 = Sheet4.Range("Per10080G.DonationAmtOther").Cells
'    ReDim Donation_other_80G(end80G)
'    Set rangecells = Sheet4.Range("Per10080G.Traref").Cells
'    Dim i As Long
'    ReDim ChequeNumber_80G(end80G)
'    Dim rangecells3 As Range
'    Set rangecells3 = Sheet4.Range("Per10080G.IFSC").Cells
'    ReDim BankIFSC_80G(end80G)
'
'    For i = 1 To end80G
'        ChequeNumber_80G(i) = rangecells.item(i).Value
'        BankIFSC_80G(i) = rangecells3.item(i).Value
'        Donation_other_80G(i) = rangecells2.item(i).Value
'
'    If rangecells2.item(i).Locked = False Then
'        If Donation_other_80G(i) > 0 Then
'            If Not chkCompulsory(ChequeNumber_80G(i)) Then
'                MsgBox_80GA = MsgBox_80GA + "* Please enter ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr. No " & i & " in Schedule 80GGC" & Chr(13)
'                ValidateChequeNumber_80G = False
'                Exit Function
'            End If
'
'            If isdropdownblank(BankIFSC_80GGC(i)) Then
'                MsgBox_80GA = MsgBox_80GA + "* Please enter ""your bank IFSC from which Contribution is made"" at Sr.No " & i & " in Schedule 80GGC." & Chr(13)
'                ValidateChequeNumber_80G = False
'                Exit Function
'            End If
'
'
'            If Len(ChequeNumber_80G(i)) > 50 Then
'                MsgBox_80GA = MsgBox_80GA + "* ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr.No " & i & " in Schedule 80GGC cannot exceed 50 characters." & Chr(13)
'                ValidateChequeNumber_80G = False
'                Exit Function
'            End If
'
'            If Not checkfieldspecialcharacter_Trans(ChequeNumber_80G(i)) Then
'                MsgBox_80GA = MsgBox_80GA + "* ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of Contribution transaction at Sr.No " & i & " in Schedule 80GGC is invalid, Only "" / "" and "" - "" special characters are allowed." & Chr(13)
'                ValidateChequeNumber_80G = False
'                Exit Function
'            End If
'
'            If Len(BankIFSC_80G(i)) > 11 Then
'                MsgBox_80GA = MsgBox_80GA + "* IFS Code at Sr.No " & i & " in Schedule 80G cannot exceed 11 characters" & Chr(13)
'                ValidateChequeNumber_80G = False
'                Exit Function
'            End If
'
'            If Not ValidateIFSCList(UCase(BankIFSC_80G(i))) Then
'                MsgBox_80GA = MsgBox_80GA + "* Invalid IFS Code at Sr.No " & i & " in Schedule 80G. Refer to your bank for valid IFS Codes." & Chr(13) & "(1st 4 Alphabets, followed by by Zero and remaining 6 should be alphanumeric)" & Chr(13)
'                ValidateChequeNumber_80G = False
'                Exit Function
'            End If
'           End If
'     End If
'     Sheet4.Activate
'
'        UpdateProgressBar
'    Next
'End Function
'



'Ankita_03/04/2026
Function ValidatePAN_Comb90() As Boolean

    errmsg80G = ""
    ValidatePAN_Comb90 = True

    Dim i As Long
    Dim comb90Val As Long
    Dim panVal As String

    'Loop through PAN column (assumed column J = 10)
    For i = 1 To Sheet4.Range("J:J").Rows.count

        panVal = Trim(Sheet4.Cells(i, 10).Value)  'PAN column
        comb90Val = Sheet4.Cells(i, "AA").Value    'comb_90 column

        If panVal <> "" Then
            If comb90Val > 0 Then
                fmsgbox "* Same PAN cannot be entered in other block." & Chr(13)
                ValidatePAN_Comb90 = False
            End If
        End If

    Next i

End Function

'Ankita_03/04/2026
Function ValidatePAN_CashRestriction() As Boolean

    Dim panCell As Range
    Dim panRanges As Range
    Dim cntOnes As Long
    Dim panVal As String

    errmsg80G = ""
    ValidatePAN_CashRestriction = True

    '-----------------------------------
    ' Combine all PAN named ranges
    '-----------------------------------
    Set panRanges = Union( _
        Sheet4.Range("Per10080G.DoneePAN"), _
        Sheet4.Range("PerNO5080G.DoneePAN"), _
        Sheet4.Range("Per5080G.DoneePAN"), _
        Sheet4.Range("PerYES10080G.DoneePAN") _
    )

    '-----------------------------------
    ' Check helper column Z
    '-----------------------------------
    cntOnes = Application.CountIf(Sheet4.Range("Z:Z"), 1)
    If cntOnes <= 1 Then Exit Function

    '-----------------------------------
    ' Loop through PAN cells
    '-----------------------------------
    For Each panCell In panRanges.Cells

        panVal = Trim(panCell.Value)

        If panVal <> "" Then
            fmsgbox "* Same PAN cannot be entered for ""Donation in cash""" & Chr(13)
            ValidatePAN_CashRestriction = False
            Exit Function

        End If
    Next panCell

End Function


