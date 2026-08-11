Attribute VB_Name = "Sch80GGA"

Option Explicit

Public MsgBox_80GGA As String

Public RelevantClauseClaimed_80GGA As Variant
Public Name_of_Donee_80GGA As Variant
Public Address_80GGA As Variant
Public City_Town_District_80GGA As Variant
Public State_Code_80GGA As Variant
Public Pincode_80GGA As Variant
Public PAN_of_donee_80GGA As Variant
Public Donation_cash_80GGA As Variant
Public Donation_other_80GGA As Variant
Public Donation_total_80GGA As Variant
Public Donation_Eligible_80GGA As Variant
Public Date_Donation, Total_DonationInCash_80GGA, Total_DonationInOThMode_80GGA, Total_TotalDonation_80GGA As Variant

Public rngname_80GGA As Variant
Public Total_Donation_Eligible_80GGA As Variant

Public end80GGA, end80GGA1, end80GGA2, end80GGA3, end80GGA4, end80GGA5, end80GGA6, end80GGA7, end80GGA8 As Long
Sub ValidateSheet80GGA_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2026-27"                'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
Validate80GGA
'MsgBox "Sheet 80GGA is OK", vbOKOnly, vbMessgaeCaption
fmsgboxoK "Sheet 80GGA is OK"
End Sub
 Sub Prev80GGA_Click()
Sheet4.Activate
End Sub
Sub Next_80GGAClick()
Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("80GGC")
    sourceSheet.Activate
End Sub


Sub Validate80GGA()
    If Not Validate80GGA_1 Then
        Sheet12.Activate
        'MsgBox MsgBox_80GGA, vbOKOnly, "Error(s)"
        fmsgbox (MsgBox_80GGA)
        CloseMsg
    End If
End Sub

Sub setTableInfo80GGA()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet12.Range("RelevantClauseClaimed_80GGA").Cells
    mIntCells = Sheet12.Range("RelevantClauseClaimed_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        ccount = ccount + 1
        End If
    Next
    end80GGA = ccount
    rngname_80GGA = "RelevantClauseClaimed_80GGA;Name_of_Donee_80GGA;Address_80GGA;City_Town_District_80GGA;State_Code_80GGA;Pincode_80GGA;PAN_of_donee_80GGA;Donation_cash_80GGA;Donation_other_80GGA;Donation_total_80GGA;Donation_Eligible_80GGA;"
End Sub

Sub setTableInfo80GGA1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet12.Range("Name_of_Donee_80GGA").Cells
    mIntCells = Sheet12.Range("Name_of_Donee_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA1 = ccount
End Sub

Sub setTableInfo80GGA2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet12.Range("Address_80GGA").Cells
    mIntCells = Sheet12.Range("Address_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA2 = ccount
End Sub

Sub setTableInfo80GGA3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet12.Range("City_Town_District_80GGA").Cells
    mIntCells = Sheet12.Range("City_Town_District_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA3 = ccount
End Sub

Sub setTableInfo80GGA4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet12.Range("State_Code_80GGA").Cells
    mIntCells = Sheet12.Range("State_Code_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        ccount = ccount + 1
        End If
    Next
    end80GGA4 = ccount
End Sub

Sub setTableInfo80GGA5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet12.Range("Pincode_80GGA").Cells
    mIntCells = Sheet12.Range("Pincode_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA5 = ccount
End Sub

Sub setTableInfo80GGA6()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet12.Range("PAN_of_donee_80GGA").Cells
    mIntCells = Sheet12.Range("PAN_of_donee_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA6 = ccount
End Sub

Sub setTableInfo80GGA7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet12.Range("Donation_cash_80GGA").Cells
    mIntCells = Sheet12.Range("Donation_cash_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA7 = ccount
End Sub

Sub setTableInfo80GGA8()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet12.Range("Donation_other_80GGA").Cells
    mIntCells = Sheet12.Range("Donation_other_80GGA").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80GGA8 = ccount
End Sub

Function ValidateRelevantClauseClaimed_80GGA() As Boolean
    ValidateRelevantClauseClaimed_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet12.Range("RelevantClauseClaimed_80GGA").Cells
    Dim i As Long
    ReDim RelevantClauseClaimed_80GGA(end80GGA)
    
    subProcCaption = "Validating 80GGA"
    noOfProcessSub = end80GGA
    
    For i = 1 To end80GGA
        RelevantClauseClaimed_80GGA(i) = rangecells.item(i).Value
        
        If isdropdownblank(RelevantClauseClaimed_80GGA(i)) Then
            'MsgBox_80GGA = MsgBox_80GGA + "* Please select relevant clause at Sr. No." & i & " from Drop-Down in Schedule 80GGA" & Chr(13)
            MsgBox_80GGA = MsgBox_80GGA + "* Please select relevant clause from Drop-Down at Sr.No. " & i & " in Schedule 80GGA" & Chr(13)  'Ankita_29/05/2025
            ValidateRelevantClauseClaimed_80GGA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateName_of_Donee_80GGA() As Boolean
    ValidateName_of_Donee_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet12.Range("Name_of_Donee_80GGA").Cells
    Dim i As Long
    ReDim Name_of_Donee_80GGA(end80GGA)
    
    For i = 1 To end80GGA
        Name_of_Donee_80GGA(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Name_of_Donee_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter name of donee at Sr. No " & i & " in Schedule 80GGA" & Chr(13)
            ValidateName_of_Donee_80GGA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateAddress_80GGA() As Boolean
    ValidateAddress_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet12.Range("Address_80GGA").Cells
    Dim i As Long
    ReDim Address_80GGA(end80GGA)
    
    For i = 1 To end80GGA
        Address_80GGA(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Address_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter address of donee at Sr. No " & i & " in Schedule 80GGA" & Chr(13)
            ValidateAddress_80GGA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateCity_Town_District_80GGA() As Boolean
    ValidateCity_Town_District_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet12.Range("City_Town_District_80GGA").Cells
    Dim i As Long
    ReDim City_Town_District_80GGA(end80GGA)
    
    For i = 1 To end80GGA
        City_Town_District_80GGA(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(City_Town_District_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter city/town/district of donee at Sr. No " & i & " in Schedule 80GGA" & Chr(13)
            ValidateCity_Town_District_80GGA = False
            Exit Function
        End If
        UpdateProgressBar
    Next
End Function

Function ValidateState_Code_80GGA() As Boolean
ValidateState_Code_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet12.Range("State_Code_80GGA").Cells
    'malli
    Set rangecells1 = Sheet12.Range("Pincode_80GGA").Cells 'Newly
    Dim i As Long
    ReDim State_Code_80GGA(end80GGA)
    For i = 1 To end80GGA
        State_Code_80GGA(i) = rangecells.item(i).Value
       ' Pincode_80GGA(i) = rangecells1.item(i).Value 'Newly
        'mmmm
  
          'ramya1234---------------------------------------------------------------
          
       If State_Code_80GGA(i) <> "" Then
          Sheets("80GGA").Activate
         Dim PIN_targetadd, state_targetadd As String
      'state_targetadd = Target.address
     state_targetadd = Replace(rangecells.item(i).Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "I", "J")
         If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet12.Range(Replace(state_targetadd, "I", "J")).Value = "" 'State_Code_80GGANew(i) = ""
    End If
    'ramya1234===================================================================================
    
        
        If ((State_Code_80GGA(i) = "(Select)") Or (State_Code_80GGA(i) = "")) Then
          MsgBox_80GGA = MsgBox_80GGA + "* Please select state of donee at Sr. No " & i & " in Schedule 80GGA" & Chr(13)
          ValidateState_Code_80GGA = False
          Exit Function
        End If
    Next
End Function

Function ValidatePincode_80GGA() As Boolean
ValidatePincode_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet12.Range("Pincode_80GGA").Cells
    'malli
    Set rangecells1 = Sheet12.Range("State_Code_80GGA").Cells 'Newly
    Dim i, x  As Long
    ReDim Pincode_80GGA(end80GGA)
    For i = 1 To end80GGA
        Pincode_80GGA(i) = rangecells.item(i).Value
        'malli
        
        
        State_Code_80GGA(i) = rangecells1.item(i).Value 'Newly
          'Malli--------------------
          
          
    If Pincode_80GGA(i) <> "" Then
         Dim PIN_targetadd, state_targetadd As String
      PIN_targetadd = Replace(rangecells.item(i).Address, "$", "")
      state_targetadd = Replace(PIN_targetadd, "J", "I")
     ' PIN_targetadd = Replace(state_targetadd, "I", "J")
         If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet12.Range(Replace(state_targetadd, "J", "I")).Value = "" 'State_Code_80GGANew(i) = ""
    End If

        '---------------------------
        
        If Not chkCompulsory(Pincode_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter pin code of donee at Sr. No " & i & "" & Chr(13)
          
            ValidatePincode_80GGA = False
            Exit Function
        End If
            
        If Mid(Pincode_80GGA(i), 1, 1) = 0 Then
            MsgBox_80GGA = MsgBox_80GGA + "* Pin code of donee at Sr. No " & i & " must be 6 digits and cannot begin with '0' in Schedule 80GGA." & Chr(13)
            ValidatePincode_80GGA = False
            Exit Function
        End If

        For x = 1 To Len(Pincode_80GGA(i))
            If Not IsNumeric(Mid(Pincode_80GGA(i), x, 1)) Then
                 ' MsgBox_80GGA = MsgBox_80GGA + "* Pin code of donee at Sr. No " & i & " must contain only digits from 0 to 9 in Schedule 80GGA." & Chr(13)
                   'Added by ramya
                  MsgBox_80GGA = MsgBox_80GGA + "* Invalid Pin code." & Chr(13)
                ValidatePincode_80GGA = False
                Exit Function
            End If
        Next
        
        If Len(Pincode_80GGA(i)) > 6 Then
            MsgBox_80GGA = MsgBox_80GGA + "* Pin code of donee at Sr. No " & i & " must be 6 digits in Schedule 80GGA." & Chr(13)
            ValidatePincode_80GGA = False
            Exit Function
        End If
        
        If Not checkfieldspecialcharacter(Pincode_80GGA(i)) Then
             MsgBox_80GGA = MsgBox_80GGA + "* Pin code of donee at Sr. No  " & i & "cannot Contain Special Characters in Schedule 80GGA." & Chr(13)
             ValidatePincode_80GGA = False
             Exit Function
        End If
    Next
End Function

Function ValidatePAN_of_donee_80GGA() As Boolean
ValidatePAN_of_donee_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet12.Range("PAN_of_donee_80GGA").Cells
    Dim i As Long
    ReDim PAN_of_donee_80GGA(end80GGA)
    For i = 1 To end80GGA
    PAN_of_donee_80GGA(i) = UCase(rangecells.item(i).Value)
        If Not chkCompulsory(PAN_of_donee_80GGA(i)) Then
            MsgBox_80GGA = MsgBox_80GGA + "* Please enter PAN of donee at Sr. No " & i & " in Schedule 80GGA." & Chr(13)
            ValidatePAN_of_donee_80GGA = False
            Exit Function
        End If
        If Not checkfieldspecialcharacter(PAN_of_donee_80GGA(i)) Then
             MsgBox_80GGA = MsgBox_80GGA + "* PAN of donee at Sr. No  " & i & " cannot Contain Special Characters in Schedule 80GGA." & Chr(13)
             ValidatePAN_of_donee_80GGA = False
             Exit Function
        End If
        
        If Not CheckDoneePAN(UCase(PAN_of_donee_80GGA(i))) Then
            ' MsgBox_80GGA = MsgBox_80GGA + "* PAN of donee at Sr. No  " & i & " is Invalid. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in Schedule 80GGA." & Chr(13)
            ' Added by ramya
            MsgBox_80GGA = MsgBox_80GGA + "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet." & Chr(13)
             ValidatePAN_of_donee_80GGA = False
             Exit Function
        End If
        
        If ((UCase(PAN_of_donee_80GGA(i)) = UCase(Sheet1.Range("sheet1.PAN").Value)) Or (UCase(PAN_of_donee_80GGA(i)) = UCase(Sheet3.Range("Ver.PAN").Value))) Then
            'Ankita_29/05/2025
            MsgBox_80GGA = MsgBox_80GGA + "* PAN of donee at Sr. No  " & i & " is Invalid." & Chr(13) & " Donee PAN cannot be same as assessee PAN or verification PAN in Schedule 80GGA." & Chr(13)
            ValidatePAN_of_donee_80GGA = False
            Exit Function
        End If
    Next
End Function

Function ValidateDonationAmt_80GGA() As Boolean
ValidateDonationAmt_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Dim rangecells1 As Range
    Set rangecells = Sheet12.Range("Donation_cash_80GGA").Cells
    Set rangecells1 = Sheet12.Range("Donation_other_80GGA").Cells
    Dim i As Long
    ReDim Donation_cash_80GGA(end80GGA)
    ReDim Donation_other_80GGA(end80GGA)
    
    For i = 1 To end80GGA
        Donation_cash_80GGA(i) = rangecells.item(i).Value
        Donation_other_80GGA(i) = rangecells1.item(i).Value
        
        If Not chkCompulsory(Donation_cash_80GGA(i)) And Not chkCompulsory(Donation_other_80GGA(i)) Then
           ' MsgBox_80GGA = MsgBox_80GGA + "* Donation Amount at Sr. No " & i & " is Mandatory in Schedule 80GGA." & Chr(13)
           ' MsgBox_80GGA = MsgBox_80GGA + "*Donation in cash and Donation in other mode is not filled" & Chr(13)
           
           'Added by Ramya
           MsgBox_80GGA = MsgBox_80GGA + "*Enter the amount of donation either in field ""Donation in cash"" or ""Donation in other mode"" at Sr. No " & i & "." & Chr(13)

            ValidateDonationAmt_80GGA = False
            Exit Function
        End If
    Next
End Function

Function ValidateDonation_total_80GGA() As Boolean
ValidateDonation_total_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet12.Range("Donation_total_80GGA").Cells
    Dim i As Long
    ReDim Donation_total_80GGA(end80GGA)
    
    For i = 1 To end80GGA
        Donation_total_80GGA(i) = rangecells.item(i).Value
        If Len(Donation_total_80GGA(i)) > 14 Then
            MsgBox_80GGA = MsgBox_80GGA + "* Total donation Amount at Sr. No " & i & " cannot exceed 14 digits in Schedule 80GGA." & Chr(13)
            ValidateDonation_total_80GGA = False
            Exit Function
        End If
    Next
End Function

Function ValidateDonation_Eligible_80GGA() As Boolean
ValidateDonation_Eligible_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet12.Range("Donation_Eligible_80GGA").Cells
    Dim i As Long
    ReDim Donation_Eligible_80GGA(end80GGA)
    
    For i = 1 To end80GGA
        Donation_Eligible_80GGA(i) = rangecells.item(i).Value
        If Len(Donation_Eligible_80GGA(i)) > 14 Then
            MsgBox_80GGA = MsgBox_80GGA + "* Eligible donation Amount at Sr. No " & i & " cannot exceed 14 digits in Schedule 80GGA." & Chr(13)
            ValidateDonation_Eligible_80GGA = False
            Exit Function
        End If
    Next
End Function

Function ValidateTotal_Donation_Eligible_80GGA() As Boolean
 ValidateTotal_Donation_Eligible_80GGA = True
 Total_Donation_Eligible_80GGA = Sheet12.Range("Total_Donation_Eligible_80GGA").Value
 
 If Len(Total_Donation_Eligible_80GGA) > 14 Then
    MsgBox_80GGA = MsgBox_80GGA + "* Total donation amount cannot be greater than 14 digits in Schedule 80GGA." & Chr(13)
    ValidateTotal_Donation_Eligible_80GGA = False
    Exit Function
End If

End Function
Function ValidateDate_Of_Donation_Cash_80GGA() As Boolean
    ValidateDate_Of_Donation_Cash_80GGA = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet12.Range("Date_Donation").Cells
    Dim rangecells1 As Range
    Set rangecells1 = Sheet12.Range("Donation_cash_80GGA").Cells
    Dim i As Long
    ReDim Date_Donation(end80GGA)
    ReDim Donation_cash_80GGA(end80GGA)
    
    For i = 1 To end80GGA
        Date_Donation(i) = rangecells.item(i).Value
        Donation_cash_80GGA(i) = rangecells1.item(i).Value
        
'        If Not chkCompulsory(Date_Donation(i)) And chkCompulsory(Donation_cash_80GGA(i)) Then
'            MsgBox_80GGA = MsgBox_80GGA + "* Please enter Date of Donation in Cash at Sr. No " & i & " in Schedule 80GGA" & Chr(13)
'            ValidateDate_Of_Donation_Cash_80GGA = False
'            Exit Function
'        End If
        UpdateProgressBar
    Next
End Function

Function Validate80GGA_1() As Boolean
    Validate80GGA_1 = True
    
    MsgBox_80GGA = "Schedule 80GGA : " & Chr(10)
        setTableInfo80GGA
        setTableInfo80GGA1
        setTableInfo80GGA2
        setTableInfo80GGA3
        setTableInfo80GGA4
        setTableInfo80GGA5
        setTableInfo80GGA6
        setTableInfo80GGA7
        setTableInfo80GGA8
        
    end80GGA = WorksheetFunction.Max(0, end80GGA, end80GGA1, end80GGA2, end80GGA3, end80GGA4, end80GGA5, end80GGA6, end80GGA7, end80GGA8)


    
    If Not ValidateRelevantClauseClaimed_80GGA Then Validate80GGA_1 = False
        If Not isdropdownblank(Sheet12.Range("RelevantClauseClaimed_80GGA").item(1).Value) Then
            If Not ValidateName_of_Donee_80GGA Then Validate80GGA_1 = False
            If Not ValidateAddress_80GGA Then Validate80GGA_1 = False
            If Not ValidateCity_Town_District_80GGA Then Validate80GGA_1 = False
            If Not ValidateState_Code_80GGA Then Validate80GGA_1 = False
            If Not ValidatePincode_80GGA Then Validate80GGA_1 = False
            If Not ValidatePAN_of_donee_80GGA Then Validate80GGA_1 = False
            If Not ValidateDonationAmt_80GGA Then Validate80GGA_1 = False
            If Not ValidateDonation_total_80GGA Then Validate80GGA_1 = False
            If Not ValidateDonation_Eligible_80GGA Then Validate80GGA_1 = False
            If Not ValidateTotal_Donation_Eligible_80GGA Then Validate80GGA_1 = False
            If Not ValidateDate_Of_Donation_Cash_80GGA Then Validate80GGA_1 = False
            'Change.28.02.2023.102.IDS.43/44/45A
            If Not ValidateTotal_Donation_InCash_80GGA Then Validate80GGA_1 = False
            If Not ValidateTotal_Donation_OtherMode_80GGA Then Validate80GGA_1 = False
            If Not ValidateTotalof_Total_Donation_80GGA Then Validate80GGA_1 = False
            'End Change IDS.43/44/45A
        End If
        
'
'        If ((end80GGA <> end80GGA1) Or (end80GGA <> end80GGA2) Or (end80GGA <> end80GGA3) Or _
'            (end80GGA <> end80GGA4) Or (end80GGA <> end80GGA5) Or (end80GGA <> end80GGA6)) Then
'        MsgBox_80GGA = MsgBox_80GGA + "Enter All Mandatory Fields."
'        Validate80GGA_1 = False
'        End If
End Function

Sub AddRows80GGA()
Dim vRows  As Long
Sheets("80GGA").Activate
EfilingCommon.DefinedgridNameRange = "RelevantClauseClaimed_80GGA||Name_of_Donee_80GGA||Address_80GGA||City_Town_District_80GGA||State_Code_80GGA||Pincode_80GGA||PAN_of_donee_80GGA||Donation_cash_80GGA||Donation_other_80GGA||Donation_total_80GGA||Donation_Eligible_80GGA"
ActiveCellRange = EfilingCommon.searchLastRow("RelevantClauseClaimed_80GGA")
vRows = EfilingCommon.insertRowUnderSectionWithFormula
End Sub
'Change.28.02.2023.102.IDS.43/44/45
Function ValidateTotal_Donation_InCash_80GGA() As Boolean
 ValidateTotal_Donation_InCash_80GGA = True
 Total_DonationInCash_80GGA = Sheet12.Range("Total_DonationInCash_80GGA").Value
 
 If Len(Total_DonationInCash_80GGA) > 14 Then
    MsgBox_80GGA = MsgBox_80GGA + "* Total amount of Donation in Cash cannot be greater than 14 digits in Schedule 80GGA." & Chr(13)
    ValidateTotal_Donation_InCash_80GGA = False
    Exit Function
End If

End Function
Function ValidateTotal_Donation_OtherMode_80GGA() As Boolean
 ValidateTotal_Donation_OtherMode_80GGA = True
 Total_DonationInOThMode_80GGA = Sheet12.Range("Total_DonationInOtherMode_80GGA").Value
 
 If Len(Total_DonationInOThMode_80GGA) > 14 Then
    MsgBox_80GGA = MsgBox_80GGA + "* Total amount of Donation in Other Mode cannot be greater than 14 digits in Schedule 80GGA." & Chr(13)
    ValidateTotal_Donation_OtherMode_80GGA = False
    Exit Function
End If

End Function
Function ValidateTotalof_Total_Donation_80GGA() As Boolean
 ValidateTotalof_Total_Donation_80GGA = True
 Total_TotalDonation_80GGA = Sheet12.Range("Total_Donation_80GGA").Value
 
 If Len(Total_TotalDonation_80GGA) > 14 Then
    MsgBox_80GGA = MsgBox_80GGA + "* Total amount Total Donation cannot be greater than 14 digits in Schedule 80GGA." & Chr(13)
    ValidateTotalof_Total_Donation_80GGA = False
    Exit Function
End If

End Function
'End Change IDS.43/44/45
