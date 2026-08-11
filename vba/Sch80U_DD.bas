Attribute VB_Name = "Sch80U_DD"

Option Explicit

Public MsgBox_80U, MsgBox_80DD As String

 
Public Nature_of_disability_80U As Variant
Public Amount_of_deduction_80U As Variant
Public Date_filingFm10IA_80U As Variant
Public AckNo_ofForm10IAfiled_80U As Variant
Public UDID_Num_80U As Variant
Public rngname_80U As Variant
Public end80U, end80U1, end80U2, end80U3, end80U4, end80U5 As Long

'malli 23/04/2024
Public end80DD8 As Variant
Public Disability_80DD_chk, Disability_80DD_chk2, AckNo_ofForm11A2filed_80DD, Disability_80U_chk, Disability_80U_chk2, AckNo_ofForm11A2filed_80U As Variant
'-----------------
'--------------------80DD
Public end80DD7, end80DD6, end80DD5, end80DD4, end80DD3, end80DD2, end80DD1, end80DD    As Long

Public Type_dependent_80DD As Variant
Public Aadhaar_dependent_80DD As Variant
Public PAN_dependent_80DD  As Variant
Public UDID_Num_80DD As Variant
Public AckNo_ofForm10IAfiled_80DD As Variant
Public Date_filingFm10IA_80DD As Variant
Public Amount_of_deduction_80DD As Variant
Public Nature_of_disability_80DD As Variant

Sub ValidateSheet80U_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2026-27"           'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
Validate80U
'MsgBox "Sheet 80U is OK", vbOKOnly, vbMessgaeCaption
fmsgboxoK "Sheet 80U is OK"
End Sub
Sub ValidateSheet80DD_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2026-27"           'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
Validate80DD
'MsgBox "Sheet 80U is OK", vbOKOnly, vbMessgaeCaption
fmsgboxoK "Sheet 80DD is OK"
End Sub
'---------------------------------------------------------------------------------------------
Sub Validate80U()
    If Not Validate80U_1 Then
        Sheet14.Activate
        'MsgBox MsgBox_80U, vbOKOnly, "Error(s)"
        fmsgbox (MsgBox_80U)
        CloseMsg
    End If
    
    'Malli------------
    If Sheet14.Range("AckNoFm10IAfiled_80U").Value <> "" And Sheet14.Range("AckNoFm10IAfiled_80DD").Value <> "" Then
    If Sheet14.Range("AckNoFm10IAfiled_80U").Value = Sheet14.Range("AckNoFm10IAfiled_80DD").Value Then
        MsgBox_80U = MsgBox_80U & "* Acknowledgement number of Form 10IA filed for self and Dependent can't be same. Please provide proper acknowledgement number" & Chr(13)
       fmsgbox MsgBox_80U
       CloseMsg
    End If
    End If
    
    '-----------------
    
End Sub



Function Validate80U_1() As Boolean
    Validate80U_1 = True
    MsgBox_80U = "Schedule 80U : " & Chr(10)
         
setTableInfo80U
'setTableInfo80U1
'setTableInfo80U2
setTableInfo80U3
setTableInfo80U4
'Malli--
setTableInfo80U5
'---------------


    end80U = WorksheetFunction.Max(0, end80U, end80U1, end80U3, end80U4, end80U5)

 If end80U > 0 Then

       If Not ValidateNature_disability_80U Then Validate80U_1 = False
       If Not ValidateType_disability_80U Then Validate80U_1 = False     'Ankita_18/04/2025
       If Not ValidateAmount_of_deduction_80U Then Validate80U_1 = False
'       If Not ValidateDate_of_filingofForm10IA_80U Then Validate80U_1 = False
       If Not ValidateAckNoFm10IAfiled_80U Then Validate80U_1 = False
       'Malli-----
     
'        If Not ValidateAckNoFm11A2filed_80U Then Validate80U_1 = False
       '--------------------------------------
       If Not ValidateUDIDNum_80U Then Validate80U_1 = False
End If
        
End Function


Sub setTableInfo80U()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Naturedisability_80U").Cells
    mIntCells = Sheet14.Range("Naturedisability_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) And Not UCase(rangecells.item(mIntCtr).Value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80U = ccount
    'rngname_80U = "Naturedisability_80U;Amtdeduction_80U;DatefilingFm10IA_80U;AckNoFm10IAfiled_80U;UDIDNum_80U;"
End Sub
Sub setTableInfo80U1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    'Set rangecells = Sheet14.Range("Amtdeduction_80U").Cells
    Set rangecells = Sheet14.Range("Amtdeduction_80DD").Cells
    mIntCells = Sheet14.Range("Amtdeduction_80U").count
    MsgBox (rangecells.item(mIntCtr).Value)
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value <> 0 Then
        ccount = ccount + 1
        End If
    Next
    end80U1 = ccount
End Sub

'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Sub setTableInfo80U2()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet14.Range("DatefilingFm10IA_80U").Cells
'    mIntCells = Sheet14.Range("DatefilingFm10IA_80U").count
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).Value = "" Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80U2 = ccount
'End Sub

Sub setTableInfo80U3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("AckNoFm10IAfiled_80U").Cells
    mIntCells = Sheet14.Range("AckNoFm10IAfiled_80U").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80U3 = ccount
End Sub

Sub setTableInfo80U4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("UDIDNum_80U").Cells
    mIntCells = Sheet14.Range("UDIDNum_80U").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        ccount = ccount + 1
        End If
    Next
    end80U4 = ccount
End Sub

'Ankita_19/04/2025
'Disability_80U
Sub setTableInfo80U5()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Disability_80U").Cells
    mIntCells = Sheet14.Range("Disability_80U").count
    For mIntCtr = 1 To mIntCells
        'If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        If Not rangecells.item(mIntCtr).Value = "" And Not rangecells.item(mIntCtr).Value = "(Select)" Then
        ccount = ccount + 1
        End If
    Next
    end80U5 = ccount
End Sub


Function ValidateNature_disability_80U() As Boolean
   ValidateNature_disability_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Naturedisability_80U").Cells
    Dim i As Long
    ReDim Nature_of_disability_80U(end80U)
    
    For i = 1 To end80U

        Nature_of_disability_80U(i) = rangecells.item(i).Value
        
       If isdropdownblank(Nature_of_disability_80U(i)) Or UCase(Nature_of_disability_80U(i)) = UCase("Select") Then
            'Ankita_29/05/2025
            MsgBox_80U = MsgBox_80U + "* Please select one of the dropdown in 'Nature of disability' in Schedule 80U" & Chr(13)
            ValidateNature_disability_80U = False
            Exit Function
            UpdateProgressBar
       End If
    Next
End Function
 
Function ValidateAmount_of_deduction_80U() As Boolean
    ValidateAmount_of_deduction_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Amtdeduction_80U").Cells
    Dim i As Long
    ReDim Amount_of_deduction_80U(end80U)
    
    For i = 1 To end80U
        Amount_of_deduction_80U(i) = rangecells.item(i).Value
        
        If Not chkCompulsory(Amount_of_deduction_80U(i)) Then
            MsgBox_80U = MsgBox_80U + "* Please enter Amount of deduction at Sr. No " & i & " in Schedule 80U" & Chr(13)
            ValidateAmount_of_deduction_80U = False
            Exit Function
        End If
        UpdateProgressBar

        If Not IsNumeric(Amount_of_deduction_80U(i)) Then
           MsgBox_80U = MsgBox_80U + "* Amount of deduction at Sr. No  " & i & "  in  Schedule 80U should be Numeric value" & Chr(13)
            ValidateAmount_of_deduction_80U = False
            Exit Function
        End If
        
        If Amount_of_deduction_80U(i) > 99999999999999# Then
           MsgBox_80U = MsgBox_80U + "* Amount of deduction at Sr. No  " & i & "  in Schedule 80U cannot exceed 14 digits" & Chr(13)
           ValidateAmount_of_deduction_80U = False
            Exit Function
        End If
Next

End Function

 
'Function ValidateDate_of_filingofForm10IA_80U() As Boolean
'
'ValidateDate_of_filingofForm10IA_80U = True
''    setTableInfo80GGA
'    Dim rangecells, rangecells1 As Range
''    Set rangecells = Sheet14.Range("DatefilingFm10IA_80U").Cells
'    Set rangecells1 = Sheet14.Range("Disability_80U").Cells
'
'
'    Dim i As Long
''    ReDim Date_filingFm10IA_80U(end80U)
'    ReDim Disability_80U_chk(end80U)
'
'    For i = 1 To end80U
''        Date_filingFm10IA_80U(i) = rangecells.item(i).Value
'        Disability_80U_chk(i) = rangecells1.item(i).Value
'
'
'If Date_filingFm10IA_80U(i) <> "" Then
'
'    If Not CheckDateddmmyyyy(Date_filingFm10IA_80U(i)) Then
'        ValidateDate_of_filingofForm10IA_80U = False
'        MsgBox_80U = MsgBox_80U + "* Date of filing of Form 10IA must be a valid dd/mm/yyyy format at Sr. No " & i & " in Schedule 80U" & Chr(13)
'        Exit Function
'    End If
'        UpdateProgressBar
'
'    If Len(Date_filingFm10IA_80U(i)) > 10 Then
'        ValidateDate_of_filingofForm10IA_80U = False
'        MsgBox_80U = MsgBox_80U + "* Date of filing of Form 10IA cannot exceed 10 digits at Sr. No " & i & " in Schedule 80U" & Chr(13)
'        Exit Function
'    End If
'Else
'
'If Disability_80U_chk(i) = "(viii) autism" Or Disability_80U_chk(i) = "(ix) cerebral palsy" Or Disability_80U_chk(i) = "(x) multiple disability" Then
'  If Date_filingFm10IA_80U(i) = "" Then
'        ValidateDate_of_filingofForm10IA_80U = False
'        MsgBox_80U = MsgBox_80U + "* Please provide date of filing of Form 10IA in schedule 80U at Sr. No " & i & "" & Chr(13)
'        Exit Function
'  End If
'End If
'End If
'
'   'Ankita_30/04/2025
'       If Not ChkMinInclusiveDate(Trim(Dformat(Date_filingFm10IA_80U(i), "yyyy-mm-dd")), "2025-04-01") Then
'           MsgBox_80U = MsgBox_80U + "* Date of filing of Form 10IA cannot before 01/04/2025 in schedule 80U" & Chr(13)
'            ValidateDate_of_filingofForm10IA_80U = False
'          Exit Function
'        End If
'Next
'End Function
'
 

Function ValidateAckNoFm10IAfiled_80U() As Boolean
    
ValidateAckNoFm10IAfiled_80U = True
'    setTableInfo80GGA
    Dim rangecells, rangecells1 As Range
    Set rangecells = Sheet14.Range("AckNoFm10IAfiled_80U").Cells
    Set rangecells1 = Sheet14.Range("Disability_80U").Cells
    
    Dim i As Long
    ReDim AckNo_ofForm10IAfiled_80U(end80U)
    ReDim Disability_80U_chk2(end80U)
    
    For i = 1 To end80U
        AckNo_ofForm10IAfiled_80U(i) = rangecells.item(i).Value
        Disability_80U_chk2(i) = rangecells1.item(i).Value
        
   If AckNo_ofForm10IAfiled_80U(i) <> "" Then
      If Not checkfieldspecialcharacter(AckNo_ofForm10IAfiled_80U(i)) Then
             MsgBox_80U = MsgBox_80U + "* Ack. No. of Form 10IA filed at Sr. No  " & i & " cannot contain special characters in Schedule 80U." & Chr(13)
             ValidateAckNoFm10IAfiled_80U = False
             Exit Function
        End If
     
        UpdateProgressBar

    If Len(AckNo_ofForm10IAfiled_80U(i)) > 15 Then
        ValidateAckNoFm10IAfiled_80U = False
        MsgBox_80U = MsgBox_80U + "*  Ack. No. of Form 10IA filed cannot exceed 15 digits at Sr. No " & i & " in Schedule 80U" & Chr(13)
        Exit Function
    End If
     
   Else
   
   'Ankita_07/05/2025
'            If Disability_80U_chk2(i) = "(viii) autism" Or Disability_80U_chk2(i) = "(ix) cerebral palsy" Or Disability_80U_chk2(i) = "(x) multiple disability" Then
'            If Disability_80U_chk2(i) = "(i) autism, cerebral palsy, or multiple disabilities and" Then
             If Disability_80U_chk2(i) = "(i) autism, cerebral palsy, or multiple disabilities" Then     'Ankita_14/05/2025

            If AckNo_ofForm10IAfiled_80U(i) = "" Then
                 ValidateAckNoFm10IAfiled_80U = False
                 MsgBox_80U = MsgBox_80U + "*  Please provide Acknowledgement of Form 10IA in schedule 80U at Sr. No " & i & "" & Chr(13)
                 Exit Function
            End If
            End If
   End If
Next
End Function
 
Function ValidateUDIDNum_80U() As Boolean
    
ValidateUDIDNum_80U = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("UDIDNum_80U").Cells
    Dim i As Long
    ReDim UDID_Num_80U(end80U)
    
    For i = 1 To end80U
        UDID_Num_80U(i) = rangecells.item(i).Value
     
      If Not checkfieldspecialcharacterUDID(UDID_Num_80U(i)) Then  'Ankita_03/06/2025
             MsgBox_80U = MsgBox_80U + "* UDID Number at Sr. No " & i & " cannot contain special characters in Schedule 80U." & Chr(13)
             ValidateUDIDNum_80U = False
             Exit Function
        End If
     
        UpdateProgressBar

    'If Len(UDID_Num_80U(i)) > 15 Then
    If Len(UDID_Num_80U(i)) > 18 Then
        ValidateUDIDNum_80U = False
        MsgBox_80U = MsgBox_80U + "*UDID Number cannot exceed 18 digits at Sr.No " & i & " in Schedule 80U" & Chr(13)
        Exit Function
    End If
     
Next
End Function

'__________________________________________________80DD
Sub setTableInfo80DD()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Naturedisability_80DD").Cells
    mIntCells = Sheet14.Range("Naturedisability_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) And Not UCase(rangecells.item(mIntCtr).Value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80DD = ccount
    'rngname_80GGA = "RelevantClauseClaimed_80GGA;Name_of_Donee_80GGA;Address_80GGA;City_Town_District_80GGA;State_Code_80GGA;Pincode_80GGA;PAN_of_donee_80GGA;Donation_cash_80GGA;Donation_other_80GGA;Donation_total_80GGA;Donation_Eligible_80GGA;"
End Sub

 
Sub setTableInfo80DD1()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Amtdeduction_80DD").Cells
    mIntCells = Sheet14.Range("Amtdeduction_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" And Not rangecells.item(mIntCtr).Value = "0" Then
        ccount = ccount + 1
        End If
    Next
    end80DD1 = ccount
End Sub

Sub setTableInfo80DD2()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Typedependent_80DD").Cells
    mIntCells = Sheet14.Range("Typedependent_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) And Not UCase(rangecells.item(mIntCtr).Value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    end80DD2 = ccount
End Sub

Sub setTableInfo80DD3()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("PANdependent_80DD").Cells
    mIntCells = Sheet14.Range("PANdependent_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
        ccount = ccount + 1
        End If
    Next
    end80DD3 = ccount
End Sub

Sub setTableInfo80DD4()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Aadhaardependent_80DD").Cells
    mIntCells = Sheet14.Range("Aadhaardependent_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD4 = ccount
End Sub

'Sub setTableInfo80DD5()
'    Dim rangecells As Range
'    Dim mIntCells  As Long
'    Dim mIntCtr  As Long
'    Dim ccount  As Long
'    ccount = 0
'    Set rangecells = Sheet14.Range("DatefilingFm10IA_80DD").Cells
'    mIntCells = Sheet14.Range("DatefilingFm10IA_80DD").count
'    For mIntCtr = 1 To mIntCells
'        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
'        ccount = ccount + 1
'        End If
'    Next
'    end80DD5 = ccount
'End Sub

Sub setTableInfo80DD6()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("AckNoFm10IAfiled_80DD").Cells
    mIntCells = Sheet14.Range("AckNoFm10IAfiled_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD6 = ccount
End Sub

Sub setTableInfo80DD7()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("UDIDNum_80DD").Cells
    mIntCells = Sheet14.Range("UDIDNum_80DD").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
        ccount = ccount + 1
        End If
    Next
    end80DD7 = ccount
End Sub

'Ankita_19/04/2025

Sub setTableInfo80DD8() 'TypeDisability80DD(1)
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet14.Range("Disability_80DD").Cells
    mIntCells = Sheet14.Range("Disability_80DD").count
    For mIntCtr = 1 To mIntCells
         'If Not isdropdownblank(rangecells.item(mIntCtr).Value) And Not UCase(rangecells.item(mIntCtr).Value) = UCase("Select") Then
         'Debug.Print rangecells.item(mIntCtr).Value
'          If rangecells.item(mIntCtr).Value <> "" And UCase(rangecells.item(mIntCtr).Value) <> UCase("(Select)") Then    'Malli
          If Not rangecells.item(mIntCtr).Value = "" And Not rangecells.item(mIntCtr).Value = "(Select)" Then
'         If Not isdropdownblankDD(rangecells.item(mIntCtr).Value) And Not UCase(rangecells.item(mIntCtr).Value) = UCase("Select") Then
          ccount = ccount + 1
        End If
    Next
    end80DD8 = ccount
'    rngname_80DD = "NatureDisability_80DD;AmtDeduction_80DD;Dependent_80DD;PanDependent_80DD;AadhaarDependent_80DD;Form10IA_80DD;AcknowledgeNum_80DD;UDIDNumber_80DD"
End Sub


'---------------------------------------------------------------------------------------------------------------------------------------------------

'-------------------------------------
Function ValidateNature_disability_80DD() As Boolean
   ValidateNature_disability_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Naturedisability_80DD").Cells
    Dim i As Long
    ReDim Nature_of_disability_80DD(end80DD)
    
    For i = 1 To end80DD

        Nature_of_disability_80DD(i) = rangecells.item(i).Value
        
        If isdropdownblank(Nature_of_disability_80DD(i)) Or UCase(Nature_of_disability_80DD(i)) = UCase("Select") Then
        'Ankita_29/05/2025
           MsgBox_80DD = MsgBox_80DD + "* Please select one of the dropdown in 'Nature of disability' in Schedule 80DD" & Chr(13)
           ValidateNature_disability_80DD = False
            Exit Function
         
        UpdateProgressBar
        
        End If
    Next
End Function

'Ankita_18/04/2025
Function ValidateType_disability_80DD() As Boolean
   ValidateType_disability_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Disability_80DD").Cells
    Dim i As Long
    ReDim Type_of_disability_80DD(end80DD)
    
    For i = 1 To end80DD

        Type_of_disability_80DD(i) = rangecells.item(i).Value
        
       ' If isdropdownblank(Type_of_disability_80DD(i)) Or UCase(Type_of_disability_80DD(i)) = UCase("Select") Then
        
         If Type_of_disability_80DD(i) = "" Or Type_of_disability_80DD(i) = "(Select)" Then  'Malli
           MsgBox_80DD = MsgBox_80DD + "*Selection of ""Type of disability"" in schedule 80DD is mandatory." & Chr(13)
           ValidateType_disability_80DD = False
            Exit Function
         
        UpdateProgressBar
        
        End If
    Next
End Function

'Ankita_18/04/2025

Function ValidateType_disability_80U() As Boolean
   ValidateType_disability_80U = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Disability_80U").Cells
    Dim i As Long
    ReDim Type_of_disability_80U(end80U)
    
    For i = 1 To end80U

        Type_of_disability_80U(i) = rangecells.item(i).Value
        
        'If isdropdownblank(Type_of_disability_80U(i)) Or UCase(Type_of_disability_80U(i)) = UCase("Select") Then
         'Malli-------
          If Type_of_disability_80U(i) = "" Or Type_of_disability_80U(i) = "(Select)" Then
           MsgBox_80U = MsgBox_80U + "* Selection of ""Type of disability"" in schedule 80U is mandatory." & Chr(13)
           ValidateType_disability_80U = False
            Exit Function
         
        UpdateProgressBar
        
        End If
    Next
End Function
'-----------------------------------------------------
Function ValidateAmount_of_deduction_80DD() As Boolean
    ValidateAmount_of_deduction_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Amtdeduction_80DD").Cells
    Dim i As Long
    ReDim Amount_of_deduction_80DD(end80DD)
    
    For i = 1 To end80DD
        
     Amount_of_deduction_80DD(i) = rangecells.item(i).Value
        
If Amount_of_deduction_80DD(i) <> 0 Then
        
If Not chkCompulsory(Amount_of_deduction_80DD(i)) Then
            MsgBox_80DD = MsgBox_80DD + "* Please enter Amount of deduction at Sr. No " & i & " in Schedule 80DD" & Chr(13)
            ValidateAmount_of_deduction_80DD = False
            Exit Function
        End If
        UpdateProgressBar

        If Not IsNumeric(Amount_of_deduction_80DD(i)) Then
           MsgBox_80DD = MsgBox_80DD + "* Amount of deduction at Sr. No  " & i & "  in  Schedule 80DD should be Numeric value" & Chr(13)
            ValidateAmount_of_deduction_80DD = False
            Exit Function
        End If
        
        If Amount_of_deduction_80DD(i) > 99999999999999# Then
           MsgBox_80DD = MsgBox_80DD + "* Amount of deduction at Sr. No  " & i & "  in Schedule 80DD cannot exceed 14 digits" & Chr(13)
           ValidateAmount_of_deduction_80DD = False
            Exit Function
        End If
        
   End If
Next

End Function

'--------------------------------------------------------------------------------------------------------------
'Function ValidateDate_of_filingofForm10IA_80DD() As Boolean
'
'ValidateDate_of_filingofForm10IA_80DD = True
'
'    Dim rangecells, rangecells1 As Range
'    Set rangecells = Sheet14.Range("DatefilingFm10IA_80DD").Cells
'    Set rangecells1 = Sheet14.Range("Disability_80DD").Cells
'
'    Dim i As Long
'    ReDim Date_filingFm10IA_80DD(end80DD)
'    ReDim Disability_80DD_chk(end80DD)
'
'    For i = 1 To end80DD
'        Date_filingFm10IA_80DD(i) = rangecells.item(i).Value
'        Disability_80DD_chk(i) = rangecells1.item(i).Value
'
'    'Malli-23/04/2025
'    If Date_filingFm10IA_80DD(i) <> "" Then
'
'    If Not CheckDateddmmyyyy(Date_filingFm10IA_80DD(i)) Then
'        ValidateDate_of_filingofForm10IA_80DD = False
'       MsgBox_80DD = MsgBox_80DD + "* Date of filing of Form 10IA must be a valid dd/mm/yyyy format at Sr. No " & i & " in Schedule 80DD" & Chr(13)
'        Exit Function
'    End If
'        UpdateProgressBar
'
'    If Len(Date_filingFm10IA_80DD(i)) > 10 Then
'        ValidateDate_of_filingofForm10IA_80DD = False
'        MsgBox_80DD = MsgBox_80DD + "* Date of filing of Form 10IA cannot exceed 10 digits at Sr. No " & i & " in Schedule 80DD" & Chr(13)
'        Exit Function
'    End If
'
'    Else
'      If Disability_80DD_chk(i) = "(viii) autism" Or Disability_80DD_chk(i) = "(ix) cerebral palsy" Or Disability_80DD_chk(i) = "(x) multiple disability" Then
'         If Date_filingFm10IA_80DD(i) = "" Then
'                ValidateDate_of_filingofForm10IA_80DD = False
'                MsgBox_80DD = MsgBox_80DD + "* Please provide date of filing of Form 10IA in schedule 80DD at Sr. No " & i & "" & Chr(13)
'                Exit Function
'         End If
'      End If
'    End If
'
'     'Ankita_30/04/2025
'      If Not ChkMinInclusiveDate(Trim(Dformat(Date_filingFm10IA_80DD(i), "yyyy-mm-dd")), "2025-04-01") Then
'           MsgBox_80DD = MsgBox_80DD + "* Date of filing of Form 10IA cannot before 01/04/2025 in schedule 80DD " & Chr(13)
'            ValidateDate_of_filingofForm10IA_80DD = False
'
'
'          Exit Function
'        End If
'Next
'End Function

'---------------------------------------------------------------------

Function ValidateAckNoFm10IAfiled_80DD() As Boolean
    
ValidateAckNoFm10IAfiled_80DD = True
 
    Dim rangecells, rangecells1 As Range
    Set rangecells = Sheet14.Range("AckNoFm10IAfiled_80DD").Cells
    Set rangecells1 = Sheet14.Range("Disability_80DD").Cells
    
    Dim i As Long
    ReDim AckNo_ofForm10IAfiled_80DD(end80DD)
    ReDim Disability_80DD_chk2(end80DD)
    
    For i = 1 To end80DD
        AckNo_ofForm10IAfiled_80DD(i) = rangecells.item(i).Value
         Disability_80DD_chk2(i) = rangecells1.item(i).Value
     'Malli------
     If AckNo_ofForm10IAfiled_80DD(i) <> "" Then
      If Not checkfieldspecialcharacter(AckNo_ofForm10IAfiled_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "* Ack. No. of Form 10IA filed at Sr. No  " & i & " cannot contain special characters in Schedule 80DD." & Chr(13)
             ValidateAckNoFm10IAfiled_80DD = False
             Exit Function
        End If
     
        UpdateProgressBar

    If Len(AckNo_ofForm10IAfiled_80DD(i)) > 15 Then
        ValidateAckNoFm10IAfiled_80DD = False
        MsgBox_80DD = MsgBox_80DD + "*  Ack. No. of Form 10IA filed cannot exceed 15 digits at Sr. No " & i & " in Schedule 80DD" & Chr(13)
        Exit Function
    End If
     Else
     
     'Ankita_07/05/2025
'        If Disability_80DD_chk2(i) = "(viii) autism" Or Disability_80DD_chk2(i) = "(ix) cerebral palsy" Or Disability_80DD_chk2(i) = "(x) multiple disability" Then
'        If Disability_80DD_chk2(i) = "(i) autism, cerebral palsy, or multiple disabilities and" Then
                If Disability_80DD_chk2(i) = "(i) autism, cerebral palsy, or multiple disabilities" Then  'Ankita_14/05/2025
         If AckNo_ofForm10IAfiled_80DD(i) = "" Then
                 ValidateAckNoFm10IAfiled_80DD = False
                 MsgBox_80DD = MsgBox_80DD + "* Please provide Acknowledgement of Form 10IA in schedule 80DD at Sr. No " & i & "" & Chr(13)
                 Exit Function
         End If
      End If
     
     End If
Next
End Function

'Malli----------------------
'Ankita_06/05/2025_Commented as per DESheet_v0.7
'Function ValidateAckNoFm11A2filed_80DD() As Boolean
'
'ValidateAckNoFm11A2filed_80DD = True
'
'    Dim rangecells, rangecells1 As Range
'    Set rangecells = Sheet14.Range("AcknowledgeNum11A2_80DD").Cells
'
'
'    Dim i As Long
'    ReDim AckNo_ofForm11A2filed_80DD(end80DD)
'
'
'    For i = 1 To end80DD
'        AckNo_ofForm11A2filed_80DD(i) = rangecells.item(i).Value
'
'
'     If AckNo_ofForm11A2filed_80DD(i) <> "" Then
'
'      If Not checkfieldspecialcharacter(AckNo_ofForm11A2filed_80DD(i)) Then
'             MsgBox_80DD = MsgBox_80DD + "* Ack no. of Form as per Rule 11A(2) filed at Sr. No  " & i & " cannot contain special characters in Schedule 80DD." & Chr(13)
'             ValidateAckNoFm11A2filed_80DD = False
'             Exit Function
'        End If
'
'        UpdateProgressBar
'
'    If Len(AckNo_ofForm11A2filed_80DD(i)) > 50 Then
'        ValidateAckNoFm11A2filed_80DD = False
'        MsgBox_80DD = MsgBox_80DD + "*  Ack no. of Form as per Rule 11A(2) filed cannot exceed 50 digits at Sr. No " & i & " in Schedule 80DD" & Chr(13)
'        Exit Function
'    End If
'
'
'     End If
'Next
'End Function

'Malli----------
'Function ValidateAckNoFm11A2filed_80U() As Boolean
'
'ValidateAckNoFm11A2filed_80U = True
'
'    Dim rangecells, rangecells1 As Range
'    Set rangecells = Sheet14.Range("AcknowledgeNum11A2_80U").Cells
'
'
'    Dim i As Long
'    ReDim AckNo_ofForm11A2filed_80U(end80U)
'
'
'    For i = 1 To end80U
'        AckNo_ofForm11A2filed_80U(i) = rangecells.item(i).Value
'
'
'     If AckNo_ofForm11A2filed_80U(i) <> "" Then
'
'      If Not checkfieldspecialcharacter(AckNo_ofForm11A2filed_80U(i)) Then
'            MsgBox_80U = MsgBox_80U + "* Ack no. of Form as per Rule 11A(2) filed at Sr. No  " & i & " cannot contain special characters in Schedule 80U." & Chr(13)
'             ValidateAckNoFm11A2filed_80U = False
'             Exit Function
'        End If
'
'        UpdateProgressBar
'
'    If Len(AckNo_ofForm11A2filed_80U(i)) > 50 Then
'        ValidateAckNoFm11A2filed_80U = False
'        MsgBox_80U = MsgBox_80U + "*  Ack no. of Form as per Rule 11A(2) filed cannot exceed 50 digits at Sr. No " & i & " in Schedule 80U" & Chr(13)
'        Exit Function
'    End If
'
'
'     End If
'Next
'End Function
'-----------------------------
'-----------------------------


'------------------------------------------------------------------------------------------------------------------------------
Function ValidateUDIDNum_80DD() As Boolean
    
ValidateUDIDNum_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("UDIDNum_80DD").Cells
    Dim i As Long
    ReDim UDID_Num_80DD(end80DD)
    
    For i = 1 To end80DD
        UDID_Num_80DD(i) = rangecells.item(i).Value
     
      If Not checkfieldspecialcharacterUDID(UDID_Num_80DD(i)) Then  'Ankita_03/06/2025
            MsgBox_80DD = MsgBox_80DD + "* UDID Number at Sr. No " & i & " cannot contain special characters in Schedule 80DD." & Chr(13)
             ValidateUDIDNum_80DD = False
             Exit Function
        End If
     
        UpdateProgressBar

    'If Len(UDID_Num_80DD(i)) > 15 Then 'Malli
    If Len(UDID_Num_80DD(i)) > 18 Then
        ValidateUDIDNum_80DD = False
        MsgBox_80DD = MsgBox_80DD + "*UDID Number cannot exceed 18 digits at Sr.No " & i & " in Schedule 80DD" & Chr(13)
        Exit Function
    End If
    
    
    
    Next
End Function

Function ValidatePANdependent_80DD() As Boolean
    
ValidatePANdependent_80DD = True
'    setTableInfo80GGA
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("PANdependent_80DD").Cells
    Dim i As Long
    ReDim PAN_dependent_80DD(end80DD)
    
    For i = 1 To end80DD
        PAN_dependent_80DD(i) = rangecells.item(i).Value
        
If Trim(PAN_dependent_80DD(i)) <> "" Then
     
      If Not checkfieldspecialcharacter(PAN_dependent_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "* PAN Number at Sr. No " & i & " cannot contain special characters in Schedule 80DD." & Chr(13)
            ValidatePANdependent_80DD = False
             Exit Function
        End If
     
        UpdateProgressBar

    If Not CheckDoneePAN_80DD(UCase(PAN_dependent_80DD(i))) Then
             MsgBox_80DD = MsgBox_80DD + "*""Invalid PAN in Schedule 80DD.""" & Chr(13)
             ValidatePANdependent_80DD = False
             Exit Function
        End If

'If PAN_dependent_80DD(i) <> o Then
        If ((UCase(PAN_dependent_80DD(i)) = UCase(Sheet1.Range("sheet1.PAN").Value)) Or (UCase(PAN_dependent_80DD(i)) = UCase(Sheet3.Range("Ver.PAN").Value))) Then
        MsgBox_80DD = MsgBox_80DD + "* ""PAN of the dependent cannot be same as assessee PAN in Part-A General Information.""" & Chr(13)
         ValidatePANdependent_80DD = False
        Exit Function
        End If
'End If
 
 End If
 
Next
End Function


Function ValidateAadhaardependent_80DD() As Boolean
    
ValidateAadhaardependent_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Aadhaardependent_80DD").Cells
    Dim i As Long
    ReDim Aadhaar_dependent_80DD(end80DD)
    
    For i = 1 To end80DD
        Aadhaar_dependent_80DD(i) = rangecells.item(i).Value
     

If Trim(Aadhaar_dependent_80DD(i)) <> "" Then


      If Not checkfieldspecialcharacter(Aadhaar_dependent_80DD(i)) Then
             MsgBox_80DD = MsgBox_80DD + "*""Invalid Aadhaar""  at Sr. No  " & i & "  in Schedule 80DD." & Chr(13)
            ValidateAadhaardependent_80DD = False
             Exit Function
        End If


      If UCase(Aadhaar_dependent_80DD(i)) = UCase(Sheet1.Range("Sheet1.Aadhaar").Value) Then
        MsgBox_80DD = MsgBox_80DD + "*""Aadhar of the dependent cannot be same as assessee Aaadhar in Part-A General Information.""" & Chr(13)
         ValidateAadhaardependent_80DD = False
        Exit Function
        End If
     
        If Not IsNumeric(Aadhaar_dependent_80DD(i)) Then
            
           MsgBox_80DD = MsgBox_80DD + "*""Invalid Aadhaar""  at Sr. No  " & i & "  in Schedule 80DD." & Chr(13)
           ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
        If Len(Aadhaar_dependent_80DD(i)) <> 12 Then
       MsgBox_80DD = MsgBox_80DD + "* ""Invalid Aadhaar"" at Sr. No  " & i & "  in Schedule 80DD." & Chr(13)
           ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
        If Aadhaar_dependent_80DD(i) = "000000000000" Then
        MsgBox_80DD = MsgBox_80DD + "* ""Invalid Aadhaar"" at Sr. No  " & i & "  in Schedule 80DD." & Chr(13)
         
            ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
        If Aadhaar_dependent_80DD(i) = "111111111111" Then
       MsgBox_80DD = MsgBox_80DD + "* ""Invalid Aadhaar"" at Sr. No  " & i & "  in Schedule 80DD." & Chr(13)
            
            ValidateAadhaardependent_80DD = False
            Exit Function
        End If
        
    End If
     
Next

UpdateProgressBar
End Function

Function ValidateTypedependent_80DD() As Boolean
  ValidateTypedependent_80DD = True
 
    Dim rangecells As Range
    Set rangecells = Sheet14.Range("Typedependent_80DD").Cells
    Dim i As Long
    end80DD = 1
    ReDim Type_dependent_80DD(end80DD)
    
    For i = 1 To end80DD

        Type_dependent_80DD(i) = rangecells.item(i).Value
        
        If isdropdownblank(Type_dependent_80DD(i)) Or UCase(Type_dependent_80DD(i)) = UCase("Select") Then
        'Ankita_03/06/2025
'           MsgBox_80DD = MsgBox_80DD + "*Please select one of the dropdown in 'Type of dependent' in schedule 80DD at Sr. No " & i & " in Schedule 80DD" & Chr(13)
           MsgBox_80DD = MsgBox_80DD + "*Please select one of the dropdown in 'Type of dependent' in schedule 80DD " & Chr(13)
 
           ValidateTypedependent_80DD = False
           Exit Function
        
        End If
        UpdateProgressBar
    Next
End Function
Sub Validate80DD()
    If Not Validate80DD_1 Then
        Sheet14.Activate
        'MsgBox MsgBox_80U, vbOKOnly, "Error(s)"
        fmsgbox (MsgBox_80DD)
        CloseMsg
    End If
    
    'Malli------23/04/2025----
    
    If Sheet14.Range("AckNoFm10IAfiled_80DD").Value <> "" And Sheet14.Range("AckNoFm10IAfiled_80U").Value <> "" Then
    If Sheet14.Range("AckNoFm10IAfiled_80DD").Value = Sheet14.Range("AckNoFm10IAfiled_80U").Value Then
        MsgBox_80DD = MsgBox_80DD & "* Acknowledgement number of Form 10IA filed for self and Dependent can't be same. Please provide proper acknowledgement number" & Chr(13)
       fmsgbox MsgBox_80DD
       CloseMsg
    End If
    End If
    
    
    '-------------------------
    
End Sub



Function Validate80DD_1() As Boolean
    Validate80DD_1 = True
    MsgBox_80DD = "Schedule 80DD : " & Chr(10)


setTableInfo80DD
'setTableInfo80DD1
setTableInfo80DD2
setTableInfo80DD3
setTableInfo80DD4
'setTableInfo80DD5
setTableInfo80DD6
setTableInfo80DD7
setTableInfo80DD8  'Malli

end80DD = WorksheetFunction.Max(0, end80DD, end80DD1, end80DD2, end80DD3, end80DD4, end80DD6, end80DD7, end80DD8)

If end80DD > 0 Then
If Not ValidateNature_disability_80DD Then Validate80DD_1 = False

'Ankita_18/04/2025
            
    If Not ValidateType_disability_80DD Then Validate80DD_1 = False
If Not ValidateAmount_of_deduction_80DD Then Validate80DD_1 = False
If Not ValidateTypedependent_80DD Then Validate80DD_1 = False
'If Not ValidateDate_of_filingofForm10IA_80DD Then Validate80DD_1 = False
If Not ValidateAckNoFm10IAfiled_80DD Then Validate80DD_1 = False
'Malli-------
' If Not ValidateAckNoFm11A2filed_80DD Then Validate80DD_1 = False
'------------
If Not ValidateUDIDNum_80DD Then Validate80DD_1 = False
If Not ValidatePANdependent_80DD Then Validate80DD_1 = False
If Not ValidateAadhaardependent_80DD Then Validate80DD_1 = False

End If
End Function

Sub Prev80U_DD_Click()
Sheet13.Activate '80GGC SHEET
End Sub

'Ankita_02/05/2025

Sub Next_80U_80DDClick()
Sheet15.Activate
End Sub
