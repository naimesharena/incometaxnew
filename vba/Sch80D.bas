Attribute VB_Name = "Sch80D"
Public FamilyMember, PreventiveHealth, SeniorCitizen As Variant
Public MsgBox_80D As Variant

Public ReceiptA1_80D, ReceiptB1_80D, ReceiptA2_80D, ReceiptB2_80D As Variant



'Ankita_09/04/2025
'TableA1
Dim end_80D As Variant
Dim end_80DNameA1 As Variant
Dim end_80DPolicyA1 As Variant
Dim end_80DReceiptA1 As Variant
Dim end_80DAmountA1 As Variant
'TableB1
Dim end_80DB1 As Variant
Dim end_80DNameB1 As Variant
Dim end_80DPolicyB1 As Variant
Dim end_80DReceiptB1 As Variant
Dim end_80DAmountB1 As Variant
'TableA2
Dim end_80DA2 As Variant
Dim end_80DNameA2 As Variant
Dim end_80DPolicyA2 As Variant
Dim end_80DReceiptA2 As Variant
Dim end_80DAmountA2 As Variant
'TableB2
Dim end_80DB2 As Variant
Dim end_80DNameB2 As Variant
Dim end_80DPolicyB2 As Variant
Dim end_80DReceiptB2 As Variant
Dim end_80DAmountB2 As Variant

Sub ValidateSheet80D_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2026-27"               'Year Changed from 2024-25 to 2025-26 by Ankita on 16/12/2024
Validate80D_All
'MsgBox "Sheet 80G is OK", vbOKOnly, vbMessgaeCaption
fmsgboxoK "Sheet 80D is OK"

End Sub

Sub Validate80D_All()
Dim vbMessgaeCaption As String
 vbMessgaeCaption = "Error"
 
   Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("80D")
    
If Not Validate_80D Then
    sourceSheet.Activate
   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
    fmsgbox (MsgBox_80D)
    CloseMsg
End If
End Sub
Sub Next_80DClick()
Dim sourceSheet As Worksheet
Set sourceSheet = ThisWorkbook.Sheets("80G")
sourceSheet.Activate
End Sub
 Sub Prev80D_Click()
  'Ankita 11/11/2024
 If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
 Sheet202.Activate
 Else:
 '-------------------------
Sheet3.Activate
End If 'Ankita 11/11/2024
End Sub
Function Validate_80D()
Validate_80D = True

end_80D = 0
end_80DB1 = 0
end_80DA2 = 0
end_80DB2 = 0

'Ankita-------------
'TableA1
setTblinfo_80DNameA1
setTblinfo_80DPolicyA1
'setTblinfo_80DReceiptA1
setTblinfo_80DAmountA1
'end_80D = WorksheetFunction.Max(0, end_80D, end_80DNameA1, end_80DPolicyA1, end_80DReceiptA1, end_80DAmountA1)
end_80D = WorksheetFunction.Max(0, end_80DNameA1, end_80DPolicyA1, end_80DAmountA1)

'--------------------
'TableB1
setTblinfo_80DNameB1
setTblinfo_80DPolicyB1
'setTblinfo_80DReceiptB1
setTblinfo_80DAmountB1
'end_80DB1 = WorksheetFunction.Max(0, end_80DB1, end_80DNameB1, end_80DPolicyB1, end_80DReceiptB1, end_80DAmountB1)
end_80DB1 = WorksheetFunction.Max(0, end_80DNameB1, end_80DPolicyB1, end_80DAmountB1)

'---------------------
'TableA2
setTblinfo_80DNameA2
setTblinfo_80DPolicyA2
'setTblinfo_80DReceiptA2
setTblinfo_80DAmountA2
'end_80DA2 = WorksheetFunction.Max(0, end_80DA2, end_80DNameA2, end_80DPolicyA2, end_80DReceiptA2, end_80DAmountA2)
end_80DA2 = WorksheetFunction.Max(0, end_80DNameA2, end_80DPolicyA2, end_80DAmountA2)


 'TableB2
setTblinfo_80DNameB2
setTblinfo_80DPolicyB2
'setTblinfo_80DReceiptB2
setTblinfo_80DAmountB2
'end_80DB2 = WorksheetFunction.Max(0, end_80DB2, end_80DNameB2, end_80DPolicyB2, end_80DReceiptB2, end_80DAmountB2)
end_80DB2 = WorksheetFunction.Max(0, end_80DNameB2, end_80DPolicyB2, end_80DAmountB2)




    If Not ChkFamilyMember Then Validate_80D = False
    If Not chkPreventiveHealth Then Validate_80D = False


'Ankita
    'TableA1
    If Not ValidateNameA1_80D Then Validate_80D = False
    If Not ValidatePolicyA1_80D Then Validate_80D = False
'    If Not ValidateReceiptA1_80D Then Validate_80D = False
    If Not ValidateAmtA1_80D Then Validate_80D = False
    '---------------
    
    'TableB1
    If Not ValidateNameB1_80D Then Validate_80D = False
    If Not ValidatePolicyB1_80D Then Validate_80D = False
'    If Not ValidateReceiptB1_80D Then Validate_80D = False
    If Not ValidateAmtB1_80D Then Validate_80D = False
    '---------------
    
    'TableA2
    If Not ValidateNameA2_80D Then Validate_80D = False
    If Not ValidatePolicyA2_80D Then Validate_80D = False
'    If Not ValidateReceiptA2_80D Then Validate_80D = False
    If Not ValidateAmtA2_80D Then Validate_80D = False
    '---------------
    
     'TableB2
    If Not ValidateNameB2_80D Then Validate_80D = False
    If Not ValidatePolicyB2_80D Then Validate_80D = False
'    If Not ValidateReceiptB2_80D Then Validate_80D = False
    If Not ValidateAmtB2_80D Then Validate_80D = False
    '---------------

End Function
Function ChkFamilyMember() As Boolean
    ChkFamilyMember = True
    FamilyMember = Sheet9.Range("DropDown_ValueOf_FamilyM_80D")
  
    If Trim(FamilyMember) = "" Or Trim(FamilyMember) = "(Select)" Then
    
    'MsgBox_80D = MsgBox_80D + "*  Please select an option from the dropdown ""Whether you or any of your family member (excluding parents) is a senior citizen?"" in Sheet 80D" & Chr(13)
    'NEWLY CHANGED BY BINDU
    MsgBox_80D = MsgBox_80D + "*Please select an option from the dropdown" & Chr(13)
    
    ChkFamilyMember = False
    Exit Function
    End If
    
    SeniorCitizen = Sheet9.Range("DropDown_ValueOf_SC_80D")
  
    If Trim(SeniorCitizen) = "" Or Trim(SeniorCitizen) = "(Select)" Then
    'MsgBox_80D = MsgBox_80D + "* Please select an option from the dropdown ""Whether any one of your parents is a senior citizen"" in Sheet 80D" & Chr(13)
    MsgBox_80D = MsgBox_80D + "* Please select an option from the dropdown" & Chr(13)

    ChkFamilyMember = False
    Exit Function
    End If
    
    If FamilyMember = "Yes" And SeniorCitizen = "No" Then
    MsgBox_80D = MsgBox_80D + "* If Self and Family is a senior citizen then Parents cannot be Non Senior Citizen" & Chr(13)
    ChkFamilyMember = False
    Exit Function
    End If
    
End Function
Function chkPreventiveHealth() As Boolean
chkPreventiveHealth = True

If (Range("Preventive_Health_80D").Value + Range("Preventive_Health_SC_80D").Value + Range("Preventive_Health2_80D").Value + Range("Preventive_Health3_80D").Value) > 5000 Then
    MsgBox_80D = MsgBox_80D + "*Preventive_Health Cannot exceed 5000" & Chr(13)
    chkPreventiveHealth = False
    Exit Function
End If
End Function


'Ankita
'***********************************************starts--------------------------------
'TableA1
'Ankita_09/04
Sub setTblinfo_80DNameA1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("NameInsurerA1.80D").count
    Set rangecells = Range("NameInsurerA1.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DNameA1 = ccount
End Sub

Sub setTblinfo_80DPolicyA1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("PolicyNumA1.80D").count
    Set rangecells = Range("PolicyNumA1.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DPolicyA1 = ccount
End Sub

'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Sub setTblinfo_80DReceiptA1()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("ReceiptNumA1.80D").count
'    Set rangecells = Range("ReceiptNumA1.80D").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80DReceiptA1 = ccount
'End Sub


Sub setTblinfo_80DAmountA1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("AmtA1.80D").count
    Set rangecells = Range("AmtA1.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DAmountA1 = ccount
End Sub


Function ValidateNameA1_80D() As Boolean
 
    ValidateNameA1_80D = True
'    setTblinfo_TCS
'setTblinfo_80DNameA1
    Dim rangecells As Range
    Dim i As Long
    Dim cellrange As String
    Set rangecells = Sheet9.Range("NameInsurerA1.80D").Cells
    ReDim NameA1_80D(end_80D)
    For i = 1 To end_80D
    cellrange = GetMergedAddressCell(rangecells, i)
   ' NameA1_80D(i) = Sheet16.Range(cellrange).Value
     NameA1_80D(i) = Sheet9.Range(cellrange).Value
 '  NameA1_80D(i) = rangecells.item(i).value
        If Not chkCompulsory(NameA1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Name of insurer is mandatory"" in schedule 80D at Sl.no. " & i & " in table 1a" & Chr(13)
            ValidateNameA1_80D = False
            Exit Function
        End If


         If Len(NameA1_80D(i)) > 125 Then
            MsgBox_80D = MsgBox_80D + "* ""Name Of insurer"" in schedule 80D at Sl.no. " & i & " in table 1a cannot be more than 125 characters." & Chr(13)
            ValidateNameA1_80D = False
            Exit Function
         End If
         
         'checkfieldSuperSpecialcharacter
         
         If Not checkfieldSuperSpecialcharacter(NameA1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Name of insurer"" in schedule 80D at Sl.no. " & i & " in table 1a should not Contain <, >, characters." & Chr(13)
            ValidateNameA1_80D = False
            Exit Function
        End If
         
 Next
End Function

Function ValidatePolicyA1_80D() As Boolean
 
    ValidatePolicyA1_80D = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Dim cellrange As String
    Set rangecells = Sheet9.Range("PolicyNumA1.80D").Cells
    ReDim PolicyA1_80D(end_80D)
    For i = 1 To end_80D
    cellrange = GetMergedAddressCell(rangecells, i)
    PolicyA1_80D(i) = Sheet9.Range(cellrange).Value
        If Not chkCompulsory(PolicyA1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Policy Number is mandatory in schedule 80D at Sl.no. " & i & " in table 1a""" & Chr(13)
            ValidatePolicyA1_80D = False
            Exit Function
        End If


         If Len(PolicyA1_80D(i)) > 75 Then
            MsgBox_80D = MsgBox_80D + "* Policy Number in schedule 80D at Sl.no. " & i & " in table 1a cannot be more than 75 characters." & Chr(13)
            ValidatePolicyA1_80D = False
            Exit Function
         End If
         
         If Not checkfieldSuperSpecialcharacter(PolicyA1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* Policy Number in schedule 80D at Sl.no. " & i & " in table 1a should not Contain <, >, characters." & Chr(13)
            ValidatePolicyA1_80D = False
            Exit Function
        End If
 Next
End Function

'Ankita_06/05/2025_Commented as per DESheet_v0.7


'Function ValidateReceiptA1_80D() As Boolean
'
'    ValidateReceiptA1_80D = True
'    'setTblinfo_TCS
'    Dim rangecells As Range
'    Dim i As Long
'    Dim cellrange As String
'    Set rangecells = Sheet9.Range("ReceiptNumA1.80D").Cells
'    'ReDim ReceiptNumA1_80D(end_80D)
'     ReDim ReceiptA1_80D(end_80D)
'
'    For i = 1 To end_80D
'    cellrange = GetMergedAddressCell(rangecells, i)
'    'ReceiptA1_80D(i) = Sheet16.Range(cellrange).Value
'    ReceiptA1_80D(i) = Sheet9.Range(cellrange).Value
'
'       'ReceiptNumA1_80D(i) = rangecells.item(i).Value
'        If Not chkCompulsory(ReceiptA1_80D(i)) Then
'             MsgBox_80D = MsgBox_80D + "* ""Receipt/document number of the premium paid is mandatory in schedule 80D at Sl.no. " & i & " in table 1a""" & Chr(13)
'            ValidateReceiptA1_80D = False
'            Exit Function
'        End If
'
'
'         If Len(ReceiptA1_80D(i)) > 75 Then
'            MsgBox_80D = MsgBox_80D + "* Receipt/document number of the premium paid in schedule 80D at Sl.no. " & i & " in table 1a cannot be more than 75 characters." & Chr(13)
'            ValidateReceiptA1_80D = False
'            Exit Function
'         End If
'
'        If Not checkfieldSuperSpecialcharacter(ReceiptA1_80D(i)) Then
'             MsgBox_80D = MsgBox_80D + "* Receipt/document number of the premium paid at Sl.no. " & i & " in table 1a should not Contain <, >, characters." & Chr(13)
'            ValidateReceiptA1_80D = False
'            Exit Function
'        End If
'
' Next
'End Function

Function ValidateAmtA1_80D() As Boolean
 
    ValidateAmtA1_80D = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet9.Range("AmtA1.80D").Cells
    ReDim AmtA1_80D(end_80D)
    For i = 1 To end_80D
        AmtA1_80D(i) = rangecells.item(i).Value
        If Not chkCompulsory(AmtA1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Amount of premium is mandatory in schedule 80D at Sl.no. " & i & " in table 1a""" & Chr(13)
            ValidateAmtA1_80D = False
            Exit Function
        End If


'         If Len(AmtA1_80D(i)) > 14 Then
'            MsgBox_80D = MsgBox_80D + "* Amount A1 at Sr. No " & i & " in Sheet 80D  cannot exceed 99999999999999" & Chr(13)
'            ValidateAmtA1_80D = False
'            Exit Function
'         End If

         'Ankita_new
         If Not IsNumeric(AmtA1_80D(i)) Then
            MsgBox_80D = MsgBox_80D + "Amount of premium in schedule 80D at Sl.no. " & i & " in table 1a should be Numeric value" & Chr(13)
            ValidateAmtA1_80D = False
            Exit Function
        End If

        If AmtA1_80D(i) > 99999999999999# Then
            MsgBox_80D = MsgBox_80D + "* Amount of premium in schedule 80D at Sl.no. " & i & " in table 1a cannot exceed 14 digits" & Chr(13)
            ValidateAmtA1_80D = False
            Exit Function
        End If

       If AmtA1_80D(i) < 0 Or AmtA1_80D(i) = 0 Then
            MsgBox_80D = MsgBox_80D + """Health Insurance Amount"" claimed in table 1a at Sl.no. " & i & " should be more than 0 in schedule 80D" & Chr(13)
            ValidateAmtA1_80D = False
            Exit Function
        End If

'Ankita_30/04/2025
     If Sheet9.Range("TotAmtA1.80D") > 99999999999999# Then
            MsgBox_80D = MsgBox_80D + "* Total of payments in schedule 80D in table 1a cannot exceed 14 digits" & Chr(13)
            ValidateAmtA1_80D = False
            Exit Function
        End If
 Next
End Function

'TableB1
'Ankita_09/04
Sub setTblinfo_80DNameB1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("NameInsurerB1.80D").count
    Set rangecells = Range("NameInsurerB1.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DNameB1 = ccount
End Sub

Sub setTblinfo_80DPolicyB1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("PolicyNumB1.80D").count
    Set rangecells = Range("PolicyNumB1.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DPolicyB1 = ccount
End Sub

'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Sub setTblinfo_80DReceiptB1()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("ReceiptNumB1.80D").count
'    Set rangecells = Range("ReceiptNumB1.80D").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80DReceiptB1 = ccount
'End Sub

Sub setTblinfo_80DAmountB1()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("AmtB1.80D").count
    Set rangecells = Range("AmtB1.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DAmountB1 = ccount
End Sub

Function ValidateNameB1_80D() As Boolean
 
    ValidateNameB1_80D = True
'    setTblinfo_TCS
'setTblinfo_80DNameA1
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet9.Range("NameInsurerB1.80D").Cells
    Dim cellrange As String
    ReDim NameB1_80D(end_80DB1)
    For i = 1 To end_80DB1
    cellrange = GetMergedAddressCell(rangecells, i)
    NameB1_80D(i) = Sheet9.Range(cellrange).Value
'        NameB1_80D(i) = rangecells.item(i).Value
        If Not chkCompulsory(NameB1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Name of insurer is mandatory"" in schedule 80D at Sl.no. " & i & " in table 1b" & Chr(13)
            ValidateNameB1_80D = False
            Exit Function
        End If


         If Len(NameB1_80D(i)) > 125 Then
            MsgBox_80D = MsgBox_80D + "* ""Name Of insurer"" in schedule 80D at Sl.no. " & i & " in table 1b cannot be more than 125 characters." & Chr(13)
            ValidateNameB1_80D = False
            Exit Function
         End If
         
         'checkfieldSuperSpecialcharacter
         
         If Not checkfieldSuperSpecialcharacter(NameB1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Name of insurer"" in schedule 80D at Sl.no. " & i & " in table 1b should not Contain <, >, characters." & Chr(13)
            ValidateNameB1_80D = False
            Exit Function
        End If
 Next
End Function

Function ValidatePolicyB1_80D() As Boolean
 
    ValidatePolicyB1_80D = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet9.Range("PolicyNumB1.80D").Cells
    
   'Ankita_21/04/2025
    Dim cellrange As String
    
    ReDim PolicyB1_80D(end_80DB1)
    For i = 1 To end_80DB1
    
   'Ankita_21/04/2025
    cellrange = GetMergedAddressCell(rangecells, i)
    PolicyB1_80D(i) = Sheet9.Range(cellrange).Value
        
        
        If Not chkCompulsory(PolicyB1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Policy Number is mandatory in schedule 80D at Sl.no. " & i & " in table 1b""" & Chr(13)
            ValidatePolicyB1_80D = False
            Exit Function
        End If


         If Len(PolicyB1_80D(i)) > 75 Then
            MsgBox_80D = MsgBox_80D + "* Policy Number in schedule 80D at Sl.no. " & i & " in table 1b cannot be more than 75 characters." & Chr(13)
            ValidatePolicyB1_80D = False
            Exit Function
         End If
         
        If Not checkfieldSuperSpecialcharacter(PolicyB1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* Policy Number at Sl.no. " & i & " in table 1b should not Contain <, >, characters." & Chr(13)
            ValidatePolicyB1_80D = False
            Exit Function
        End If

 Next
End Function

'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Function ValidateReceiptB1_80D() As Boolean
'
'    ValidateReceiptB1_80D = True
'    'setTblinfo_TCS
'    Dim rangecells As Range
'    Dim i As Long
'    Set rangecells = Sheet9.Range("ReceiptNumB1.80D").Cells
'    'Ankita_18/04/2025
'    'ReDim ReceiptNumB1_80D(end_80DB1)
'     ReDim ReceiptB1_80D(end_80DB1)
'    For i = 1 To end_80DB1
'    cellrange = GetMergedAddressCell(rangecells, i)
'    ReceiptB1_80D(i) = Sheet9.Range(cellrange).Value
''        ReceiptNumB1_80D(i) = rangecells.item(i).Value
'
'        'If Not chkCompulsory(ReceiptNumB1_80D(i)) Then
'         If Not chkCompulsory(ReceiptB1_80D(i)) Then
'             MsgBox_80D = MsgBox_80D + "* ""Receipt/document number of the premium paid is mandatory in schedule 80D at Sl.no. " & i & " in table 1b""" & Chr(13)
'            ValidateReceiptB1_80D = False
'            Exit Function
'        End If
'
'
'         'If Len(PolicyB1_80D(i)) > 75 Then
'          If Len(ReceiptB1_80D(i)) > 75 Then
'            MsgBox_80D = MsgBox_80D + "* Receipt/document number of the premium paid in schedule 80D at Sl.no. " & i & " in table 1b cannot be more than 75 characters." & Chr(13)
'            ValidateReceiptB1_80D = False
'            Exit Function
'         End If
'
'
'        If Not checkfieldSuperSpecialcharacter(ReceiptB1_80D(i)) Then
'             MsgBox_80D = MsgBox_80D + "* Receipt/document number of the premium paid at Sl.no. " & i & " in table 1b should not Contain <, >, characters." & Chr(13)
'            ValidateReceiptB1_80D = False
'            Exit Function
'        End If
'
' Next
'End Function
'

Function ValidateAmtB1_80D() As Boolean
 
    ValidateAmtB1_80D = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet9.Range("AmtB1.80D").Cells
    ReDim AmtB1_80D(end_80DB1)
    For i = 1 To end_80DB1
        AmtB1_80D(i) = rangecells.item(i).Value
        If Not chkCompulsory(AmtB1_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Amount of premium is mandatory in schedule 80D at Sl.no. " & i & " in table 1b""" & Chr(13)
            ValidateAmtB1_80D = False
            Exit Function
        End If


'         If Len(AmtB1_80D(i)) > 14 Then
'            MsgBox_80D = MsgBox_80D + "* Amount B1 at Sr. No " & i & " in Sheet 80D  cannot exceed 99999999999999" & Chr(13)
'            ValidateAmtB1_80D = False
'            Exit Function
'         End If

'Ankita_new
         If Not IsNumeric(AmtB1_80D(i)) Then
            MsgBox_80D = MsgBox_80D + "Amount of premium in schedule 80D at Sl.no. " & i & " in table 1b should be Numeric value" & Chr(13)
            ValidateAmtB1_80D = False
            Exit Function
        End If
        
        If AmtB1_80D(i) > 99999999999999# Then
            MsgBox_80D = MsgBox_80D + "* Amount of premium in schedule 80D at Sl.no. " & i & " in table 1b cannot exceed 14 digits" & Chr(13)
            ValidateAmtB1_80D = False
            Exit Function
        End If
        
        If AmtB1_80D(i) < 0 Or AmtB1_80D(i) = 0 Then
            MsgBox_80D = MsgBox_80D + """Health Insurance Amount"" claimed in table 1b at Sl.no. " & i & " should be more than 0 in schedule 80D" & Chr(13)
            ValidateAmtB1_80D = False
            Exit Function
        End If
        
        
        'Ankita_30/04/2025
        
        If Sheet9.Range("TotAmtB1.80D") > 99999999999999# Then
            MsgBox_80D = MsgBox_80D + "* Total of payments in schedule 80D in table 1b cannot exceed 14 digits" & Chr(13)
            ValidateAmtB1_80D = False
            Exit Function
        End If
        
        
 Next
End Function

'TableA2
'Ankita_09/04
Sub setTblinfo_80DNameA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("NameInsurerA2.80D").count
    Set rangecells = Range("NameInsurerA2.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DNameA2 = ccount
End Sub


Sub setTblinfo_80DPolicyA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("PolicyNumA2.80D").count
    Set rangecells = Range("PolicyNumA2.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DPolicyA2 = ccount
End Sub

'Ankita_06/05/2025_Commented as per DESheet_v0.7


'Sub setTblinfo_80DReceiptA2()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("ReceiptNumA2.80D").count
'    Set rangecells = Range("ReceiptNumA2.80D").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80DReceiptA2 = ccount
'End Sub


Sub setTblinfo_80DAmountA2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("AmtA2.80D").count
    Set rangecells = Range("AmtA2.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DAmountA2 = ccount
End Sub


Function ValidateNameA2_80D() As Boolean
 
    ValidateNameA2_80D = True
'    setTblinfo_TCS
'setTblinfo_80DNameA1
    Dim rangecells As Range
    Dim i As Long
    Dim cellrange As String
    Set rangecells = Sheet9.Range("NameInsurerA2.80D").Cells
    ReDim NameA2_80D(end_80DA2)
    For i = 1 To end_80DA2
    cellrange = GetMergedAddressCell(rangecells, i)
    NameA2_80D(i) = Sheet9.Range(cellrange).Value
 '  NameA2_80D(i) = rangecells.item(i).value
        If Not chkCompulsory(NameA2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Name of insurer is mandatory"" in schedule 80D at Sl.no. " & i & " in table 2a" & Chr(13)
            ValidateNameA2_80D = False
            Exit Function
        End If


         If Len(NameA2_80D(i)) > 125 Then
            MsgBox_80D = MsgBox_80D + "* ""Name Of insurer"" in schedule 80D at Sl.no. " & i & " in table 2a cannot be more than 125 characters." & Chr(13)
            ValidateNameA2_80D = False
            Exit Function
         End If
         
         'checkfieldSuperSpecialcharacter
         
         If Not checkfieldSuperSpecialcharacter(NameA2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Name of insurer"" in schedule 80D at Sl.no. " & i & " in table 2a should not Contain <, >, characters." & Chr(13)
            ValidateNameA2_80D = False
            Exit Function
        End If
        
 Next
End Function

Function ValidatePolicyA2_80D() As Boolean
 
    ValidatePolicyA2_80D = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Dim cellrange As String
    Set rangecells = Sheet9.Range("PolicyNumA2.80D").Cells
    ReDim PolicyA2_80D(end_80DA2)
    For i = 1 To end_80DA2
        cellrange = GetMergedAddressCell(rangecells, i)
        PolicyA2_80D(i) = Sheet9.Range(cellrange).Value
       ' PolicyA2_80D(i) = rangecells.item(i).value
        If Not chkCompulsory(PolicyA2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Policy Number is mandatory in schedule 80D at Sl.no. " & i & " in table 2a""" & Chr(13)
            ValidatePolicyA2_80D = False
            Exit Function
        End If


         If Len(PolicyA2_80D(i)) > 75 Then
            MsgBox_80D = MsgBox_80D + "* Policy Number in schedule 80D at Sl.no. " & i & " in table 2a cannot be more than 75 characters." & Chr(13)
            ValidatePolicyA2_80D = False
            Exit Function
         End If
         
        If Not checkfieldSuperSpecialcharacter(PolicyA2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* Receipt/document number of the premium paid at Sl.no. " & i & " in table 2a should not Contain <, >, characters." & Chr(13)
            ValidatePolicyA2_80D = False
            Exit Function
        End If

 Next
End Function

'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Function ValidateReceiptA2_80D() As Boolean
'
'    ValidateReceiptA2_80D = True
'    'setTblinfo_TCS
'    Dim rangecells As Range
'    Dim i As Long
'    Dim cellrange As String
'    Set rangecells = Sheet9.Range("ReceiptNumA2.80D").Cells
'    'ReDim ReceiptNumA2_80D(end_80DA2)
'     ReDim ReceiptA2_80D(end_80DA2)
'
'    For i = 1 To end_80DA2
'    cellrange = GetMergedAddressCell(rangecells, i)
'    ReceiptA2_80D(i) = Sheet9.Range(cellrange).Value
''       ReceiptNumA2_80D(i) = rangecells.item(i).Value
'
'        'If Not chkCompulsory(ReceiptNumA2_80D(i)) Then
'         If Not chkCompulsory(ReceiptA2_80D(i)) Then
'             MsgBox_80D = MsgBox_80D + "* ""Receipt/document number of the premium paid is mandatory in schedule 80D at Sl.no. " & i & " in table 2a""" & Chr(13)
'            ValidateReceiptA2_80D = False
'            Exit Function
'        End If
'
'
'         'If Len(ReceiptNumA2_80D(i)) > 75 Then
'          If Len(ReceiptA2_80D(i)) > 75 Then
'            MsgBox_80D = MsgBox_80D + "* Receipt A2 at Sr. No " & i & " in Sheet 80D  less than 25 characters" & Chr(13)
'            ValidateReceiptA2_80D = False
'            Exit Function
'         End If
'
'        If Not checkfieldSuperSpecialcharacter(ReceiptA2_80D(i)) Then
'             MsgBox_80D = MsgBox_80D + "* Receipt/document number of the premium paid at Sl.no. " & i & " in table 2a should not Contain <, >, characters." & Chr(13)
'            ValidateReceiptA2_80D = False
'            Exit Function
'        End If
'
' Next
'End Function

Function ValidateAmtA2_80D() As Boolean
 
    ValidateAmtA2_80D = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet9.Range("AmtA2.80D").Cells
    ReDim AmtA2_80D(end_80DA2)
    For i = 1 To end_80DA2
        AmtA2_80D(i) = rangecells.item(i).Value
        If Not chkCompulsory(AmtA2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Amount of premium is mandatory in schedule 80D at Sl.no. " & i & " in table 2a""" & Chr(13)
            ValidateAmtA2_80D = False
            Exit Function
        End If


        If Not IsNumeric(AmtA2_80D(i)) Then
            MsgBox_80D = MsgBox_80D + "Amount of premium in schedule 80D at Sl.no. " & i & " in table 2a should be Numeric value" & Chr(13)
            ValidateAmtA2_80D = False
            Exit Function
        End If
        
        If AmtA2_80D(i) > 99999999999999# Then
            MsgBox_80D = MsgBox_80D + "* Amount of premium in schedule 80D at Sl.no. " & i & " in table 2a cannot exceed 14 digits" & Chr(13)
            ValidateAmtA2_80D = False
            Exit Function
        End If
        
        If AmtA2_80D(i) < 0 Or AmtA2_80D(i) = 0 Then
            MsgBox_80D = MsgBox_80D + """Health Insurance Amount"" claimed in table 2a at Sl.no. " & i & "  should be more than 0 in schedule 80D" & Chr(13)
            ValidateAmtA2_80D = False
            Exit Function
        End If
        
        'Ankita_30/04/2025
        If Sheet9.Range("TotAmtA2.80D") > 99999999999999# Then
            MsgBox_80D = MsgBox_80D + "* Total of payments in schedule 80D in table 2a cannot exceed 14 digits" & Chr(13)
            ValidateAmtA2_80D = False
            Exit Function
        End If
 Next
End Function


'TableB2
'Ankita_09/04
Sub setTblinfo_80DNameB2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("NameInsurerB2.80D").count
    Set rangecells = Range("NameInsurerB2.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DNameB2 = ccount
End Sub

Sub setTblinfo_80DPolicyB2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("PolicyNumB2.80D").count
    Set rangecells = Range("PolicyNumB2.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DPolicyB2 = ccount
End Sub

'Sub setTblinfo_80DReceiptB2()
'    Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("ReceiptNumB2.80D").count
'    Set rangecells = Range("ReceiptNumB2.80D").Cells
'    For mIntCtr = 1 To mIntCells
'            If Not rangecells.item(mIntCtr).Value = "" Then
'                ccount = ccount + 1
'            End If
'    Next
'    end_80DReceiptB2 = ccount
'End Sub


Sub setTblinfo_80DAmountB2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("AmtB2.80D").count
    Set rangecells = Range("AmtB2.80D").Cells
    For mIntCtr = 1 To mIntCells
            If Not rangecells.item(mIntCtr).Value = "" Then
                ccount = ccount + 1
            End If
    Next
    end_80DAmountB2 = ccount
End Sub


Function ValidateNameB2_80D() As Boolean
 
    ValidateNameB2_80D = True
'    setTblinfo_TCS
'setTblinfo_80DNameA1
    Dim rangecells As Range
    Dim i As Long
    Dim cellrange As String
    
    Set rangecells = Sheet9.Range("NameInsurerB2.80D").Cells
    ReDim NameB2_80D(end_80DB2)
    For i = 1 To end_80DB2
    cellrange = GetMergedAddressCell(rangecells, i)
    NameB2_80D(i) = Sheet9.Range(cellrange).Value
'        NameB2_80D(i) = rangecells.item(i).Value
        If Not chkCompulsory(NameB2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Name of insurer is mandatory"" in schedule 80D at Sl.no. " & i & " in table 2b" & Chr(13)
            ValidateNameB2_80D = False
            Exit Function
        End If


         If Len(NameB2_80D(i)) > 125 Then
            MsgBox_80D = MsgBox_80D + "* ""Name Of insurer"" in schedule 80D at Sl.no. " & i & " in table 2b cannot be more than 125 characters." & Chr(13)
            ValidateNameB2_80D = False
            Exit Function
         End If
         
         'checkfieldSuperSpecialcharacter
         
         If Not checkfieldSuperSpecialcharacter(NameB2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Name of insurer"" in schedule 80D at Sl.no. " & i & " in table 2b should not Contain <, >, characters." & Chr(13)
            ValidateNameB2_80D = False
            Exit Function
        End If
 Next
End Function

Function ValidatePolicyB2_80D() As Boolean
 
    ValidatePolicyB2_80D = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Dim cellrange As String
    
    Set rangecells = Sheet9.Range("PolicyNumB2.80D").Cells
    ReDim PolicyB2_80D(end_80DB2)
    For i = 1 To end_80DB2
    cellrange = GetMergedAddressCell(rangecells, i)
    PolicyB2_80D(i) = Sheet9.Range(cellrange).Value
'   PolicyB2_80D(i) = rangecells.item(i).value
        If Not chkCompulsory(PolicyB2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Policy Number is mandatory in schedule 80D at Sl.no. " & i & " in table 2b""" & Chr(13)
            ValidatePolicyB2_80D = False
            Exit Function
        End If


         If Len(PolicyB2_80D(i)) > 75 Then
            MsgBox_80D = MsgBox_80D + "* Policy Number in schedule 80D at Sl.no. " & i & " in table 2b cannot be more than 75 characters." & Chr(13)
            ValidatePolicyB2_80D = False
            Exit Function
         End If
         
         If Not checkfieldSuperSpecialcharacter(PolicyB2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* Receipt/document number of the premium paid at Sl.no. " & i & " in table 2b should not Contain <, >, characters." & Chr(13)
            ValidatePolicyB2_80D = False
            Exit Function
        End If
 Next
End Function

'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Function ValidateReceiptB2_80D() As Boolean
'
'    ValidateReceiptB2_80D = True
'    'setTblinfo_TCS
'    Dim rangecells As Range
'    Dim i As Long
'    Dim cellrange As String
'    Set rangecells = Sheet9.Range("ReceiptNumB2.80D").Cells
'    'ReDim ReceiptNumB2_80D(end_80DB2)
'    ReDim ReceiptB2_80D(end_80DB2)
'
'    For i = 1 To end_80DB2
'    cellrange = GetMergedAddressCell(rangecells, i)
'    ReceiptB2_80D(i) = Sheet9.Range(cellrange).Value
''        ReceiptNumB2_80D(i) = rangecells.item(i).Value
'
'        'If Not chkCompulsory(ReceiptNumB2_80D(i)) Then
'        If Not chkCompulsory(ReceiptB2_80D(i)) Then
'             MsgBox_80D = MsgBox_80D + "* ""Receipt/document number of the premium paid is mandatory in schedule 80D at Sl.no. " & i & " in table 2b""" & Chr(13)
'            ValidateReceiptB2_80D = False
'            Exit Function
'        End If
'
'
'         If Len(ReceiptB2_80D(i)) > 75 Then
'            MsgBox_80D = MsgBox_80D + "* Receipt/document number of the premium paid in schedule 80D at Sl.no. " & i & " in table 2b cannot be more than 75 characters." & Chr(13)
'            ValidateReceiptB2_80D = False
'            Exit Function
'         End If
'
'        If Not checkfieldSuperSpecialcharacter(ReceiptB2_80D(i)) Then
'             MsgBox_80D = MsgBox_80D + "* Receipt/document number of the premium paid at Sl.no. " & i & " in table 2b should not Contain <, >, characters." & Chr(13)
'            ValidateReceiptB2_80D = False
'            Exit Function
'        End If
'
' Next
'End Function

Function ValidateAmtB2_80D() As Boolean
 
    ValidateAmtB2_80D = True
    'setTblinfo_TCS
    Dim rangecells As Range
    Dim i As Long
    Set rangecells = Sheet9.Range("AmtB2.80D").Cells
    ReDim AmtB2_80D(end_80DB2)
    For i = 1 To end_80DB2
        AmtB2_80D(i) = rangecells.item(i).Value
        If Not chkCompulsory(AmtB2_80D(i)) Then
             MsgBox_80D = MsgBox_80D + "* ""Amount of premium is mandatory in schedule 80D at Sl.no. " & i & " in table 2b""" & Chr(13)
            ValidateAmtB2_80D = False
            Exit Function
        End If


'         If Len(AmtB2_80D(i)) > 14 Then
'            MsgBox_80D = MsgBox_80D + "* Amount B2 at Sr. No " & i & " in Sheet 80D  cannot exceed 99999999999999" & Chr(13)
'            ValidateAmtB2_80D = False
'            Exit Function
'         End If

'Ankita_new
         If Not IsNumeric(AmtB2_80D(i)) Then
            MsgBox_80D = MsgBox_80D + "Amount of premium in schedule 80D at Sl.no. " & i & " in table 2b should be Numeric value" & Chr(13)
            ValidateAmtB2_80D = False
            Exit Function
        End If
        
        If AmtB2_80D(i) > 99999999999999# Then
            MsgBox_80D = MsgBox_80D + "* Amount of premium in schedule 80D at Sl.no. " & i & " in table 2b cannot exceed 14 digits" & Chr(13)
            ValidateAmtB2_80D = False
            Exit Function
        End If
        
        If AmtB2_80D(i) < 0 Or AmtB2_80D(i) = 0 Then
            MsgBox_80D = MsgBox_80D + """Health Insurance Amount"" claimed in table 2b at Sl.no. " & i & " should be more than 0 in schedule 80D" & Chr(13)
            ValidateAmtB2_80D = False
            Exit Function
        End If
        
        'Ankita_30/04/2025
        
        If Sheet9.Range("TotAmtB2.80D") > 99999999999999# Then
            MsgBox_80D = MsgBox_80D + "* Total of payments in schedule 80D in table 2b cannot exceed 14 digits" & Chr(13)
            ValidateAmtB2_80D = False
            Exit Function
        End If
 Next
End Function










