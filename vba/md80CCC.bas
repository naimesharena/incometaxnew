Attribute VB_Name = "md80CCC"
'Public MsgBox_80CCC As Variant
'Dim end_80CCC As Variant
'Dim end_80CCCNameInsurer As Variant
'Dim end_80CCCPolicy_Number As Variant
'Dim end_80CCCAmount As Variant
'
'
'Sub ValidateSheet80C2_Click()
'Dim vbMessgaeCaption As String
''Dim vbMessgaeCaption As String
''vbMessgaeCaption = "ITR 4: AY: 2022-23" dpk1201
'' vbMessgaeCaption = "ITR 4: AY: 2023-24"
'vbMessgaeCaption = "ITR 1: AY: 2025-26"
'Validate80C2_All
''MsgBox "Sheet 80G is OK", vbOKOnly, vbMessgaeCaption
'fmsgboxStatus "Sheet 80CCC is OK"
'End Sub
'Sub Validate80C2_All()
'Dim vbMessgaeCaption As String
'vbMessgaeCaption = "Error"
'   Dim sourceSheet As Worksheet
'    Set sourceSheet = ThisWorkbook.Sheets("80C")
''If Not Validate_80CCC2 Then
''    sourceSheet.Activate
''   ' MsgBox (MsgBox_80GA), vbOKOnly, vbMessgaeCaption
''    fmsgboxStatus MsgBox_80CCC
''    CloseMsg
'End If
'End Sub
'Function Validate_80CCC2()
'Validate_80CCC2 = True
''Ankita-------------
''TableA1
'setTblinfo_80CCCNameInsurer
'
'setTblinfo_80CCCPolicy_Number
'setTblinfo_80CCCAmount
'end_80CCC = WorksheetFunction.Max(0, end_80CCC, end_80CCCNameInsurer, end_80CCCPolicy_Number, end_80CCCAmount)
'    If Not ValidateNameInsurer_80CCC Then Validate_80CCC2 = False
'    If Not ValidateAmount_80CCC Then Validate_80CCC2 = False
'    If Not ValidatePolicy_Number_80CCC Then Validate_80CCC2 = False
'    If Not Validategreater_80CCC Then Validate_80CCC2 = False 'Newly added by sai on 22/04/2025
'    '---------------
'End Function
'
''Ankita_06/05/2025_Commented as per DESheet_v0.7
'
''Sub setTblinfo_80CCCNameInsurer()
''    Dim rangecells As Range
''    Dim mIntCells As Long
''    Dim mIntCtr As Long
''    Dim ccount As Long
''    ccount = 0
''    mIntCells = Range("Name_of_insurer.80CCC").count
''    Set rangecells = Range("Name_of_insurer.80CCC").Cells
''    For mIntCtr = 1 To mIntCells
''            If Not rangecells.item(mIntCtr).Value = "" Then
''                ccount = ccount + 1
''            End If
''    Next
''    end_80CCCNameInsurer = ccount
''End Sub
''Sub setTblinfo_80CCCAmount()
''    Dim rangecells As Range
''    Dim mIntCells As Long
''    Dim mIntCtr As Long
''    Dim ccount As Long
''    ccount = 0
''    mIntCells = Range("Amount.80CCC").count
''    Set rangecells = Range("Amount.80CCC").Cells
''    For mIntCtr = 1 To mIntCells
''            If Not rangecells.item(mIntCtr).Value = "" Then
''                ccount = ccount + 1
''            End If
''    Next
''    end_80CCCAmount = ccount
''End Sub
'
''Ankita_06/05/2025_Commented as per DESheet_v0.7
'
''Sub setTblinfo_80CCCPolicy_Number()
''    Dim rangecells As Range
''    Dim mIntCells As Long
''    Dim mIntCtr As Long
''    Dim ccount As Long
''    ccount = 0
''    mIntCells = Range("Policy_Document_Number.80CCC").count
''    Set rangecells = Range("Policy_Document_Number.80CCC").Cells
''    For mIntCtr = 1 To mIntCells
''            If Not rangecells.item(mIntCtr).Value = "" Then
''                ccount = ccount + 1
''            End If
''    Next
''    end_80CCCPolicy_Number = ccount
''End Sub
'
''Ankita_06/05/2025_Commented as per DESheet_v0.7
'
''Function ValidateNameInsurer_80CCC() As Boolean
''    ValidateNameInsurer_80CCC = True
'''    setTblinfo_TCS
'''setTblinfo_80DNameA1
''    Dim rangecells As Range
''    Dim i As Long
''    Set rangecells = Sheet15.Range("Name_of_insurer.80CCC").Cells
''    ReDim Nature_80CCC(end_80CCC)
''    For i = 1 To end_80CCC
''        Nature_80CCC(i) = rangecells.item(i).Value
''        If Not chkCompulsory(Nature_80CCC(i)) Then
'''             MsgBox_80CCC = MsgBox_80CCC + "*  Name of insurer at A1 at Sr. No " & i & "  in Sheet 80D" & Chr(13)
'''modified by sai on 22/04/2025
''              MsgBox_80CCC = MsgBox_80CCC + "*  ""Name of the insurer"" is mandatory in schedule 80CCC at Sr. No " & i & "" & Chr(13)
''              ValidateNameInsurer_80CCC = False
''            Exit Function
''        End If
''
''         If Len(Nature_80CCC(i)) > 125 Then
''            MsgBox_80CCC = MsgBox_80CCC + "* ""Name of the insurer"" cannot be more than 125 characters at Sr. No " & i & "" & Chr(13)
''            ValidateNameInsurer_80CCC = False
''            Exit Function
''         End If
''Next
''End Function
'
''Function ValidateAmount_80CCC() As Boolean
''    ValidateAmount_80CCC = True
'''    setTblinfo_TCS
'''setTblinfo_80DNameA1
''    Dim rangecells As Range
''    Dim i As Long
''    Set rangecells = Sheet15.Range("Amount.80CCC").Cells
''    ReDim Amount_80CCC(end_80CCC)
''    For i = 1 To end_80CCC
''        Amount_80CCC(i) = rangecells.item(i).Value
''        If Not chkCompulsory(Amount_80CCC(i)) Then
'''              MsgBox_80CCC = MsgBox_80CCC + "* Please enter the Amount at A1 at Sr. No " & i & "  in Sheet 80D" & Chr(13)
''
'''Modified by sai on 22/04/2025
''               MsgBox_80CCC = MsgBox_80CCC + "*  ""Deduction u/s 80CCC is mandatory in schedule 80CCC"" at Sr. No " & i & "" & Chr(13)
''            ValidateAmount_80CCC = False
''            Exit Function
''        End If
''         If Len(Amount_80CCC(i)) > 125 Then
'''            MsgBox_80CCC = MsgBox_80CCC + "* Amount at A1as Sr. No " & i & " in Sheet 80D cannot be more than 125 characters." & Chr(13)
''             MsgBox_80CCC = MsgBox_80CCC + "* Deduction Amount u/s 80CCC cannot be more than 14 digits in schedule 80CCC at Sr. No " & i & "" & Chr(13)
''            ValidateAmount_80CCC = False
''            Exit Function
''         End If
''Next
''End Function
'
''Ankita_06/05/2025_Commented as per DESheet_v0.7
'
''Function ValidatePolicy_Number_80CCC() As Boolean
''    ValidatePolicy_Number_80CCC = True
'''    setTblinfo_TCS
'''setTblinfo_80DNameA1
''    Dim rangecells As Range
''    Dim i As Long
''    Set rangecells = Sheet15.Range("Policy_Document_Number.80CCC").Cells
''    ReDim Identification_Number_80CCC(end_80CCC)
''    For i = 1 To end_80CCC
''        Identification_Number_80CCC(i) = rangecells.item(i).Value
''        If Not chkCompulsory(Identification_Number_80CCC(i)) Then
'''             MsgBox_80CCC = MsgBox_80CCC + "* Please enter the Identification_Number at A1 at Sr. No " & i & "  in Sheet 80D" & Chr(13)
'''Modified by sai on 22/04/2025
''              MsgBox_80CCC = MsgBox_80CCC + "*  ""Policy document number"" is mandatory in schedule 80CCC at Sr. No " & i & "" & Chr(13)
''            ValidatePolicy_Number_80CCC = False
''            Exit Function
''        End If
''
''         If Len(Identification_Number_80CCC(i)) > 125 Then
'''            MsgBox_80CCC = MsgBox_80CCC + "* Identification_Number at A1as Sr. No " & i & " in Sheet 80D cannot be more than 125 characters." & Chr(13)
''             MsgBox_80CCC = MsgBox_80CCC + "* ""Policy document number"" cannot be more than 100 characters in schedule 80CCC at Sr. No " & i & "" & Chr(13)
''            ValidatePolicy_Number_80CCC = False
''            Exit Function
''         End If
''Next
''End Function
''Newly added by sai on 22/04/2025
'Function Validategreater_80CCC() As Boolean
'    Validategreater_80CCC = True
'
'         If (Len(Sheet15.Range("TotAmount.80CC").Value) > 14) Then
'            MsgBox_80CCC = MsgBox_80CCC + "*  Total of deduction u/s 80CCC amount cannot exceed 14 Digits." & Chr(13)
'            Validategreater_80CCC = False
'            Exit Function
'         End If
'
'End Function
'
'
'
'
