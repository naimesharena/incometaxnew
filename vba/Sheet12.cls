Attribute VB_Name = "Sheet12"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit



Private Sub Worksheet_Activate()
On Error Resume Next
    ActiveSheet.Shapes("80GGA").Left = Range("Button_80GGA").Left + 5
    ActiveSheet.Shapes("80GGA").Top = Range("Button_80GGA").Top + 5

'    ThisWorkbook.Protect Password:=getmsgstate
End Sub


'Private Sub Worksheet_Deactivate()
'ThisWorkbook.Unprotect Password:=getmsgstate
'End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub

Private Sub Worksheet_Calculate()

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo exit1
sPassword = EfilingCommon.getmsgstate
Sheet12.Unprotect Password:=sPassword
Sheet1.Range("taxcheck").Value = ""
Sheet12.Protect Password:=sPassword
    Application.EnableEvents = False
    If Target.Validation.Type = 3 Then
         GoTo exit1
    End If

    Target.Value = UCase(Target.Value)
Dim var As Integer
'ramya12 commented===========================================================

' If Target.Column = 10 Then
'        Dim PinCode As Range
'        For Each PinCode In Sheet5.Range("PincodeList")
'            Dim code As String
'            code = Left(CStr(Target.Value), 3)
'            If CStr(PinCode.Value) = CStr(code) Then
'                Range("I" & Target.row).Value = Sheet5.Range("V" & PinCode.row).Value
'                Exit For
'            End If
'         Next PinCode
'    End If
 'ramya12 commented=======================================================================
 
 
    ' Itr-6 code


'ramya1234====================================================================================================

'If Not Application.Intersect(Sheet12.Range("State_Code_80GGA"), Sheet12.Range(Target.Address)) Is Nothing Then
            'If Sheet1.Range("sheet1.PinCode").value <> "" Then
'           If Sheet12.Range(Replace(Target.Address, "I", "J")).Value <> "" Then
                'if state and pincode doesn't match, empty the pincode.
'                Dim State1 As Variant
                'Sheet2.Range(Replace(Target.address, "I", "G")).value
'                State1 = Application.VLookup(Sheet12.Range(Replace(Target.Address, "I", "J")).Value, Sheet5.Range("All_PinCode_V"), 2, False)
'                If Not (State1 = UCase(Mid(Target.Value, 4)) Or EfilingCommon.StateMatchesPin(Target.Value, Sheet12.Range(Replace(Target.Address, "I", "J")).Value)) Then
'                       Sheet12.Range(Replace(Target.Address, "I", "J")).Value = ""
'                End If
'            End If
'        End If
'
'
'
'If Not Application.Intersect(Sheet12.Range("Pincode_80GGA"), Sheet12.Range(Target.Address)) Is Nothing Then
         'If the pincode is not in the duplicate list, populate, else don't populate---
'         Dim DupFlag As Boolean
'         Dim PinCode As Variant
'
'         DupFlag = False
'                 'If the pincode is not in the duplicate list, populate, else don't populate---
'                 For Each PinCode In Sheet5.Range("Duplicate_Pincode_List")
'                    If Target.Value = PinCode Then
'                        DupFlag = True
'                        Exit For
'                    End If
'                 Next PinCode
'                 If DupFlag = True Then
'                    If Sheet12.Range(Replace(Target.Address, "J", "I")).Value <> "" Then
'                        If Not StateMatchesPin(Sheet12.Range(Replace(Target.Address, "J", "I")).Value, Target.Value) Then
'                            Sheet12.Range(Replace(Target.Address, "J", "I")).Value = "(Select)"
'                        End If
'                    End If
'                 End If
'                 If DupFlag = False Then
'                       Dim StateName
'                       Dim StateCode
''
'                    StateName = Application.VLookup(Target.Value, Sheet5.Range("All_PinCode_V"), 2, False)
'                    If IsError(StateName) Then
'                        If Target.Value <> "" Then
'                        MsgBox ("Invalid Pincode")
'                        End If
'                        'Sheet1.Range("sheet1.StateCode1").value = "(Select)"
'                        Application.EnableEvents = False
'                        Target.Value = ""
'                        Application.EnableEvents = True
'                    Else
'                        StateCode = Application.VLookup(StateName, Sheet5.Range("CS:CT"), 2, False)
'                          Sheet12.Range(Replace(Target.Address, "J", "I")).Value = StateCode
                         'sheet2.Range(Replace(Target.address, "J", "I")).value = "91-INDIA"
                         Sheet12.Unprotect Password:=EfilingCommon.getmsgstate
'                        'If Mid(Sheet12.Range(Replace(Target.address, "J", "H")).value, 1, 2) = "99" Then
'                                sheet2.Range(Replace(Target.address, "J", "K")).Interior.Color = (&HCCFFCC)
'                                sheet2.Range(Replace(Target.address, "J", "K")).value = ""
'                                sheet2.Range(Replace(Target.address, "J", "K")).Locked = False
'                        ElseIf Mid(sheet2.Range(Replace(Target.address, "J", "H")).value, 1, 2) <> "99" Then
'                                sheet2.Range(Replace(Target.address, "J", "K")).Interior.Color = (&HD8D8D8)
'                                sheet2.Range(Replace(Target.address, "J", "K")).value = ""
'                                sheet2.Range(Replace(Target.address, "J", "K")).Locked = True
'                        End If
 
                                
                    'End If
'            End If
'
'End If
'End If
'ramya1234=========================================================================================================








If Not Application.Intersect(Sheet12.Range("State_Code_80GGA"), Sheet12.Range(Target.Address)) Is Nothing Then
    Dim PIN_targetadd, state_targetadd As String
      'state_targetadd = Target.address
      state_targetadd = Replace(Target.Address, "$", "")
      PIN_targetadd = Replace(state_targetadd, "I", "J")

    If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet12.Range(Replace(state_targetadd, "I", "J")).Value = ""
End If
 

If Not Application.Intersect(Sheet12.Range("Pincode_80GGA"), Sheet12.Range(Target.Address)) Is Nothing Then
     Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
     Dim ws As Worksheet
      'pinrng_chk = Target.address
      pinrng_chk = Replace(Target.Address, "$", "")
      staterng_chk = Replace(pinrng_chk, "J", "I")
'      cntryrng_chk = Replace(pinrng_chk, "J", "I")
'      ZIPrng_chk = Replace(pinrng_chk, "J", "K")
      Set ws = Worksheets("80GGA")
     ' Call PinState_codes_validation(ws, pinrng_chk, staterng_chk)
      
Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
End If












'sai1234================================================================================================
    
    If Target.Column = 11 Then
    
        If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Range("K" & Target.row).Value) Then
        
            If Not Range("M" & Target.row).Value = "" And Not Range("K" & Target.row).Value = "" Then
            
                MsgBox "* Duplicate Pan value is not allowed"
                Range("K" & Target.row).Value = ""
                Range("M" & Target.row).Value = ""
                
            End If
        End If
    End If

    If Target.Column = 13 Then
        If Not Range("M" & Target.row).Value = "" Then
        var = Application.CountIf(Range("K1:K" & Rows.count), Range("K" & Target.row).Value)
            If 1 < var Then
                 MsgBox "* Duplicate Pan value is not allowed"
                 Range("K" & Target.row).Value = ""
                 Range(Target.Address).Value = ""
            End If
        End If
        
        
        If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Range("K" & Target.row).Value) Then
            If Not Range("M" & Target.row).Value = "" And Not Range("K" & Target.row).Value = "" Then
            
'             MsgBox "* Duplicate Pan value is not allowed"
'             Range("K" & Target.Row).Value = ""
'             Range("M" & Target.Row).Value = ""
            End If
        End If
      
        If Not Range("K" & Target.row).Value = "" Then
        
        var = Application.CountIf(Range("M1:M" & Rows.count), Range("M" & Target.row).Value)
      
           If Range("M" & Target.row).Value > 2000 Then
      
                  MsgBox "* If Donation in cash is more than Rs. 2000, 80GGA Deduction is not allowed"
                 'Range("M" & Target.Row).Value = ""
                 'Range(Target.Address).Value = ""
            End If
        End If
    End If
    If Target.Column = 14 Then
        If Not Range("N" & Target.row).Value = "" Then
        var = Application.CountIf(Range("K1:K" & Rows.count), Range("K" & Target.row).Value)
            If 1 < var Then
            
                 MsgBox "* Duplicate Pan value is not allowed"
                 Range("K" & Target.row).Value = ""
                 Range(Target.Address).Value = ""
            End If
        End If
    End If
   If Target.Column = 12 Then
   
        If Not CheckDateBefore1(Range("L" & Target.row).Value) Then
            MsgBox "* 01/04/2022 to 31/03/2023 is allowed in date of donation"
            Target.Value = ""
        End If
        
        If Not ChkMaxDOBDate(Range("L" & Target.row).Value, "31/03/2023") Then
            MsgBox "* 01/04/2022 to 31/03/2023 is allowed in date of donation"
            Target.Value = ""
        End If
        
'          If Len(Range("M" & Target.Row).Value) > 0 Then
'            If CDate("31/05/2021") >= CDate(Range("L" & Target.Row).Value) Then
'                If Range("M" & Target.Row).Value > 10000 Then
'                    MsgBox "* Please note that the donation made in cash more than 10000 on or before 31.05.2021 will not be eligible for 80GGA deduction."
'                    'Range("M" & Target.Row).Value = ""
'                End If
'            End If
'         End If

          If Len(Range("M" & Target.row).Value) > 0 Then
            If CDate("01/04/2022") <= CDate(Range("L" & Target.row).Value) Then
                If Range("M" & Target.row).Value > 2000 Then

                    MsgBox "* If Donation in cash is more than Rs. 2000, 80GGA Deduction is not allowed"
                    Range("M" & Target.row).Value = ""
                End If
            End If
         End If
         
    End If
    
     If Target.Column = 13 Then
'         If Len(Target.Value) > 0 Then
'            If CDate("31/05/2021") >= CDate(Range("L" & Target.Row).Value) Then
'                If Target.Value > 10000 Then
'                    MsgBox "* Please note that the donation made in cash more than 10000 on or before 31.05.2021 will not be eligible for 80GGA deduction."
'                    'Target.Value = ""
'                End If
'            End If
'         End If
          If Len(Target.Value) > 0 Then
            If CDate("01/04/2022") <= CDate(Range("L" & Target.row).Value) Then
                If Target.Value > 2000 Then

                    MsgBox "* If Donation in cash is more than Rs. 2000, 80GGA Deduction is not allowed"
                    'Target.Value = ""
                End If
            End If
         End If
         If Len(Range("N" & Target.row).Value) > 0 And Len(Range("M" & Target.row).Value) = 0 Then
'             MsgBox "* Date of donation is only applicable for donation in dash."
'             Range("L" & Target.Row).Value = ""
          End If
     End If
'ramya12----------------------------------------------------------------------

'Malli------------------------
 
 
'If Not Application.Intersect(Sheet12.Range("State_Code_80GGA"), Sheet12.Range(Target.Address)) Is Nothing Then
'    Dim PIN_targetadd, state_targetadd As String
'      'state_targetadd = Target.address
'      state_targetadd = Replace(Target.Address, "$", "")
'      PIN_targetadd = Replace(state_targetadd, "I", "J")
'
'    If state_Validation(PIN_targetadd, state_targetadd) = False Then Sheet12.Range(Replace(state_targetadd, "I", "J")).Value = ""
'End If
'
'
'If Not Application.Intersect(Sheet12.Range("Pincode_80GGA"), Sheet12.Range(Target.Address)) Is Nothing Then
'     Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
'     Dim ws As Worksheet
'      'pinrng_chk = Target.address
'      pinrng_chk = Replace(Target.Address, "$", "")
'      staterng_chk = Replace(pinrng_chk, "J", "I")
'      'cntryrng_chk = Replace(pinrng_chk, "J", "I")
'      'ZIPrng_chk = Replace(pinrng_chk, "J", "K")
'      Set ws = Worksheets("80GGA")
'      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
'End If
'

 If Not Application.Intersect(Sheet12.Range("Donation_cash_80GGA"), Sheet12.Range(Target.Address)) Is Nothing Then


 End If
 Sheet12.Protect Password:=sPassword

'ramya12---------------------------------------------------------------------------
exit1:
Application.EnableEvents = True
Sheet12.Activate
End Sub

Private Sub Worksheet_Changeold3(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False
Sheet1.Range("taxcheck").Value = ""


                 

            
            
If Not IsEmpty(Trim(Target.Value)) Then

    If Target.Validation.Type = 6 Then
'        Target.Value = UCase(Target.Value)

If Not Intersect(Range("PAN_of_donee_80GGA"), Range(Target.Address)) Is Nothing Then
Range(Target.Address) = UCase(Range(Target.Address))
If Not checkfieldSuperSpecialcharacterDot(Range(Target.Address).Value) Then
fmsgbox "* Invalid PAN. PAN should contain First 5 alphabets, next 4 digits, then alphabet"
Range(Target.Address).Value = ""
End If
End If

        If Not checkfieldSuperSpecialcharacter(Target.Value) Then
            fmsgbox ("* Please don't use Special Characters < and  > " & Chr(13))
            Target.Value = ""
            GoTo endline
        End If
    End If
End If

'For Each cell80G1 In Range("PAN_of_donee_80GGA")
'If Not cell80G1.Value = "" Then
 If Target.Column = 11 Then
    If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Range("K" & Target.row).Value) Then
        If Not Range("M" & Target.row).Value = "" And Not Range("K" & Target.row).Value = "" Then
        
         fmsgbox "* Duplicate Pan value is not allowed"
         Range("K" & Target.row).Value = ""
         Range("M" & Target.row).Value = ""
        End If
    End If
End If

If Target.Column = 13 Then
    If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Range("K" & Target.row).Value) Then
        If Not Range("M" & Target.row).Value = "" And Not Range("K" & Target.row).Value = "" Then
        
         fmsgbox "* Duplicate Pan value is not allowed"
         Range("K" & Target.row).Value = ""
         Range("M" & Target.row).Value = ""
        End If
    End If
End If

'  Exit For
'      End If
       
'    Next cell80G1
    
   If Target.Column = 12 Then
        If Not CheckDateBefore1(Range("L" & Target.row).Value) Then
            fmsgbox "*01/04/2022 to 31/03/2023 is allowed in date of donation"
            Target.Value = ""
        End If
        If Not ChkMaxDOBDate(Range("L" & Target.row).Value, "31/03/2023") Then
            fmsgbox "*01/04/2022 to 31/03/2023 is allowed in date of donation"
            Target.Value = ""
        End If
          If Len(Range("M" & Target.row).Value) > 0 Then
            If CDate("31/05/2022") >= CDate(Range("L" & Target.row).Value) Then
                If Range("M" & Target.row).Value > 10000 Then
                    fmsgbox "* Please note that the donation made in cash more than 10000 on or before 31.05.2022 will not be eligible for 80GGA deduction."
                    Range("M" & Target.row).Value = ""
                End If
            End If
         End If
          If Len(Range("M" & Target.row).Value) > 0 Then
            If CDate("01/06/2022") <= CDate(Range("L" & Target.row).Value) Then
                If Range("M" & Target.row).Value > 2000 Then
                    fmsgbox "* Please note that the donation made in cash more than 2000 after 31.05.2022 will not be eligible for 80GGA deduction."
                    Range("M" & Target.row).Value = ""
                End If
            End If
         End If
         
          If Len(Range("N" & Target.row).Value) > 0 And Len(Range("M" & Target.row).Value) <= 0 Then
             Target.Value = ""
          End If
    End If
    
     If Target.Column = 13 Then
         If Len(Target.Value) > 0 Then
            If CDate("31/05/2022") >= CDate(Range("L" & Target.row).Value) Then
                If Target.Value > 10000 Then
                    fmsgbox "* Please note that the donation made in cash more than 10000 on or before 31.05.2022 will not be eligible for 80GGA deduction."
                    Target.Value = ""
                End If
            End If
         End If
          If Len(Target.Value) > 0 Then
            If CDate("01/06/2022") <= CDate(Range("L" & Target.row).Value) Then
                If Target.Value > 2000 Then
                    fmsgbox "* Please note that the donation made in cash more than 2000 after 31.05.2022 will not be eligible for 80GGA deduction."
                    Target.Value = ""
                End If
            End If
         End If
         
     End If
     
     If Target.Column = 14 Then
        If Range("L" & Target.row).Value <> "" And Range("N" & Target.row).Value <> "" Then
            Range("L" & Target.row).Value = ""
        End If
     End If
     
     If Target.Column = 10 Then
        Dim PinCode As Range
        For Each PinCode In Sheet5.Range("PincodeList")
          Dim code As String
            code = Left(CStr(Target.Value), 3)
            If CStr(PinCode.Value) = CStr(code) Then
                Sheet12.Range("I" & Target.row).Value = Sheet5.Range("V" & PinCode.row).Value
                Exit For
            End If
         Next PinCode
    End If

    
' If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Target.Value) Then
' If Not Range("M" & Target.Row).Value = "" Then
'  fmsgbox "* Duplicate Pan value is not allowed"
'  End If
'  End If

endline:
Application.EnableEvents = True
End Sub

Private Sub Worksheet_Changeold2(ByVal Target As Range)
On Error GoTo exit1

    Sheet1.Range("taxcheck").Value = ""
    Application.EnableEvents = False
    If Target.Validation.Type = 3 Then
         GoTo exit1
    End If

    Target.Value = UCase(Target.Value)
Dim var As Integer


 If Target.Column = 10 Then
        Dim PinCode As Range
        For Each PinCode In Sheet5.Range("PincodeList")
            Dim code As String
            code = Left(CStr(Target.Value), 3)
            If CStr(PinCode.Value) = CStr(code) Then
                Range("J" & Target.row).Value = Sheet5.Range("V" & PinCode.row).Value
                Exit For
            End If
         Next PinCode
    End If
    
    
    If Target.Column = 11 Then
        If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Range("K" & Target.row).Value) Then
            If Not Range("M" & Target.row).Value = "" And Not Range("K" & Target.row).Value = "" Then
            
                'MsgBox "* Duplicate Pan value is not allowed"
                'Range("K" & Target.Row).Value = ""
                'Range("M" & Target.Row).Value = ""
                
            End If
        End If
    End If

    If Target.Column = 13 Then
        If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Range("K" & Target.row).Value) Then
            If Not Range("M" & Target.row).Value = "" And Not Range("K" & Target.row).Value = "" Then
            
             'MsgBox "* Duplicate Pan value is not allowed"
             'Range("K" & Target.Row).Value = ""
             'Range("M" & Target.Row).Value = ""
            End If
        End If
    End If
    
   If Target.Column = 12 Then
        If Not CheckDateBefore1(Range("L" & Target.row).Value) Then
            MsgBox "* 01/04/2022 to 31/03/2023 is allowed in date of donation"
            Target.Value = ""
        End If
        If Not ChkMaxDOBDate(Range("L" & Target.row).Value, "31/03/2023") Then
            MsgBox "* 01/04/2022 to 31/03/2023 is allowed in date of donation"
            Target.Value = ""
        End If
          If Len(Range("M" & Target.row).Value) > 0 Then
            If CDate("31/05/2022") >= CDate(Range("L" & Target.row).Value) Then
                If Range("M" & Target.row).Value > 10000 Then
                    MsgBox "* Please note that the donation made in cash more than 10000 on or before 31.05.2022 will not be eligible for 80GGA deduction."
                    'Range("M" & Target.Row).Value = ""
                End If
            End If
         End If
          If Len(Range("M" & Target.row).Value) > 0 Then
            If CDate("01/06/2022") <= CDate(Range("L" & Target.row).Value) Then
                If Range("M" & Target.row).Value > 2000 Then
                    MsgBox "* Please note that the donation made in cash more than 2000 after 31.05.2022 will not be eligible for 80GGA deduction."
                    'Range("M" & Target.Row).Value = ""
                End If
            End If
         End If
         
    End If
    
     If Target.Column = 13 Then
         If Len(Target.Value) > 0 Then
            If CDate("31/05/2022") >= CDate(Range("L" & Target.row).Value) Then
                If Target.Value > 10000 Then
                    MsgBox "* Please note that the donation made in cash more than 10000 on or before 31.05.2022 will not be eligible for 80GGA deduction."
                    'Target.Value = ""
                End If
            End If
         End If
          If Len(Target.Value) > 0 Then
            If CDate("01/06/2022") <= CDate(Range("L" & Target.row).Value) Then
                If Target.Value > 2000 Then
                    MsgBox "* Please note that the donation made in cash more than 2000 after 31.05.2022 will not be eligible for 80GGA deduction."
                    'Target.Value = ""
                End If
            End If
         End If
         If Len(Range("N" & Target.row).Value) > 0 And Len(Range("M" & Target.row).Value) = 0 Then
             MsgBox "* Date of donation is only applicable for donation in dash."
             Range("L" & Target.row).Value = ""
          End If
     End If
     
     
exit1:
Application.EnableEvents = True
Sheet12.Activate
End Sub


Private Sub Worksheet_Changeold1(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False
Sheet1.Range("taxcheck").Value = ""

If Not IsEmpty(Trim(Target.Value)) Then

    If Target.Validation.Type = 6 Then
'        Target.Value = UCase(Target.Value)

If Not Intersect(Range("PAN_of_donee_80GGA"), Range(Target.Address)) Is Nothing Then
Range(Target.Address) = UCase(Range(Target.Address))
If Not checkfieldSuperSpecialcharacterDot(Range(Target.Address).Value) Then
fmsgbox "* Invalid PAN. PAN should contain First 5 alphabets, next 4 digits, then alphabet"
Range(Target.Address).Value = ""
End If
End If

        If Not checkfieldSuperSpecialcharacter(Target.Value) Then
            fmsgbox ("* Please don't use Special Characters < and  > " & Chr(13))
            Target.Value = ""
            GoTo endline
        End If
    End If
End If

'For Each cell80G1 In Range("PAN_of_donee_80GGA")
'If Not cell80G1.Value = "" Then
 If Target.Column = 11 Then
    If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Range("K" & Target.row).Value) Then
        If Not Range("M" & Target.row).Value = "" And Not Range("K" & Target.row).Value = "" Then
        
         fmsgbox "* Duplicate Pan value is not allowed"
         Range("K" & Target.row).Value = ""
         Range("M" & Target.row).Value = ""
        End If
    End If
End If

If Target.Column = 13 Then
    If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Range("K" & Target.row).Value) Then
        If Not Range("M" & Target.row).Value = "" And Not Range("K" & Target.row).Value = "" Then
        
         fmsgbox "* Duplicate Pan value is not allowed"
         Range("K" & Target.row).Value = ""
         Range("M" & Target.row).Value = ""
        End If
    End If
End If

'  Exit For
'      End If
       
'    Next cell80G1
    
   If Target.Column = 12 Then
        If Not CheckDateBefore1(Range("L" & Target.row).Value) Then
            fmsgbox "* 1/04/2022 to 31/03/2023 is allowed in date of donation"
            Target.Value = ""
        End If
        If Not ChkMaxDOBDate(Range("L" & Target.row).Value, "31/03/2023") Then
            fmsgbox "* 1/04/2022 to 31/03/2023 is allowed in date of donation"
            Target.Value = ""
        End If
          If Len(Range("M" & Target.row).Value) > 0 Then
            If CDate("31/05/2022") >= CDate(Range("L" & Target.row).Value) Then
                If Range("M" & Target.row).Value > 10000 Then
                    fmsgbox "* You cannot donate more than Rs 10000 in cash"
                    Range("M" & Target.row).Value = ""
                End If
            End If
         End If
          If Len(Range("M" & Target.row).Value) > 0 Then
            If CDate("01/06/2022") <= CDate(Range("L" & Target.row).Value) Then
                If Range("M" & Target.row).Value > 2000 Then
                    fmsgbox "* You cannot donate more than Rs 2000 in cash"
                    Range("M" & Target.row).Value = ""
                End If
            End If
         End If
         
          If Len(Range("N" & Target.row).Value) > 0 Then
             Target.Value = ""
          End If
    End If
    
     If Target.Column = 13 Then
         If Len(Target.Value) > 0 Then
            If CDate("31/05/2022") >= CDate(Range("L" & Target.row).Value) Then
                If Target.Value > 10000 Then
                    fmsgbox "* You cannot donate more than Rs 10000 in cash"
                    Target.Value = ""
                End If
            End If
         End If
          If Len(Target.Value) > 0 Then
            If CDate("01/06/2022") <= CDate(Range("L" & Target.row).Value) Then
                If Target.Value > 2000 Then
                    fmsgbox "* You cannot donate more than Rs 2000 in cash"
                    Target.Value = ""
                End If
            End If
         End If
     End If
     
     If Target.Column = 14 Then
        If Range("L" & Target.row).Value <> "" And Range("N" & Target.row).Value <> "" Then
            Range("L" & Target.row).Value = ""
        End If
     End If
     
     If Target.Column = 10 Then
        Dim PinCode As Range
        For Each PinCode In Sheet5.Range("PincodeList")
          Dim code As String
            code = Left(CStr(Target.Value), 3)
            If CStr(PinCode.Value) = CStr(code) Then
                Sheet12.Range("I" & Target.row).Value = Sheet5.Range("V" & PinCode.row).Value
                Exit For
            End If
         Next PinCode
    End If

    
' If 1 < Application.CountIf(Range("PAN_of_donee_80GGA"), Target.Value) Then
' If Not Range("M" & Target.Row).Value = "" Then
'  fmsgbox "* Duplicate Pan value is not allowed"
'  End If
'  End If

endline:
Application.EnableEvents = True
End Sub

Sub TESTT()
Application.EnableEvents = True
End Sub

