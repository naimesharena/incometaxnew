Attribute VB_Name = "Sheet4"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
''Application.OnKey "{TAB}"
''MsgBox ActiveCell.Address
'
'
'If Not Range(ActiveCell.Address).Locked = True Then
'       If ActiveCell.Validation.Type = 6 Then
'    Application.OnKey "%{DOWN}", ""
'    'Sheet1.Range(Target.Address).Value = ""
'ElseIf ActiveCell.Validation.Type = 1 Then
' Application.OnKey "%{DOWN}", ""
'ElseIf ActiveCell.Validation.Type = 3 Then
'
'    Application.OnKey "%{DOWN}"
'End If
'Else
'Application.OnKey "%{DOWN}", ""
'End If
'End Sub

Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False
ActiveWindow.DisplayHeadings = False
    ActiveSheet.Shapes("Button 4").Left = Range("Sch80G.Button1").Left + 1
    ActiveSheet.Shapes("Button 4").Top = Range("Sch80G.Button1").Top + 10


    ActiveSheet.Shapes("Button 5").Left = Range("Sch80G.Button2").Left + 1
    ActiveSheet.Shapes("Button 5").Top = Range("Sch80G.Button2").Top + 10

    ActiveSheet.Shapes("Button 6").Left = Range("Sch80G.Button3").Left + 1
    ActiveSheet.Shapes("Button 6").Top = Range("Sch80G.Button3").Top + 10

    ActiveSheet.Shapes("Button 7").Left = Range("Sch80G.Button4").Left + 1
    ActiveSheet.Shapes("Button 7").Top = Range("Sch80G.Button4").Top + 10

endline:
Application.EnableEvents = True
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo exit1

    sPassword = EfilingCommon.getmsgstate
    Sheet1.Unprotect Password:=sPassword
    Sheet1.Range("taxcheck").Value = ""
    Sheet1.Protect Password:=sPassword
    Application.EnableEvents = False
    If Target.Validation.Type = 3 Then
         GoTo exit1
    End If
    
    
    'Ankita_03/04/2026==========
    
    
    
'Change.27.01.2023.102.80GARN0.2
    Sheet4.Unprotect Password:=sPassword
        If Not Application.Intersect(Sheet4.Range("Per5080G.ArnNbr"), Sheet4.Range(Target.Address)) Is Nothing Then
            If Range(Target.Address).Value <> "" Then
                    If Not isAlphanumeric(Range(Target.Address).Value) Then
                            MsgBox ("Only Alphanumeric and 25 Characters allowed.")
                           Range(Target.Address).Value = ""
                           Range(Target.Address).Select
                    End If
            End If
        End If
    
'End Chhange==
    
    Target.Value = UCase(Target.Value)
    
    Dim var As Integer
    If Target.Column = 10 Then
        If Trim(Range("J" & Target.row).Value) = "AAAAR1077P" Then
'            If 1 < Application.CountIf(Range("Per10080G.DoneePAN"), Range("J" & Target.row).Value) Or 1 < Application.CountIf(Range("PerNO5080G.DoneePAN"), Range("J" & Target.row).Value) Or 1 < Application.CountIf(Range("PerYES10080G.DoneePAN"), Range("J" & Target.row).Value) Then 'Or 1 < Application.CountIf(Range("Per5080G.DoneePAN"), Range("J" & Target.row).Value)
'                 MsgBox "* Duplicate Pan AAAAR1077P is not allowed for same table"
'                 Range(Target.Address).Value = ""
'                 Range("K" & Target.row).Value = ""
'                 Range("L" & Target.row).Value = ""
'            End If
        Else
        
        'TABLE A
'        If Not Application.Intersect(Sheet4.Range("Per10080G.DoneePAN"), Range(Target.Address)) Is Nothing Then
'           var = Application.CountIf(Sheet4.Range("Per10080G.DoneePAN"), Range("j" & Target.row).Value)
'
'           If var > 1 Then
'              MsgBox "*Please enter aggregate amount of donation made in cash to same Donee once"
'              Range(Target.Address).MergeArea.ClearContents
'           ElseIf Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.row).Value) > 1 Then
'              MsgBox "Same PAN cannot be entered for different donation type"
'              Range(Target.Address).MergeArea.ClearContents
'           End If
'       End If
        
        'TABLE B
'       If Not Application.Intersect(Sheet4.Range("PerNO5080G.DoneePAN"), Range(Target.Address)) Is Nothing Then
'           var = Application.CountIf(Sheet4.Range("PerNO5080G.DoneePAN"), Range("j" & Target.row).Value)
'
'           If var > 1 Then
'              MsgBox "*Please enter aggregate amount of donation made in cash to same Donee once"
'              Range(Target.Address).MergeArea.ClearContents
'           ElseIf Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.row).Value) > 1 Then
'              MsgBox "Same PAN cannot be entered for different donation type"
'              Range(Target.Address).MergeArea.ClearContents
'           End If
'       End If
'       'TABLE C
'       If Not Application.Intersect(Sheet4.Range("PerYES10080G.DoneePAN"), Range(Target.Address)) Is Nothing Then
'           var = Application.CountIf(Sheet4.Range("PerYES10080G.DoneePAN"), Range("j" & Target.row).Value)
'
'           If var > 1 Then
'              MsgBox "*Please enter aggregate amount of donation made in cash to same Donee once"
'              Range(Target.Address).MergeArea.ClearContents
'           ElseIf Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.row).Value) > 1 Then
'              MsgBox "Same PAN cannot be entered for different donation type"
'              Range(Target.Address).MergeArea.ClearContents
'           End If
'       End If
       'TABLE D
'        If Not Application.Intersect(Sheet4.Range("Per5080G.DoneePAN"), Range(Target.Address)) Is Nothing Then
'           var = Application.CountIf(Sheet4.Range("Per5080G.DoneePAN"), Range("j" & Target.row).Value)
'
'           If var > 1 Then
'              MsgBox "*Please enter aggregate amount of donation made in cash to same Donee once"
'              Range(Target.Address).ClearContents
'           ElseIf Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.row).Value) > 1 Then
'              MsgBox "Same PAN cannot be entered for different donation type"
'              Range(Target.Address).ClearContents
'           End If
'       End If
        
        'Ankita_06/03/2026==============
'If Not Application.Intersect(Sheet4.Range("Per10080G.DonationAmt"), Sheet4.Range(Target.Address)) Is Nothing Then
' If Target.Value > 0 And Not IsEmpty(Target.Value) Then
'        Sheet4.Range(Replace(Target.Address, "L", "M")).ClearContents
'        Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = True
'        Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
'
'        Sheet4.Range(Replace(Target.Address, "L", "N")).ClearContents
'        Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = True
'        Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HD8D8D8"
'
'        Sheet4.Range(Replace(Target.Address, "L", "O")).ClearContents
'        Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = True
'        Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HD8D8D8"
'
'       Else
'
'        Sheet4.Range(Replace(Target.Address, "L", "M")).ClearContents
'        Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = False
'        Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
'
'        Sheet4.Range(Replace(Target.Address, "L", "N")).ClearContents
'        Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = False
'        Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HCCFFCC"
'
'        Sheet4.Range(Replace(Target.Address, "L", "O")).ClearContents
'        Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = False
'        Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HCCFFCC"
'
'    End If
'End If

If Not Application.Intersect(Sheet4.Range("Per10080G.DonationAmtOther"), Sheet4.Range(Target.Address)) Is Nothing Then
   If (Target.Value > 0 And Not IsEmpty(Target.Value)) Or (Not IsEmpty(Range(Replace(Target.Address, "M", "N")).Value) And Len(Range(Replace(Target.Address, "M", "N")).Value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "M", "O")).Value) And Len(Range(Replace(Target.Address, "M", "O")).Value) > 0) Then
        Sheet4.Unprotect Password:=sPassword
        Sheet4.Range(Replace(Target.Address, "M", "N")).ClearContents
        Sheet4.Range(Replace(Target.Address, "M", "N")).Locked = True
        Sheet4.Range(Replace(Target.Address, "M", "N")).Interior.Color = "&HD8D8D8"
    Else
        Sheet4.Range(Replace(Target.Address, "M", "N")).Locked = False
        Sheet4.Range(Replace(Target.Address, "M", "N")).Interior.Color = "&HCCFFCC"
    End If
End If

        '===============================
        'Change-ANK, 28.01.2023
'            If Application.Intersect(Sheet4.Range("Per5080G.DoneePAN"), Range(Target.Address)) Is Nothing Then
'                var = Application.CountIf(Range("Per10080G.DoneePAN"), Range("J" & Target.row).Value) + Application.CountIf(Range("PerNO5080G.DoneePAN"), Range("J" & Target.row).Value) + Application.CountIf(Range("PerYES10080G.DoneePAN"), Range("J" & Target.row).Value)
'                If 1 <= var And Range("J" & Target.row).Value <> "" Then
'                    'MsgBox "* Same PAN cannot be entered for different donation type"
'                    'Malli AY_2024_25 Change Added
'                    MsgBox "* Please enter aggregate amount of donation made in cash to same Donee once"
'                    Range(Target.Address).Value = ""
''                    Range("K" & Target.row).Value = ""
''                    Range("L" & Target.row).Value = ""
'                End If
'
'            Else
'                var = Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.row).Value)
'                If 1 < var Then
'                     'MsgBox "* Same PAN cannot be entered for different donation type"
'                     'Malli AY_2024_25 Change Added
'                     MsgBox "* Please enter aggregate amount of donation made in cash to same Donee once"
'                     Range(Target.Address).Value = ""
'                     Range("K" & Target.row).Value = ""
'                     Range("L" & Target.row).Value = ""
'                End If
'            End If
        End If
    End If
'    If Target.Column = 10 Then
'        If (Not Range("K" & Target.Row).Value = "") And Trim(Range("J" & Target.Row).Value) <> "AAAAR1077P" Then
'        var = Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.Row).Value)
'            If 1 < var Then
'
'                 MsgBox "* Duplicate Pan value is not allowed"
'                 Range(Target.Address).Value = ""
'                 Range("K" & Target.Row).Value = ""
'            End If
'        End If
'    End If
'
'    If Target.Column = 11 Then
'        If (Not Range("K" & Target.Row).Value = "") And Trim(Range("J" & Target.Row).Value) <> "AAAAR1077P" Then
'        var = Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.Row).Value)
'            If 1 < var Then
'
'                 MsgBox "* Duplicate Pan value is not allowed"
'                 Range("J" & Target.Row).Value = ""
'                 Range(Target.Address).Value = ""
'            End If
'        End If
'    End If
' Change.25.01.2023.102.80GA
'If Target.Column = 11 Then
'
'        If Not Range("J" & Target.row).Value = "" Then
'
'        var = Application.CountIf(Range("K1:K" & Rows.count), Range("K" & Target.row).Value)
''           If Range("K" & Target.Row).Value > 2000 Then
''                  MsgBox "* If Donation in cash is more than Rs. 2000, 80G Deduction is not allowed"
''                 Range("K" & Target.Row).Value = ""
''                 Range(Target.Address).Value = ""
''            End If
'        End If
'End If

'Ankita_07/03/2026=========
'For A
If Not Application.Intersect(Sheet4.Range("Per10080G.IFSC"), Sheet4.Range(Target.Address)) Is Nothing Then

Dim IfsCode As Range
Set IfsCode = Sheet4.Range("Per10080G.IFSC")

    Dim sourceAddress, Address As Variant
    Dim rangecell As Range
    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
         For Each rangecell In Sheet4.Range("Per10080G.IFSC")
            rangecell.Value = UCase(rangecell.Value)
     
     If rangecell.Value <> "" Then
           
           If Not CheckIFSC(UCase(Target.Value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet4.Unprotect Password:=getmsgstate
                    Target.Value = ""
                    GoTo exit1
           
            ElseIf Not ValidateIFSCList(UCase(Target.Value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Target.Value = ""
                    GoTo exit1
            Else:
                       Address = rangecell.Address
                     End If
                If Mid(rangecell, 1, 4) <> "GIFT" Then
                    If (Sheet4.Range(Address).Value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress, "K", "J")
                     Sheet4.Range(Address).Value = ""
                    GoTo exit1
                    End If
                End If
                
            End If
    Next
    GoTo exit1
End If
End If
'------

'For B
If Not Application.Intersect(Sheet4.Range("PerNO5080G.IFSC"), Sheet4.Range(Target.Address)) Is Nothing Then

Dim IfsCode1 As Range
Set IfsCode1 = Sheet4.Range("PerNO5080G.IFSC")

    Dim sourceAddress1, Address1 As Variant
    Dim rangecell1 As Range
'    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode1, Range(Target.Address)) Is Nothing Then
         For Each rangecell1 In Sheet4.Range("PerNO5080G.IFSC")
            rangecell1.Value = UCase(rangecell1.Value)
     
     If rangecell1.Value <> "" Then
           
           If Not CheckIFSC(UCase(Target.Value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet4.Unprotect Password:=getmsgstate
                    Target.Value = ""
                    GoTo exit1
           
            ElseIf Not ValidateIFSCList(UCase(Target.Value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Target.Value = ""
                    GoTo exit1
     
            Else:
                       Address1 = rangecell1.Address
                     
                     End If
                
                If Mid(rangecell1, 1, 4) <> "GIFT" Then
                    If (Sheet4.Range(Address1).Value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress1, "K", "J")
                     Sheet4.Range(Address1).Value = ""
                    GoTo exit1
                    End If
                End If
                
            End If
    Next
    GoTo exit1
End If
End If
'------
'For C
If Not Application.Intersect(Sheet4.Range("PerYES10080G.IFSC"), Sheet4.Range(Target.Address)) Is Nothing Then

Dim IfsCode2 As Range
Set IfsCode2 = Sheet4.Range("PerYES10080G.IFSC")

    Dim sourceAddress2, Address2  As Variant
    Dim rangecell2 As Range
'    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode2, Range(Target.Address)) Is Nothing Then
         For Each rangecell2 In Sheet4.Range("PerYES10080G.IFSC")
            rangecell2.Value = UCase(rangecell2.Value)
     
     If rangecell2.Value <> "" Then
           
           If Not CheckIFSC(UCase(Target.Value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet4.Unprotect Password:=getmsgstate
                    Target.Value = ""
                    GoTo exit1
           
            ElseIf Not ValidateIFSCList(UCase(Target.Value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Target.Value = ""
                    GoTo exit1
     
            Else:
                       Address2 = rangecell2.Address
                     
                     End If
                
                If Mid(rangecell, 1, 4) <> "GIFT" Then
                    If (Sheet4.Range(Address2).Value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress2, "K", "J")
                     Sheet4.Range(Address2).Value = ""
                    GoTo exit1
                    End If
                End If
                
            End If
    Next
    GoTo exit1
End If
End If
'------
'For D
If Not Application.Intersect(Sheet4.Range("Per5080G.IFSC"), Sheet4.Range(Target.Address)) Is Nothing Then

Dim IfsCode3 As Range
Set IfsCode3 = Sheet4.Range("Per5080G.IFSC")

    Dim sourceAddress3, Address3 As Variant
    Dim rangecell3 As Range
'    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode3, Range(Target.Address)) Is Nothing Then
         For Each rangecell3 In Sheet4.Range("Per5080G.IFSC")
            rangecell3.Value = UCase(rangecell3.Value)
     
     If rangecell3.Value <> "" Then
           
           If Not CheckIFSC(UCase(Target.Value)) Then
                    fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet4.Unprotect Password:=getmsgstate
                    Target.Value = ""
                    GoTo exit1
           
            ElseIf Not ValidateIFSCList(UCase(Target.Value)) Then
                    fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Target.Value = ""
                    GoTo exit1
     
            Else:
                       Address3 = rangecell3.Address
                     
                     End If
                
                If Mid(rangecell3, 1, 4) <> "GIFT" Then
                    If (Sheet4.Range(Address3).Value = "") Then
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress3, "K", "J")
                     Sheet4.Range(Address3).Value = ""
                    GoTo exit1
                    End If
                End If
                
            End If
    Next
    GoTo exit1
End If
End If
'------

'==========================

If Target.Column = 12 Then

        If Not Range("L" & Target.row).Value = "" Then
        
'        If Range("L" & Target.row).Value > 2000 Then
'          MsgBox "* If Donation in cash is more than Rs. 2000, 80G Deduction is not allowed"
'        End If
        
       ' var = Application.CountIf(Range("L1:L" & Rows.count), Range("L" & Target.row).Value)
'           If Range("K" & Target.Row).Value > 2000 Then
'                  MsgBox "* If Donation in cash is more than Rs. 2000, 80G Deduction is not allowed"
'                 Range("K" & Target.Row).Value = ""
'                 Range(Target.Address).Value = ""
'            End If
        End If
End If
'End Change=====
'    If Target.Column = 9 Then
'           Dim PinCode As Range
'           Dim found As Variant
'            If (Target.Value <> "") Then
'           found = False
'
'           For Each PinCode In Sheet5.Range("PincodeList")
'               Dim code As String
'               code = Left(CStr(Target.Value), 3)
'               If CStr(PinCode.Value) = CStr(code) Then
'                   Range("H" & Target.row).Value = Sheet5.Range("V" & PinCode.row).Value
'                   found = True
'                   Exit For
'               End If
'            Next PinCode
'            If found = False Then
'            'MsgBox "Enter correct Pincode"
'            MsgBox "Invalid Pin code"
'            Target.Value = ""
'            End If
'            End If
'            End If
   '=====================================================================
    
 ' by Sai 12/09/2024================================================================
   If Target.Column = 8 Then
    
    Dim TargtVal, PIN_targetadd, state_targetadd As Variant
      state_targetadd = Target.Address
      state_targetadd = Replace(state_targetadd, "$", "")
      PIN_targetadd = Replace(state_targetadd, "H", "I")
    
    Call state_Validation(PIN_targetadd, state_targetadd)
   End If
   
  If Target.Column = 9 Then
            
            If (Target.Value <> "") Then
            
            Dim pinrng_chk, staterng_chk  As Variant
 
      pinrng_chk = Target.Address
      pinrng_chk = Replace(pinrng_chk, "$", "")
      staterng_chk = Replace(pinrng_chk, "I", "H")
       
      Call PinState_codes_validation_80G(pinrng_chk, staterng_chk)
           
             End If
      End If
      
'End by sai===============================================
    
'        If Intersect(Target, Range("Per10080G.PinCode")) Then
'            If Not (ValidatePinCode_Per10080G()) Then
'                MsgBox "Invalid PinCode"
'            End If
'            GoTo exit1
'        End If
'
'        If Intersect(Target, Range("PerNO5080G.PinCode")) Then
'            If Not (ValidatePinCode_PerNO5080G()) Then
'                MsgBox "Invalid PinCode"
'            End If
'            GoTo exit1
'        End If
'
'        If (getRangeName(Target) = "Per5080G.DoneePAN") Then
'            If Not (ValidateDoneePAN_Per5080G()) Then
'                MsgBox "Invalid DoneePAN"
'            End If
'            GoTo exit1
'        End If
'
'        If (getRangeName(Target) = "Per5080G.PinCode") Then
'            If Not (ValidatePinCode_Per5080G()) Then
'                MsgBox "Invalid PinCode"
'            End If
'            GoTo exit1
'        End If

If Not Intersect(Range("Per10080G.DonationAmt"), Range(Target.Address)) Is Nothing Then
  Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).Value = 0 Then

            Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"

            Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HCCFFCC"
            
            
            Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HCCFFCC"
            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
        
            Sheet4.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "L", "N")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "L", "O")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HD8D8D8"
            Sheet4.Protect Password:=getmsgstate

    End If
End If

If Not Intersect(Range("PerNO5080G.DonationAmt"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).Value = 0 Then
        
            Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"


            Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HCCFFCC"
            
            
            Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HCCFFCC"
            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
        
            Sheet4.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
            
            Sheet4.Range(Replace(Target.Address, "L", "N")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "L", "O")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HD8D8D8"
            
            Sheet4.Protect Password:=getmsgstate

    End If
End If

If Not Intersect(Range("PerYES10080G.DonationAmt"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).Value = 0 Then

            Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"

            Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HCCFFCC"
            
            Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HCCFFCC"
            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
        
            Sheet4.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "L", "N")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "L", "O")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HD8D8D8"
            Sheet4.Protect Password:=getmsgstate

    End If
End If

If Not Intersect(Range("Per5080G.DonationAmt"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).Value = 0 Then
        
            Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HCCFFCC"
            
            Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HCCFFCC"
            
            Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = False
            Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HCCFFCC"
            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
        
            Sheet4.Range(Replace(Target.Address, "L", "M")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "M")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "M")).Interior.Color = "&HD8D8D8"
            
            Sheet4.Range(Replace(Target.Address, "L", "N")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "N")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "N")).Interior.Color = "&HD8D8D8"

            Sheet4.Range(Replace(Target.Address, "L", "O")).ClearContents
            Sheet4.Range(Replace(Target.Address, "L", "O")).Locked = True
            Sheet4.Range(Replace(Target.Address, "L", "O")).Interior.Color = "&HD8D8D8"
            Sheet4.Protect Password:=getmsgstate

    End If
End If

'Ankita===
If Not Intersect(Range("Per10080G.DonationAmtOther"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).Value = 0 Then
        
            Sheet4.Range(Replace(Target.Address, "M", "L")).Locked = False
            Sheet4.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"

            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
            Sheet4.Range(Replace(Target.Address, "M", "L")).ClearContents
            Sheet4.Range(Replace(Target.Address, "M", "L")).Locked = True
            Sheet4.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
            Sheet4.Protect Password:=getmsgstate

    End If
End If


If Not Intersect(Range("PerNO5080G.DonationAmtOther"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).Value = 0 Then
        
            Sheet4.Range(Replace(Target.Address, "M", "L")).Locked = False
            Sheet4.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"

            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
            Sheet4.Range(Replace(Target.Address, "M", "L")).ClearContents
            Sheet4.Range(Replace(Target.Address, "M", "L")).Locked = True
            Sheet4.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
            Sheet4.Protect Password:=getmsgstate

    End If
End If


If Not Intersect(Range("PerYES10080G.DonationAmtOther"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).Value = 0 Then
        
            Sheet4.Range(Replace(Target.Address, "M", "L")).Locked = False
            Sheet4.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"

            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
            Sheet4.Range(Replace(Target.Address, "M", "L")).ClearContents
            Sheet4.Range(Replace(Target.Address, "M", "L")).Locked = True
            Sheet4.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
            Sheet4.Protect Password:=getmsgstate

    End If
End If


If Not Intersect(Range("Per5080G.DonationAmtOther"), Range(Target.Address)) Is Nothing Then
Sheet4.Unprotect Password:=getmsgstate
        If Sheet4.Range(Target.Address).Value = 0 Then
        
            Sheet4.Range(Replace(Target.Address, "M", "L")).Locked = False
            Sheet4.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HCCFFCC"

            Sheet4.Protect Password:=getmsgstate
        Else
        Sheet4.Unprotect Password:=getmsgstate
            Sheet4.Range(Replace(Target.Address, "M", "L")).ClearContents
            Sheet4.Range(Replace(Target.Address, "M", "L")).Locked = True
            Sheet4.Range(Replace(Target.Address, "M", "L")).Interior.Color = "&HD8D8D8"
            Sheet4.Protect Password:=getmsgstate

    End If
End If

'Ankita_31/03/2026=======

'Dim var As Integer
Dim isPAN As Boolean
Dim isCash As Boolean
If Target.Column = 10 Or Target.Column = 12 Then
isPAN = _
    Not Application.Intersect(Target, Sheet4.Range("Per10080G.DoneePAN")) Is Nothing _
Or Not Application.Intersect(Target, Sheet4.Range("PerNO5080G.DoneePAN")) Is Nothing _
Or Not Application.Intersect(Target, Sheet4.Range("Per5080G.DoneePAN")) Is Nothing _
Or Not Application.Intersect(Target, Sheet4.Range("PerYES10080G.DoneePAN")) Is Nothing
isCash = _
    Not Application.Intersect(Target, Sheet4.Range("Per10080G.DonationAmt")) Is Nothing _
Or Not Application.Intersect(Target, Sheet4.Range("PerNO5080G.DonationAmt")) Is Nothing _
Or Not Application.Intersect(Target, Sheet4.Range("PerYES10080G.DonationAmt")) Is Nothing _
Or Not Application.Intersect(Target, Sheet4.Range("Per5080G.DonationAmt")) Is Nothing
        If Trim(Range("J" & Target.row).Value) = "AAAAR1077P" Then
'            If 1 < Application.CountIf(Range("Per10080G.DoneePAN"), Range("J" & Target.row).Value) Or 1 < Application.CountIf(Range("PerNO5080G.DoneePAN"), Range("J" & Target.row).Value) Or 1 < Application.CountIf(Range("PerYES10080G.DoneePAN"), Range("J" & Target.row).Value) Then 'Or 1 < Application.CountIf(Range("Per5080G.DoneePAN"), Range("J" & Target.row).Value)
'                 MsgBox "* Duplicate Pan AAAAR1077P is not allowed for same table"
'                 Range(Target.Address).Value = ""
'                 Range("K" & Target.row).Value = ""
'                 Range("L" & Target.row).Value = ""
'            End If
   ElseIf _
       isPAN Or isCash Then
        Dim rngCash As Range
        Dim rngPAN As Range
        Dim rngVal As Range
        Dim cntOnes As Long
        Dim cell As Range

        'Combine all cash blocks
        Set rngCash = Union( _
            Sheet4.Range("Per10080G.DonationAmt"), _
            Sheet4.Range("PerNO5080G.DonationAmt"), _
            Sheet4.Range("PerYES10080G.DonationAmt"), _
            Sheet4.Range("Per5080G.DonationAmt") _
        )
        Set rngPAN = Union( _
            Sheet4.Range("Per10080G.DoneePAN"), _
            Sheet4.Range("PerNO5080G.DoneePAN"), _
            Sheet4.Range("Per5080G.DoneePAN"), _
            Sheet4.Range("PerYES10080G.DoneePAN") _
        )
        If isPAN Then
            Set ranVal = rngPAN
        ElseIf isCash Then
            Set ranVal = rngCash
        End If
        Application.EnableEvents = False
        Application.Calculate
        'Column V
        cntOnes = Application.WorksheetFunction.CountIf(Sheet4.Range("Z:Z"), 1)
        If cntOnes > 1 Then
            'Clear
            For Each cell In Application.Intersect(Target, ranVal)
                If isPAN Then
                    If cell.Value <> "" Then
                        cell.MergeArea.ClearContents
                    End If
                    ElseIf isCash Then
                    If cell.Value <> "" And cell.Value <> 0 Then
                        cell.ClearContents
                    End If
                End If
            Next cell
            MsgBox "Same PAN cannot be entered for ""Donation in cash"""
        End If
        Application.EnableEvents = True
End If
End If

Dim comb90Val As Long

If Target.Column = 10 Then   'PAN column
    comb90Val = Sheet4.Cells(Target.row, "U").Value
    If Trim(Target.Value) <> "" Then
        If Trim(Range("J" & Target.row).Value) = "AAAAR1077P" Then
        ElseIf comb90Val > 0 Then
            Application.EnableEvents = False
            MsgBox """Same PAN cannot be entered in other block"""
            Target.Value = ""
            Application.EnableEvents = True
        End If
    End If
End If

 
 'Ankita_02/04/2026==
If Target.Column = 12 Then
    If Range("L" & Target.row).Value > 2000 Then
        fmsgbox "* ""If Donation in cash is more than Rs. 2000, 80G Deduction is not allowed"""
        Range(Target.Address).Value = ""
    End If
End If


'========================

'==


exit1:
Application.EnableEvents = True
'Change.27.01.2023.102.80GARN0.2A
Sheet4.Protect Password:=sPassword
'End Chhange==

Sheet4.Activate
End Sub

Private Sub Worksheet_Changeold1(ByVal Target As Range)

On Error GoTo endline7

    Dim cell80G1 As Range
    Dim cell80G2 As Range
    Dim cell80G3 As Range
    Dim cell80G4 As Range
    Dim DonationAmountA As Range
    Dim DonationAmountB As Range
    Dim DonationAmountC As Range
    Dim DonationAmountD As Range

Application.ScreenUpdating = False
Application.EnableEvents = False

 'To Change TAn number to upper Case of 1st Grid
Sheet1.Range("taxcheck").Value = ""
If Not Intersect(Range("Per10080G.DoneePAN"), Range(Target.Address)) Is Nothing Then
   Range(Target.Address).Value = UCase(Range(Target.Address))
End If

If Not Intersect(Range("PerNO5080G.DoneePAN"), Range(Target.Address)) Is Nothing Then
   Range(Target.Address).Value = UCase(Range(Target.Address))
End If

If Not Intersect(Range("PerYES10080G.DoneePAN"), Range(Target.Address)) Is Nothing Then
   Range(Target.Address).Value = UCase(Range(Target.Address))
End If

If Not Intersect(Range("Per5080G.DoneePAN"), Range(Target.Address)) Is Nothing Then
   Range(Target.Address).Value = UCase(Range(Target.Address))
End If



'For Each cell80G1 In Range("Per10080G.DoneePAN")
'If Not cell80G1.Value = "" Then
' If 1 < Application.CountIf(Range("Per10080G.DoneePAN"), cell80G1.Value) Then
' If Not Range("K" & cell80G1.Row).Value = "" Then
'  fmsgbox "* Duplicate Pan value is not allowed"
'  Exit For
'       End If
'        End If
'       End If
'    Next cell80G1
'
'    For Each cell80G1 In Range("PerNO5080G.DoneePAN")
'    If Not cell80G1.Value = "" Then
' If 1 < Application.CountIf(Range("PerNO5080G.DoneePAN"), cell80G1.Value) Then
' If Not Range("K" & cell80G1.Row).Value = "" Then
'  fmsgbox "* Duplicate Pan value is not allowed"
'  Exit For
'  End If
'       End If
'       End If
'    Next cell80G1
'
'    For Each cell80G1 In Range("PerYES10080G.DoneePAN")
'    If Not cell80G1.Value = "" Then
' If 1 < Application.CountIf(Range("PerYES10080G.DoneePAN"), cell80G1.Value) Then
' If Not Range("K" & cell80G1.Row).Value = "" Then
'  fmsgbox "* Duplicate Pan value is not allowed"
'  Exit For
'  End If
'       End If
'       End If
'    Next cell80G1
'
'    For Each cell80G1 In Range("Per5080G.DoneePAN")
'    If Not cell80G1.Value = "" Then
' If 1 < Application.CountIf(Range("Per5080G.DoneePAN"), cell80G1.Value) Then
' If Not Range("K" & cell80G1.Row).Value = "" Then
'  fmsgbox "* Duplicate Pan value is not allowed"
'  Exit For
'  End If
'       End If
'       End If
'    Next cell80G1

'For i = 1 To Rows.Count
'     If Not Range("J" & i).Value = "" Then
'        If Not Range("J" & i).Value = "PAN of Donee" Then
'            If Not Range("K" & i).Value = "" Then
'                If 1 < Application.CountIf(Range("Z1:Z" & Rows.Count), Range("J" & i).Value, Range("K" & i).Value, "<>") Then
'                    fmsgbox "* Duplicate Pan value is not allowed"
'                    Exit For
'                Else
'                    Range("Z" & i).Value = Range("J" & i).Value
'                End If
'            End If
'        End If
'    End If
'Next i

'For i = 1 To Rows.Count
'    Range("Z" & i).Value = ""
'Next i

'If Target.Column = 10 Then
'    If Not Range("K" & Target.Row).Value = "" Then
'        If 1 >= Application.CountIf(Range("Z1:Z" & Rows.Count), Target.Value) Then
'             fmsgbox "* Duplicate Pan value is not allowed"
'             Range(Target.Address).Value = ""
'             Range("K" & Target.Row).Value = ""
'        Else
'             Range("Z" & i).Value = Target.Value
'        End If
'    End If
'End If

'If Target.Column = 11 Then
'    If Not Range("K" & Target.Row).Value = "" Then
'        If 1 >= Application.CountIf(Range("Z1:Z" & Rows.Count), Range("J" & Target.Row).Value) Then
'             fmsgbox "* Duplicate Pan value is not allowed"
'             Range("J" & Target.Row).Value = ""
'             Range(Target.Address).Value = ""
'        Else
'             Range("Z" & i).Value = Range("J" & Target.Row).Value
'        End If
'    End If
'End If
If Target.Column = 10 Then
    If Not Range("K" & Target.row).Value = "" Then
        If 1 < Application.CountIf(Range("J1:J" & Rows.count), Target.Value) Then
        
             fmsgbox "* Same PAN cannot be entered for different donation type"
             Range(Target.Address).Value = ""
             Range("K" & Target.row).Value = ""
        End If
    End If
End If

If Target.Column = 11 Then
    If Not Range("K" & Target.row).Value = "" Then
    Dim var As Integer
    var = Application.CountIf(Range("J1:J" & Rows.count), Range("J" & Target.row).Value)
        If 1 < var Then
        
             fmsgbox "* Same PAN cannot be entered for different donation type"
             Range("J" & Target.row).Value = ""
             Range(Target.Address).Value = ""
        End If
    End If
End If

    If Target.Column = 9 Then
        Dim PinCode As Range
        For Each PinCode In Sheet5.Range("PincodeList")
            Dim code As String
            code = Left(CStr(Target.Value), 3)
            If CStr(PinCode.Value) = CStr(code) Then
                Sheet4.Range("H" & Target.row).Value = Sheet5.Range("V" & PinCode.row).Value
                Exit For
            End If
         Next PinCode
    End If
    

  '  For Each cell80G1 In Range("Per10080G.DoneePAN")
  '      If Len(cell80G1) > 0 Then cell80G1 = UCase(cell80G1)
  '  Next cell80G1


   ' For Each cell80G2 In Range("PerNO5080G.DoneePAN")
   '     If Len(cell80G2) > 0 Then cell80G2 = UCase(cell80G2)
   ' Next cell80G2

   ' For Each cell80G3 In Range("PerYES10080G.DoneePAN")
   '     If Len(cell80G3) > 0 Then cell80G3 = UCase(cell80G3)
   ' Next cell80G3


   ' For Each cell80G4 In Range("Per5080G.DoneePAN")
   '     If Len(cell80G4) > 0 Then cell80G4 = UCase(cell80G4)
   ' Next cell80G4


   ' For Each DonationAmountA In Range("Per10080G.DonationAmt")
   ''     If Len(DonationAmountA) > 14 Or Not IsNumeric(DonationAmountA) Then DonationAmountA = ""
   ' Next
   '
   ' For Each DonationAmountB In Range("PerNO5080G.DonationAmt")
   '     If Len(DonationAmountB) > 14 Or Not IsNumeric(DonationAmountB) Then DonationAmountB = ""
   ' Next
   '
   ' For Each DonationAmountC In Range("PerYES10080G.DonationAmt")
   '     If Len(DonationAmountC) > 14 Or Not IsNumeric(DonationAmountC) Then DonationAmountC = ""
   ' Next


   ' For Each DonationAmountD In Range("Per5080G.DonationAmt")
   '     If Len(DonationAmountD) > 14 Or Not IsNumeric(DonationAmountD) Then DonationAmountD = ""
   ' Next


endline7:
Application.EnableEvents = True
Application.ScreenUpdating = True
End Sub



'Ankita--------07/03/2026----
Function GetBankName1(IFSC As Variant) As String
On Error Resume Next
GetBankName1 = ""
Dim rangecells As Range
Dim Address As Variant
For Each rangecells In Sheet10.Range("IFSC_StartLetter")

If UCase(rangecells.Value) = UCase(IFSC) Then
Address = rangecells.Address
Address = Replace(Address, "A", "B")
GetBankName1 = Sheet10.Range(Address).Value
Exit Function
End If

Next

'If GetBankName = "" Then
'GetBankName = False
'End If

End Function





Sub TESTT()
Application.EnableEvents = True
End Sub
