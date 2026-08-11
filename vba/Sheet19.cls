Attribute VB_Name = "Sheet19"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Dim MsgBx_TenantPan As String

Dim ActiveCellRange As String

Private Sub Worksheet_Activate()
On Error Resume Next
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet19.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
Sheet19.Shapes("Button 14").Left = Range("Button_HP").Left
Sheet19.Shapes("Button 14").Top = Range("Button_HP").Top

Sheet19.Shapes("Button 17").Left = Range("Button_HP2").Left
Sheet19.Shapes("Button 17").Top = Range("Button_HP2").Top

Sheet19.Shapes("Button 75").Left = Range("Button_HP2").Left + 180
Sheet19.Shapes("Button 75").Top = Range("Button_HP2").Top
Sheet19.Protect Password:=sPassword

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

Sheet1.Range("taxcheck").Value = "Y"

Sheet19.Activate

Dim validateType As String
Dim cell As Variant
            Dim PinCode As Range
            Dim DupFlag As Boolean
            Dim StateName
            Dim StateCode
            Dim state1 As Variant
   validateType = ""
   Application.EnableEvents = False

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    Dim HpCount As Long

       For HpCount = 1 To Sheet19.Range("PropertySectionCount").Value

       For Each cell In Sheet19.Range("HP.Co.Pan" & HpCount).Cells
        Sheet19.Range("" & cell.Address).Value = UCase(Sheet19.Range("" & cell.Address).Value)
       Next

       For Each cell In Sheet19.Range("HP.PANofTenant" & HpCount).Cells
        Sheet19.Range("" & cell.Address).Value = UCase(Sheet19.Range("" & cell.Address).Value)
       Next
       
       For Each cell In Sheet19.Range("HP.AadharofTenant" & HpCount).Cells
        Sheet19.Range("" & cell.Address).Value = UCase(Sheet19.Range("" & cell.Address).Value)
       Next


        For Each cell In Sheet19.Range("HP.TANofTenant" & HpCount).Cells
        Sheet19.Range("" & cell.Address).Value = UCase(Sheet19.Range("" & cell.Address).Value)
       Next


    If Not Application.Intersect(Range("HP.Co.Pan" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
        Range(Target.Address).Value = UCase(Range(Target.Address).Value)
        
            If Not ValidateTenantPan(Range(Target.Address).Value) Then
                fmsgboxStatus "Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule HP"
                Range(Target.Address).Value = ""
                Range(Target.Address).Select
            End If
        End If

       If Not Application.Intersect(Range("HP.PANofTenant" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
        Range(Target.Address).Value = UCase(Range(Target.Address).Value)
            If Not ValidateTenantPan(Range(Target.Address).Value) Then
                fmsgboxStatus """Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule HP"""
                Range(Target.Address).Value = ""
                Range(Target.Address).Select
            End If
        End If

'Error message changed as per V0.4
'Ankita_24/02/2026================
        If Not Application.Intersect(Range("HP.TANofTenant" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
         If Not (CheckTAN(Range(Target.Address).Value) Or ValidateTenantPan(Range(Target.Address).Value)) Then
'               fmsgboxStatus """Invalid PAN/TAN. PAN/TAN format should be First Four Alphabets, next 5 digits, then 1 Alphabet or First Five Alphabets, next 4 digits, then 1 Alphabet in schedule HP."""
                fmsgboxStatus """Invalid PAN/TAN. PAN/TAN format should be First Four Alphabets, next 5 digits, then 1 Alphabet in case of  TAN OR First Five Alphabets, next 4 digits, then 1 Alphabet in schedule HP in case of PAN"""
                Range(Target.Address).Value = ""
                Range(Target.Address).Select
            End If
        End If
    
       If Not Application.Intersect(Range("HP.AadharofTenant" & HpCount).Cells, Range(Target.Address)) Is Nothing Then
         If Trim(Range(Target.Address).Value) <> "" Then
    
               If Not (IsNumeric(Range(Target.Address).Value)) Then
    
                   fmsgboxStatus """Invalid Aadhaar. Please enter the valid Aadhaar Number."""
                   Range(Target.Address).Value = ""
                   Range(Target.Address).Select
               End If
           End If
       End If
       
 '===================
'    If Target.Validation.Type = 2 Then
'      If Not Intersect(Sheet19.Range("HP.Co.Share" & HpCount), Range(Target.Address)) Is Nothing Then
'       LockUnlockCoOwnersDetails (HpCount)
'      End If
'    End If

      If Not Intersect(Sheet19.Range("HP.CoOwnedYN" & HpCount), Range(Target.Address)) Is Nothing Or _
         Not Intersect(Sheet19.Range("HP.Co.Share" & HpCount), Range(Target.Address)) Is Nothing Then
       LockUnlockCoOwnersDetails (HpCount)
      End If
      
'        If (Target.name.name) = "HP.CoOwnedYN" & HpCount Then
'           LockUnlockCoOwnersDetails (HpCount)
'        End If

'===========================================
''    If Target.Validation.Type = 3 Then
''         If (Target.name.name) = "HP.StateCode" & HpCount Or (Target.name.name) = "HP.CountryCode" & HpCount Then
'            Sheet19.Unprotect getmsgstate
'            If isdropdownblank(Sheet19.Range("HP.StateCode" & HpCount).Value) Then
'                Sheet19.Range("HP.CountryCode" & HpCount).Value = "(Select)"
'
'                Sheet19.Range("HP.ZipCode" & HpCount).ClearContents
'                Sheet19.Range("HP.ZipCode" & HpCount).Locked = False
'                Sheet19.Range("HP.ZipCode" & HpCount).Interior.Color = "&HCCFFCC"
'
'                Sheet19.Range("HP.PinCode" & HpCount).ClearContents
'                Sheet19.Range("HP.PinCode" & HpCount).Locked = False
'                Sheet19.Range("HP.PinCode" & HpCount).Interior.Color = "&HCCFFCC"
'            ElseIf Mid(Sheet19.Range("HP.StateCode" & HpCount).Value, 1, 2) = "99" Then
'
'              If Mid(Sheet19.Range("HP.CountryCode" & HpCount).Value, 1, 2) = "91" Then
'                Sheet19.Range("HP.CountryCode" & HpCount).Value = "(Select)"
'                End If
'
'                Sheet19.Range("HP.ZipCode" & HpCount).ClearContents
'                Sheet19.Range("HP.ZipCode" & HpCount).Locked = False
'                Sheet19.Range("HP.ZipCode" & HpCount).Interior.Color = "&HCCFFCC"
'
'                Sheet19.Range("HP.PinCode" & HpCount).ClearContents
'                Sheet19.Range("HP.PinCode" & HpCount).Locked = True
'                Sheet19.Range("HP.PinCode" & HpCount).Interior.Color = "&HD8D8D8"
'             ElseIf Mid(Sheet19.Range("HP.StateCode" & HpCount).Value, 1, 2) <> "99" Then
'
'                If (Target.name.name) = "HP.CountryCode" & HpCount And Mid(Sheet19.Range("HP.CountryCode" & HpCount).Value, 1, 2) <> "91" Then
'                 fmsgboxStatus """Country cannot be other than India as you have selected an Indian state in schedule HP."""
'                End If
'
'                Sheet19.Range("HP.CountryCode" & HpCount).Value = "91-INDIA"
'                Sheet19.Range("HP.ZipCode" & HpCount).ClearContents
'                Sheet19.Range("HP.ZipCode" & HpCount).Locked = True
'                Sheet19.Range("HP.ZipCode" & HpCount).Interior.Color = "&HD8D8D8"
'
'                If Sheet19.Range("HP.PinCode" & HpCount).Value = "" Then  'Malli
'                Sheet19.Range("HP.PinCode" & HpCount).ClearContents
'                Sheet19.Range("HP.PinCode" & HpCount).Locked = False
'                Sheet19.Range("HP.PinCode" & HpCount).Interior.Color = "&HCCFFCC"
'                End If
'                End If
'                End If
'            End If
            
    
    
'=====================================
 If Not Intersect(Range("HP.ifLetOut" & HpCount), Range(Target.Address)) Is Nothing Then
           If Mid(Sheet19.Range("HP.ifLetOut" & HpCount).Value, 1, 1) = "S" Or Mid(Sheet19.Range("HP.ifLetOut" & HpCount).Value, 1, 1) = "D" Then
                
                Sheet19.Range("HP.NameofTenant" & HpCount).Locked = True
                Sheet19.Range("HP.NameofTenant" & HpCount).ClearContents
                Sheet19.Range("HP.NameofTenant" & HpCount).Interior.Color = (&HD8D8D8)

                Sheet19.Range("HP.PANofTenant" & HpCount).Locked = True
                Sheet19.Range("HP.PANofTenant" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet19.Range("HP.PANofTenant" & HpCount).ClearContents
                
                Sheet19.Range("HP.AadharofTenant" & HpCount).Locked = True
                Sheet19.Range("HP.AadharofTenant" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet19.Range("HP.AadharofTenant" & HpCount).ClearContents
                
                Sheet19.Range("HP.TANofTenant" & HpCount).Locked = True
                Sheet19.Range("HP.TANofTenant" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet19.Range("HP.TANofTenant" & HpCount).ClearContents
                If Mid(Sheet19.Range("HP.ifLetOut" & HpCount).Value, 1, 1) = "S" Then
                    Sheet19.Range("HP.AnnualLetableValue" & HpCount).Locked = True
                    Sheet19.Range("HP.AnnualLetableValue" & HpCount).Interior.Color = (&HD8D8D8)
                    Sheet19.Range("HP.AnnualLetableValue" & HpCount).ClearContents
    
                    Sheet19.Range("HP.RentNotRealized" & HpCount).Locked = True
                    Sheet19.Range("HP.RentNotRealized" & HpCount).Interior.Color = (&HD8D8D8)
                    Sheet19.Range("HP.RentNotRealized" & HpCount).ClearContents
    
                    Sheet19.Range("HP.LocalTaxes" & HpCount).Locked = True
                    Sheet19.Range("HP.LocalTaxes" & HpCount).Interior.Color = (&HD8D8D8)
                    Sheet19.Range("HP.LocalTaxes" & HpCount).ClearContents
                End If
                '---------------------------------------------------------------------------
                Sheet19.Range("HP.PANofTenant" & HpCount).Locked = True
                Sheet19.Range("HP.PANofTenant" & HpCount).Interior.Color = (&HD8D8D8)
                Sheet19.Range("HP.PANofTenant" & HpCount).ClearContents
                
            Else
                Sheet19.Range("HP.NameofTenant" & HpCount).Locked = False
                Sheet19.Range("HP.NameofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet19.Range("HP.NameofTenant" & HpCount).ClearContents

                Sheet19.Range("HP.PANofTenant" & HpCount).Locked = False
                Sheet19.Range("HP.PANofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet19.Range("HP.PANofTenant" & HpCount).ClearContents
                
                Sheet19.Range("HP.AadharofTenant" & HpCount).Locked = False
                Sheet19.Range("HP.AadharofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet19.Range("HP.AadharofTenant" & HpCount).ClearContents
                
                Sheet19.Range("HP.TANofTenant" & HpCount).Locked = False
                Sheet19.Range("HP.TANofTenant" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet19.Range("HP.TANofTenant" & HpCount).ClearContents
                
                Sheet19.Range("HP.AnnualLetableValue" & HpCount).Locked = False
                Sheet19.Range("HP.AnnualLetableValue" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet19.Range("HP.AnnualLetableValue" & HpCount).ClearContents

                Sheet19.Range("HP.RentNotRealized" & HpCount).Locked = False
                Sheet19.Range("HP.RentNotRealized" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet19.Range("HP.RentNotRealized" & HpCount).ClearContents

                Sheet19.Range("HP.LocalTaxes" & HpCount).Locked = False
                Sheet19.Range("HP.LocalTaxes" & HpCount).Interior.Color = (&HCCFFCC)
                Sheet19.Range("HP.LocalTaxes" & HpCount).ClearContents
            End If
        End If
        
        '===========
        

        '========
        If Not Intersect(Sheet19.Range("HP.OwnerProperty" & HpCount), Range(Target.Address)) Is Nothing Then
           If Mid(Sheet19.Range("HP.OwnerProperty" & HpCount).Value, 1, 1) = "O" Then
                Sheet19.Range("HP.OwnerPropertyDescription" & HpCount).Locked = False
                Sheet19.Range("HP.OwnerPropertyDescription" & HpCount).Interior.Color = (&HCCFFCC)
           Else
                Sheet19.Range("HP.OwnerPropertyDescription" & HpCount).Locked = True
                Sheet19.Range("HP.OwnerPropertyDescription" & HpCount).ClearContents
                Sheet19.Range("HP.OwnerPropertyDescription" & HpCount).Interior.Color = (&HD8D8D8)
         End If
        End If
        '========
        
    If Not Application.Intersect(Sheet19.Range("HP.PinCode" & HpCount), Sheet19.Range(Target.Address)) Is Nothing Then

    Dim pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk As String
     Dim ws As Worksheet
      pinrng_chk = Replace(Target.Address, "$", "")
      staterng_chk = Replace(pinrng_chk, "J", "H")
      cntryrng_chk = Replace(pinrng_chk, "J", "I")
      ZIPrng_chk = Replace(pinrng_chk, "J", "K")

      Set ws = Worksheets("HP")

      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)
      Call PinState_codes_validation(ws, pinrng_chk, staterng_chk, cntryrng_chk, ZIPrng_chk)

    End If


'===========
If Not Intersect(Range("HP.StateCode" & HpCount), Range(Target.Address)) Is Nothing Or _
   Not Intersect(Range("HP.CountryCode" & HpCount), Range(Target.Address)) Is Nothing Or _
   Not Intersect(Range("HP.PinCode" & HpCount), Range(Target.Address)) Is Nothing Or _
   Not Intersect(Range("HP.ZipCode" & HpCount), Range(Target.Address)) Is Nothing Then


'Ankita_03/02/2026====

'        If (Target.name.name) = "HP.StateCode" & HpCount Or _
'        (Target.name.name) = "HP.CountryCode" & HpCount Or _
'        (Target.name.name) = "HP.PinCode" & HpCount Or _
'        (Target.name.name) = "HP.ZipCode" & HpCount Then

            Sheet19.Unprotect getmsgstate
            If isdropdownblank(Sheet19.Range("HP.StateCode" & HpCount).Value) Then
'                Sheet19.Range("HP.CountryCode" & HpCount).Value = "(Select)"
                Sheet19.Range("HP.CountryCode" & HpCount).Value = "91-INDIA"
                Sheet19.Range("HP.ZipCode" & HpCount).ClearContents
                Sheet19.Range("HP.ZipCode" & HpCount).Locked = False
                Sheet19.Range("HP.ZipCode" & HpCount).Interior.Color = "&HCCFFCC"
                
                Sheet19.Range("HP.PinCode" & HpCount).ClearContents
                Sheet19.Range("HP.PinCode" & HpCount).Locked = False
                Sheet19.Range("HP.PinCode" & HpCount).Interior.Color = "&HCCFFCC"
            ElseIf Mid(Sheet19.Range("HP.StateCode" & HpCount).Value, 1, 2) = "99" Then
                If Sheet19.Range("HP.CountryCode" & HpCount).Value = "" Or Mid(Sheet19.Range("HP.CountryCode" & HpCount).Value, 1, 2) = "91" Then
                Sheet19.Range("HP.CountryCode" & HpCount).Value = "(Select)"
                End If
                
                If Sheet19.Range("HP.ZipCode" & HpCount).Value = "" Then
                Sheet19.Range("HP.ZipCode" & HpCount).ClearContents
                Sheet19.Range("HP.ZipCode" & HpCount).Locked = False
                Sheet19.Range("HP.ZipCode" & HpCount).Interior.Color = "&HCCFFCC"
                End If
                Sheet19.Range("HP.PinCode" & HpCount).ClearContents
                Sheet19.Range("HP.PinCode" & HpCount).Locked = True
                Sheet19.Range("HP.PinCode" & HpCount).Interior.Color = "&HD8D8D8"
             ElseIf Mid(Sheet19.Range("HP.StateCode" & HpCount).Value, 1, 2) <> "99" Then
                Sheet19.Range("HP.CountryCode" & HpCount).Value = "91-INDIA"
                Sheet19.Range("HP.ZipCode" & HpCount).ClearContents
                Sheet19.Range("HP.ZipCode" & HpCount).Locked = True
                Sheet19.Range("HP.ZipCode" & HpCount).Interior.Color = "&HD8D8D8"
                
                If Sheet19.Range("HP.PinCode" & HpCount).Value = "" Then
                Sheet19.Range("HP.PinCode" & HpCount).ClearContents
                Sheet19.Range("HP.PinCode" & HpCount).Locked = False
                Sheet19.Range("HP.PinCode" & HpCount).Interior.Color = "&HCCFFCC"

                End If
 End If

            Sheet19.Protect getmsgstate

            'GoTo endline

          End If

'    End If     'Ankita_03/02/2026
 



'===========
    Next
        
        
    
Dim r As Long
For r = 1 To Sheet19.Range("PropertySectionCount").Value
            If Range("HP.ifLetOut" & r).Value = "Let Out" Then
                If Range("HP.TotalIncomeChargeableUnHP").Value < -200000 Then
                    MsgBox "* Note : Maximum Loss from House property that can be set-off is INR 2,00,000. in schedule HP" ', vbOKOnly, mdCommon.vbMessageCaption
                End If
            End If
 
   If Mid(Range("HP.ifLetOut" & r).Value, 1, 1) = "S" Then
        Dim hpInterest1, salIncome1 As Variant
           hpInterest1 = Range("HP.IntOnBorwCap" & r).Value
            If Range("HP.ifLetOut" & r).Value = "Self Occupied" Then
             
                     If Not SchHP.CheckHousePropertyIncome(hpInterest1) Then
                     hpInterest1 = Range("HP.IntOnBorwCap" & r).Value
                     End If
                     
                     Dim CheckSelfOccupied_count As Variant
                     Dim CheckSelfOccupied_IntOnBorwCap_count As Variant
                     CheckSelfOccupied_count = CheckSelfOccupied_count + 1
                     CheckSelfOccupied_IntOnBorwCap_count = CheckSelfOccupied_IntOnBorwCap_count + hpInterest1
                     
    
                     If CheckSelfOccupied_count > 2 Then
                        fmsgboxsmall "* ""Self occupied property cannot be selected more than twice""" & Chr(13)
                        Range("HP.ifLetOut" & r).Value = "(Select)"
                        CheckSelfOccupied_count = 0
                     End If
    
                     If CheckSelfOccupied_IntOnBorwCap_count > 200000 Then
                        fmsgboxsmall "*""Amount of deduction for interest u/s 24(b) in case of Self Occupied House Property or aggregate of such properties  cannot exceed 2 lakh. Therefore, interest is being restrcited to 2 lakh""" & Chr(13)
                     Range("HP.IntOnBorwCap" & r).Value = ""
                        CheckSelfOccupied_IntOnBorwCap_count = 0
                     End If
                End If
    End If
    

        If Sheet19.Range("HP.AnnualLetableValue" & r).Value = "" Or Sheet19.Range("HP.AnnualLetableValue" & r).Value = 0 Then
        If Sheet19.Range("HP.LocalTaxes" & r).Value > 0 Then
        fmsgboxsmall "Tax paid to local authorities can be claimed only if income from house property is declared in schedule HP" & Chr(13)
        Target.Value = ""
        End If
        End If
Next


'====================================

    Dim c As Long
    For c = 1 To Sheet19.Range("PropertySectionCOunt").Value
    If Not Application.Intersect(Range("HP.ifLetOut" & c).Cells, Range(Target.Address)) Is Nothing Then
            If Sheet5.Range("bacValue").Value = 1 And Mid(Sheet19.Range("HP.ifLetOut" & c).Value, 1, 1) = "S" Then
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).Locked = True
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).ClearContents
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("bankName.24b" & c).Locked = True
                Sheet19.Range("bankName.24b" & c).ClearContents
                Sheet19.Range("bankName.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("loanAccNum.24b" & c).Locked = True
                Sheet19.Range("loanAccNum.24b" & c).ClearContents
                Sheet19.Range("loanAccNum.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("loanDate.24b" & c).Locked = True
                Sheet19.Range("loanDate.24b" & c).ClearContents
                Sheet19.Range("loanDate.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("loanAmt.24b" & c).Locked = True
                Sheet19.Range("loanAmt.24b" & c).ClearContents
                Sheet19.Range("loanAmt.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("loanOutstanding.24b" & c).Locked = True
                Sheet19.Range("loanOutstanding.24b" & c).ClearContents
                Sheet19.Range("loanOutstanding.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("Intrst.24b" & c).Locked = True
                Sheet19.Range("Intrst.24b" & c).ClearContents
                Sheet19.Range("Intrst.24b" & c).Interior.Color = (&HD8D8D8)
                
            Else
                
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).Locked = False
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).ClearContents
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("bankName.24b" & c).Locked = False
                Sheet19.Range("bankName.24b" & c).ClearContents
                Sheet19.Range("bankName.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("loanAccNum.24b" & c).Locked = False
                Sheet19.Range("loanAccNum.24b" & c).ClearContents
                Sheet19.Range("loanAccNum.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("loanDate.24b" & c).Locked = False
                Sheet19.Range("loanDate.24b" & c).ClearContents
                Sheet19.Range("loanDate.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("loanAmt.24b" & c).Locked = False
                Sheet19.Range("loanAmt.24b" & c).ClearContents
                Sheet19.Range("loanAmt.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("loanOutstanding.24b" & c).Locked = False
                Sheet19.Range("loanOutstanding.24b" & c).ClearContents
                Sheet19.Range("loanOutstanding.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("Intrst.24b" & c).Locked = False
                Sheet19.Range("Intrst.24b" & c).ClearContents
                Sheet19.Range("Intrst.24b" & c).Interior.Color = (&HCCFFCC)
                
            End If
        End If
    Next
'--end

Dim s As Long
Dim SR As Long
SR = 0
Sheet19.Unprotect Password:=sPassword
For s = 1 To Sheet19.Range("PropertySectionCOunt").Value

    If Sheet5.Range("bacValue").Value = 1 And Mid(Sheet19.Range("HP.ifLetOut" & s).Value, 1, 1) = "S" Then
           Sheet19.Range("HP.IntOnBorwCap" & s).Value = 0
    ElseIf Sheet5.Range("bacValue").Value = 2 And Mid(Sheet19.Range("HP.ifLetOut" & s).Value, 1, 1) = "S" Then
            Sheet19.Range("HP.IntOnBorwCap" & s).Value = Application.WorksheetFunction.Min(200000 - SR, Sheet19.Range("TotAmt.24b" & s).Value)
            SR = SR + Sheet19.Range("HP.IntOnBorwCap" & s).Value
    ElseIf (Sheet19.Range("HP.AnnualLetableValue" & s).Value = 0 Or Sheet19.Range("HP.AnnualLetableValue" & s).Value = "") And Mid(Sheet19.Range("HP.ifLetOut" & s).Value, 1, 1) <> "S" Then
            Sheet19.Range("HP.IntOnBorwCap" & s).Value = 0
    Else
            Sheet19.Range("HP.IntOnBorwCap" & s).Value = Sheet19.Range("TotAmt.24b" & s).Value
    End If

Next

'====================================
'    Dim c As Long
    For c = 1 To Sheet19.Range("PropertySectionCOunt").Value
    If Not Application.Intersect(Range("HP.ifLetOut" & c).Cells, Range(Target.Address)) Is Nothing Then
'        If Target.name.name = ("HP.ifLetOut" & c) Then
            If Sheet5.Range("bacValue").Value = 1 And Mid(Sheet19.Range("HP.ifLetOut" & c).Value, 1, 1) = "S" Then
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).Locked = True
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).ClearContents
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("bankName.24b" & c).Locked = True
                Sheet19.Range("bankName.24b" & c).ClearContents
                Sheet19.Range("bankName.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("loanAccNum.24b" & c).Locked = True
                Sheet19.Range("loanAccNum.24b" & c).ClearContents
                Sheet19.Range("loanAccNum.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("loanDate.24b" & c).Locked = True
                Sheet19.Range("loanDate.24b" & c).ClearContents
                Sheet19.Range("loanDate.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("loanAmt.24b" & c).Locked = True
                Sheet19.Range("loanAmt.24b" & c).ClearContents
                Sheet19.Range("loanAmt.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("loanOutstanding.24b" & c).Locked = True
                Sheet19.Range("loanOutstanding.24b" & c).ClearContents
                Sheet19.Range("loanOutstanding.24b" & c).Interior.Color = (&HD8D8D8)
                
                Sheet19.Range("Intrst.24b" & c).Locked = True
                Sheet19.Range("Intrst.24b" & c).ClearContents
                Sheet19.Range("Intrst.24b" & c).Interior.Color = (&HD8D8D8)
                
            Else
                
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).Locked = False
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).ClearContents
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("bankName.24b" & c).Locked = False
                Sheet19.Range("bankName.24b" & c).ClearContents
                Sheet19.Range("bankName.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("loanAccNum.24b" & c).Locked = False
                Sheet19.Range("loanAccNum.24b" & c).ClearContents
                Sheet19.Range("loanAccNum.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("loanDate.24b" & c).Locked = False
                Sheet19.Range("loanDate.24b" & c).ClearContents
                Sheet19.Range("loanDate.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("loanAmt.24b" & c).Locked = False
                Sheet19.Range("loanAmt.24b" & c).ClearContents
                Sheet19.Range("loanAmt.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("loanOutstanding.24b" & c).Locked = False
                Sheet19.Range("loanOutstanding.24b" & c).ClearContents
                Sheet19.Range("loanOutstanding.24b" & c).Interior.Color = (&HCCFFCC)
                
                Sheet19.Range("Intrst.24b" & c).Locked = False
                Sheet19.Range("Intrst.24b" & c).ClearContents
                Sheet19.Range("Intrst.24b" & c).Interior.Color = (&HCCFFCC)
                
            End If
        End If
    Next

'====================================
Sheet19.Protect Password:=sPassword
endline:
'----------------Lock Password-------------------START---
   ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
End Sub

Sub AddTotIncomeUnderHouseProperty()
    Dim TotIncomeRangeName As String
    Dim i As Variant
    Dim newDefinedName As String
    Dim tot, Totsections As Long

    tot = 0
    TotIncomeRangeName = "HP.IncomeOfHP"
    Totsections = Sheet19.Range("PropertySectionCOunt").Value
        For i = 1 To Totsections
        newDefinedName = TotIncomeRangeName & i
        tot = tot + Sheet19.Range(newDefinedName)
        Next
    Sheet19.Range("HP.TotalIncomeChargeableUnHP").Value = tot + Sheet19.Range("HP.PassTroughIncome").Value '+ Sheet19.Range("HP.RentArearsSec25BAfter30pcDeduct").value
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True
End Sub

'Ankita_20/01/2026
Sub LockUnlockCoOwnersDetails(HpCount As Long)
   On Error Resume Next
   If Sheet19.Range("HP.CoOwnedYN" & HpCount).Value = "No" Then
                 Sheet19.Range("Co_OwnnerBlock" & HpCount).ClearContents
                 Sheet19.Range("Co_OwnnerBlock" & HpCount).Locked = True
                 Sheet19.Range("Co_OwnnerBlock" & HpCount).Interior.Color = (&HD8D8D8)
                 Sheet19.Range("HP.SharePercent" & HpCount).Value = "100"
                 Sheet19.Range("HP.SharePercent" & HpCount).Locked = True
                 Sheet19.Range("HP.SharePercent" & HpCount).Interior.ColorIndex = 2
  ElseIf Sheet19.Range("HP.CoOwnedYN" & HpCount).Value = "Yes" Then
                 Sheet19.Range("Co_OwnnerBlock" & HpCount).Locked = False
                 Sheet19.Range("Co_OwnnerBlock" & HpCount).Interior.Color = (&HCCFFCC)
                 Sheet19.Range("HP.SharePercent" & HpCount).Value = WorksheetFunction.Max(0, 100 - WorksheetFunction.Sum(Sheet19.Range("HP.Co.Share" & HpCount).Value))
                 Sheet19.Range("HP.SharePercent" & HpCount).Locked = True
                 Sheet19.Range("HP.SharePercent" & HpCount).Interior.ColorIndex = 2
   ElseIf Sheet19.Range("HP.CoOwnedYN" & HpCount).Value <> "No" Then
                 Sheet19.Range("Co_OwnnerBlock" & HpCount).ClearContents
                 Sheet19.Range("Co_OwnnerBlock" & HpCount).Locked = False
                 Sheet19.Range("Co_OwnnerBlock" & HpCount).Interior.Color = (&HCCFFCC)
                 Sheet19.Range("HP.SharePercent" & HpCount).Value = ""
                 Sheet19.Range("HP.SharePercent" & HpCount).Locked = False
                 Sheet19.Range("HP.SharePercent" & HpCount).Interior.Color = (&HCCFFCC)
    End If
End Sub


'Ankita_29/01/2026===================
Sub LockUnlock_24bHP()
On Error Resume Next
'Application.EnableEvents = False
    Sheet19.Unprotect Password:=getmsgstate
    Dim HpCount1 As Long
 
       For HpCount1 = 1 To Sheet19.Range("PropertySectionCount").Value
    'If Not Application.Intersect(Range("HP.ifLetOut" & HpCount1).Cells, Range(Target.Address)) Is Nothing Then
            If Sheet5.Range("bacValue").Value = 1 And Mid(Sheet19.Range("HP.ifLetOut" & HpCount1).Value, 1, 1) = "S" Then
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & HpCount1).Locked = True
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & HpCount1).Interior.Color = (&HD8D8D8)

                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("bankName.24b" & HpCount1).Locked = True
                Sheet19.Range("bankName.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("bankName.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanAccNum.24b" & HpCount1).Locked = True
                Sheet19.Range("loanAccNum.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanAccNum.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanDate.24b" & HpCount1).Locked = True
                Sheet19.Range("loanDate.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanDate.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanAmt.24b" & HpCount1).Locked = True
                Sheet19.Range("loanAmt.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanAmt.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanOutstanding.24b" & HpCount1).Locked = True
                Sheet19.Range("loanOutstanding.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanOutstanding.24b" & HpCount1).Interior.Color = (&HD8D8D8)
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("Intrst.24b" & HpCount1).Locked = True
                Sheet19.Range("Intrst.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("Intrst.24b" & HpCount1).Interior.Color = (&HD8D8D8)
            Else
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & HpCount1).Locked = False
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("LoanfrmBankOrInstitute.24b" & HpCount1).Interior.Color = (&HCCFFCC)

                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("bankName.24b" & HpCount1).Locked = False
                Sheet19.Range("bankName.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("bankName.24b" & HpCount1).Interior.Color = (&HCCFFCC)
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanAccNum.24b" & HpCount1).Locked = False
                Sheet19.Range("loanAccNum.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanAccNum.24b" & HpCount1).Interior.Color = (&HCCFFCC)

                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanDate.24b" & HpCount1).Locked = False
                Sheet19.Range("loanDate.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanDate.24b" & HpCount1).Interior.Color = (&HCCFFCC)

                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanAmt.24b" & HpCount1).Locked = False
                Sheet19.Range("loanAmt.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanAmt.24b" & HpCount1).Interior.Color = (&HCCFFCC)
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanOutstanding.24b" & HpCount1).Locked = False
                Sheet19.Range("loanOutstanding.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("loanOutstanding.24b" & HpCount1).Interior.Color = (&HCCFFCC)
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("Intrst.24b" & HpCount1).Locked = False
                Sheet19.Range("Intrst.24b" & HpCount1).ClearContents
                Sheet19.Unprotect Password:=getmsgstate
                Sheet19.Range("Intrst.24b" & HpCount1).Interior.Color = (&HCCFFCC)
            End If
        'End If
    Next
 
 
    Sheet19.Protect Password:=getmsgstate
End Sub
 
 
Sub test()
Application.EnableEvents = True
End Sub

