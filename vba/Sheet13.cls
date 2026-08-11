Attribute VB_Name = "Sheet13"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Worksheet_Activate()
On Error Resume Next
Application.EnableEvents = False
ActiveWindow.DisplayHeadings = False

Sheet13.Unprotect Password:=EfilingCommon.getmsgstate
fmsgboxoK "Please fill date wise contribution details. If Contribution is made in cash and in other mode on same day, please fill details in separate rows"
Sheet13.Protect Password:=EfilingCommon.getmsgstate
Application.EnableEvents = True
ActiveWindow.DisplayHeadings = True

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline

Application.EnableEvents = False

Dim IfsCode As Range
Dim sPassword

Dim IffCode As Range
Dim TargtVal, targetadd As Variant

TargtVal = Target.Value
targetadd = Target.Address
targetadd = Replace(targetadd, "$", "")


sPassword = EfilingCommon.getmsgstate
Sheet13.Unprotect Password:=sPassword

'Newly added by Malli

If Not Application.Intersect(Sheet13.Range("DateofDonation_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
        
        Dim DOC1, DOC2 As Variant   'added by Chetan C M on 01/02/2025 for AY 2026-27
        Dim Specified_Datei As Variant
        Specified_Datei = Target.Value
        
         'added by Chetan C M on 01/02/2025 for AY 2026-27
        'start--
        DOC1 = Sheet5.Range("DateOfCont1").Value
        DOC2 = Sheet5.Range("DateOfCont2").Value
        '--end

'        If ChkMinInclusiveDate(Dformat(Specified_Datei, "yyyy-mm-dd"), "2024-04-01") And Specified_Datei <> "" Then
'             fmsgbox "* ""Please enter the correct date, the date cannot be greater than 31/03/2024. """
         
         'changed by Chetan C M on 30/12/2025 for AY 2026-27
         '--start
         ''             Ankita_25-26
        'If ChkMinInclusiveDate(Dformat(Specified_Datei, "yyyy-mm-dd"), "2025-04-01") And Specified_Datei <> "" Then
            ' fmsgbox "* ""Please enter the correct date, the date cannot be greater than 31/03/2025. """
         If ChkMinInclusiveDate(Dformat(Specified_Datei, "yyyy-mm-dd"), Dformat(Sheet5.Range("MaxDate").Value, "")) And Specified_Datei <> "" Then
            fmsgbox ("* ""Please enter the correct date, the date cannot be greater than " & Dformat3(DOC1, "")) & Chr(13)
         '--end
             'Added by sudha on 21st Jan 25
             Target.Value = ""     'Added by sudha on 21st Jan 25
'Commented by Sudha on 21/01/2025
'             Sheet13.Range("DateofDonation_80GGC").Value = ""
'             Sheet13.Range("DateofDonation_80GGC").Activate
'        End If
        

'         If ChkMaxInclusiveDate(Dformat(Specified_Datei, "yyyy-mm-dd"), "2023-03-31") And Specified_Datei <> "" Then
'             fmsgbox "* ""Please enter the correct date, the date cannot be less than 01/04/2023. """
'             Ankita_25-26
Else '       'Added by sudha on 21st Jan 25
        
        'changed by Chetan C M on 30/12/2025 for AY 2026-27
        '--start
         'If ChkMaxInclusiveDate(Dformat(Specified_Datei, "yyyy-mm-dd"), "2024-03-31") And Specified_Datei <> "" Then
               'fmsgbox "* ""Please enter the correct date, the date cannot be less than 01/04/2024. """
           If ChkMaxInclusiveDate(Dformat(Specified_Datei, "yyyy-mm-dd"), Dformat(Sheet5.Range("MinDate").Value, "")) And Specified_Datei <> "" Then
                fmsgbox ("* ""Please enter the correct date, the date cannot be less than " & Dformat3(DOC2, "")) & Chr(13)
        '--end
        
           'Added by sudha on 21st Jan 25
                Target.Value = ""
                'Commented by Sudha on 21/01/2025
'               Sheet13.Range("DateofDonation_80GGC").Value = ""
'             Sheet13.Range("DateofDonation_80GGC").Activate
        End If
    End If
End If  'Added by sudha on 21st Jan 25


'--------------------------
'Newly added by Bindu

If Not Application.Intersect(Sheet13.Range("Chequeno_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
   If Not checkfieldspecialcharacter_Trans(Target.Value) Then
       fmsgbox "* ""Transaction Reference number/Cheque number/IMPS/NEFT/RTGS"" of contribution transaction in Schedule 80GGC is invalid, Only "" / "" and "" - "" special characters are allowed."
       Target.ClearContents
   End If
End If



If Not Application.Intersect(Sheet13.Range("DateofDonation_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then

    If Target.Value <> "00/00/0000" And Target.Value <> "" Then
        'If Not FormatNCheckDate(Trim(Target.Value)) Then
        'CheckDateddmmyyyy
        If Not CheckDateddmmyyyy(Trim(Target.Value)) Then
            'fmsgbox "* Date Of filing Orignal Return Is Invalid in Sheet Income Details. Please enter a valid Date in dd/mm/yyyy format."        ', vbOKOnly, mdCommon.vbMessageCaption
            fmsgbox "* Date is invalid in Sheet 80GGC. Please enter a valid Date in dd/mm/yyyy format." ', vbOKOnly, mdCommon.vbMessageCaption
            Target.ClearContents
        End If
    End If
     
End If

If Not Intersect(Range("PAN_80GGC"), Range(Target.Address)) Is Nothing Then
        Sheet13.Unprotect Password:=getmsgstate
    If Sheet13.Range(Target.Address).Value <> "" Then
        
     Call CheckPolitical_PAN_80GGC(Sheet13.Range(Target.Address).Value, Target.Address)
    End If

End If
            
Set IfsCode = Range("IFSC_80GGC")
If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
        Call CheckIFSC_80GGC(TargtVal, targetadd)
        'GoTo Endline
End If
       
Application.EnableEvents = False
sPassword = EfilingCommon.getmsgstate
Sheet13.Unprotect Password:=sPassword
If Not Application.Intersect(Sheet13.Range("Donationincash_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then

    If Target.Value > 0 And Not IsEmpty(Target.Value) Then
        'DONATION IN OTHER MODE
        Sheet13.Range(Replace(Target.Address, "F", "G")).ClearContents
        Sheet13.Range(Replace(Target.Address, "F", "G")).Locked = True
        Sheet13.Range(Replace(Target.Address, "F", "G")).Interior.Color = "&HD8D8D8"
        
        
'        'NATURE OF TRANSACTION
'        Sheet13.Range(Replace(Target.Address, "F", "J")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "F", "J")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "F", "J")).Interior.Color = "&HD8D8D8"
        
        
        'REF/CHEQUE NUMBER
        Sheet13.Range(Replace(Target.Address, "F", "L")).ClearContents
        Sheet13.Range(Replace(Target.Address, "F", "L")).Locked = True
        Sheet13.Range(Replace(Target.Address, "F", "L")).Interior.Color = "&HD8D8D8"
        
        'IFSC
        Sheet13.Range(Replace(Target.Address, "F", "M")).ClearContents
        Sheet13.Range(Replace(Target.Address, "F", "M")).Locked = True
        Sheet13.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HD8D8D8"
        
        
'        'BANK NAME OF DONoR
'        Sheet13.Range(Replace(Target.Address, "F", "M")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "F", "M")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HD8D8D8"
'
'
'        'BANK ACCOUNT OF Donor
'        Sheet13.Range(Replace(Target.Address, "F", "N")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "F", "N")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "F", "N")).Interior.Color = "&HD8D8D8"
        
    Else
    'DONATION IN OTHER MODE
        Sheet13.Range(Replace(Target.Address, "F", "G")).ClearContents
        Sheet13.Range(Replace(Target.Address, "F", "G")).Locked = False
        Sheet13.Range(Replace(Target.Address, "F", "G")).Interior.Color = "&HCCFFCC"
        
        
'        'NATURE OF TRANSACTION
'        Sheet13.Range(Replace(Target.Address, "F", "J")).Value = "(Select)"
'        Sheet13.Range(Replace(Target.Address, "F", "J")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "F", "J")).Interior.Color = "&HCCFFCC"
        
        
        'REF/CHEQUE NUMBER
        Sheet13.Range(Replace(Target.Address, "F", "L")).ClearContents
        Sheet13.Range(Replace(Target.Address, "F", "L")).Locked = False
        Sheet13.Range(Replace(Target.Address, "F", "L")).Interior.Color = "&HCCFFCC"
        
        'IFSC
        Sheet13.Range(Replace(Target.Address, "F", "M")).ClearContents
        Sheet13.Range(Replace(Target.Address, "F", "M")).Locked = False
        Sheet13.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HCCFFCC"
        
        
'        'BANK NAME OF DONoR
'        Sheet13.Range(Replace(Target.Address, "F", "M")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "F", "M")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "F", "M")).Interior.Color = "&HCCFFCC"
'
'
'        'BANK ACCOUNT OF Donor
'        Sheet13.Range(Replace(Target.Address, "F", "N")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "F", "N")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "F", "N")).Interior.Color = "&HCCFFCC"
        
        
    End If
End If



''DONATION IN OTHER MODE
'If Not Application.Intersect(Sheet13.Range("Donationinothermode_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
'     Sheet13.Unprotect Password:=sPassword
'    'If (Target.Value > 0 And Not IsEmpty(Target.Value)) Or ((IsEmpty(Range(Replace(Target.Address, "G", "J")).Value) And UCase(Range(Replace(Target.Address, "G", "J")).Value <> UCase("(Select)")))) Or Len(Range(Replace(Target.Address, "G", "K")).Value) > 0 Or Len(Range(Replace(Target.Address, "G", "L")).Value) > 0 Then
'    If (Target.Value > 0 And Not IsEmpty(Target.Value)) Or Not Len(Range(Replace(Target.Address, "G", "J")).Value) > 0 Or Not Len(Range(Replace(Target.Address, "G", "K")).Value) > 0 Then
'
'        'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "G", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "G", "F")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "G", "F")).Interior.Color = "&HD8D8D8"
'    Else
'
'       'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "G", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "G", "F")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "G", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If

'Malli
'DONATION IN OTHER MODE
If Not Application.Intersect(Sheet13.Range("Donationinothermode_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
     'If (Target.Value > 0 And Not IsEmpty(Target.Value)) Or ((Not IsEmpty(Range(Replace(Target.Address, "G", "J")).Value) And UCase(Range(Replace(Target.Address, "G", "J")).Value <> UCase("(Select)")))) Or Len(Range(Replace(Target.Address, "G", "K")).Value) > 0 Or Len(Range(Replace(Target.Address, "G", "L")).Value) > 0 Then

'if (target.Value>0 and not isempty(target.Value)) or ((not isempty(range(replace(target.Address,g,j)).Value)) or ((not isempty(range(replace(target.Address,G,K)).Value))) then
If (Target.Value > 0 And Not IsEmpty(Target.Value)) Or (Not IsEmpty(Range(Replace(Target.Address, "G", "j")).Value) And Len(Range(Replace(Target.Address, "G", "J")).Value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "G", "k")).Value) And Len(Range(Replace(Target.Address, "G", "k")).Value) > 0) Then
'    If (Target.value > 0 And Not IsEmpty(Target.value)) Or Len(Range(Replace(Target.Address, "G", "J")).value) > 0 Or Len(Range(Replace(Target.Address, "G", "K")).value) > 0 Or Len(Range(Replace(Target.Address, "G", "L")).value) > 0 Or Len(Range(Replace(Target.Address, "G", "M")).value) > 0 Or Len(Range(Replace(Target.Address, "G", "N")).value) > 0 Then
    Sheet13.Unprotect Password:=sPassword
    'DONATION IN CASH
        Sheet13.Range(Replace(Target.Address, "G", "F")).ClearContents
        Sheet13.Range(Replace(Target.Address, "G", "F")).Locked = True
        Sheet13.Range(Replace(Target.Address, "G", "F")).Interior.Color = "&HD8D8D8"
'        Sheet11.Range(Replace(Target.Address, "G", "F")).value = 0
    Else
'DONATION IN CASH
        'Sheet11.Range(Replace(Target.Address, "G", "F")).ClearContents
        Sheet13.Range(Replace(Target.Address, "G", "F")).Locked = False
        Sheet13.Range(Replace(Target.Address, "G", "F")).Interior.Color = "&HCCFFCC"
    End If
              
End If




                        
 ' Malli comented
''Nature TRANSACTION
'If Not Application.Intersect(Sheet13.Range("NatureofTransaction_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
'
'    If ((Not IsEmpty(Target.Value) And UCase(Target.Value) <> UCase("(Select)"))) Or (Range(Replace(Target.Address, "J", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "J", "G")).Value)) Or Len(Range(Replace(Target.Address, "J", "K")).Value) > 0 Or Len(Range(Replace(Target.Address, "J", "L")).Value) > 0 Then
'
'
'        'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HD8D8D8"
'    Else
'
'       'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If


''Transaction number
'If Not Application.Intersect(Sheet13.Range("Chequeno_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
'      Sheet13.Unprotect Password:=sPassword
'    If Len(Target.Value) > 1 Or Not (Range(Replace(Target.Address, "J", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "J", "G")).Value)) Or Not Len(Range(Replace(Target.Address, "J", "L")).Value) > 0 Then
'
'
'        'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "j", "F")).Interior.Color = "&HD8D8D8"
'    Else
'
'       'DONATION IN CASH
'        'Sheet13.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If
'Malli------------------------------------------------------
''Transaction number
'If Not Application.Intersect(Sheet13.Range("Chequeno_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
'      Sheet13.Unprotect Password:=sPassword
'    'If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "J", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "J", "G")).Value)) And ((Not IsEmpty(Range(Replace(Target.Address, "K", "J")).Value)) Or UCase(Range(Replace(Target.Address, "K", "J")).Value <> UCase("(Select)"))) Or Len(Range(Replace(Target.Address, "K", "L")).Value) > 0 Then
'    'If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "J", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "J", "G")).Value)) And ((Not IsEmpty(Range(Replace(Target.Address, "K", "J")).Value)) Or UCase(Range(Replace(Target.Address, "K", "J")).Value <> UCase("(Select)"))) Or Len(Range(Replace(Target.Address, "J", "L")).Value) > 0 Then
'    'If (Len(Target.Value) > 0 And Not IsEmpty(Target.Value)) Or (Not IsEmpty(Range(Replace(Target.Address, "J", "G")).Value) And Len(Range(Replace(Target.Address, "J", "G")).Value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "J", k)).Value) And Len(Range(Replace(Target.Address, "J", k)).Value) > 0) Then
'If (Len(Target.Value) > 1) Or (Not IsEmpty(Range(Replace(Target.Address, "J", "G")).Value) And Len(Range(Replace(Target.Address, "J", "G")).Value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "J", k)).Value) And Len(Range(Replace(Target.Address, "J", k)).Value) > 0) Then
'
'
'    Sheet13.Unprotect Password:=sPassword
'        'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "J", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HD8D8D8"
''        Sheet11.Range(Replace(Target.Address, "K", "F")).value = 0
'    Else
'       'DONATION IN CASH
'        'Sheet11.Range(Replace(Target.Address, "K", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "J", "F")).Interior.Color = "&HCCFFCC"
'    End If
'
'End If
'-----------------------------------------------------
If Not Application.Intersect(Sheet13.Range("Chequeno_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then

'MsgBox Target.Value
     sPassword = EfilingCommon.getmsgstate
     Sheet13.Unprotect Password:=sPassword
    'If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "L", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "L", "G")).Value)) Or ((IsEmpty(Range(Replace(Target.Address, "L", "J")).Value)) And UCase(Range(Replace(Target.Address, "L", "J")).Value <> UCase("(Select)"))) Or Len(Range(Replace(Target.Address, "L", "K")).Value) > 0 Then
     'If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "K", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "L", "G")).Value)) Or ((Not IsEmpty(Range(Replace(Target.Address, "K", "J")).Value)) And UCase(Range(Replace(Target.Address, "K", "J")).Value <> UCase("(Select)"))) Or Len(Range(Replace(Target.Address, "L", "K")).Value) > 0 Then
   ' If (Len(Target.Value) > 0 And Not IsEmpty(Target.Value)) Or (Not IsEmpty(Range(Replace(Target.Address, "K", "G")).Value) And Len(Range(Replace(Target.Address, "K", "G")).Value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "K", "J")).Value) And Len(Range(Replace(Target.Address, "K", "J")).Value) > 0) Then

'If (Len(Target.Value) > 1) Or (Not IsEmpty(Range(Replace(Target.Address, "J", "G")).Value) And Len(Range(Replace(Target.Address, "J", "G")).Value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "J", k)).Value) And Len(Range(Replace(Target.Address, "J", k)).Value) > 0) Then
If (Len(Target.Value) > 1) Or (Not IsEmpty(Range(Replace(Target.Address, "L", "G")).Value) And Len(Range(Replace(Target.Address, "L", "G")).Value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "L", "M")).Value) And Len(Range(Replace(Target.Address, "L", "M")).Value) > 0) Then

sPassword = EfilingCommon.getmsgstate
    Sheet13.Unprotect Password:=sPassword
    'Ankita_16/02/2026_Changed J to L
        'DONATION IN CASH
        Sheet13.Range(Replace(Target.Address, "L", "F")).ClearContents
        Sheet13.Range(Replace(Target.Address, "L", "F")).Locked = True
        Sheet13.Range(Replace(Target.Address, "L", "F")).Interior.Color = "&HD8D8D8"
'        Sheet11.Range(Replace(Target.Address, "L", "F")).value = 0
    Else
       'DONATION IN CASH
        'Sheet11.Range(Replace(Target.Address, "L", "F")).ClearContents
        Sheet13.Range(Replace(Target.Address, "L", "F")).Locked = False
        Sheet13.Range(Replace(Target.Address, "L", "F")).Interior.Color = "&HCCFFCC"
    End If
              
End If



'IFSC
'If Not Application.Intersect(Sheet13.Range("IFSC_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
'     Sheet13.Unprotect Password:=sPassword
'    'If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "L", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "L", "G")).Value)) Or ((IsEmpty(Range(Replace(Target.Address, "L", "J")).Value)) And UCase(Range(Replace(Target.Address, "L", "J")).Value <> UCase("(Select)"))) Or Len(Range(Replace(Target.Address, "L", "K")).Value) > 0 Then
'     If Len(Target.Value) > 1 Or Not (Range(Replace(Target.Address, "K", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "L", "G")).Value)) Or Not Len(Range(Replace(Target.Address, "L", "J")).Value) > 0 Then
'
'        'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "L", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "L", "F")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "L", "F")).Interior.Color = "&HD8D8D8"
'    Else
'
'       'DONATION IN CASH
'        'Sheet13.Range(Replace(Target.Address, "L", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "L", "F")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "L", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If

If Not Application.Intersect(Sheet13.Range("IFSC_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
     Sheet13.Unprotect Password:=sPassword
    'If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "L", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "L", "G")).Value)) Or ((IsEmpty(Range(Replace(Target.Address, "L", "J")).Value)) And UCase(Range(Replace(Target.Address, "L", "J")).Value <> UCase("(Select)"))) Or Len(Range(Replace(Target.Address, "L", "K")).Value) > 0 Then
     'If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "K", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "L", "G")).Value)) Or ((Not IsEmpty(Range(Replace(Target.Address, "K", "J")).Value)) And UCase(Range(Replace(Target.Address, "K", "J")).Value <> UCase("(Select)"))) Or Len(Range(Replace(Target.Address, "L", "K")).Value) > 0 Then
   ' If (Len(Target.Value) > 0 And Not IsEmpty(Target.Value)) Or (Not IsEmpty(Range(Replace(Target.Address, "K", "G")).Value) And Len(Range(Replace(Target.Address, "K", "G")).Value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "K", "J")).Value) And Len(Range(Replace(Target.Address, "K", "J")).Value) > 0) Then
If (Len(Target.Value) > 1) Or (Not IsEmpty(Range(Replace(Target.Address, "M", "G")).Value) And Len(Range(Replace(Target.Address, "M", "G")).Value) > 0) Or (Not IsEmpty(Range(Replace(Target.Address, "M", "L")).Value) And Len(Range(Replace(Target.Address, "M", "L")).Value) > 0) Then
    'Ankita_16/02/2026_Changed K to M
    Sheet13.Unprotect Password:=sPassword
        'DONATION IN CASH
        Sheet13.Range(Replace(Target.Address, "M", "F")).ClearContents
        Sheet13.Range(Replace(Target.Address, "M", "F")).Locked = True
        Sheet13.Range(Replace(Target.Address, "M", "F")).Interior.Color = "&HD8D8D8"
'        Sheet11.Range(Replace(Target.Address, "L", "F")).value = 0
    Else
       'DONATION IN CASH
        'Sheet11.Range(Replace(Target.Address, "L", "F")).ClearContents
        Sheet13.Range(Replace(Target.Address, "M", "F")).Locked = False
        Sheet13.Range(Replace(Target.Address, "M", "F")).Interior.Color = "&HCCFFCC"
    End If
              
End If
          
''NAME OF BANK
' If Not Application.Intersect(Sheet13.Range("NameofDonor_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
'
'    If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "M", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "M", "G")).Value)) Or Len(Range(Replace(Target.Address, "M", "J")).Value) > 0 Or Len(Range(Replace(Target.Address, "M", "K")).Value) > 0 Or Len(Range(Replace(Target.Address, "M", "L")).Value) > 0 Or Len(Range(Replace(Target.Address, "M", "N")).Value) > 0 Then
'
'
'        'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "M", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "M", "F")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "M", "F")).Interior.Color = "&HD8D8D8"
'    Else
'
'       'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "M", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "M", "F")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "M", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If
            
''AccountofDonor_80GGC
' If Not Application.Intersect(Sheet13.Range("AccountofDonor_80GGC"), Sheet13.Range(Target.Address)) Is Nothing Then
'
'    If Len(Target.Value) > 1 Or (Range(Replace(Target.Address, "N", "G")).Value > 0 And Not IsEmpty(Range(Replace(Target.Address, "N", "G")).Value)) Or Len(Range(Replace(Target.Address, "N", "J")).Value) > 0 Or Len(Range(Replace(Target.Address, "N", "K")).Value) > 0 Or Len(Range(Replace(Target.Address, "N", "L")).Value) > 0 Or Len(Range(Replace(Target.Address, "N", "M")).Value) > 0 Then
'
'
'        'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "N", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "N", "F")).Locked = True
'        Sheet13.Range(Replace(Target.Address, "N", "F")).Interior.Color = "&HD8D8D8"
'    Else
'
'       'DONATION IN CASH
'        Sheet13.Range(Replace(Target.Address, "N", "F")).ClearContents
'        Sheet13.Range(Replace(Target.Address, "N", "F")).Locked = False
'        Sheet13.Range(Replace(Target.Address, "N", "F")).Interior.Color = "&HCCFFCC"
'
'    End If
'
'End If


endline:

Application.EnableEvents = True


sPassword = EfilingCommon.getmsgstate
Sheet13.Protect Password:=sPassword
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
Sub TESST()
Application.EnableEvents = True
End Sub
Function CheckPolitical_PAN_80GGC(Tval As Variant, Tadd As Variant) '80GGC_C5
On Error Resume Next
Application.EnableEvents = False
 'sPassword = getmsgstate
   Sheet11.Unprotect Password:=getmsgstate


If Tval <> "" Then
    
     If Not CheckPAN_80GGC(UCase(Tval)) Then
        fmsgbox "Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet  in schedule 80GGC."
           Sheet13.Range(Tadd) = ""
           'Sheet11.Range(Tadd).Offset(0, 1) = ""
           GoTo endfd
     End If
     
     Dim PAN_Ucase As Variant
    PAN_Ucase = UCase(Sheet13.Range(Tadd).Value)
    Sheet13.Range(Tadd).Value = PAN_Ucase
    
 End If
 

endfd:
Sheet11.Protect Password:=getmsgstate
Application.EnableEvents = True

End Function
Function CheckPAN_80GGC(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckPAN_80GGC = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
                
        If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
            
        
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPAN_80GGC = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacterdot(PAN) Then
        CheckPAN_80GGC = False
            Exit Function
        End If
               
    End If
End Function


