Attribute VB_Name = "Sheet14"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
Sheet14.Protect Password:=EfilingCommon.getmsgstate

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

On Error GoTo endline
Application.EnableEvents = False


sPassword = EfilingCommon.getmsgstate
Sheet14.Unprotect Password:=sPassword


'Ankita_06/05/2025_Commented as per DESheet_v0.7

'If Not Application.Intersect(Sheet14.Range("DatefilingFm10IA_80U"), Sheet14.Range(Target.Address)) Is Nothing Then
'
'    If Target.Value <> "00/00/0000" And Target.Value <> "" Then
'       If Not CheckDateddmmyyyy(Trim(Target.Value)) Then
'            fmsgbox "* Date of filing of Form 10IA is invalid in Sheet 80U. Please enter a valid Date in dd/mm/yyyy format." ', vbOKOnly, mdCommon.vbMessageCaption
'            Target.ClearContents
'        End If
'    End If
'End If



'If Not Intersect(Range("DatefilingFm10IA_80U"), Range(Target.Address)) Is Nothing Then
'             If Trim(Range("DatefilingFm10IA_80U").Value) <> "" Then
''If Dformat(Trim(Range("DatefilingFm10IA_80U").Value), "yyyy-mm-dd") > Dformat(Trim(Sheet3.Range("Ver.Date").Value), "yyyy-mm-dd") Then
'             If Dformat(Trim(Range("DatefilingFm10IA_80U").Value), "yyyy-mm-dd") > Dformat(Trim(Sheet3.Range("Ver.Date").Value), "yyyy-mm-dd") Then
'                 fmsgbox "* Date of filing of Form 10IA can not be more than system date in Schedule 80U"
'                       Range("DatefilingFm10IA_80U").Select
'                        Range("DatefilingFm10IA_80U").Value = ""
'              End If
'              End If
'End If
 

'-------------80DD



''Ankita_18/04/2025
''Malli---------23/04/2025
'If Not Intersect(Range("Disability_80DD"), Range(Target.Address)) Is Nothing Then
'Sheet14.Unprotect Password:=getmsgstate
'        'If Sheet14.Range(Target.Address).Value = "(viii) autism" Then
'          If Sheet14.Range(Target.Address).Value <> "(Select)" Then
'
''        If Sheet14.Range(Target.Address).Value <> "(viii) autism" And _
''        Sheet14.Range(Target.Address).Value <> "(ix) cerebral palsy" And _
''        Sheet14.Range(Target.Address).Value <> "(x) multiple disability" Then
'        If Sheet14.Range(Target.Address).Value <> "(i)autism, cerebral palsy, or multiple disabilities and" Then
'
'
''            Sheet14.Range(Replace(Target.Address, "D", "I")).ClearContents
''            Sheet14.Range(Replace(Target.Address, "D", "I")).Locked = True
''            Sheet14.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HD8D8D8"
'
'            Sheet14.Range(Replace(Target.Address, "D", "J")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Locked = True
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HD8D8D8"
'
''            Sheet14.Range(Replace(Target.Address, "D", "K")).Locked = False
''            Sheet14.Range(Replace(Target.Address, "D", "K")).Interior.Color = "&HCCFFCC"
'
'
'        Else
'        Sheet14.Unprotect Password:=getmsgstate
'
''            Sheet14.Range(Replace(Target.Address, "D", "I")).ClearContents
''            Sheet14.Range(Replace(Target.Address, "D", "I")).Locked = False
''            Sheet14.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HCCFFCC"
'
'            Sheet14.Range(Replace(Target.Address, "D", "J")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HCCFFCC"
'
''            Sheet14.Range(Replace(Target.Address, "D", "K")).ClearContents
''            Sheet14.Range(Replace(Target.Address, "D", "K")).Locked = True
''            Sheet14.Range(Replace(Target.Address, "D", "K")).Interior.Color = "&HD8D8D8"
'
'
'    End If
'
'    Else
'            Sheet14.Unprotect Password:=getmsgstate
''            Sheet14.Range(Replace(Target.Address, "D", "I")).ClearContents
''            Sheet14.Range(Replace(Target.Address, "D", "I")).Locked = False
''            Sheet14.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HCCFFCC"
'
'            Sheet14.Range(Replace(Target.Address, "D", "J")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HCCFFCC"
'
''            Sheet14.Range(Replace(Target.Address, "D", "K")).ClearContents
''            Sheet14.Range(Replace(Target.Address, "D", "K")).Locked = False
''            Sheet14.Range(Replace(Target.Address, "D", "K")).Interior.Color = "&HCCFFCC"
'    End If
'
'End If


'Ankita_07/05/2025

If Not Intersect(Range("Disability_80DD"), Range(Target.Address)) Is Nothing Then
Sheet14.Unprotect Password:=getmsgstate
        'If Sheet14.Range(Target.Address).Value = "(viii) autism" Then
          If Sheet14.Range(Target.Address).Value <> "(Select)" Then
'        If Sheet14.Range(Target.Address).Value <> "(viii) autism" And _
'        Sheet14.Range(Target.Address).Value <> "(ix) cerebral palsy" And _
'        Sheet14.Range(Target.Address).Value <> "(x) multiple disability" Then
'        If Sheet14.Range(Target.Address).Value <> "(i) autism, cerebral palsy, or multiple disabilities and" Then

        If Sheet14.Range(Target.Address).Value <> "(i) autism, cerebral palsy, or multiple disabilities" Then  'Ankita_14/05/2025
             'Ankita_28/05/2025
         If Sheet14.Range(Target.Address).Value = "(ii) others" And Sheet14.Range(Replace(Target.Address, "D", "I")).Value <> "" Then
            MsgBox """Type of disability shall be selected as ""(i) autism, cerebral palsy, or multiple disabilities"" when acknowledgement no. of form 10IA is provided"""
            Sheet14.Range(Target.Address).Value = "(Select)"
         Else:
             Sheet14.Range(Replace(Target.Address, "D", "I")).ClearContents
             Sheet14.Range(Replace(Target.Address, "D", "I")).Locked = True
             Sheet14.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HD8D8D8"
             End If
'            Sheet14.Range(Replace(Target.Address, "D", "J")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Locked = True
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HD8D8D8"
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Interior.Color = "&HCCFFCC"

        Else
             Sheet14.Unprotect Password:=getmsgstate
'            Sheet14.Range(Replace(Target.Address, "D", "I")).ClearContents      'Ankita_28/05/2025
             Sheet14.Range(Replace(Target.Address, "D", "I")).Locked = False
             Sheet14.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HCCFFCC"
'            Sheet14.Range(Replace(Target.Address, "D", "J")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HCCFFCC"
'            Sheet14.Range(Replace(Target.Address, "D", "K")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Locked = True
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Interior.Color = "&HD8D8D8"

 
    End If
    Else
             Sheet14.Unprotect Password:=getmsgstate
'            Sheet14.Range(Replace(Target.Address, "D", "I")).ClearContents     'Ankita_28/05/2025
             Sheet14.Range(Replace(Target.Address, "D", "I")).Locked = False
             Sheet14.Range(Replace(Target.Address, "D", "I")).Interior.Color = "&HCCFFCC"
'            Sheet14.Range(Replace(Target.Address, "D", "J")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "J")).Interior.Color = "&HCCFFCC"
'
'            Sheet14.Range(Replace(Target.Address, "D", "K")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "K")).Interior.Color = "&HCCFFCC"
    End If
End If




'80U------------------------
'Malli---------23/04/2025
If Not Intersect(Range("Disability_80U"), Range(Target.Address)) Is Nothing Then
Sheet14.Unprotect Password:=getmsgstate
        'If Sheet14.Range(Target.Address).Value = "(viii) autism" Then
          If Sheet14.Range(Target.Address).Value <> "(Select)" Then
        
'        If Sheet14.Range(Target.Address).Value <> "(viii) autism" And _
'        Sheet14.Range(Target.Address).Value <> "(ix) cerebral palsy" And _
'        Sheet14.Range(Target.Address).Value <> "(x) multiple disability" Then
             
'        If Sheet14.Range(Target.Address).Value <> "(i) autism, cerebral palsy, or multiple disabilities and" Then
         If Sheet14.Range(Target.Address).Value <> "(i) autism, cerebral palsy, or multiple disabilities" Then  'Ankita_14/05/2025
            'Ankita_28/05/2025
            If Sheet14.Range(Target.Address).Value = "(ii) others" And Sheet14.Range(Replace(Target.Address, "D", "F")).Value <> "" Then
            MsgBox """Type of disability shall be selected as ""(i) autism, cerebral palsy, or multiple disabilities"" when acknowledgement no. of form 10IA is provided"""
            Sheet14.Range(Target.Address).Value = "(Select)"
            Else:
            Sheet14.Range(Replace(Target.Address, "D", "F")).ClearContents
            Sheet14.Range(Replace(Target.Address, "D", "F")).Locked = True
            Sheet14.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HD8D8D8"
            End If
'            Sheet14.Range(Replace(Target.Address, "D", "G")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Locked = True
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HD8D8D8"
            
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Interior.Color = "&HCCFFCC"
            
            
        Else
        
        Sheet14.Unprotect Password:=getmsgstate
        
'            Sheet14.Range(Replace(Target.Address, "D", "F")).ClearContents 'Ankita_28/05/2025
            Sheet14.Range(Replace(Target.Address, "D", "F")).Locked = False
            Sheet14.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
            
'            Sheet14.Range(Replace(Target.Address, "D", "G")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
            
'            Sheet14.Range(Replace(Target.Address, "D", "H")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Locked = True
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Interior.Color = "&HD8D8D8"
            Sheet14.Protect Password:=getmsgstate
             

    End If
    
    Else
            Sheet14.Unprotect Password:=getmsgstate
'            Sheet14.Range(Replace(Target.Address, "D", "F")).ClearContents  'Ankita_28/05/2025
            Sheet14.Range(Replace(Target.Address, "D", "F")).Locked = False
            Sheet14.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
''
'            Sheet14.Range(Replace(Target.Address, "D", "G")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
'
'            Sheet14.Range(Replace(Target.Address, "D", "H")).ClearContents
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Locked = False
'            Sheet14.Range(Replace(Target.Address, "D", "H")).Interior.Color = "&HCCFFCC"
    End If
    
End If


'---------------------------


'---------------------------

'---------------------------
'If Not Application.Intersect(Sheet14.Range("DatefilingFm10IA_80DD"), Sheet14.Range(Target.Address)) Is Nothing Then
'
'    If Target.Value <> "00/00/0000" And Target.Value <> "" Then
'         If Not CheckDateddmmyyyy(Trim(Target.Value)) Then
'             'fmsgbox "* Date of filing of Form 10IA is invalid in Sheet 80DD. Please enter a valid Date in dd/mm/yyyy format." ', vbOKOnly, mdCommon.vbMessageCaption
'            fmsgbox "* Date of filing of Form 10IA is invalid in Sheet 80DD. Please enter date in valid format." ', vbOKOnly, mdCommon.vbMessageCaption
'
'            Target.ClearContents
'        End If
'    End If
'End If

'If Not Intersect(Range("DatefilingFm10IA_80DD"), Range(Target.Address)) Is Nothing Then
'              If Trim(Range("DatefilingFm10IA_80DD").Value) <> "" Then
'              If Dformat(Trim(Range("DatefilingFm10IA_80DD").Value), "yyyy-mm-dd") > Dformat(Trim(Sheet3.Range("Ver.Date").Value), "yyyy-mm-dd") Then
'                 fmsgbox "* Date of filing of Form 10IA can not be more than system date in Schedule 80DD"
'                       Range("DatefilingFm10IA_80DD").Select
'                        Range("DatefilingFm10IA_80DD").Value = ""
'              End If
'              End If
'End If
'If Not Intersect(Range("DatefilingFm10IA_80DD"), Range(Target.Address)) Is Nothing Then
'              If Trim(Range("DatefilingFm10IA_80DD").Value) <> "" Then
'              If Dformat(Trim(Range("DatefilingFm10IA_80DD").Value), "yyyy-mm-dd") > Dformat(Trim(Sheet3.Range("Ver.Date").Value), "yyyy-mm-dd") Then
'                 fmsgbox "* Date of filing of Form 10IA can not be more than system date in Schedule 80DD"
'                       Range("DatefilingFm10IA_80DD").Select
'                        Range("DatefilingFm10IA_80DD").Value = ""
'              End If
             ' End If
'End If

'sudha_2025-26

     
'              If Trim(Range("DatefilingFm10IA_80DD").Value) <> "" Then
'              If ChkMaxDate(Trim(Sheet14.Range("DatefilingFm10IA_80DD").Value), "") Then
'                 fmsgbox "* Date of filing of Form 10IA can not be before 01/04/2025 for AY 2025-26 in Schedule 80DD"
'                       Range("DatefilingFm10IA_80DD").Select
'                        Range("DatefilingFm10IA_80DD").Value = ""
'              End If
'              End If
              
              
'End If
'End If


'=======================




If Not Intersect(Range("Aadhaardependent_80DD"), Range(Target.Address)) Is Nothing Then
              If Trim(Range("Aadhaardependent_80DD").Value) <> "" Then
              If UCase(Trim(Range("Aadhaardependent_80DD").Value)) = UCase(Trim(Sheet1.Range("Sheet1.Aadhaar").Value)) Then
                 'fmsgbox "*Aadhar of the dependent cannot be same as assessee Aadhar in Part-A General Information."
                 fmsgbox "*""Aadhaar of the dependent cannot be same as assessee Aadhaar in Part-A General Information."""
                       Range("Aadhaardependent_80DD").Select
                        Range("Aadhaardependent_80DD").Value = ""
              End If
              End If
              If Len(Range("Aadhaardependent_80DD").Value) >= 12 Then
                 If IsNumeric(Range("Aadhaardependent_80DD").Value) Then
                    If Not checkfieldSuperSpecialcharacterDot(Range("Aadhaardependent_80DD").Value) Then
                       fmsgbox "* ""Invalid Aadhaar in Schedule 80DD."""
                       Range(Target.Address).ClearContents
                    End If
                  Else
                     fmsgbox "* ""Invalid Aadhaar in Schedule 80DD."""
                     Range(Target.Address).ClearContents
                 End If
              End If

                          
End If

If Not Intersect(Range("PANdependent_80DD"), Range(Target.Address)) Is Nothing Then
              If Trim(Range("PANdependent_80DD").Value) <> "" Then
              If UCase(Trim(Range("PANdependent_80DD").Value)) = UCase(Trim(Sheet1.Range("sheet1.PAN").Value)) Then
                 fmsgbox "*""PAN of the dependent cannot be same as assessee PAN in Part-A General Information."""
                       Range("PANdependent_80DD").Select
                        Range("PANdependent_80DD").Value = ""
              End If
              End If
End If

If Not Intersect(Range("PANdependent_80DD"), Range(Target.Address)) Is Nothing Then
         'If Not CheckDoneePAN_80DD(UCase(Trim("PANdependent_80DD"))) Then
         If Not CheckDoneePAN_80DD(UCase(Trim(Target.Value))) Then
            fmsgbox "*""Invalid PAN"" in Schedule 80DD" ', vbOKCancel, mdCommon.vbMessageCaption
            Sheet14.Range("PANdependent_80DD").Select
            Sheet14.Range("PANdependent_80DD").Value = ""
            GoTo endline
        End If
End If

'Newly added by Bindu
If Not Intersect(Range("AckNoFm10IAfiled_80DD"), Range(Target.Address)) Is Nothing Then
   If Not checkfieldspecialcharacter(Target.Value) Then
     fmsgbox " Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80DD."
     Range(Target.Address).ClearContents
   End If
End If

'Newly added by Bindu
If Not Intersect(Range("UDIDNum_80DD"), Range(Target.Address)) Is Nothing Then
   If Not checkfieldspecialcharacter(Target.Value) Then
     fmsgbox " UDID Number cannot contain special characters in Schedule 80DD."
     Range(Target.Address).ClearContents
   End If
End If


'Newly added by Bindu
If Not Intersect(Range("AckNoFm10IAfiled_80U"), Range(Target.Address)) Is Nothing Then
   If Not checkfieldspecialcharacter(Target.Value) Then
     fmsgbox " Ack. No. of Form 10IA filed cannot contain special characters in Schedule 80U."
     Range(Target.Address).ClearContents
   End If
End If


'Newly added by Bindu
If Not Intersect(Range("UDIDNum_80U"), Range(Target.Address)) Is Nothing Then
   If Not checkfieldspecialcharacter(Target.Value) Then
     fmsgbox " UDID Number cannot contain special characters in Schedule 80U."
     Range(Target.Address).ClearContents
   End If
End If


sPassword = EfilingCommon.getmsgstate
Sheet14.Protect Password:=sPassword
endline:
Application.EnableEvents = True

End Sub
 
 Sub app()
 Application.EnableEvents = True
 End Sub



