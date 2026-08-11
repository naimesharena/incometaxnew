Attribute VB_Name = "AY23_24Changes"
Option Explicit
Sub a()
Application.EnableEvents = False
End Sub
Sub b()
Application.EnableEvents = True
End Sub

Sub c()

Sheet2.Protect Password:=getmsgstate

End Sub
Function ChkMaxDOBDate23_24(dob As Variant, maxDefinedDOB As Variant) As Boolean
On Error Resume Next
     ChkMaxDOBDate23_24 = True
     If Len(dob) > 0 Then
        'PAG_E1
        ''If val(Mid(dob, 7, 4)) >= 2023 Then
        
        'PAG_C1 - 2024-25 Bindu
        
        'Ankita_25-26
'       If val(Mid(dob, 7, 4)) >= 2024 Then
'        If val(Mid(dob, 7, 4)) >= 2025 Then
            If val(Mid(dob, 7, 4)) >= CInt(Sheet5.Range("DOB_Year").Value) Then 'Ankita_27/12/2025====
             If val(Mid(dob, 4, 2)) >= 4 Then
                If val(Mid(dob, 1, 2)) >= 1 Then
                    ChkMaxDOBDate23_24 = False
                    Exit Function
                End If
        Else
               'PAG_E1
               'If val(Mid(dob, 4, 2)) <= 3 And val(Mid(dob, 7, 4)) <= 2023 Then
               'PAG_C1 2024-25 Naresh Agarwal
               'Ankita_25-26
'              If val(Mid(dob, 4, 2)) <= 3 And val(Mid(dob, 7, 4)) <= 2024 Then
              If val(Mid(dob, 4, 2)) <= 3 And val(Mid(dob, 7, 4)) <= CInt(Sheet5.Range("DOB_Year").Value) Then 'Ankita_27/12/2025====
              Else
                ChkMaxDOBDate23_24 = False
                    Exit Function
             End If
            End If
        Else
        End If
     End If
End Function
Function CheckDOB23_24(dob As Variant) As Boolean
On Error Resume Next
'The DOB should be in DD/MM/YYYY format only.

    CheckDOB23_24 = True
    If Trim(dob) = "" Or Not IsEmpty(dob) Then
        If Not FormatNCheckDate(dob) Then
            CheckDOB23_24 = False
           fmsgbox ("* Date of Birth in Sheet Income Details  must be a valid dd/mm/yyyy format")
            Exit Function
        End If
        
        'PAG_E1
        'If Not ChkMaxDOBDate23_24(dob, "31/03/2023") Then
        
        'PAG_C1 2024-25 Naresh Agarwal
        
        'Ankita_25-26
'        If Not ChkMaxDOBDate23_24(dob, "31/03/2024") Then
'         If Not ChkMaxDOBDate23_24(dob, "31/03/2025") Then

            'PAG_E1
            'fmsgbox ("* Date of Birth in Sheet Income Details  should not be on or after 01/04/2023 for A.Y.2023-24.")
            
            'PAG_C1  2024-25 Naresh Agarwal
'            fmsgbox ("* Date of Birth in Sheet Income Details  should not be on or after 01/04/2024 for A.Y.2024-25.")
'             fmsgbox ("* Date of Birth in Sheet Income Details should not be on or after 01/04/2025 for A.Y.2025-26.")
             
             'Ankita_27/12/2025=============
        Dim cutoff As Date
        cutoff = CDate(Sheet5.Range("DOB_1").Value)
        If Not ChkMaxDOBDate23_24(dob, Sheet5.Range("DOB_1").Value) Then
                 fmsgbox ("* Date of birth in sheet income details should not be on or after " & Dformat1(cutoff, "dd/mm/yyyy"))
             CheckDOB23_24 = False
            Exit Function
        Else
            AssesseeDob = dob
        End If
    End If
 End Function


Sub PartAGen_Country()
If UCase(Sheet1.Range("sheet1.Country")) = "91-INDIA" Then
          'Ankita 31/03/2026
            If Sheet1.Range("sheet1.StateCode1").Value = "" Then
            Sheet1.Range("sheet1.StateCode1") = "(Select)"
            End If
            If Sheet1.Range("sheet1.PinCode").Value = "" Then
            Sheet1.Range("sheet1.PinCode") = ""
            End If
            '-----------------------

    Sheet1.Range("sheet1.StateCode1") = "(Select)"
    Sheet1.Range("sheet1.PinCode") = ""
    Sheet1.Range("sheet1.MobileCountryCode").Value = "91"
    If UCase(Sheet1.Range("sheet1.StateCode1")) = "99-FOREIGN" Then
        MsgBox "Country Should not be India for Selected State" & Chr(13)
        Sheet1.Range("sheet1.Country").Value = "(Select)"
        'Sheet1.Range("sheet1.Country").Select
        'Sheet1.Range("sheet1.StateCode1").value = ""
    End If

            'Ankita 31/03/2026
            If Not Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False Then
            Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
            Sheet1.Range("sheet1.PinCode").MergeArea.ClearContents
            Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)
            End If
            
            If Not Sheet1.Range("HASZIP").MergeArea.Locked = True Then
            Sheet1.Range("HASZIP").MergeArea.ClearContents
            Sheet1.Range("HASZIP").MergeArea.Locked = True
            Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
            End If
            
            If Not Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True Then
            Sheet1.Range("sheet1.ZipCode").MergeArea.ClearContents
            Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
            Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
            End If
            '----------------------


'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("sheet1.PinCode").Value = ""
'    Sheet1.Range("sheet1.PinCode").Locked = False
'    Sheet1.Range("sheet1.PinCode").Interior.Color = (&HCCFFCC)
'
'    Sheet1.Range("HASZIP").MergeArea.Value = ""
'    Sheet1.Range("HASZIP").MergeArea.Locked = True
'    Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
'
'    Sheet1.Range("sheet1.ZipCode").MergeArea.Value = ""
'    Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
'    Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
    Sheet1.Protect Password:=getmsgstate

ElseIf Sheet1.Range("sheet1.Country").Value = "(Select)" Then
    Sheet1.Range("sheet1.StateCode1") = "(Select)"

    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("sheet1.PinCode").Value = ""
    Sheet1.Range("sheet1.PinCode").Locked = False
    Sheet1.Range("sheet1.PinCode").Interior.Color = (&HCCFFCC)
    
    Sheet1.Range("HASZIP").MergeArea.Value = ""
    Sheet1.Range("HASZIP").MergeArea.Locked = True
    Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)
    
    Sheet1.Range("sheet1.ZipCode").MergeArea.Value = ""
    Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
    Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
    Sheet1.Protect Password:=getmsgstate
ElseIf UCase(Sheet1.Range("sheet1.Country").Value) <> "91-INDIA" Then
    'Change-19.11.2022.101.02B
    If UCase(Sheet1.Range("sheet1.StateCode1").Value) <> "99-FOREIGN" And Sheet1.Range("sheet1.StateCode1").Value <> "(Select)" And Sheet1.Range("sheet1.StateCode1").Value <> "" Then
        'MsgBox "Country Should be India for Selected State" & Chr(13)
        MsgBox "Country cannot be other than India as you have selected an Indian state" & Chr(13) 'Malli
        Sheet1.Range("sheet1.Country").Value = ""
    End If
'    If UCase(Sheet1.Range("sheet1.StateCode1").Value) <> "99-FOREIGN" And Sheet1.Range("sheet1.StateCode1").Value <> "(Select)" And Sheet1.Range("sheet1.StateCode1").Value <> "" Then
'        MsgBox "Country Should be India for Selected State" & Chr(13)
'        Sheet1.Range("sheet1.Country").Value = ""
'    End If
    If (Sheet1.Range("sheet1.StateCode1").Value) = "(Select)" Or (Sheet1.Range("sheet1.StateCode1").Value) = "" Then
        Sheet1.Range("sheet1.StateCode1").Value = "99-FOREIGN"
    End If
    '---end
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("sheet1.PinCode").MergeArea.Value = ""
    Sheet1.Range("sheet1.PinCode").MergeArea.Locked = True
    Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HD8D8D8)
    
    Sheet1.Range("HASZIP").MergeArea.Value = ""
    Sheet1.Range("HASZIP").MergeArea.Locked = False
    Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HCCFFCC)
    
    Sheet1.Range("sheet1.ZipCode").MergeArea.Value = ""
    Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
    Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
    Sheet1.Protect Password:=getmsgstate
End If
End Sub

'Ankita_14/01/2026_V0.2==========
Sub PartAGen_Country1()
If UCase(Sheet1.Range("sheet1.Country1")) = "91-INDIA" Then
    Sheet1.Range("sheet1.StateCode2") = "(Select)"
    Sheet1.Range("sheet1.PinCode1") = ""
    Sheet1.Range("sheet1.MobileCountryCode").Value = "91"
    If UCase(Sheet1.Range("sheet1.StateCode2")) = "99-FOREIGN" Then
        MsgBox "Country Should not be India for Selected State" & Chr(13)
        Sheet1.Range("sheet1.Country1").Value = "(Select)"
    End If
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("sheet1.PinCode1").Value = ""
    Sheet1.Range("sheet1.PinCode1").Locked = False
    Sheet1.Range("sheet1.PinCode1").Interior.Color = (&HCCFFCC)
    Sheet1.Range("HASZIP1").MergeArea.Value = ""
    Sheet1.Range("HASZIP1").MergeArea.Locked = True
    Sheet1.Range("HASZIP1").MergeArea.Interior.Color = (&HD8D8D8)
    Sheet1.Range("sheet1.ZipCode1").MergeArea.Value = ""
    Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = True
    Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = (&HD8D8D8)
    Sheet1.Protect Password:=getmsgstate
ElseIf Sheet1.Range("sheet1.Country1").Value = "(Select)" Then
    Sheet1.Range("sheet1.StateCode2") = "(Select)"
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("sheet1.PinCode1").Value = ""
    Sheet1.Range("sheet1.PinCode1").Locked = False
    Sheet1.Range("sheet1.PinCode1").Interior.Color = (&HCCFFCC)
    Sheet1.Range("HASZIP1").MergeArea.Value = ""
    Sheet1.Range("HASZIP1").MergeArea.Locked = True
    Sheet1.Range("HASZIP1").MergeArea.Interior.Color = (&HD8D8D8)
    Sheet1.Range("sheet1.ZipCode1").MergeArea.Value = ""
    Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = True
    Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = (&HD8D8D8)
    Sheet1.Protect Password:=getmsgstate
ElseIf UCase(Sheet1.Range("sheet1.Country1").Value) <> "91-INDIA" Then
    If UCase(Sheet1.Range("sheet1.StateCode2").Value) <> "99-FOREIGN" And Sheet1.Range("sheet1.StateCode2").Value <> "(Select)" And Sheet1.Range("sheet1.StateCode2").Value <> "" Then
        MsgBox "Country cannot be other than India as you have selected an Indian state" & Chr(13)
        Sheet1.Range("sheet1.Country1").Value = ""
    End If
    If (Sheet1.Range("sheet1.StateCode2").Value) = "(Select)" Or (Sheet1.Range("sheet1.StateCode2").Value) = "" Then
        Sheet1.Range("sheet1.StateCode2").Value = "99-FOREIGN"
    End If
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("sheet1.PinCode1").MergeArea.Value = ""
    Sheet1.Range("sheet1.PinCode1").MergeArea.Locked = True
    Sheet1.Range("sheet1.PinCode1").MergeArea.Interior.Color = (&HD8D8D8)
    Sheet1.Range("HASZIP1").MergeArea.Value = ""
    Sheet1.Range("HASZIP1").MergeArea.Locked = False
    Sheet1.Range("HASZIP1").MergeArea.Interior.Color = (&HCCFFCC)
    Sheet1.Range("sheet1.ZipCode1").MergeArea.Value = ""
    Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = False
    Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = (&HCCFFCC)
    Sheet1.Protect Password:=getmsgstate
End If
End Sub

'================================
         
Sub PartAGen_FilingStatusCode1()
If Range("sheet1.ReturnFileSec").Value = "139(4)-Belated" Then
    Dim vdate As Variant
    vdate = Sheet3.Range("Ver.Date").Value
    'If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2023-08-01") Then
    'PAG_C3 AY 2024-25 change

'    If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2025-09-16") Then   'Ankita_22/07/2025 'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
      'Ayush_DueDate_08/09/2025
      'Ankita_09/01/2026========
          Dim dueDt As Date
          dueDt = Sheet5.Range("DueDate1").Value

      If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), Dformat(Sheet5.Range("DueDate11").Value, "yyyy-mm-dd")) Then
        'fmsgbox "* Filing section u/s. 139(4) cannot be selected till 31st July, 2023 or extended due date "
        'PAG_C3 AY 2024-25 change
        'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
'         fmsgbox "* Filing section u/s. 139(4) cannot be selected till 31st July, 2025 or extended due date "
          fmsgbox "* Filing section u/s. 139(4) cannot be selected till " & FormatDateWithOrdinal1(dueDt) & " or extended due date "
          Range("sheet1.ReturnFileSec").Value = "(Select)"
    End If
End If
End Sub

Sub PartAGen_FilingStatusCode2()
If Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
    Dim vdate As Variant
    vdate = Sheet3.Range("Ver.Date").Value
    'PAG_C3 AY 2024-25 change
    'If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2024-01-01") Then
    'Ankita_UR
    
'    If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2025-01-01") Then
         'PAG_C3 AY 2024-25 change
'         fmsgbox "* Filing section u/s. 139(8A) cannot be selected till 31st Dec, 2023 or extended due date "
'         fmsgbox "* Filing section u/s. 139(8A) cannot be selected till 31st Dec, 2024 or extended due date "

    'Added by Ankita on 12/12/2024
    'Ankita_09/01/2026=================
    If Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
              Dim cutoff_139_8A As Date
              cutoff_139_8A = CDate(Sheet5.Range("Date_8A").Value)
        If Not ChkMinInclusiveDate(Dformat5(vdate, "yyyy-mm-dd"), Sheet5.Range("Date_8A").Value) Then
'        If Not ChkMinInclusiveDate(Dformat5(vdate, "yyyy-mm-dd") >= Dformat5(Sheet5.Range("Date_8A").Value, "yyyy-mm-dd")) Then
    '        fmsgbox "* Filing section u/s. 139(8A) cannot be selected till 31st Dec, 2025 or extended due date "
             fmsgbox "* Filing section u/s. 139(8A) cannot be selected till " & Dformat5(cutoff_139_8A, "dd/mm/yyyy") & " or extended due date "
             Range("sheet1.ReturnFileSec").Value = "(Select)"
        End If
      End If
    End If
End Sub


'method to calculate the age from dateOfBirth
Function calculateAge23_24(dob As Variant) As Long

On Error Resume Next
    'calculateAge23_24 = (2023) - val(Mid(dob, 7, 4))
    'PAG_C1 AY 2024-25 Change
    'Changed year from 2024 to 2025 by Ankita on 16/12/2024
    calculateAge23_24 = (2025) - val(Mid(dob, 7, 4))
    If 4 < val(Mid(dob, 4, 2)) Then
        calculateAge23_24 = calculateAge23_24 - 1
    ElseIf val(Mid(dob, 4, 2)) = 4 And 1 < val(Mid(dob, 1, 2)) Then
        calculateAge23_24 = calculateAge23_24 - 1
    End If

End Function
Sub Noticedate23_24()
    'If EfilingCommon.checkFirstDateBefore(Range("sheet1.NoticeDate").Value, "31/03/2023") Then
    If EfilingCommon.checkFirstDateBefore(Range("sheet1.NoticeDate").Value, "31/03/2025") Then  'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
       'EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*  Date of Notice/Order in Sheet Income Details cannot be prior to 01/04/2023 " & Chr(13)
        EfilingCommon.MsgPISheet = EfilingCommon.MsgPISheet + "*  Date of Notice/Order cannot be prior to 01/04/2025 " & Chr(13)     'Year Changed from 2024 to 2025 by Ankita on 16/12/2024
  End If
End Sub

Function CheckNoticeDateBefore23_24(dob As Variant) As Boolean
On Error Resume Next
    'The DOB should be in DD/MM/YYYY format only.
    CheckNoticeDateBefore23_24 = True
    If Len(dob) > 0 Then
        If Not FormatNCheckDate(dob) Then
            CheckNoticeDateBefore23_24 = False
            'fmsgbox ("* Date of Notice/Order in Sheet Income Details  must be a valid dd/mm/yyyy format")
             fmsgbox ("*""Please enter a date in ""Date of notice /order"" field in dd/mm/yyyy format""")
            Exit Function
        End If
        'PAG_E8
        'If EfilingCommon.checkFirstDateBefore(dob, "31/03/2023") Then
        'fmsgbox ("* Date of Notice/Order cannot be prior to 01/04/2023")
        
        
        'PAG_C8 AY 2024-25 change
'        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2024") Then
'            fmsgbox ("* Date of Notice/Order cannot be prior to 01/04/2024")
'       Year Changed by Ankita on 13/12/2024

'        If EfilingCommon.checkFirstDateBefore(dob, "31/03/2025") Then
'Ankita_17/12/2025==============

Dim cuttoff As Date
cuttoff = CDate(Sheet5.Range("DOB_extended").Value)
        If EfilingCommon.checkFirstDateBefore(dob, Sheet5.Range("Date_Notice").Value) Then
            fmsgbox ("* ""Date of Notice/Order cannot be prior to " & Dformat1(cuttoff, "dd/mm/yyyy") & ".""")
            CheckNoticeDateBefore23_24 = False
            Exit Function
        Else
            NoticeDate = dob
        End If
    End If
End Function

Sub Check_GrossSalary23_24(ByVal Target As Range)
    Sheet1.Unprotect Password:=getmsgstate
    Application.EnableEvents = False
    
    If Target.Value = "Not Applicable (eg. Family pension etc)" Then
        Range("IncD.Allowances").MergeArea.ClearContents
        Range("IncD.Allowances").MergeArea.Locked = True
        Range("IncD.Allowances").MergeArea.Interior.Color = (&HD8D8D8)

        Range("IncD.Perquisites").MergeArea.ClearContents
        Range("IncD.Perquisites").MergeArea.Locked = True
        Range("IncD.Perquisites").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("IncD.Profits").MergeArea.ClearContents
        Range("IncD.Profits").MergeArea.Locked = True
        Range("IncD.Profits").MergeArea.Interior.Color = (&HD8D8D8)
        
        'Ankita_16/01/2026======
'        Range("IncomeNotified89A_AmountUS").MergeArea.ClearContents
'        Range("IncomeNotified89A_AmountUS").MergeArea.Locked = True
'        Range("IncomeNotified89A_AmountUS").MergeArea.Interior.Color = (&HD8D8D8)
                
        'Ankita_16/01/2026======
'        Range("IncomeNotified89A_AmountUK").MergeArea.ClearContents
'        Range("IncomeNotified89A_AmountUK").MergeArea.Locked = True
'        Range("IncomeNotified89A_AmountUK").MergeArea.Interior.Color = (&HD8D8D8)
        
        'Ankita_16/01/2026======
'        Range("IncomeNotified89A_AmountCan").MergeArea.ClearContents
'        Range("IncomeNotified89A_AmountCan").MergeArea.Locked = True
'        Range("IncomeNotified89A_AmountCan").MergeArea.Interior.Color = (&HD8D8D8)
        
        'Ankita_16/01/2026======
'        Range("IncomeNotifiedOther89A").MergeArea.ClearContents
'        Range("IncomeNotifiedOther89A").MergeArea.Locked = True
'        Range("IncomeNotifiedOther89A").MergeArea.Interior.Color = (&HD8D8D8)
        
        
'        Range("Increliefus89A").MergeArea.ClearContents
'        Range("Increliefus89A").MergeArea.Locked = True
'        Range("Increliefus89A").MergeArea.Interior.Color = (&HD8D8D8)
        
        'Malli
       
       ' Range("IncD.Deduction16ia").MergeArea.Locked = True
        'Range("IncD.Deduction16ia").MergeArea.Interior.Color = (&HD8D8D8)
       ' -----------------------------------------
        
        
        Range("IncD.Deduction16").MergeArea.ClearContents
        Range("IncD.Deduction16").MergeArea.Locked = True
        Range("IncD.Deduction16").MergeArea.Interior.Color = (&HD8D8D8)
        
        
        Range("IncD.Deduction16ic").MergeArea.ClearContents
        Range("IncD.Deduction16ic").MergeArea.Locked = True
        Range("IncD.Deduction16ic").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("Others.NOI_1").ClearContents
        Sheet1.Range("Others.NOI_1").Locked = True
        Sheet1.Range("Others.NOI_1").Interior.Color = "&HD8D8D8"
'        Sheet1.Range("Nature_Others_1").ClearContents
'        Sheet1.Range("Nature_Others_1").Locked = True
'        Sheet1.Range("Nature_Others_1").Interior.Color = "&HD8D8D8"
        Sheet1.Range("Others.Amount_1").ClearContents
        Sheet1.Range("Others.Amount_1").Locked = True
        Sheet1.Range("Others.Amount_1").Interior.Color = "&HD8D8D8"

    Else
        Range("IncD.Allowances").MergeArea.Locked = False
        Range("IncD.Allowances").MergeArea.Interior.Color = (&HCCFFCC)
        

        Range("IncD.Perquisites").MergeArea.Locked = False
        Range("IncD.Perquisites").MergeArea.Interior.Color = (&HCCFFCC)


        Range("IncD.Profits").MergeArea.Locked = False
        Range("IncD.Profits").MergeArea.Interior.Color = (&HCCFFCC)

        'Ankita_16/01/2026======
'        Range("IncomeNotified89A_AmountUS").MergeArea.Locked = False
'        Range("IncomeNotified89A_AmountUS").MergeArea.Interior.Color = (&HCCFFCC)

      'Ankita_16/01/2026======
'        Range("IncomeNotified89A_AmountUK").MergeArea.Locked = False
'        Range("IncomeNotified89A_AmountUK").MergeArea.Interior.Color = (&HCCFFCC)
   
        'Ankita_16/01/2026======
'        Range("IncomeNotified89A_AmountCan").MergeArea.Locked = False
'        Range("IncomeNotified89A_AmountCan").MergeArea.Interior.Color = (&HCCFFCC)
        
        'Ankita_16/01/2026======
'        Range("IncomeNotifiedOther89A").MergeArea.Locked = False
'        Range("IncomeNotifiedOther89A").MergeArea.Interior.Color = (&HCCFFCC)


'        Range("Increliefus89A").MergeArea.Locked = False
'        Range("Increliefus89A").MergeArea.Interior.Color = (&HCCFFCC)
       ' Range("Increliefus89A").Value = 0
        
        
        'Malli
'
'        Range("IncD.Deduction16ia").MergeArea.Locked = True
'        Range("IncD.Deduction16ia").MergeArea.Interior.Color = (&HCCFFCC)
''

        '----------------
'

   If Sheet5.Range("BacValue").Value <> 1 Then    '03-02-2026  Malli SIT-109371
        Range("IncD.Deduction16").MergeArea.Locked = False
        Range("IncD.Deduction16").MergeArea.Interior.Color = (&HCCFFCC)

        Range("IncD.Deduction16ic").MergeArea.Locked = False
        Range("IncD.Deduction16ic").MergeArea.Interior.Color = (&HCCFFCC)
    End If
      '  Range("IncD.Deduction16ic").Value = 0
        
        Sheet1.Range("Others.NOI_1").Locked = False
        Sheet1.Range("Others.NOI_1").Interior.Color = "&HCCFFCC"
''        Sheet1.Range("Nature_Others_1").Locked = False
'        'Sheet1.Range("Nature_Others_1").Interior.Color = "&HD8D8D8"
       Sheet1.Range("Others.Amount_1").Locked = False
      Sheet1.Range("Others.Amount_1").Interior.Color = "&HCCFFCC"

    End If
    
    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate
End Sub


Sub Nature_ExemptDropdown()
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate

'Ankita_11/02/2026===========
'changed as per V0.4====
    If Sheet5.Range("BacValue").Value <> 1 Then
        If Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "State Government" Then
'        If Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "State Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958" Then
            Sheet1.Range("Others.NOI_1").Validation.Delete
            Sheet1.Range("Others.NOI_1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Dropdown115BAC_OTHY_CG_SG"
        ElseIf Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - State Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Public sector undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Others" Then
            Sheet1.Range("Others.NOI_1").Validation.Delete
            Sheet1.Range("Others.NOI_1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Dropdown115BAC_OTHY_CGSG_PCGSG"
        Else
            Sheet1.Range("Others.NOI_1").Validation.Delete
            Sheet1.Range("Others.NOI_1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Dropdown_115BAC_OTHY"
        End If
    ElseIf Sheet5.Range("BacValue").Value = 1 Then
        If Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "State Government" Then
'         If Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "State Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958" Then
            Sheet1.Range("Others.NOI_1").Validation.Delete
            Sheet1.Range("Others.NOI_1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Dropdown_115BAC_Y_CG_SG"
        ElseIf Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - State Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Public sector undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Others" Then
            Sheet1.Range("Others.NOI_1").Validation.Delete
            Sheet1.Range("Others.NOI_1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Dropdown_115BAC_Y_CGSG_PCGSG"
        Else
            Sheet1.Range("Others.NOI_1").Validation.Delete
            Sheet1.Range("Others.NOI_1").Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Dropdown_115BAC_Y"
        End If
    End If

Application.EnableEvents = True
Sheet1.Protect Password:=sPassword
End Sub

Sub Nature_Amount23_24(ByVal Target As Range)
      'Changed as per DE sheet v0.8 by Ankita
       If Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "State Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - State Government" Then
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(10)-Death-cum-retirement gratuity received " Then
        '        If (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
            If Sheet1.Range(Target.Address).Value > 2500000 Then  'Changed as per DE sheet v0.8 by Ankita
'                fmsgbox ("*For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed 20 Lakhs")
                 fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 25 Lakhs where nature of employment is selected as CG, SG, SG-Pensioners, CG-Pensioners" & Chr(13))
                 Sheet1.Range(Target.Address).Value = ""
                 Sheet1.Range(Target.Address).Select
            End If
        '        End If
        End If 'Changed as per DE sheet v0.8 by Ankita
        Else
          If Sheet1.Range("sheet1.EmployerCategory1").Value = "Public Sector Undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Public sector undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Others" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Others" Then
           If Trim(Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(10)-Death-cum-retirement gratuity received" Then
               If Sheet1.Range(Target.Address).Value > 2000000 Then
                   fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 20 Lakhs where nature of employment is selected as Public Sector Undertaking, PSU - Pensioners, others - Pensioners, Others" & Chr(13))
                   Range(Target.Address).Value = ""
                   Range(Target.Address).Select
               End If
           End If
          End If
        End If
        
        
        '=======================================================================
        'Added by Shrutika(17/04/2025)NewDev
        
'        If (Sheet1.Range(Replace(Target.Address, "AO", "J")).Value) = "Sec 10(13A)-Allowance to meet expenditure incurred on house rent" Then
'            If Sheet1.Range(Target.Address).Value > 2000000 Then
'                fmsgbox ("*""If claim of deduction under exempt allowance under Section 10(13A) is more than Rs 20,00,000, you are required to file return in ITR 2 or 3. Please refer Rule 12 for further details.""")
'                Sheet1.Range(Target.Address).Value = ""
'                Sheet1.Range(Target.Address).Select
'            End If
'
'        End If
        '-----------------------------------
        
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
        
            If Sheet1.Range(Target.Address).Value > Sheet1.Range("IncD.Allowances").Value Then
                fmsgbox ("*Amount u/s 10(10AA) can not exceed Salary as per sec 17(1).")
                Sheet1.Range(Target.Address).Value = ""
        'If (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
           ' ElseIf (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Pub" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
            ElseIf (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Oth") Then
                'INC_E3
                'If Sheet1.Range(Target.Address).Value > 300000 Then
                
                    'fmsgbox ("*Amount u/s 10(10AA) can not exceed Salary as per sec 17(1)")
                    'fmsgbox ("Deduction u/s 10(10AA) shall be restricted to Rs. 3 Lakh in case of  employer category is selected as other than Central or  state Government, CG-Pensioners, SG-Pensioners, PSU or PSU-Pensioners")
                    
                'INC_C3 2024-25 Bindu
               ' If Sheet1.Range(Target.Address).Value > 250000 Then  'Malli comented
               If Sheet1.Range(Target.Address).Value > 2500000 Then
              ' Ankita_30/05/2025
                    fmsgbox ("Deduction u/s 10(10AA) shall be restricted to Rs. 25 Lakh in case of  employer category is selected as other than Central or state Government, CG-Pensioners, SG-Pensioners")
                    Sheet1.Range(Target.Address).Value = ""
                End If
               ElseIf (Range("Sheet1.EmployerCategory1").Value = "Pensioners - Public sector undertaking" Or Range("Sheet1.EmployerCategory1").Value = "Public Sector Undertaking") Then
                   ' Ankita_30/05/2025
                    fmsgbox ("* ""Deduction u/s 10(10AA) shall be restricted to Rs. 25 Lakh in case of  employer category is selected as other than Central or state Government, CG-Pensioners, SG-Pensioners""")
            End If
        End If
        
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Section 10(14)(ii) -  Transport allowance granted to certain physically handicapped assessee" Then
        '        If (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
            If Sheet5.Range("BacValue").Value = 1 Then
                If Sheet1.Range(Target.Address).Value > 38400 Then
                ' Ankita_08/05/2025_Changed as per DE sheet v0.7
'                   fmsgbox ("*For Section 10(14)(ii) -  Transport allowance granted to certain physically handicapped assessee is selected than amount should not exceed 38400")
                    fmsgbox ("* ""Exempt allowance under Section 10(14)(ii) cannot exceed Rs 38,400""")
                    Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
                End If
            End If
        '        End If
        End If
        
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
        
            If Sheet1.Range(Target.Address).Value > 500000 Then
                fmsgbox ("*For Sec 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.5 lakhs")
                Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
            End If
        End If
        
        'INC_C3 2024-25 Bindu
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
        
            If Sheet1.Range(Target.Address).Value > 500000 Then
                fmsgbox ("*For 10(10B) Second proviso - Compensation under scheme approved by the Central Government can not exceed Rs.5 lakhs")
                Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
            End If
        End If
        
        'Commented by Shrutika(07/5/2025)
'        If (Sheet1.Range(Replace(Target.Address, "AO", "J")).Value) = "Sec 10(13A)-Allowance to meet expenditure incurred on house rent" Then
'            If Sheet1.Range(Target.Address).Value > (1 / 3 * Sheet1.Range("IncD.Allowances").Value) Then
'            'Change-13.02.2023.103.ID.11
'                'fmsgbox ("*Sec 10(13A)-Allowance to meet expenditure incurred on house rent can not exceed 33.33% of Salary 17(1)")
'                'INC_E3
'                'fmsgbox ("*Sec 10(13A)-Allowance to meet expenditure incurred on house rent can not exceed 50% of Salary 17(1)")
'                'INC_C3 2024-25 Bindu
'              '  fmsgbox ("*Exempt allowance u/s 10(13A) cannot be more than the limit prescribed.")
'
'                Sheet1.Range(Target.Address).Value = ""
'                Sheet1.Range(Target.Address).Select
'            End If
'        End If
'
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(5)-Leave Travel concession/assistance" Then
                
            If Sheet1.Range(Target.Address).Value > Sheet1.Range("IncD.Allowances").Value Then
                fmsgbox ("*Amount u/s 10(5) can not exceed Salary as per sec 17(1).")
                Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
            End If
        End If
        
        
        
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(10A)-Commuted value of pension received" Then
                
            If Sheet1.Range(Target.Address).Value > Sheet1.Range("IncD.Allowances").Value Then
                fmsgbox ("*Amount u/s 10(10A) can not exceed Salary as per sec 17(1).")
                Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
            End If
        End If
        
        
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
                
            If Sheet1.Range(Target.Address).Value > Sheet1.Range("IncD.Allowances").Value Then
'            Change.27.02.2023.102.IDS.47
'                fmsgbox ("*For Sec 10(10AA)-Earned leave encashment can not be more than Salary as per section 17(1).")
                fmsgbox ("*Amount u/s 10(10AA) can not exceed Salary as per sec 17(1)")
'            End ChangeIDS
                Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
            End If
        End If
        
        
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service" Then
                
            If Sheet1.Range(Target.Address).Value > 500000 Then
'            Amount u/s 10(10C) can not exceed Rs. 5 lakhs
'                fmsgbox ("*For Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service can not exceed Rs. 5,00,000.")
                fmsgbox ("*Amount u/s 10(10C) can not exceed Rs. 5 lakhs.")
                Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
            End If
        End If
        
        
        
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(10CC)-Tax paid by employer on non-monetary perquisite" Then
                
            If Sheet1.Range(Target.Address).Value > Sheet1.Range("IncD.Perquisites").Value Then
'            Amount u/s 10(10CC) can not exceed Value of perquisites as per sec 17(2)
'                fmsgbox ("*For Sec 10(10CC)-Tax paid by employer on non-monetary perquisite can not be more than Value of perquisites as per section 17(2).")
                fmsgbox ("*Amount u/s 10(10CC) can not exceed Value of perquisites as per sec 17(2).")
                Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
            End If
        End If
        
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(14)(i)-Prescribed Allowances or benefits (not in a nature of perquisite) specifically granted to meet expenses wholly, necessarily and exclusively and to the extent actually incurred, in performance of duties of office or employment" Then
                
            If Sheet1.Range(Target.Address).Value > Sheet1.Range("IncD.Allowances").Value Then
'                fmsgbox ("*For Sec 10(14)(i) Prescribed Allowances or benefits (not in a nature of perquisite) specifically granted to meet expenses wholly, necessarily and exclusively and to the extent actually incurred, in performance of duties of office or employment can not be more than Salary as per section 17(1).")
                fmsgbox ("*Amount u/s 10(14)(i) can not exceed Salary as per sec 17(1).")
                Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
            End If
        End If
        
        If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(14)(ii)-Prescribed Allowances or benefits granted to meet personal expenses in performance of duties of office or employment or to compensate him for increased cost of living" Then
                
            If Sheet1.Range(Target.Address).Value > Sheet1.Range("IncD.Allowances").Value Then
                'Amount u/s 10(14)(ii) can not exceed Salary as per sec 17(1)
'                fmsgbox ("*For Sec 10(14)(ii)  Prescribed Allowances or benefits granted to meet personal expenses in performance of duties of office or employment or to compensate him for increased cost of living can not be more than Salary as per section 17(1)")
                fmsgbox ("*Amount u/s 10(14)(ii) can not exceed Salary as per sec 17(1)")
                Sheet1.Range(Target.Address).Value = ""
                Sheet1.Range(Target.Address).Select
            End If
        End If

End Sub

'Added by Ankita on 12/12/2024 =====================
'Ankita_29/11

Sub Nature_Amount23_24_2(ByVal Target As String)
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(10)-Death-cum-retirement gratuity received " Then

'         If (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
         If Sheet1.Range("sheet1.EmployerCategory1").Value = "Public Sector Undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Public sector undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Others" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Others" Then
            If Sheet1.Range(Target).Value > 2000000 Then
'              fmsgbox ("*For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed 20 Lakhs")
               fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 20 Lakhs where nature of employment is selected as Public Sector Undertaking, PSU - Pensioners, others - Pensioners, Others" & Chr(13))
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
                End If
        End If

'Ankita_25/06/2025=================================
'         If (Sheet1.Range(Replace(Target, "AO", "J")).Value) = "Sec 10(10)-Death-cum-retirement gratuity received " Then
'          If (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
'            If Sheet1.Range(Target).Value > 2500000 Then
'              fmsgbox ("*For Sec 10(10)-Death-cum-retirement gratuity is selected than amount should not exceed 20 Lakhs")
'               fmsgboxsmall ("* Amount u/s 10(10) can not exceed Rs. 25 Lakhs where nature of employment is selected as Public Sector Undertaking, PSU - Pensioners, others - Pensioners, Others" & Chr(13))
'                Sheet1.Range(Target).Value = ""
'                Sheet1.Range(Target).Select
'            End If
'                End If
'        End If
'
'==================================================

        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
        
            If Sheet1.Range(Target).Value > Sheet1.Range("IncD.Allowances").Value Then
                fmsgbox ("*Amount u/s 10(10AA) can not exceed Salary as per sec 17(1).")
                Sheet1.Range(Target).Value = ""
        'If (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
           ' ElseIf (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Pub" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
            ElseIf (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Oth") Then
                'INC_E3
                'If Sheet1.Range(Target).Value > 300000 Then
                
                    'fmsgbox ("*Amount u/s 10(10AA) can not exceed Salary as per sec 17(1)")
                    'fmsgbox ("Deduction u/s 10(10AA) shall be restricted to Rs. 3 Lakh in case of  employer category is selected as other than Central or  state Government, CG-Pensioners, SG-Pensioners, PSU or PSU-Pensioners")
                    
                'INC_C3 2024-25 Bindu
               ' If Sheet1.Range(Target).Value > 250000 Then  'Malli comented
               If Sheet1.Range(Target).Value > 2500000 Then
                    fmsgbox ("Deduction u/s 10(10AA) shall be restricted to Rs. 25 Lakh in case of  employer category is selected as other than Central or  state Government, CG-Pensioners, SG-Pensioners")
                    Sheet1.Range(Target).Value = ""
                End If
               ElseIf (Range("Sheet1.EmployerCategory1").Value = "Pensioners - Public sector undertaking" Or Range("Sheet1.EmployerCategory1").Value = "Public Sector Undertaking") Then
                    fmsgbox ("Deduction u/s 10(10AA) shall be restricted to Rs. 25 Lakh in case of employer category is selected as other than Central or  state Government, CG-Pensioners, SG-Pensioners")
            End If
        End If
        
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Section 10(14)(ii) -  Transport allowance granted to certain physically handicapped assessee" Then
        '        If (Mid(Sheet1.Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
            If Sheet5.Range("BacValue").Value = 1 Then
                If Sheet1.Range(Target).Value > 38400 Then
                    fmsgbox ("*For Section 10(14)(ii) -  Transport allowance granted to certain physically handicapped assessee is selected than amount should not exceed 38400")
                    Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
                End If
            End If
        '        End If
        End If
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(10B) First proviso - Compensation limit notified by CG in the Official Gazette" Then
        
            If Sheet1.Range(Target).Value > 500000 Then
                fmsgbox ("*For Sec 10(10B)-First Proviso-  Compensation limit notified by CG in the Official Gazette can not exceed Rs.5 lakhs")
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If
        
        'INC_C3 2024-25 Bindu
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
        
            If Sheet1.Range(Target).Value > 500000 Then
                fmsgbox ("*For 10(10B) Second proviso - Compensation under scheme approved by the Central Government can not exceed Rs.5 lakhs")
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If
        
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(13A)-Allowance to meet expenditure incurred on house rent" Then
            If Sheet1.Range(Target).Value > (1 / 3 * Sheet1.Range("IncD.Allowances").Value) Then
            'Change-13.02.2023.103.ID.11
                'fmsgbox ("*Sec 10(13A)-Allowance to meet expenditure incurred on house rent can not exceed 33.33% of Salary 17(1)")
                'INC_E3
                'fmsgbox ("*Sec 10(13A)-Allowance to meet expenditure incurred on house rent can not exceed 50% of Salary 17(1)")
                'INC_C3 2024-25 Bindu
                fmsgbox ("*Exempt allowance u/s 10(13A) cannot be more than the limit prescribed.")
                
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(5)-Leave Travel concession/assistance" Then
                
            If Sheet1.Range(Target).Value > Sheet1.Range("IncD.Allowances").Value Then
                fmsgbox ("*Amount u/s 10(5) can not exceed Salary as per sec 17(1).")
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If
        
        
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(10A)-Commuted value of pension received" Then
                
            If Sheet1.Range(Target).Value > Sheet1.Range("IncD.Allowances").Value Then
                fmsgbox ("*Amount u/s 10(10A) can not exceed Salary as per sec 17(1).")
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If
        
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(10AA)-Earned leave encashment on Retirement" Then
                
            If Sheet1.Range(Target).Value > Sheet1.Range("IncD.Allowances").Value Then
'            Change.27.02.2023.102.IDS.47
'                fmsgbox ("*For Sec 10(10AA)-Earned leave encashment can not be more than Salary as per section 17(1).")
                fmsgbox ("*Amount u/s 10(10AA) can not exceed Salary as per sec 17(1)")
'            End ChangeIDS
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If
        
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service" Then
                
            If Sheet1.Range(Target).Value > 500000 Then
'            Amount u/s 10(10C) can not exceed Rs. 5 lakhs
'                fmsgbox ("*For Sec 10(10C)-Amount received/receivable on voluntary retirement or termination of service can not exceed Rs. 5,00,000.")
                fmsgbox ("*Amount u/s 10(10C) can not exceed Rs. 5 lakhs.")
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If
        
        
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(10CC)-Tax paid by employer on non-monetary perquisite" Then
                
            If Sheet1.Range(Target).Value > Sheet1.Range("IncD.Perquisites").Value Then
'            Amount u/s 10(10CC) can not exceed Value of perquisites as per sec 17(2)
'                fmsgbox ("*For Sec 10(10CC)-Tax paid by employer on non-monetary perquisite can not be more than Value of perquisites as per section 17(2).")
                fmsgbox ("*Amount u/s 10(10CC) can not exceed Value of perquisites as per sec 17(2).")
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(14)(i)-Prescribed Allowances or benefits (not in a nature of perquisite) specifically granted to meet expenses wholly, necessarily and exclusively and to the extent actually incurred, in performance of duties of office or employment" Then
                
            If Sheet1.Range(Target).Value > Sheet1.Range("IncD.Allowances").Value Then
'                fmsgbox ("*For Sec 10(14)(i) Prescribed Allowances or benefits (not in a nature of perquisite) specifically granted to meet expenses wholly, necessarily and exclusively and to the extent actually incurred, in performance of duties of office or employment can not be more than Salary as per section 17(1).")
                fmsgbox ("*Amount u/s 10(14)(i) can not exceed Salary as per sec 17(1).")
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If
        
        If (Sheet1.Range(Replace(Target, "Z", "J")).Value) = "Sec 10(14)(ii)-Prescribed Allowances or benefits granted to meet personal expenses in performance of duties of office or employment or to compensate him for increased cost of living" Then
                
            If Sheet1.Range(Target).Value > Sheet1.Range("IncD.Allowances").Value Then
                'Amount u/s 10(14)(ii) can not exceed Salary as per sec 17(1)
'                fmsgbox ("*For Sec 10(14)(ii)  Prescribed Allowances or benefits granted to meet personal expenses in performance of duties of office or employment or to compensate him for increased cost of living can not be more than Salary as per section 17(1)")
                fmsgbox ("*Amount u/s 10(14)(ii) can not exceed Salary as per sec 17(1)")
                Sheet1.Range(Target).Value = ""
                Sheet1.Range(Target).Select
            End If
        End If

End Sub


Sub LockUnlockTRP_Details()
On Error Resume Next
Application.EnableEvents = False
    Sheet3.Unprotect Password:=getmsgstate
    If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
    
        Sheet3.Range("Sheet2.IdentificationNoOfTRP").ClearContents
        Sheet3.Range("Sheet2.IdentificationNoOfTRP").Locked = True
        Sheet3.Range("Sheet2.IdentificationNoOfTRP").Interior.Color = "&HD8D8D8"
        Sheet3.Range("Sheet2.IdentificationNoOfTRP").Value = ""

        Sheet3.Range("Sheet2.NameOfTRP").ClearContents
        Sheet3.Range("Sheet2.NameOfTRP").Locked = True
        Sheet3.Range("Sheet2.NameOfTRP").Interior.Color = "&HD8D8D8"
        Sheet3.Range("Sheet2.NameOfTRP").Value = ""

        Sheet3.Range("Sheet2.ReImbFrmGov").MergeArea.ClearContents
        Sheet3.Range("Sheet2.ReImbFrmGov").MergeArea.Locked = True
        Sheet3.Range("Sheet2.ReImbFrmGov").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet3.Range("Sheet2.ReImbFrmGov").Value = ""
    Else
        Sheet3.Range("Sheet2.IdentificationNoOfTRP").Locked = False
        Sheet3.Range("Sheet2.IdentificationNoOfTRP").Interior.Color = "&HCCFFCC"

        Sheet3.Range("Sheet2.NameOfTRP").Locked = False
        Sheet3.Range("Sheet2.NameOfTRP").Interior.Color = "&HCCFFCC"

        Sheet3.Range("Sheet2.ReImbFrmGov").MergeArea.Locked = False
        Sheet3.Range("Sheet2.ReImbFrmGov").MergeArea.Interior.Color = "&HCCFFCC"
    End If



    Application.EnableEvents = True
    Sheet3.Protect Password:=getmsgstate
End Sub
Sub Table12_139_23_24Dropdown(RowIndex As Long)
On Error GoTo endline
Application.EnableEvents = False
Sheet201.Unprotect Password:=getmsgstate



'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = "2024-25" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).Value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2024", "") <= Dformat(Sheet3.Range("Ver.Date").Value, "")) And (Dformat("31/03/2025", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2023", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If
'
'
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = "2025-26" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).Value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
''        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("01/01/2025", "") <= Dformat(Sheet3.Range("Ver.Date").Value, "")) And (Dformat("31/03/202", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AY23Return"
'        ElseIf (Dformat("31/12/2024", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        End If
'    End If
'End If

'Ankita_04/11/2025===========
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = "2025-26" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).Value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("31/12/2025", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        ElseIf (Dformat("01/01/2026", "") < Dformat(Sheet3.Range("Ver.Date").Value, "")) Or (Dformat("31/03/2026", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
'        End If
'    End If
'End If
'
'If Sheet201.Cells(RowIndex, Sheet201.Range("U_UnabsorbedDepreciationYear").Column).Value = "2026-27" Then
'    If Sheet201.Cells(RowIndex, Sheet201.Range("U_RevisedReturnFile").Column).Value = "No" Then
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.ClearContents
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = True
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HD8D8D8"
'    Else
'
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Locked = False
'        Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).MergeArea.Interior.Color = "&HCCFFCC"
'        If (Dformat("31/12/2026", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=BeforeDateReturn"
'        ElseIf (Dformat("01/01/2027", "") < Dformat(Sheet3.Range("Ver.Date").Value, "")) Or (Dformat("31/03/2027", "") >= Dformat(Sheet3.Range("Ver.Date").Value, "")) Then
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.Delete
'            Sheet201.Cells(RowIndex, Sheet201.Range("U_UpdatedReturnFile").Column).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=AfterDateReturn"
'        End If
'    End If
'End If

endline:
Sheet201.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Sub

Sub PartBATI_IT1DateValidation(ByVal Target As Range)
Sheet202.Unprotect Password:=getmsgstate
Application.EnableEvents = False
If Not Application.Intersect(Sheet202.Range("U_DateDep1"), Sheet202.Range(Target.Address)) Is Nothing Then
            Dim temp
            temp = Dformat(Sheet202.Range(Target.Address).Value, "")
          '  If Not ChkMinInclusiveDate(temp, "2024-01-01") Then
            'Ankita_UR
'            If Not ChkMinInclusiveDate(temp, "2025-01-01") Then
  '          fmsgboxsmall ("Date of Deposit in Sheet IT can not be prior to 01/01/2024")
'            fmsgboxsmall ("""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/01/2025 in schedule IT.""") & Chr(13)
             
             'Ankita_03/11/2025==================
            If Not ChkMinInclusiveDate(temp, "2026-01-01") Then
                fmsgboxsmall ("""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/01/2026 in schedule IT.""") & Chr(13)
                Sheet202.Range(Target.Address).Value = ""
            End If
End If



Application.EnableEvents = True
Sheet202.Protect Password:=getmsgstate

End Sub
Sub PartBATI_IT2DateValidation(ByVal Target As Range)
Sheet202.Unprotect Password:=getmsgstate
Application.EnableEvents = False
If Not Application.Intersect(Sheet202.Range("U_DateDep2"), Sheet202.Range(Target.Address)) Is Nothing Then

                    Dim temp1
                temp1 = Dformat(Sheet202.Range(Target.Address).Value, "")
              '  If Not ChkMinInclusiveDate(temp1, "2022-04-01") Then
                  'Ankita_UR
'                   If Not ChkMinInclusiveDate(temp1, "2023-04-01") Then
                  ' fmsgboxsmall ("Date of Deposit in Sheet IT can not be prior to 01/04/2022")
'                   fmsgboxsmall ("""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/04/2023 in schedule IT.""") & Chr(13)
               
               'Ankita_03/11/2025==========================
                If Not ChkMinInclusiveDate(temp1, "2024-04-01") Then
                    fmsgboxsmall ("""Please enter valid date in dd/mm/yyyy format in schedule IT. Also date of credit to Central government cannot be prior to 01/04/2024 in schedule IT.""") & Chr(13)
                    Sheet202.Range(Target.Address).Value = ""
                End If

 Application.EnableEvents = True
Sheet202.Protect Password:=getmsgstate

End If
End Sub
'Change28.11.2022.102.01ITB
Public Sub fmsgboxsmall(IntMsg2 As Variant)
With MessageBox
     .Height = 255
     .LMessagebox.Height = 174
     .CommandButton1.Top = 191
     .LMessagebox = IntMsg2
     .LMessagebox.TextAlign = fmTextAlignLeft
     .Show
End With
End Sub
'---end
Function ChkDescription23_24(field As Variant) As Boolean
    Dim i As Long
    ChkDescription23_24 = True
    
    
    For i = 1 To Len(field)
    
        If ((asc(Mid(field, i, i)) = 40) Or (asc(Mid(field, i, i)) = 41)) Or (asc(Mid(field, i, i)) = 44) Or (asc(Mid(field, i, i)) = 45) Or (asc(Mid(field, i, i)) = 92) Or (asc(Mid(field, i, i)) = 95) Then
        
        ElseIf ((asc(Mid(field, i, i)) > 46) And (asc(Mid(field, i, i)) < 60)) Then
        
        ElseIf ((asc(Mid(field, i, i)) > 63) And (asc(Mid(field, i, i)) < 91)) Then
        
        ElseIf ((asc(Mid(field, i, i)) > 96) And (asc(Mid(field, i, i)) < 123)) Then
    
        Else
        
            ChkDescription23_24 = False
            Exit Function
        End If
        
    Next

    
End Function
Sub IC_LockUnlockGrossIncome(ByVal Target As Range)
    If Target.Value = "Not Applicable (eg. Family pension etc)" Then
        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("IncD.Allowances").MergeArea.Locked = True
        Sheet1.Range("IncD.Allowances").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Perquisites").MergeArea.Locked = True
        Sheet1.Range("IncD.Perquisites").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Profits").MergeArea.Locked = True
        Sheet1.Range("IncD.Profits").MergeArea.Interior.Color = "&HD8D8D8"
       'Ankita_16/01/2026======
'        Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.Locked = True
'        Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.Locked = True       'Ankita_16/01/2026======
'        Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.Locked = True 'Ankita_16/01/2026======
'        Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.Interior.Color = "&HD8D8D8"
'        Sheet1.Range("IncomeNotifiedOther89A").MergeArea.Locked = True 'Ankita_16/01/2026======
'        Sheet1.Range("IncomeNotifiedOther89A").MergeArea.Interior.Color = "&HD8D8D8"

        Sheet1.Range("Others.NOI_1").Locked = True
        Sheet1.Range("Others.NOI_1").Interior.Color = "&HD8D8D8"
'        Sheet1.Range("Nature_Others_1").Locked = True  'Ankita_07/03/2026
'        Sheet1.Range("Nature_Others_1").Interior.Color = "&HD8D8D8"
        Sheet1.Range("Others.Amount_1").Locked = True
        Sheet1.Range("Others.Amount_1").Interior.Color = "&HD8D8D8"
'        Sheet1.Range("Increliefus89A").MergeArea.Locked = True 'Ankita_16/01/2026=====
'        Sheet1.Range("Increliefus89A").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Deduction16").MergeArea.Locked = True
        Sheet1.Range("IncD.Deduction16").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.Deduction16ic").MergeArea.Locked = True
        Sheet1.Range("IncD.Deduction16ic").MergeArea.Interior.Color = "&HD8D8D8"

        Sheet1.Protect Password:=getmsgstate
    Else
        Sheet1.Unprotect Password:=getmsgstate

        Sheet1.Range("IncD.Allowances").MergeArea.Locked = False
        Sheet1.Range("IncD.Allowances").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IncD.Perquisites").MergeArea.Locked = False
        Sheet1.Range("IncD.Perquisites").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IncD.Profits").MergeArea.Locked = False
        Sheet1.Range("IncD.Profits").MergeArea.Interior.Color = "&HCCFFCC"

        'Ankita_16/01/2026======

'        Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.Locked = False
'        Sheet1.Range("IncomeNotified89A_AmountUS").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.Locked = False       'Ankita_16/01/2026======
'        Sheet1.Range("IncomeNotified89A_AmountUK").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.Locked = False 'Ankita_16/01/2026======
'        Sheet1.Range("IncomeNotified89A_AmountCan").MergeArea.Interior.Color = "&HCCFFCC"
'        Sheet1.Range("IncomeNotifiedOther89A").MergeArea.Locked = False  'Ankita_16/01/2026======
'        Sheet1.Range("IncomeNotifiedOther89A").MergeArea.Interior.Color = "&HCCFFCC"

        Sheet1.Range("Others.NOI_1").Locked = False
        Sheet1.Range("Others.NOI_1").Interior.Color = "&HCCFFCC"
'        Sheet1.Range("Nature_Others_1").Locked = False
        'Sheet1.Range("Nature_Others_1").Interior.Color = "&HD8D8D8"
        Sheet1.Range("Others.Amount_1").Locked = False
        Sheet1.Range("Others.Amount_1").Interior.Color = "&HCCFFCC"
'        Sheet1.Range("Increliefus89A").MergeArea.Locked = False  'Ankita_16/01/2026=====
'        Sheet1.Range("Increliefus89A").MergeArea.Interior.Color = "&HCCFFCC"

     If Sheet5.Range("BacValue").Value <> 1 Then    '03-02-2026  Malli SIT-109371
        Sheet1.Range("IncD.Deduction16").MergeArea.Locked = False
        Sheet1.Range("IncD.Deduction16").MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range("IncD.Deduction16ic").MergeArea.Locked = False
        Sheet1.Range("IncD.Deduction16ic").MergeArea.Interior.Color = "&HCCFFCC"
        
       End If
        
        Sheet1.Protect Password:=getmsgstate

    End If
End Sub
Function isAlphanumeric(field As Variant) As Boolean
    Dim i As Long
    isAlphanumeric = True
    
    For i = 1 To Len(field)
        If ((asc(Mid(field, i, i)) > 47) And (asc(Mid(field, i, i)) < 58)) Or ((asc(Mid(field, i, i)) > 64) And (asc(Mid(field, i, i)) < 91)) Or ((asc(Mid(field, i, i)) > 96) And (asc(Mid(field, i, i)) < 123)) Then
        Else
            isAlphanumeric = False
            Exit Function
        End If
    Next

End Function

Function ValidatePAN_ARN_80GD() As Boolean
ValidatePAN_ARN_80GD = True
'if pan and arn are same, then show error
Dim i, j As Long
For i = 1 To Range("Per5080G.ArnNbr").count
    'If Range("Per5080G.DoneePAN").Cells(i, 1).Value <> "" And Range("Per5080G.ArnNbr").Cells(i, 1).Value <> "" Then
    If Range("Per5080G.DoneePAN").Cells(i, 1).Value <> "" Then
    For j = i + 1 To Range("Per5080G.ArnNbr").count
        If Range("Per5080G.DoneePAN").Cells(i, 1).Value = Range("Per5080G.DoneePAN").Cells(j, 1).Value And _
            Range("Per5080G.ArnNbr").Cells(i, 1).Value = Range("Per5080G.ArnNbr").Cells(j, 1).Value Then
'            Change.28.02.2023.102.IDS49
            'MsgBox_80GD = MsgBox_80GD + "*Donation to same donee cannot be entered more than once, please enter different ARN, if applicable."
            MsgBox_80GD = MsgBox_80GD + "*Donation to same donee cannot be entered more than once, kindly enter gross amount of donation or enter different ARN, if applicable."
'            End ChangeIDS
            ValidatePAN_ARN_80GD = False
            Exit Function
        End If
    Next
    End If
Next
End Function
Public Function ChkMaxDate(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 & PAG_C10 AY 2024-25 Change
        If Year > 2025 Then      'Year changed by Ankita on 13/12/2024
            ChkMaxDate = False
            Exit Function
        Else
            If Year = 2025 Then   'Year changed by Ankita on 13/12/2024
                If month > 4 Then
                    ChkMaxDate = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                            ChkMaxDate = False
                            Exit Function
                        Else
                            If dat = 1 Then
                                ChkMaxDate = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function


Sub AnyOtherDeductions80CCH(ByVal Target As Range)
'    If (Sheet1.Range("sheet1.DOB").Value) = "" Then
'    GoTo Endline
'    End If
   ' Dim age
    'age = calculateAge23_24(Sheet1.Range("sheet1.DOB").Value)
'Ankita
'   Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    
    'INC_E30
    'If (Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government") And (age >= 17 And age <= 23) Then
                  
          'Ankita_25-26
'      If Range("sheet1.EmployerCategory1").Value = "Central Government" And (Dformat((Sheet1.Range("sheet1.DOB").Value), "") >= "2001-03-31" And Dformat((Sheet1.Range("sheet1.DOB").Value), "") <= "2007-04-01") Then
'       If Range("sheet1.EmployerCategory1").Value = "Central Government" And (Dformat((Sheet1.Range("sheet1.DOB").Value), "") >= "2001-03-31" And Dformat((Sheet1.Range("sheet1.DOB").Value), "") <= "2008-04-01") Then
       'Changes done according to new DE sheet v0.4_1 by Ankita on 13/03/2025
'        If Range("sheet1.EmployerCategory1").Value = "Central Government" And (Dformat((Sheet1.Range("sheet1.DOB").Value), "") >= "1998-03-31" And Dformat((Sheet1.Range("sheet1.DOB").Value), "") <= "2008-04-01") Then
        
    'Ankita_29/12/2025==============
    If Range("sheet1.EmployerCategory1").Value = "Central Government" And (Dformat((Sheet1.Range("sheet1.DOB").Value), "") >= "1999-03-31" And Dformat((Sheet1.Range("sheet1.DOB").Value), "") <= "2009-04-01") Then
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Locked = False
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Interior.Color = "&HCCFFCC"
    Else
    Sheet1.Unprotect Password:=getmsgstate
     
     'Ankita
        'Sheet1.Range("IncD.AnyOtherDeductions").Value = 0
        'Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Value = "(Select)"
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Locked = True
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Value = 0
       ' Ankita
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Locked = True
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Interior.Color = "&HD8D8D8"
    End If

''INC_C30 2024-25 Bindu
'    If (Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government") Then
'        'Sheet1.Range("IncD.AnyOther").MergeArea.Locked = False
'        'Sheet1.Range("IncD.AnyOther").MergeArea.Interior.Color = "&HCCFFCC"
'
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Locked = False
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Interior.Color = "&HCCFFCC"
'
'    Else
''        Sheet1.Range("IncD.AnyOther").MergeArea.Value = "(Select)"
''        Sheet1.Range("IncD.AnyOther").MergeArea.Locked = True
''        Sheet1.Range("IncD.AnyOther").MergeArea.Interior.Color = "&HD8D8D8"
'
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Value = 0
'        Sheet1.Unprotect Password:=getmsgstate
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Locked = True
'        Sheet1.Range("IncD.AnyOtherDeductions").MergeArea.Interior.Color = "&HD8D8D8"
'    End If

endline:
    Sheet1.Protect Password:=getmsgstate
    Application.EnableEvents = True
End Sub
'AY_2024_25 'Malli
Public Function ChkMaxDate1(dateEntered As Variant, maxDefinedDate As Variant) As Boolean
On Error Resume Next
Dim Year, month, dat As Variant

     ChkMaxDate1 = True
     If Len(dateEntered) > 0 Then

     Year = val(Mid(dateEntered, 7, 4))
     month = val(Mid(dateEntered, 4, 2))
     dat = val(Mid(dateEntered, 1, 2))

'        If Year > 2023 Then
        'PAG_C1 & PAG_C10 AY 2024-25 Change
        If Year > 2023 Then
            ChkMaxDate1 = False
            Exit Function
        Else
            If Year = 2023 Then
                If month > 4 Then
                    ChkMaxDate1 = False
                    Exit Function
                Else
                    If month = 4 Then
                        If dat > 1 Then
                            ChkMaxDate1 = False
                            Exit Function
                        Else
                            If dat = 1 Then
                                ChkMaxDate1 = False
                            End If
                        End If
                    End If
                End If
            End If
        End If
     End If
End Function
  
'Added by Ankita on 12/12/2024

Sub resetNatureexempt()
   Dim nature1
            Dim nature2
            Dim NatureRow
            Dim mIntCtrr As Variant
            Dim mIntCtr As Variant
            Dim mIntCells As Variant
            Dim rangecells As Variant
            Set NatureRow = Sheet1.Range("Others.NOI_1")
            
            
            Application.EnableEvents = False
            
            Sheet1.Unprotect Password:=getmsgstate
            
            For Each nature1 In NatureRow.Rows
                
            If Sheet1.Range("J" & nature1.row).Value <> "(Select)" Or Sheet1.Range("J" & nature1.row).Value <> "" Or Sheet1.Range("J" & nature1.row).Value <> "Any Other" Then
               If (Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "State Government") And Sheet5.Range("BacValue").Value <> 1 Then
'                If (Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "State Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Judge as defined in The Supreme Court Judges (Salaries and Conditions of Service) Act, 1958") And Sheet5.Range("BacValue").Value <> 1 Then
                   nature2 = Application.IsError(Application.VLookup(Sheet1.Range("J" & nature1.row).Value, Sheet5.Range("Dropdown115BAC_OTHY_CG_SG"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                      Sheet1.Unprotect Password:=getmsgstate
                      nature1.Value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "J", "Z")).Value = ""
                    End If
                    
                    ElseIf (Sheet1.Range("sheet1.EmployerCategory1").Value = "Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "State Government") And Sheet5.Range("BacValue").Value = 1 Then
                   nature2 = Application.IsError(Application.VLookup(Sheet1.Range("J" & nature1.row).Value, Sheet5.Range("Dropdown_115BAC_Y_CG_SG"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                      Sheet1.Unprotect Password:=getmsgstate
                      nature1.Value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "J", "Z")).Value = ""
                    End If
                    
                    ElseIf (Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - State Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Public sector undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Others") And Sheet5.Range("BacValue").Value <> 1 Then
                        nature2 = Application.IsError(Application.VLookup(Sheet1.Range("J" & nature1.row).Value, Sheet5.Range("Dropdown115BAC_OTHY_CGSG_PCGSG"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                      Sheet1.Unprotect Password:=getmsgstate
                      nature1.Value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "J", "Z")).Value = ""
                    End If
                   
                   ElseIf (Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Central Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - State Government" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Public sector undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Pensioners - Others") And Sheet5.Range("BacValue").Value = 1 Then
                        nature2 = Application.IsError(Application.VLookup(Sheet1.Range("J" & nature1.row).Value, Sheet5.Range("Dropdown_115BAC_Y_CGSG_PCGSG"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                      Sheet1.Unprotect Password:=getmsgstate
                      nature1.Value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "J", "Z")).Value = ""
                    End If
                    
                    ElseIf (Sheet1.Range("sheet1.EmployerCategory1").Value = "Public Sector Undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Others") And Sheet5.Range("BacValue").Value <> 1 Then
                        nature2 = Application.IsError(Application.VLookup(Sheet1.Range("J" & nature1.row).Value, Sheet5.Range("Dropdown_115BAC_OTHY"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                      Sheet1.Unprotect Password:=getmsgstate
                      nature1.Value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "J", "Z")).Value = ""
                    End If
                    
                    ElseIf (Sheet1.Range("sheet1.EmployerCategory1").Value = "Public Sector Undertaking" Or Sheet1.Range("sheet1.EmployerCategory1").Value = "Others") And Sheet5.Range("BacValue").Value = 1 Then
                        nature2 = Application.IsError(Application.VLookup(Sheet1.Range("J" & nature1.row).Value, Sheet5.Range("Dropdown_115BAC_Y"), 1, False))
                   Debug.Print nature2
                   If nature2 = True Then
                      Sheet1.Unprotect Password:=getmsgstate
                      nature1.Value = "(Select)"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range(Replace(nature1.Address, "J", "Z")).Value = ""
                    End If
                    
                    End If
                    End If
                    Nature_Amount23_24_2 ("Z" & nature1.row)
'                    Sheet1.Protect Password:=getmsgstate
                    Next
                    
        'SIT- 81028 added by Chetan C M
                   
                    mIntCells = Range("Others.NOI_1").Rows.count
                    Set rangecells = Range("Others.NOI_1").Cells
                    Dim countrycd As Variant
                    For mIntCtr = 1 To mIntCells
                         mIntCtrr = mIntCtr + 1
                            If (rangecells.item(mIntCtr, 1).Value = "" Or rangecells.item(mIntCtr, 1).Value = "(Select)") Then
                                                 
                                 If Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "J", "Z")) < Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z")) Then
                                    Sheet1.Unprotect Password:=getmsgstate
                                    rangecells.item(mIntCtr, 1).Value = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "J"))
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "J", "Z")).Value = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z"))
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "J", "Z")) = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z"))
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "J")) = "(Select)"
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z")) = ""
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z")) = ""
                                    Sheet1.Unprotect Password:=getmsgstate
                            End If
                            End If
                    Next mIntCtr
                    
                    For mIntCtr = 1 To mIntCells
                         mIntCtrr = mIntCtr + 1
                            If (rangecells.item(mIntCtr, 1).Value = "" Or rangecells.item(mIntCtr, 1).Value = "(Select)") Then
                                                 
                                 If Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "J", "Z")) < Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z")) Then
                                    Sheet1.Unprotect Password:=getmsgstate
                                    rangecells.item(mIntCtr, 1).Value = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "J"))
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "J", "Z")).Value = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z"))
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtr, 1).Address, "J", "Z")) = Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z"))
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "J")) = "(Select)"
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z")) = ""
                                    Sheet1.Unprotect Password:=getmsgstate
                                    Sheet1.Range(Replace(rangecells.item(mIntCtrr, 1).Address, "J", "Z")) = ""
                                 
                            End If
                            End If
                    Next mIntCtr
                    
        'end Chetan C M

End Sub

 'Ankita_09/01/2026=====================

' === Formats a date as: 31st July, 2026 ===
Private Function FormatDateWithOrdinal1(ByVal dt As Date) As String
    Dim d As Integer, suffix As String
    d = day(dt)
    If d Mod 100 >= 11 And d Mod 100 <= 13 Then
        suffix = "th"
    Else
        Select Case d Mod 10
            Case 1: suffix = "st"
            Case 2: suffix = "nd"
            Case 3: suffix = "rd"
            Case Else: suffix = "th"
        End Select
    End If
    FormatDateWithOrdinal1 = CStr(d) & suffix & " " & Format(dt, "mmmm, yyyy")
End Function



Sub Filesection_autopopulated()
Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False


Dim due1, due2 As String
due1 = "2025-09-15"
due2 = "2025-09-16"

Dim sheet1_ReturnFileSec_duplicate As Variant
sheet1_ReturnFileSec_duplicate = Sheet1.Range("sheet1.ReturnFileSec").Value

If Dformat(Sheet3.Range("Ver.Date"), "") <= due1 Then
    Sheet1.Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date"
    

'    If Sheet1.Range("sheet1.ReturnFileSec").Value <> "139(5)-Revised" Then
'        ElseIf Dformat(Sheet3.Range("Ver.Date"), "") > due1 Then
'        Sheet1.Range("sheet1.ReturnFileSec").Value = sheet1_ReturnFileSec_duplicate
'    End If

Else
    If Dformat(Sheet3.Range("Ver.Date"), "") >= due2 Or Sheet1.Range("sheet1.ReturnFileSec").Value <> "139(5)-Revised" Then
        Sheet1.Range("sheet1.ReturnFileSec").Value = "139(4)-Belated"
        
            ElseIf Dformat(Sheet3.Range("Ver.Date"), "") >= due2 Then
            Sheet1.Range("sheet1.ReturnFileSec").Value = sheet1_ReturnFileSec_duplicate
    
    End If
End If


Application.EnableEvents = True
'Sheet1.Protect Password:=getmsgstate
End Sub

'Ankita_18/03/2026========
Sub DPMAddComment45(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet1.Unprotect Password:=getmsgstate
    TargtVal = Target.Value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
With Sheet1.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption is allowed to the extent such income does not exceed one thousand five hundred rupees in respect of each minor child whose income is so includible"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet1.Protect Password:=getmsgstate
End Sub


Sub DPMAddComment46(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet1.Unprotect Password:=getmsgstate

    TargtVal = Target.Value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
With Sheet1.Range(targetadd)
.AddComment
With .Comment
.text text:="The exemption under 10(26) is available only to certain categories of tax payers in NER and Ladakh"""
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet1.Protect Password:=getmsgstate
End Sub

Sub DPMAddComment47(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet1.Unprotect Password:=getmsgstate
    TargtVal = Target.Value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
With Sheet1.Range(targetadd)
.AddComment
With .Comment
.text text:="The exemption under 10(26AAA) is available only to certain categories of Sikkimese tax payers"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet1.Protect Password:=getmsgstate
End Sub

Sub DPMAddComment48(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet1.Unprotect Password:=getmsgstate

    TargtVal = Target.Value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
With Sheet1.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed sixty per cent of the total amount payable  at the time of such closure or opting out of the scheme"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet1.Protect Password:=getmsgstate
End Sub
Sub DPMAddComment49(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet1.Unprotect Password:=getmsgstate

    TargtVal = Target.Value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
With Sheet1.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent that it does not exceed sixty per cent. of the individual corpus"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet1.Protect Password:=getmsgstate
End Sub
'Ankita_18/03/2026=====
Sub DPMAddComment50(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet1.Unprotect Password:=getmsgstate

    TargtVal = Target.Value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
With Sheet1.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed twenty-five per cent of the amount of contributions made by him"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet1.Protect Password:=getmsgstate
End Sub

Sub DPMAddComment51(Target As Range)
On Error Resume Next
Dim TargtVal, targetadd As Variant
sPassword = getmsgstate
Sheet1.Unprotect Password:=getmsgstate

    TargtVal = Target.Value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
With Sheet1.Range(targetadd)
.AddComment
With .Comment
.text text:="Exemption allowed to the extent it does not exceed twenty-five per cent of the amount of contributions made by him"
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet1.Protect Password:=getmsgstate
End Sub

Sub DPMADeleteComment45(Target As Range)
Dim TargtVal, targetadd As Variant
On Error Resume Next
sPassword = getmsgstate
Sheet1.Unprotect Password:=getmsgstate
    TargtVal = Target.Value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
Sheet1.Protect Password:=getmsgstate
End Sub




