Attribute VB_Name = "Sheet16"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'
''Function GetBankName(IFSC As Variant) As String
''On Error Resume Next
''GetBankName = ""
''Dim rangecells As Range
''Dim Address As Variant
'''For Each rangecells In Sheet8.Range("IFSC_StartLetter")
''For Each rangecells In Sheet10.Range("IFSC_StartLetter")
''
''If UCase(rangecells.Value) = UCase(IFSC) Then
''Address = rangecells.Address
''Address = Replace(Address, "A", "B")
'''GetBankName = Sheet8.Range(Address).Value
''GetBankName = Sheet10.Range(Address).Value
''Exit Function
''End If
''
''Next
''
'''If GetBankName = "" Then
'''GetBankName = False
'''End If
''
''End Function
''
''Sub LockUnlockIFSC()
''On Error Resume Next
''Application.EnableEvents = False
''Sheet16.Unprotect Password:=getmsgstate
''    If Mid(Sheet16.Range("LoanfrmBankOrInstitute.24b").value, 1, 1) = "O" Then
''    Sheet16.Unprotect Password:=getmsgstate
''        Sheet16.Range("IFSC.24b").ClearContents
''       ' Sheet16.Range("IFSC.24b").Locked = True
''        Sheet16.Range("IFSC.24b").Interior.Color = "&HD8D8D8"
''        Sheet16.Range("IFSC.24b").value = ""
''        Sheet16.Range("IFSC.24b").Locked = True
''
''    Else
''    Sheet16.Unprotect Password:=getmsgstate
''        Sheet16.Range("IFSC.24b").ClearContents
''        Sheet16.Range("IFSC.24b").Locked = False
''        Sheet16.Range("IFSC.24b").Interior.Color = "&HCCFFCC"
''
''    End If
''Sheet16.Protect Password:=getmsgstate
''Application.EnableEvents = True
''End Sub
'
''Ankita_19/04/2025
'
'Function ValidateDate_24b(ByVal Target As Range) As Boolean
'    Dim rangecells As Range
'    Dim i As Long
'    Dim DateDep_IT
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'
'
'    mIntCells = Sheet16.Range("loanDate.24b").count
'    ValidateDate_24b = True
'
'    Dim end_24
'    end_24 = 0
'    Set rangecells = Sheet16.Range("loanDate.24b").Cells
'
'    For mIntCtr = 1 To mIntCells
'        If Not rangecells.item(mIntCtr).Value = "" Then
'            end_24 = end_24 + 1
'        End If
'    Next
'
'    ReDim Date_24b(end_24)
'
'    For i = 1 To end_24
'        Date_24b(i) = rangecells.item(i).Value
'        If Not chkCompulsory(Date_24b(i)) Then
'        'Ankita_05/05/2025_Commented as per DESheet_v0.7
'            'fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is mandatory") '
'            fmsgboxsmall ("* ""Please provide Date of sanction of loan in schedule 24(b)"" at Sr. No  " & i & "")
'            ValidateDate_24b = False
'            Exit Function
'        End If
'
'        If Not CheckDateddmmyyyy(Date_24b(i)) Then
'            ValidateDate_24b = False
''            fmsgboxsmall ("Date of Deposit at Sr. No  " & i & "  in Sheet IT  is invalid. Pl enter in dd/mm/yyyy format")
''Ankita_10/02
''              fmsgboxsmall ("Please enter a valid date in dd/mm/yyyy format at Sr. No  " & i & "  in Sheet IT ")
''                fmsgboxsmall ("Date of taking loan - Please enter valid date in valid format at Sr. No  " & i & "")
'               fmsgboxsmall ("* ""Please enter date in valid format"" at Sr. No  " & i & "")
'            Target.Value = ""
'            Exit Function
'        Else
'            Date_24b(i) = Dformat(Date_24b(i), "")
'
'          'Ankita_23/12/2025=================
'            Dim cutoff4 As Date, d As Date
'            cutoff4 = DateValue(CDate(Sheet5.Range("Date_24b").Cells(1, 1).Value))
'            If Not ChkMaxDate_24b(Trim(Date_24b(i)), (Sheet5.Range("DOB_1").Value)) Then
'                 fmsgboxsmall ("* Date can not be on or after " & Format(cutoff4, "dd/mm/yyyy") & " at Sr. No " & i & ".") & Chr$(13)
'                Target.Value = ""
'                ValidateDate_24b = False
'            Exit Function
'            End If
'          '================================
'        End If
'    Next
'
'End Function
'
'
'Private Sub Worksheet_Activate()
'Sheet16.Activate
'Sheet16.Protect Password:=getmsgstate
'
'End Sub
'
'Private Sub Worksheet_Change(ByVal Target As Range)
'On Error GoTo endline
'Dim TargtVal, targetadd As Variant
'Dim IffCode As Range
'Dim IfsCode As Range
'
'
'Application.EnableEvents = False
'Application.ScreenUpdating = False
'
'Sheet1.Range("taxcheck").Value = "Y"
'Sheet16.Unprotect Password:=getmsgstate
'
''Ankita_18/04
'
''Ankita_12/05/2025
'TargtVal = Target.Value
'    targetadd = Target.Address
'    targetadd = Replace(targetadd, "$", "")
'
'
'If Not Application.Intersect(Sheet16.Range("loanDate.24b"), Sheet16.Range(Target.Address)) Is Nothing Then
'
'        If Not (ValidateDate_24b(Target)) Then
'
'         End If
'           GoTo endline
'End If
'
''Commented by Ankita_12/05/2025 for 0.7 changes
'
''Ankita---------------16/04/2025
'
''If Not Application.Intersect(Sheet16.Range("LoanfrmBankOrInstitute.24b"), Sheet16.Range(Target.Address)) Is Nothing Then
''
''If Sheet16.Range(Target.Address).Value <> "" Or Sheet11.Range(Target.Address).Value <> "(Select)" Then
''
''            If Mid(Sheet16.Range(Target.Address).Value, 1, 1) = "O" Then
''
''            Sheet16.Unprotect Password:=getmsgstate
''
''            Sheet16.Range(Replace(Target.Address, "D", "E")).ClearContents
''
''            Sheet16.Range(Replace(Target.Address, "D", "E")).Locked = True
''
''            Sheet16.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HD8D8D8"
''
''            Sheet16.Unprotect Password:=getmsgstate
''
''            Sheet16.Range(Replace(Target.Address, "D", "G")).ClearContents
''
''            'Sheet16.Range(Replace(Target.Address, "D", "G")).Locked = True
''
''            Sheet16.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
''            Sheet16.Unprotect Password:=getmsgstate
''
''            Sheet16.Range(Replace(Target.Address, "D", "F")).ClearContents
''
''            Sheet16.Range(Replace(Target.Address, "D", "F")).Locked = False
''
''            Sheet16.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
''
''
''            ElseIf Mid(Sheet16.Range(Target.Address).Value, 1, 1) = "B" Then
''
''            'Malli-----------
''            Sheet16.Unprotect Password:=getmsgstate
''            Sheet16.Range(Replace(Target.Address, "D", "E")).ClearContents
''            Sheet16.Unprotect Password:=getmsgstate
''            Sheet16.Range(Replace(Target.Address, "D", "E")).Locked = False
''            Sheet16.Unprotect Password:=getmsgstate
''            Sheet16.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
''            '----------------
''
''            Sheet16.Unprotect Password:=getmsgstate
''            Sheet16.Range(Replace(Target.Address, "D", "F")).ClearContents
''            Sheet16.Unprotect Password:=getmsgstate
''            Sheet16.Range(Replace(Target.Address, "D", "F")).Locked = True
''            Sheet16.Unprotect Password:=getmsgstate
''            Sheet16.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HD8D8D8"
''
''
''            Else
''
''            Sheet16.Unprotect Password:=getmsgstate
''
''            Sheet16.Range(Replace(Target.Address, "D", "E")).ClearContents
''
''            Sheet16.Range(Replace(Target.Address, "D", "E")).Locked = False
''
''            Sheet16.Range(Replace(Target.Address, "D", "E")).Interior.Color = "&HCCFFCC"
''
''            Sheet16.Unprotect Password:=getmsgstate
''
''            Sheet16.Range(Replace(Target.Address, "D", "G")).ClearContents
''
''            Sheet16.Range(Replace(Target.Address, "D", "G")).Locked = False
''
''            Sheet16.Range(Replace(Target.Address, "D", "G")).Interior.Color = "&HCCFFCC"
''
''
''
''            Sheet16.Unprotect Password:=getmsgstate
''
''            Sheet16.Range(Replace(Target.Address, "D", "F")).ClearContents
''
''            Sheet16.Range(Replace(Target.Address, "D", "F")).Locked = False
''
''            Sheet16.Range(Replace(Target.Address, "D", "F")).Interior.Color = "&HCCFFCC"
''
''
''            End If
''End If
''End If
'
''-----------------------
'
''Ankita_05/05/2025_Commented as per DESheet_v0.7
''If Not Application.Intersect(Sheet16.Range("PAN.24b"), Sheet16.Range(Target.Address)) Is Nothing Then
''If (Target.name.name) = "sheet1.PANRepAssessee" Then
'
''
''    Dim TargetPAN1  As String
''    TargetPAN1 = ""
''    TargetPAN1 = Sheet16.Range(Target.Address).Value
''    'Target.value = UCase(Target.value)
''    If TargetPAN1 <> "" Then
''        'If Not mdIncomeDetails.CheckPAN1(UCase(Trim(Range(TargetPAN1).value))) Then
''        If Not EfilingCommon.CheckPAN(UCase(Trim(TargetPAN1))) Then
''            'Ankita_PAN
''            'fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." & Chr(10) & Chr(13) & "4th letter must be        'P' in Sheet Income Details" ', vbOKCancel, mdCommon.vbMessageCaption
''            fmsgboxsmall "* Invalid PAN." & Chr(10) & Chr(13) & "PAN format should be First 5 Alphabets, Next 4 digits, Then 1 Alphabet." ', vbOKCancel, mdCommon.vbMessageCaption
''Sheet16.Unprotect Password:=getmsgstate
''           Sheet16.Range(Target.Address).ClearContents
''      '      TargetPAN1 = ""
'''            Sheet16.Range(TargetPAN1) = ""
'''                    Sheet16.Range(TargetPAN1).Offset(0, 1) = ""
''
''            GoTo endline
''        End If
'''       ' TargetPAN1 = UCase(TargetPAN1)'
''       Sheet16.Range(Target.Address).Value = UCase(TargetPAN1)
''
'''        If Range(TargetPAN1).value = Sheet1.Range("sheet1.PAN").value Then
'''            fmsgboxsmall "* PAN of the representative assessee And assessee (under personal information) cannot be same in Sheet Income Details"        ', vbOKCancel, mdCommon.vbMessageCaption
''
'''            Sheet16.Range(TargetPAN1).Select
'''            Sheet16.Range(TargetPAN1).value = ""
'''            GoTo endline
'''        End If
''    End If
''End If
'
''17/04/2025_Ankita
'  'Dim TargtVal, targetadd As Variant
'  'Ankita_05/05/2025_Commented as per DESheet_v0.7
''TargtVal = Target.Value
''targetadd = Target.Address
''targetadd = Replace(targetadd, "$", "")
''
''
''
''If Target.Validation.Type = 6 Then
''    Target.Value = UCase(Target.Value)
'
'
''--------17/12/2024----
''17/12/2024
' 'Ankita_05/05/2025_Commented as per DESheet_v0.7
'' Set IfsCode = Range("IFSC.24b")
''
''
''    Dim sourceAddress, Address, address1 As Variant
''    Dim rangecell As Range
''    Dim BankNameFromIFSC As Variant
''
''    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
''         For Each rangecell In Sheet16.Range("IFSC.24b")
''            rangecell.Value = UCase(rangecell.Value)
''     'TargtVal = UCase(rangecell.Value)
''     If rangecell.Value <> "" Then
''
''           If Not CheckIFSC(UCase(TargtVal)) Then
''                    'fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
''                     fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
''                    Sheet16.Range(targetadd) = ""
''                    Sheet16.Range(targetadd).Offset(0, 1) = ""
''                    GoTo endline
''
''            ElseIf Not ValidateIFSCList(UCase(TargtVal)) Then
''                   ' fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
''                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
''                    Sheet16.Range(targetadd) = ""
''                    Sheet16.Range(targetadd).Offset(0, 1) = ""
''                    GoTo endline
''
''                  'Call CheckIFSC1(TargtVal, targetadd)
''                  'GoTo endline
''            Else:
''                       sourceAddress = rangecell.Address
''                       Address = Replace(sourceAddress, "E", "G")
''
''                     If Mid(rangecell, 1, 4) = "GIFT" Then
''                             BankNameFromIFSC = Sheet3.GetBankName(rangecell)
''                     Else
''                             BankNameFromIFSC = Sheet3.GetBankName(Mid(rangecell, 1, 4))
''                     End If
''
''
''                If Mid(rangecell, 1, 4) = "GIFT" And BankNameFromIFSC <> "" Then
''                   Sheet16.Range(Address).Value = BankNameFromIFSC
''                ElseIf Mid(rangecell, 1, 4) <> "GIFT" Then
''                   Sheet16.Range(Address).Value = BankNameFromIFSC
''                End If
''
''                If Mid(rangecell, 1, 4) <> "GIFT" Then
''                    If (Sheet16.Range(Address).Value = "") Then
''                     'MsgBox "IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits)"
''                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
''                     Address = Replace(sourceAddress, "K", "J")
''                     Sheet16.Range(Address).Value = ""
''                    GoTo endline
''                    End If
''                End If
''
''            End If
''         End If
''    Next
''    GoTo endline
''End If
''End If
'endline:
''
'Application.EnableEvents = True
'Sheet16.Protect Password:=sPassword
'Application.ScreenUpdating = True
'End Sub
'
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'
'End Sub
'
