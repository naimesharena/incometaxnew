Attribute VB_Name = "Sheet3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Dim fn As String
Dim mn As String
Dim ln As String
Public currCnt1, currCnt2, currCnt3 As Variant

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
    Dim validateType As String
    Dim TargetPAN As String
    Dim TargetIFSC As String
    Dim rangenm As Variant
    Dim TargtVal, targetadd As Variant
    Dim Value_24, Value_21, Value_22, Value_23 As Double
    Dim IffCode As Range
    Dim IfsCode As Range

TargtVal = Target.Value
targetadd = Target.Address
targetadd = Replace(targetadd, "$", "")

    TargetPAN = ""
    TargetIFSC = ""
    validateType = ""

   Application.ScreenUpdating = False
  Application.EnableEvents = False

sPassword = EfilingCommon.getmsgstate
Sheet1.Unprotect Password:=sPassword
Sheet1.Range("taxcheck").Value = ""
Sheet1.Protect Password:=sPassword

 If Target.Validation.Type = 6 Then
        Target.Value = UCase(Target.Value)

            Set IfsCode = Range("SchBA.IFSC")
'            If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
'
'                Call CheckIFSC1(TargtVal, targetadd)
'                GoTo Endline
'            End If
'Ankita_18/12/2024

    Dim sourceAddress, Address, Address1 As Variant
    Dim rangecell As Range
    Dim BankNameFromIFSC As Variant
    
    If Not Application.Intersect(IfsCode, Range(Target.Address)) Is Nothing Then
         For Each rangecell In Sheet3.Range("SchBA.IFSC")
            rangecell.Value = UCase(rangecell.Value)
     
     If rangecell.Value <> "" Then
           
           If Not CheckIFSC(UCase(TargtVal)) Then
                    'fmsgboxsmall "* Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                     fmsgbox ("* ""Invalid IFS Code.IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet3.Range(targetadd) = ""
                    Sheet3.Range(targetadd).Offset(0, 1) = ""
                    GoTo endline
           
            ElseIf Not ValidateIFSCList(UCase(TargtVal)) Then
                   ' fmsgboxsmall "* Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets."
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                    Sheet3.Range(targetadd) = ""
                    Sheet3.Range(targetadd).Offset(0, 1) = ""
                    GoTo endline
     
                  'Call CheckIFSC1(TargtVal, targetadd)
                  'GoTo endline
            Else:
                       sourceAddress = rangecell.Address
                       Address = Replace(sourceAddress, "F", "G")
                
                     If Mid(rangecell, 1, 4) = "GIFT" Then
                             BankNameFromIFSC = GetBankName(rangecell)
                     Else
                             BankNameFromIFSC = GetBankName(Mid(rangecell, 1, 4))
                     End If
                
                
                If Mid(rangecell, 1, 4) = "GIFT" And BankNameFromIFSC <> "" Then
                   Sheet3.Range(Address).Value = BankNameFromIFSC
                ElseIf Mid(rangecell, 1, 4) <> "GIFT" Then
                   Sheet3.Range(Address).Value = BankNameFromIFSC
                End If
                
                If Mid(rangecell, 1, 4) <> "GIFT" Then
                    If (Sheet3.Range(Address).Value = "") Then
                     'MsgBox "IFS Code: Please enter valid IFS Code. Refer to your bank for valid codes. (1st 4 Alpha, followed by 0 and 6 digits)"
                     fmsgbox ("* ""Invalid IFS Code. IFS Code should be exactly 11 characters, First 4 characters should be alphabets, 5th character must be zero (0) and remaining 6 should be either numeric or alphabets.""" & Chr(13))
                     Address = Replace(sourceAddress, "K", "J")
                     Sheet3.Range(Address).Value = ""
                    GoTo endline
                    End If
                End If
                
            End If
         End If
    Next
    GoTo endline
End If


'-----------------
            If Not Intersect(Range("SchBA.BankName"), Range(Target.Address)) Is Nothing Then
            If Not checkfieldSuperSpecialcharactername(Range(Target.Address).Value) Then
                fmsgbox ("Name of Bank should not contain Special characters < , >" & ",&" & Chr(13))
               Range(Target.Address).Value = ""
               
            End If
            End If

          If (Target.name.name) = "Ver.AssesseeVerName" Then
                If Len(Sheet3.Range("Ver.AssesseeVerName")) > 0 Then
                 If Not checkfieldSuperSpecialcharactername(Sheet3.Range("Ver.AssesseeVerName").Value) Then
                    fmsgbox ("Assessee name should not Contain <, > " & Chr(34) & "& characters" & Chr(13))
                    Sheet3.Range("Ver.AssesseeVerName").Value = ""
                  
                Else
                 Sheet3.Range("Ver.AssesseeVerName").Value = (Sheet3.Range("Ver.AssesseeVerName").Value)
                End If
                End If


            End If

             If (Target.name.name) = "Ver.Place" Then
                If Len(Sheet3.Range("Ver.Place")) > 0 Then
                 If Not checkfieldSuperSpecialcharactername(Sheet3.Range("Ver.Place").Value) Then
                    fmsgbox ("Place name should not Contain <, > ,& characters" & Chr(13))
                    Sheet3.Range("Ver.Place").Value = ""
                Else
                 Sheet3.Range("Ver.Place").Value = (Sheet3.Range("Ver.Place").Value)
                End If
                End If
            End If

            If (Target.name.name) = "Sheet2.NameOfTRP" Then
                If Len(Sheet3.Range("Sheet2.NameOfTRP")) > 0 Then
                 If Not checkfieldSuperSpecialcharactername(Sheet3.Range("Sheet2.NameOfTRP").Value) Then
                    fmsgbox (" Name of TRP should not Contain <, > ,& characters" & Chr(13))
                    Sheet3.Range("Sheet2.NameOfTRP").Value = ""
                Else
                 Sheet3.Range("Sheet2.NameOfTRP").Value = (Sheet3.Range("Sheet2.NameOfTRP").Value)
                End If
                End If
            End If

          If (Target.name.name) = "Ver.FatherName" Then
                If Len(Sheet3.Range("Ver.FatherName")) > 0 Then
                 If Not checkfieldSuperSpecialcharactername(Sheet3.Range("Ver.FatherName").Value) Then
                    fmsgbox ("Father name should not  Contain <, > ,& characters" & Chr(13))
                    Sheet3.Range("Ver.FatherName").Value = ""
                Else
                 Sheet3.Range("Ver.FatherName").Value = (Sheet3.Range("Ver.FatherName").Value)
                End If
                End If
            End If



        'PAN Validation: On Data Entry
         If (Target.name.name) = "Ver.PAN" Then
              TargetPAN = Target.name.name
              If Not mIncmDtls.CheckPAN(UCase(Trim(Range(TargetPAN).Value))) Then
                 'fmsgbox "Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet. 4th character must be 'P' only."
                 'SIT-66374_change 'Malli
'                fmsgbox "Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet. 4th letter must be 'P'."
                 fmsgbox "Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet."    'Ankita_31/01/2025
                 Sheet3.Range(TargetPAN).Select
                 Sheet3.Range(TargetPAN).Value = ""
               Else
               If (Sheet1.Range("sheet1.DOB") <> "") And (calculateAge(Trim(Sheet1.Range("sheet1.DOB").Value)) < 18) And (UCase(Sheet3.Range(TargetPAN).Value) = UCase(Sheet1.Range("sheet1.PAN").Value)) Then
                fmsgbox "* Please Enter PAN of Representative"
                Sheet3.Range(TargetPAN).Value = ""
                Else
                Sheet3.Range(TargetPAN).Value = UCase(Range(TargetPAN).Value)
                End If
               End If
              
'                If calculateAge(Trim(Sheet1.Range("sheet1.DOB").Value)) < 18 Then
'                  If Sheet3.Range("Ver.PAN").Value = Sheet1.Range("Sheet1.pan").Value Then
'                    fmsgbox "*If date of birth is on or after 2 April 2002 then in verification PAN cannot be same as mentioned in general information."
'                    Sheet3.Range("Ver.PAN").Value = ""
'                  End If
'                End If
          End If



          If Target.name.name = "Ver.Date" Then
                If Trim(Range("Ver.Date").Value) <> "00/00/0000" Then
                    Range("Ver.Date").Value = Replace(Range("Ver.Date").Value, "-", "/")
                        If Not mIncmDtls.CheckDateBefore(Trim(Range("Ver.Date").Value)) Then
                            Sheet3.Range("Ver.Date").Select
                            Sheet3.Range("Ver.Date").Value = ""
                        Else
                            Range("Ver.Date").Value = Range("Ver.Date").Value
                        End If
                End If
          End If
 End If

  If (Target.name.name) = "Ver.capacity" Then
                If Len(Sheet3.Range("Ver.capacity")) > 0 Then
                 If Not checkfieldSuperSpecialcharactername(Sheet3.Range("Ver.capacity").Value) Then
                    fmsgbox ("Verification capacity should not Contain <, > ,& characters" & Chr(13))
                    Sheet3.Range("Ver.capacity").Value = ""
                Else
                    Sheet3.Range("Ver.capacity").Value = (Sheet3.Range("Ver.capacity").Value)
                End If
                End If

'                If calculateAge(Trim(Sheet1.Range("sheet1.DOB").Value)) < 18 Then
'                  If Sheet3.Range("Ver.Capacity").Value <> "Representative" Then
'                    fmsgbox "*If date of birth is on or after 2 April 2002 then in verification drop down cannot be selected as self."
'                    Sheet3.Range("Ver.Capacity").Value = ""
'                  End If
'                End If
'dpk
'                MsgBox Sheet3.Range("Ver.Capacity").Value
'                If Sheet3.Range("Ver.Capacity").Value = "Self" Then
'
'                Sheet3.Unprotect Password:=sPassword
'                Sheet3.Range("Ver.AssesseeVerName").MergeArea.Interior.Color = "&HCCFFCC"
'                Sheet3.Range("Ver.AssesseeVerName").MergeArea.Locked = False
'                Else
'
'                Sheet3.Unprotect Password:=sPassword
'                Sheet3.Range("Ver.AssesseeVerName").MergeArea.Interior.ColorIndex = 2
'                Sheet3.Range("Ver.AssesseeVerName").MergeArea.Locked = True
'                Sheet3.Protect Password:=sPassword
'                End If
                
                If Sheet3.Range("Ver.Capacity").Value = "Representative" Or Sheet3.Range("Ver.Capacity").Value = "(Select)" Then
                    Sheet3.Unprotect Password:=sPassword
                    Sheet3.Range("Ver.AssesseeVerName").MergeArea.Value = ""
                    Sheet3.Range("Ver.PAN").MergeArea.Value = ""
                    Sheet3.Range("Ver.AssesseeVerName").MergeArea.Interior.Color = "&HCCFFCC"
                    Sheet3.Range("Ver.PAN").MergeArea.Interior.Color = "&HCCFFCC"
                    Sheet3.Range("Ver.AssesseeVerName").MergeArea.Locked = False
                    Sheet3.Range("Ver.PAN").MergeArea.Locked = False
                    Sheet3.Protect Password:=sPassword
                Else
                Sheet3.Unprotect Password:=sPassword
                    Sheet3.Range("Ver.AssesseeVerName").MergeArea.Interior.Color = "&HCCFFCC"
                    Sheet3.Range("Ver.AssesseeVerName").MergeArea.Locked = False
                    Sheet3.Range("Ver.PAN").MergeArea.Interior.Color = "&HCCFFCC"
                    Sheet3.Range("Ver.PAN").MergeArea.Locked = False
                    Sheet3.Range("Ver.PAN").MergeArea.Value = Sheet1.Range("sheet1.PAN").MergeArea.Value
                    
'                    Sheet3.Protect Password:=sPassword
'                Sheet3.Unprotect Password:=sPassword
                 Sheet3.Range("Ver.AssesseeVerName").MergeArea.Value = Trim(UCase(Sheet1.Range("sheet1.FirstName").Value + " " + Sheet1.Range("sheet1.MiddleName").Value + " " + Sheet1.Range("sheet1.SurNameOrOrgName").Value))
                
                
                If (Sheet1.Range("sheet1.MiddleName").Value = "") Then
                 Sheet3.Range("Ver.AssesseeVerName").Value = Trim(UCase(Sheet1.Range("sheet1.FirstName").Value + "" + Sheet1.Range("sheet1.MiddleName").Value + " " + Sheet1.Range("sheet1.SurNameOrOrgName").Value))
                End If
                 Sheet3.Protect Password:=sPassword
                End If

            End If

endline:
 Application.EnableEvents = True
 Application.ScreenUpdating = True
 'Sheet3.Protect Password:=sPassword
End Sub

Sub protectAll()
ThisWorkbook.Protect Password:=EfilingCommon.getmsgstate
Sheet1.Protect Password:=EfilingCommon.getmsgstate
Sheet2.Protect Password:=EfilingCommon.getmsgstate
Sheet3.Protect Password:=EfilingCommon.getmsgstate
Sheet4.Protect Password:=EfilingCommon.getmsgstate
Sheet5.Protect Password:=EfilingCommon.getmsgstate
Sheet6.Protect Password:=EfilingCommon.getmsgstate
Sheet7.Protect Password:=EfilingCommon.getmsgstate
Sheet8.Protect Password:=EfilingCommon.getmsgstate
Sheet9.Protect Password:=EfilingCommon.getmsgstate
Sheet10.Protect Password:=EfilingCommon.getmsgstate
Sheet11.Protect Password:=EfilingCommon.getmsgstate
Sheet12.Protect Password:=EfilingCommon.getmsgstate
End Sub

Sub Worksheetprotect()
ThisWorkbook.Protect Password:=EfilingCommon.getmsgstate
Sheet5.Visible = False

End Sub


Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False


ActiveWindow.DisplayHeadings = False
If calculateAge(Trim(Sheet1.Range("sheet1.DOB").Value)) > 18 Then
If Sheet3.Range("Ver.PAN").Value = "" Then
    Sheet3.Range("Ver.PAN").Value = Sheet1.Range("Sheet1.pan").Value

End If
End If

If Sheet3.Range("Ver.Capacity").Value <> "Representative" Then
If Sheet3.Range("Ver.AssesseeVerName").Value = "" Then
    Sheet3.Range("Ver.AssesseeVerName").Value = Trim(UCase(Sheet1.Range("sheet1.FirstName").Value + " " + Sheet1.Range("sheet1.MiddleName").Value + " " + Sheet1.Range("sheet1.SurNameOrOrgName").Value))
    If Sheet1.Range("sheet1.MiddleName").Value = "" Then
     Sheet3.Range("Ver.AssesseeVerName").Value = Trim(UCase(Sheet1.Range("sheet1.FirstName").Value + "" + Sheet1.Range("sheet1.MiddleName").Value + " " + Sheet1.Range("sheet1.SurNameOrOrgName").Value))
    End If
End If
End If

ActiveSheet.Shapes("Button 84").Left = Range("TPV.Button1").Left + 10
ActiveSheet.Shapes("Button 84").Top = Range("TPV.Button1").Top + 10


ActiveSheet.Shapes("Button 94").Left = Range("TPV.Button2").Left
ActiveSheet.Shapes("Button 94").Top = Range("TPV.Button2").Top

ActiveSheet.Shapes("Button 131").Left = Range("Button_epaytax").Left
ActiveSheet.Shapes("Button 131").Top = Range("Button_epaytax").Top
   ' ActiveSheet.Shapes("Button 105").Visible = False

    If Sheet3.Range("IncD.BalTaxPayable").Value > 0 Then
        Sheet3.Shapes("Button 131").Visible = True
          Else
         Sheet3.Shapes("Button 131").Visible = False
    End If
'End If
endline:
Application.EnableEvents = True
Sheet3.Protect Password:=EfilingCommon.getmsgstate
End Sub



'Ankita--------17/12/2024----
Function GetBankName(IFSC As Variant) As String
On Error Resume Next
GetBankName = ""
Dim rangecells As Range
Dim Address As Variant
For Each rangecells In Sheet10.Range("IFSC_StartLetter")

If UCase(rangecells.Value) = UCase(IFSC) Then
Address = rangecells.Address
Address = Replace(Address, "A", "B")
GetBankName = Sheet10.Range(Address).Value
Exit Function
End If

Next

'If GetBankName = "" Then
'GetBankName = False
'End If

End Function


