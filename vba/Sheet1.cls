Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Public currCnt1, currCnt2, currCnt3, currCnt4, currCnt5, currCnt6, currCnt7, currCnt8, currCnt9, currCnt10, currCnt11, currCnt12, currCnt13, currCnt14, currCnt15, currCnt16, currCnt17 As Variant
Public currCnt22, currCnt33, currCnt44, currCnt66, currCnt67, currCnt68, currCnt69, currCnt70 As Variant

Function StateMatchesPin(StateCode As Variant) As Boolean
    StateMatchesPin = False
    Dim state1, state2 As Variant
    state1 = Application.VLookup(Sheet1.Range("sheet1.PinCode").Value, Sheet5.Range("Duplicate_V"), 2, False)
    state2 = Application.VLookup(Sheet1.Range("sheet1.PinCode").Value, Sheet5.Range("Duplicate_V"), 3, False)
    If IsError(state1) Or IsError(state2) Then
    StateMatchesPin = False
    Exit Function
    End If
    If UCase(Mid(StateCode, 4)) = state1 Or UCase(Mid(StateCode, 4)) = state2 Then
    StateMatchesPin = True
    End If
    
End Function
'Ayush_16
Function StateMatchesPin1(StateCode2 As Variant) As Boolean
    StateMatchesPin1 = False
    Dim state1, state2 As Variant
    state1 = Application.VLookup(Sheet1.Range("sheet1.PinCode").Value, Sheet5.Range("Duplicate_V"), 2, False)
    state2 = Application.VLookup(Sheet1.Range("sheet1.PinCode").Value, Sheet5.Range("Duplicate_V"), 3, False)
    If IsError(state1) Or IsError(state2) Then
    StateMatchesPin1 = False
    Exit Function
    End If
    If UCase(Mid(StateCode2, 4)) = state1 Or UCase(Mid(StateCode2, 4)) = state2 Then
    StateMatchesPin1 = True
    End If
    
End Function
'--------------------
Sub b()
Application.EnableEvents = True
End Sub
Sub colorchange()
Sheet1.Unprotect Password:=getmsgstate
                Sheet1.Range("IncD.IntrstPayUs234C").MergeArea.Locked = True
                Sheet1.Range("IncD.IntrstPayUs234C").MergeArea.Interior.ColorIndex = 2
                Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
                Sheet1.Range("IncD.IntrstPayUs234F").MergeArea.Interior.ColorIndex = 2
Sheet1.Protect Password:=getmsgstate
End Sub

Sub LockUnlockReturnfilesec()
On Error Resume Next
Sheet1.Unprotect Password:=getmsgstate
Application.EnableEvents = False
   If (Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 11 Or Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 12) Then

        Range("sheet1.ReceiptNo").MergeArea.ClearContents
        Range("sheet1.ReceiptNo").MergeArea.Locked = True
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.OrigRetFiledDate").MergeArea.ClearContents
        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.NoticeNo").MergeArea.ClearContents
        Range("sheet1.NoticeNo").MergeArea.Locked = True
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.NoticeDate").MergeArea.ClearContents
        Range("sheet1.NoticeDate").MergeArea.Locked = True
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HD8D8D8)
        
    ElseIf (Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 13 Or Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 14 Or Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 15 Or Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 16 Or Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 20) Then
        
        Range("sheet1.ReceiptNo").MergeArea.ClearContents
        Range("sheet1.ReceiptNo").MergeArea.Locked = True
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.OrigRetFiledDate").MergeArea.ClearContents
        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = True
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.NoticeNo").MergeArea.Locked = False
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.NoticeDate").MergeArea.Locked = False
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)

    ElseIf Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 17 Or Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 21 Then

        Range("sheet1.ReceiptNo").MergeArea.Locked = False
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HCCFFCC)
        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.NoticeNo").MergeArea.ClearContents
        Range("sheet1.NoticeNo").MergeArea.Locked = True
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HD8D8D8)

        Range("sheet1.NoticeDate").MergeArea.ClearContents
        Range("sheet1.NoticeDate").MergeArea.Locked = True
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HD8D8D8)
    
    ElseIf (Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = 18) Then
        
        Range("sheet1.ReceiptNo").MergeArea.Locked = False
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.NoticeNo").MergeArea.Locked = False
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.NoticeDate").MergeArea.Locked = False
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)
        ElseIf (Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = "(S") Then
        Range("sheet1.ReceiptNo").MergeArea.Locked = False
        Range("sheet1.ReceiptNo").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.OrigRetFiledDate").MergeArea.Locked = False
        Range("sheet1.OrigRetFiledDate").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.NoticeNo").MergeArea.Locked = False
        Range("sheet1.NoticeNo").MergeArea.Interior.Color = (&HCCFFCC)

        Range("sheet1.NoticeDate").MergeArea.Locked = False
        Range("sheet1.NoticeDate").MergeArea.Interior.Color = (&HCCFFCC)

End If
Sheet1.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub
'Chandru
Private Sub CommandButton1_Click()
ActiveWorkbook.FollowHyperlink Address:="https://eportal.incometax.gov.in/iec/foservices/#/TaxCalc/calculator"

End Sub

Private Sub Worksheet_Activate()
ActiveWindow.DisplayHeadings = False

'Ankita_UR
Sheet1.Activate
Sheet1.Unprotect Password:=getmsgstate

    Call FitComments_DynamicHeight 'Ankita_23/02/2026=======


'ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 0  'Ankita_29/01/2026
'ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 1
 
ActiveSheet.Shapes("Button 618").Left = Range("Button_nature1").Left + 5
ActiveSheet.Shapes("Button 618").Top = Range("Button_nature1").Top + 5
ActiveSheet.Shapes("Button 492").Left = Range("Button_nature").Left + 5
ActiveSheet.Shapes("Button 492").Top = Range("Button_nature").Top + 5
ActiveSheet.Shapes("Button 47").ControlFormat.Enabled = False
ActiveSheet.Shapes("Button 51").ControlFormat.Enabled = False
Sheet1.Protect Password:=getmsgstate
End Sub
Private Sub Worksheet_Activate1()
ActiveWindow.DisplayHeadings = False
ActiveSheet.Shapes("Button 549").Left = Range("Button_nature_1").Left + 5
ActiveSheet.Shapes("Button 549").Top = Range("Button_nature_1").Top + 5
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub
Sub LockUnlock1398A_Gen()
On Error Resume Next
Application.EnableEvents = False
    Sheet1.Unprotect Password:=getmsgstate
    If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.ClearContents
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = True
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("sheet1.SeventhProvisoFlag").Value = ""
'Ankita_05/11/24
        Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)
        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)
        
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)
        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)
        
        Sheet1.Range("clauseiv7provisio139iFlg").MergeArea.ClearContents
        Sheet1.Range("clauseiv7provisio139iFlg").MergeArea.Locked = True
        Sheet1.Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range("clauseiv7provisio139iFlg").Value = ""
        
        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)
        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)
        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)
    Else
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Locked = False
        Sheet1.Range("sheet1.SeventhProvisoFlag").MergeArea.Interior.Color = "&HCCFFCC"
    End If
    If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
'    'Ankita_UR
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Rows("48:48").Hidden = False
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet201.Visible = xlSheetVisible
    Sheet202.Visible = xlSheetVisible
    ThisWorkbook.Protect Password:=getmsgstate
    Else
'    'Ankita_UR
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Rows("48:48").Hidden = True
    ThisWorkbook.Unprotect Password:=getmsgstate
    Sheet201.Visible = xlSheetHidden
    Sheet202.Visible = xlSheetHidden
    ThisWorkbook.Protect Password:=getmsgstate
    End If
    
    Application.EnableEvents = True
    Sheet1.Protect Password:=getmsgstate
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
 On Error GoTo endline
    Dim validateType As String
    validateType = ""
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets("Income Details")
Application.ScreenUpdating = False
Application.EnableEvents = False

Sheet1.Unprotect Password:=getmsgstate

Rows("14").RowHeight = 80

ActiveSheet.Shapes("Button 47").ControlFormat.Enabled = False
ActiveSheet.Shapes("Button 51").ControlFormat.Enabled = False


 Application.Calculation = xlCalculationAutomatic  'Ankita_20/04/2026=========
 
'Ankita_12/01/2026_'Email Address 2 : On data Entry=========================

If Not Application.Intersect(Sheet1.Range("sheet1.EmailAddress"), Sheet1.Range(Target.Address)) Is Nothing Then
     If Not mIncmDtls.CheckEmailAddress(Trim(Range("sheet1.EmailAddress").Value), validateType) Then
         If validateType = "1" Then fmsgbox "* ""Primary Email ID of the taxpayer, contains special characters in Sheet Income Details."""
         If validateType = "2" Then fmsgbox "* ""Primary Email ID of the taxpayer cannot be more than 125 characters in Sheet Income Details."""
         If validateType = "3" Then fmsgbox "* ""Invalid Primary Email ID of the taxpayer. Please retry."""
             Range("sheet1.EmailAddress").Select
             Range("sheet1.EmailAddress").Value = ""
     End If
 End If

If Not Application.Intersect(Sheet1.Range("sheet1.EmailAddress1"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Not mIncmDtls.CheckEmailAddress(Trim(Range("sheet1.EmailAddress1").Value), validateType) Then
                    If validateType = "1" Then fmsgbox "* ""Secondary Email ID, contains special characters in Sheet Income Details."""
                    If validateType = "2" Then fmsgbox "* ""Secondary Email ID cannot be more than 125 characters in Sheet Income Details."""
'                   If validateType = "3" Then fmsgbox "* Secondary Email ID entered is invalid in Sheet Income Details."
                    If validateType = "3" Then fmsgbox "* ""Invalid Secondary Email ID, Please retry."""
                        Range("sheet1.EmailAddress1").Select
                        Range("sheet1.EmailAddress1").Value = ""
                End If
            End If

'Ankita_29/01/2026===========

        If Not Application.Intersect(Sheet1.Range("sheet1.EmailRepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then
'                If Not mIncmDtls.CheckEmailAddress(Trim(Range("sheet1.EmailRepAssessee").Value), validateType) Then
                 If Not mIncmDtls.CheckEmailAddress_new1(Trim(Range("sheet1.EmailRepAssessee").Value), validateType) Then
                    If validateType = "1" Then fmsgbox "* ""Email Address of representative assessee, contains special characters in Sheet Income Details."""
                    If validateType = "2" Then fmsgbox "* ""Email Address of representative assessee cannot be more than 125 characters in Sheet Income Details."""
                    If validateType = "3" Then fmsgbox "* ""Email Address of representative assessee entered is invalid in Sheet Income Details."""
                        Range("sheet1.EmailRepAssessee").Select
                        Range("sheet1.EmailRepAssessee").Value = ""
                End If
            End If

'Ankita_14/01/2026_V0.2
'If Not Application.Intersect(Sheet1.Range("Secondary_Address"), Sheet1.Range(Target.Address)) Is Nothing Then
'    LockUnlock_Secondary_Address
'End If

'Ankita_30/03/2026======

'If Not Application.Intersect(Sheet1.Range("Secondary_Address"), Sheet1.Range(Target.Address)) Is Nothing Then
'    AutoPopulateSecondaryAddress
'End If


'newly added by Ankita on 29/11/2024
'===========================================================================================================
If Not Application.Intersect(Sheet1.Range("IncD.LessDeduction57"), Sheet1.Range(Target.Address)) Is Nothing Then
   If Range("Temp5").Value > 0 Then
       If Sheet1.Range("IncD.LessDeduction57").Value > Sheet1.Range("IncD.LessDeduction57New").Value Then
          If Sheet5.Range("BacValue").Value = 2 Then
           fmsgbox "* ""Deduction u/s 57(iia) cannot exceed lower of 1/3rd of Family pension or Rs. 15,000""."
            ElseIf Sheet5.Range("BacValue").Value = 1 Then
            fmsgboxsmall "* ""Deduction u/s 57(iia) cannot exceed lower of 1/3rd of Family pension Or Rs. 25,000""."
          End If
           Range("IncD.LessDeduction57").Value = ""
       End If
    End If
 
End If
If Not Application.Intersect(Sheet1.Range("sheet1.PinCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
         If Not (mIncmDtls.ValidatePinCode1()) Then
                 Sheet1.Range("sheet1.PinCode1").Value = ""
                End If
                If Sheet1.Range("sheet1.Country1").Value = "(Select)" Then
                 Sheet1.Range("sheet1.Country1") = "91-INDIA"
                 End If
                  Dim PinCode3 As Range
                 For Each PinCode3 In Sheet5.Range("PincodeList")
                    Dim code2 As String
                    code2 = Left(CStr(Sheet1.Range("sheet1.PinCode1").Value), 3)
                    If CStr(PinCode3.Value) = CStr(code2) Then
                        Sheet1.Range("sheet1.StateCode2").Value = Sheet5.Range("V" & PinCode3.row).Value
                        Sheet1.Range("sheet1.MobileCountryCode").Value = "91"
                        Exit For
                    End If
                Next PinCode3
End If

PartAGenChangeSub1 Target
PartAGenChangeSub2 Target
PartAGenChangeSub3 Target
PartAGenChangeSub4 Target

Dim PinCode As Range
        If Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
            If Sheet1.Range("sheet1.PinCode").Value <> "" Then

                Dim state1 As Variant
                state1 = Application.VLookup(Sheet1.Range("sheet1.PinCode").Value, Sheet5.Range("All_Pincode_V"), 2, False)
                If Not (state1 = UCase(Mid(Sheet1.Range("sheet1.StateCode1").Value, 4)) Or StateMatchesPin(Target.Value)) Then
                        Sheet1.Range("sheet1.PinCode").Value = ""

                End If
            End If
        End If
        If Not Application.Intersect(Sheet1.Range("sheet1.PinCode"), Sheet1.Range(Target.Address)) Is Nothing Then
         Dim DupFlag As Boolean
         DupFlag = False
                 For Each PinCode In Sheet5.Range("Duplicate_Pincode_List")
                    If Sheet1.Range("sheet1.PinCode").Value = PinCode Then
                        DupFlag = True
                        Exit For
                    End If
                 Next PinCode
                If DupFlag = False Then
                    Dim StateName
                    Dim StateCode

                    StateName = Application.VLookup(Sheet1.Range("sheet1.PinCode").Value, Sheet5.Range("All_Pincode_V"), 2, False)
                    If IsError(StateName) Then

                       Application.EnableEvents = False
                        Sheet1.Range("sheet1.PinCode").Value = ""
                     Application.EnableEvents = True

                    Else
                        StateCode = Application.VLookup(StateName, Sheet5.Range("CS:CT"), 2, False)
                        Sheet1.Range("sheet1.StateCode1").Value = StateCode
                        Sheet1.Range("sheet1.Country").Value = "91-INDIA"

                    End If
                 End If

                 If DupFlag = True Then
                 Application.EnableEvents = False

                    If Sheet1.Range("sheet1.StateCode1").Value <> "" Then
                        If Not StateMatchesPin(Sheet1.Range("sheet1.StateCode1").Value) Then
                            Sheet1.Range("sheet1.StateCode1").Value = "(Select)"
                        End If
                    End If
                 End If

         End If
         

'Ankita_14/01/2026_V0.2==========

'================================'ramya=================================================================================================

'Ankita_22/01/2026===
'If Not Intersect(Range("OSIncomeNotified89A_AmountUS"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("OSIncomeNotified89A_AmountUK"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("OSIncomeNotified89A_AmountCan"), Range(Target.Address)) Is Nothing Or Not Intersect(Range("OSIncreliefus89A"), Range(Target.Address)) Is Nothing Then
'
'    If (Sheet1.Range("OSIncomeNotified89A").Value <> "") Or (Sheet1.Range("OSIncreliefus89A").Value <> "" And Sheet1.Range("OSIncreliefus89A").Value <> 0) Then
'
'        If Sheet1.Range("OSIncreliefus89A").Value > Sheet1.Range("OSIncomeNotified89A").Value Then
'
'          Sheet1.Range("IncD_q1OS1").Value = Sheet1.Range("OSIncomeNotified89A")
'        Else
'            Sheet1.Range("IncD_q1OS1").Value = (Sheet1.Range("OSIncomeNotified89A").Value - Sheet1.Range("OSIncreliefus89A").Value)
'        End If
'    End If
'
'End If

If Not Application.Intersect(Sheet1.Range("IncD.Cost_LTCG"), Sheet1.Range(Target.Address)) Is Nothing Then
  If Sheet1.Range("IncD.Sale_LTCG") > 0 Then
            If Sheet1.Range("IncD.Cost_LTCG") > 0 And (Sheet1.Range("IncD.Sale_LTCG") - Sheet1.Range("IncD.Cost_LTCG")) > 125000 Then
                fmsgbox """Long term capital gains as per sec 112A should not exceed 1,25,000"""
                Sheet1.Range("IncD.Cost_LTCG").ClearContents
            End If
    End If
End If

If Not Application.Intersect(Sheet1.Range("IncD.Sale_LTCG"), Sheet1.Range(Target.Address)) Is Nothing Then
   If Sheet1.Range("IncD.Cost_LTCG") > 0 Then
        If Sheet1.Range("IncD.Sale_LTCG") > 0 And (Sheet1.Range("IncD.Sale_LTCG") - Sheet1.Range("IncD.Cost_LTCG")) > 125000 Then
            fmsgbox """In ITR 1, the maximum gains as per Section 112A can be INR 1,25,000/-. Please file ITR 2/3 if gains u/s 112A is more than INR 1,25,000/-"""
            Sheet1.Range("IncD.Sale_LTCG").ClearContents
        End If
    End If
End If
If Not Intersect(Range("sheet1.ReturnFileSec"), Range(Target.Address)) Is Nothing Then
LockUnlock1398A_Gen
LockUnlockTRP_Details
End If
'Ankita_07/03/2026===
'If Not Application.Intersect(Sheet1.Range("Nature_Others_1"), Sheet1.Range(Target.Address)) Is Nothing Then
'If Not ChkDescription23_24(Range(Target.Address).Value) Then
'                fmsgbox ("* Description should be Alpha numeric, not more than 125 characters and permitted Special characters(- , @ ;  / \ (  )  _ :)  in Sheet Income Details" & Chr(13))
'               Range(Target.Address).Value = ""
'        End If
'End If
If Not Intersect(Sheet1.Range("Others.Amount_2"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Range(Target.Address).Offset(0, -31).Value = "Family pension" Then
        If Sheet1.Range("IncD.LessDeduction57").Locked = False Then
            Sheet1.Range("IncD.LessDeduction57").MergeArea.ClearContents
        End If
    End If
End If
If Not Intersect(Sheet1.Range("IncD.Deduction16ic"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Target.Value > 5000 Then
        fmsgbox ("* Maximum deduction of Professional tax u/s 16(iii) can be claimed of Rs. 5,000." & Chr(13))
        Range(Target.Address).Value = 0
    End If
End If
If Not Application.Intersect(Sheet1.Range("sheet1.EmployerCategory1"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Sheet1.Range("sheet1.EmployerCategory1").Value = "Not Applicable (eg. Family pension etc)" Then
         MsgBox ("Since Nature of employment selected is ""Not Applicable"", Schedule ""Salary"" is greyed off. If you have Salary income, kindly select correct dropdown in Nature of employment in General Information.")
    End If
End If
If Not Intersect(Range("IncD.Allowances"), Range(Target.Address)) Is Nothing Then
    fmsgbox ("""Nature of Exempt Allowance(s) will get refreshed if the field ""Salary as per Section 17(1)"" is modified.""")
    Sheet1.Range("Others.Amount_1").ClearContents
End If
If Not Application.Intersect(Sheet1.Range("IncD.Div"), Sheet1.Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet1.Range("IncD_q1div"), Sheet1.Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet1.Range("IncD_q2div"), Sheet1.Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet1.Range("IncD_q3div"), Sheet1.Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet1.Range("IncD_q4div"), Sheet1.Range(Target.Address)) Is Nothing Or Not Application.Intersect(Sheet1.Range("IncD_q5div"), Sheet1.Range(Target.Address)) Is Nothing Then
If Not Application.Intersect(Sheet1.Range("IncD.Div"), Sheet1.Range(Target.Address)) Is Nothing Then
If Range("IncD.Div").Value > 0 Then
              Range("IncD_q1div").Value = Range("IncD.Div").Value
              fmsgbox ("If the benefit of quarterly computation of interest u/s 234C is sought, same may be claimed by making necessary modifications in below Table")
End If
End If
End If
If Not Intersect(Sheet1.Range("sheet1.EmployerCategory1"), Sheet1.Range(Target.Address)) Is Nothing Then
Check_GrossSalary23_24 Target
Nature_ExemptDropdown
End If
If Not Intersect(Range("sheet1.EmployerCategory1"), Range(Target.Address)) Is Nothing Then
            AnyOtherDeductions80CCH Target
        Sheet1.Unprotect Password:=getmsgstate
resetNatureexempt
Sheet1.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_3"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_3"), Sheet1.Range(Target.Address)) Is Nothing Then
        age = calculateAge23_24(Sheet1.Range("sheet1.DOB").Value)
       If age >= 60 Then
            If (Sheet1.Range("clauseiv7provisio139iAmount_3").Value < 50000 Or Sheet1.Range("clauseiv7provisio139iAmount_3").Value = "") Then
            fmsgbox "* Please select radio button as 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for senior citizen) " & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_3").Value = ""
            End If
        Else
            If (Sheet1.Range("clauseiv7provisio139iAmount_3").Value < 25000 Or Sheet1.Range("clauseiv7provisio139iAmount_3").Value = "") Then
            fmsgbox "* Please select radio button as 'No' if aggregate of tax deducted at source and tax collected at source during the previous year, in the case of the person, is less than twenty-five thousand rupees (or fifty-thousand for senior citizen) " & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_3").Value = ""
            End If
       End If
End If
End If

If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iAmount_4"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("clauseiv7provisio139iAmount_4").Value < 5000000 Or Sheet1.Range("clauseiv7provisio139iAmount_4").Value = "") Then
            fmsgbox "* Please select 'No' if total deposits in a savings bank account is less than fifty lakh rupees in the previous year" & Chr(10) & Chr(13)
            Sheet1.Range("clauseiv7provisio139iAmount_4").Value = ""
        End If
End If

'Ankita_19/01/2026=======
'If Not Application.Intersect(Sheet1.Range("OSIncreliefus89A"), Sheet1.Range(Target.Address)) Is Nothing Then
'    Sheet1.Unprotect Password:=getmsgstate
'    If Range("OSIncreliefus89A").Value > 0 Then
'        fmsgbox "* Please ensure to submit Form 10EE."
'    End If

'    If Sheet1.Range("OSIncreliefus89A").Value > Sheet1.Range("OSIncomeNotified89A").Value Then
'    fmsgbox "* Relief u/s 89A cannot be claimed more than income offered in pension accrued in a pension fund maintained in a notified country u/s 89A."
'    Sheet1.Range("OSIncreliefus89A").Value = 0
'    End If
'End If
Sheet1.Protect Password:=getmsgstate

'Ankita_16/01/2026=====
'If Not Application.Intersect(Sheet1.Range("Increliefus89A"), Sheet1.Range(Target.Address)) Is Nothing Then
'    If Range("Increliefus89A").Value > 0 Then
'        fmsgbox "* Please ensure to submit Form 10EE."
'    End If
    'Ankita_16/01/2026============
'    If Sheet1.Range("Increliefus89A").Value > Sheet1.Range("IncomeNotified89A").Value Then
'    fmsgbox "* Income claimed for relief from taxation u/s 89A  cannot be claimed more than income entered in income from retirement benefit account maintained in a notified country u/s 89A"
'    Sheet1.Range("Increliefus89A").Value = 0
'    End If
'End If

If Not Application.Intersect(Sheet1.Range("Others.NOI_1"), Sheet1.Range(Target.Address)) Is Nothing Then
    Dim nOfCountnature  As Long
    nOfCountnature = 0
    Dim noi
    For Each noi In Sheet1.Range("Others.NOI_1").Cells
        If Trim(CStr(noi.Value)) = CStr(Trim(Target.Value)) And (CStr(Trim(Target.Value)) <> "Any Other") Then
            nOfCountnature = nOfCountnature + 1
            
        End If
    Next noi
    If nOfCountnature > 1 Then
        fmsgbox "* Nature of Exempt Allowance has already been selected in Sheet Income Details"
        Target.Value = "(Select)"
    End If
    If (Sheet1.Range("Count4").Value > 1) Then
    fmsgbox " *You can only Claimed one out of the 3 Exempt allowances from u/s 10(10B)-First proviso, 10(10B)-Second Proviso and 10(10C) "
    Target.Value = "(Select)"
    End If
End If

'If Target.Validation.Type = 1 Then 'Ankita_27/01/2026

If Not Application.Intersect(Sheet1.Range("Others.Amount_1"), Sheet1.Range(Target.Address)) Is Nothing Then
Sheet1.Unprotect Password:=getmsgstate
'Change-23.11.2022.102.11A
        Nature_Amount23_24 Target
'Ankita_09/06/2025
     If (Sheet1.Range(Replace(Target.Address, "Z", "J")).Value) = "Sec 10(10B) Second proviso - Compensation under scheme approved by the Central Government" Then
        If Sheet1.Range(Target.Address).Value > 500000 Then
        fmsgbox ("*For Sec 10(10B) Second Proviso- Compensation under scheme approved by Central Government cannot exceed Rs.5 lakhs")
        Sheet1.Range(Target.Address).Value = ""
        Sheet1.Range(Target.Address).Select
        End If
     End If
     
 Sheet1.Protect Password:=getmsgstate
End If

If Not Application.Intersect(Sheet1.Range("Others.Amount"), Sheet1.Range(Target.Address)) Is Nothing Then
Sheet1.Unprotect Password:=getmsgstate
            If Mid(Sheet1.Range(Replace(Target.Address, "AH", "T")).Value, 1, 5) = "10(1)" Then
                 If (Sheet1.Range(Target.Address).Value > 5000) Then
                     fmsgbox ("* If sec 10(1) agricultural income is more than Rs 5000/- then use ITR 3 or 5 in Income Details")
                     Sheet1.Range(Target.Address).Value = ""
                     Sheet1.Range(Target.Address).Select
                 End If
            End If

            If Mid(Sheet1.Range(Replace(Target.Address, "AH", "H")).Value, 1, 10) = "Sec 10(34)" Then
                 If (Sheet1.Range(Target.Address).Value > 1000000) Then
                     fmsgbox ("* Sec 10(34) (Exempted Dividend Income) can not exceed Rs.1000000 in Income Details.")
                     Sheet1.Range(Target.Address).Value = ""
                     Sheet1.Range(Target.Address).Select
                 End If
            End If
'Change-25.11.2022.102.26A
            If Mid(Sheet1.Range(Replace(Target.Address, "AH", "H")).Value, 1, 4) = "LTCG" Then
                 If (Sheet1.Range(Target.Address).Value > 100000) Then
                     fmsgbox ("* LTCG u/s 112A not exceeding  Rs. 1 Lakh' should not exceed 1 lakhs")
                     Sheet1.Range(Target.Address).Value = ""
                     Sheet1.Range(Target.Address).Select
                 End If
            End If
'---end
   Sheet1.Protect Password:=getmsgstate
 End If
'End If 'Ankita_27/01/2026


'Changed by Riyaz on 05/04/2026
If Not Application.Intersect(Sheet1.Range("Nature_Others"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Target.Value = "10(32)-Minor child’s income—small exemption" Then
        DPMAddComment45 Target
    ElseIf Target.Value = "10(26)-Any income as referred to in section 10(26)" Then
        DPMAddComment46 Target
    ElseIf Target.Value = "10(26AAA)-Any income as referred to in section 10(26AAA)" Then
        DPMAddComment47 Target
    ElseIf Target.Value = "10(12A)-Any payment from the National Pension System Trust to an assessee" Then
        DPMAddComment48 Target
    ElseIf Target.Value = "10(12AA)-any payment from the National Pension System Trust " Then
        DPMAddComment49 Target
    ElseIf Target.Value = "10(12B)-Any payment from the National Pension System Trust to an Central Govt. Employee" Then
        DPMAddComment50 Target
    ElseIf Target.Value = "10(12BA)-partial withdrawal made from the National Pension System" Then
        DPMAddComment51 Target
    Else
        DPMADeleteComment45 Target
    End If
End If
'End Change IDS98/99

'Added on 05/12/2024
If Sheet5.Range("BacValue").Value = 2 Then
      If (Range("IncD.Deduction16ia").Value) > 50000 Then
                fmsgbox "* Deduction u/s 16(ia) can not exceed Rs. 50,000 or Net Salary, whichever is lower in Sheet Income Details"
                Range("IncD.Deduction16ia").Value = ""
                Range("IncD.Deduction16ia").Select
      End If

                Dim temp16a, temp16aa As Variant
                temp16a = Sheet1.Range("IncD.Deduction16ia").Value
                temp16aa = (Sheet1.Range("IncD.IncomeFromSal").Value + Sheet1.Range("IncD.Allowances").Value + Sheet1.Range("IncD.Perquisites").Value + Sheet1.Range("IncD.Profits").Value)

       If Sheet1.Range("IncD.Deduction16ia").Value > Sheet1.Range("Net_salary").Value Then
                fmsgbox "* Deduction u/s 16(ia) can not exceed Rs. 50,000 or Net Salary, whichever is lower in Sheet Income Details"
                Range("IncD.Deduction16ia").Value = ""
       End If
End If

'Added on 05/12/2024

If Sheet5.Range("BacValue").Value = 1 Then
       If (Range("IncD.Deduction16ia").Value) > 75000 Then
                fmsgbox "* Deduction u/s 16(ia) can not exceed Rs. 75,000 or Net Salary, whichever is lower in Sheet Income Details"
                Range("IncD.Deduction16ia").Value = ""
                Range("IncD.Deduction16ia").Select
       End If

                Dim temp16b, temp16bb As Variant
                temp16b = Sheet1.Range("IncD.Deduction16ia").Value
                temp16bb = (Sheet1.Range("IncD.IncomeFromSal").Value + Sheet1.Range("IncD.Allowances").Value + Sheet1.Range("IncD.Perquisites").Value + Sheet1.Range("IncD.Profits").Value)

       If Sheet1.Range("IncD.Deduction16ia").Value > Sheet1.Range("Net_salary").Value Then
                fmsgbox "* Deduction u/s 16(ia) can not exceed Rs. 75,000 or Net Salary, whichever is lower in Sheet Income Details"
                Range("IncD.Deduction16ia").Value = ""
       End If
End If

'================
'Ankita_28/01/2026========
'If Target.Validation.Type = 3 Then
    If Not Application.Intersect(Sheet1.Range("Others.NOI_1"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Not checkfieldSuperSpecialcharactername(Range(Target.Address).Value) Then
                fmsgbox ("* Nature of income should not contain Special characters < , > in Sheet Income Details" & Chr(13))
               Range(Target.Address).Value = ""
        End If

        Sheet1.Range("Z" & Target.row & ":AW" & Target.row).Value = ""

'Ankita_07/03/2026=========
'        If Sheet1.Range(Target.Address).Value = "Any Other" Then
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range(Replace(Target.Address, "J", "Z")).Value = ""
'            Sheet1.Range(Replace(Target.Address, "J", "Z")).Interior.Color = (&HCCFFCC)  ', "AU"
'            Sheet1.Range(Replace(Target.Address, "J", "Z")).MergeArea.Locked = False
'            Sheet1.Range(Replace(Target.Address, "Z", "J")).MergeArea.Locked = False
'            Sheet1.Protect Password:=getmsgstate
'        Else
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range(Replace(Target.Address, "J", "Z")).Value = "Not Applicable"
'            Sheet1.Range(Replace(Target.Address, "J", "Z")).Interior.Color = (&HD8D8D8)
'            Sheet1.Range(Replace(Target.Address, "J", "Z")).MergeArea.Locked = True
'             Sheet1.Range(Replace(Target.Address, "Z", "J")).MergeArea.Locked = False
'            Sheet1.Protect Password:=getmsgstate
'    End If
End If

If Not Application.Intersect(Sheet1.Range("Others.NOI_2"), Sheet1.Range(Target.Address)) Is Nothing Then

'Ankita_25-26
   If Sheet5.Range("BacValue").Value = 2 Then
        If Not checkfieldSuperSpecialcharactername(Range(Target.Address).Value) Then
                fmsgbox ("* Nature of income should not contain Special characters < , > in Sheet Income Details" & Chr(13))
               Range(Target.Address).Value = ""
        End If
        If Range("Temp5").Value > 0 And Not isdropdownblank(Range(Target.Address).Value) Then
            Sheet1.Unprotect Password:=getmsgstate
            Range("IncD.LessDeduction57").MergeArea.Locked = False
            Range("IncD.LessDeduction57").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Range("IncD.LessDeduction57").MergeArea.ClearContents
            Sheet1.Protect Password:=getmsgstate
            
           'Ankita_02/03/2026
            Else
                Sheet1.Unprotect Password:=getmsgstate
                Range("IncD.LessDeduction57").MergeArea.Locked = True
                Range("IncD.LessDeduction57").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Range("IncD.LessDeduction57").MergeArea.ClearContents
                Sheet1.Protect Password:=getmsgstate

        End If
    End If 'enabled by Ankita on 29/11/2024
        If Sheet1.Range(Target.Address).Value = "Dividend" Then
         Call insertRowUnderSectionWithFormula(5, False, 0)
        End If

        If Sheet1.Range(Target.Address).Value = "Any Other" Then
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(Target.Address, "J", "Z")).Value = ""
            Sheet1.Range(Replace(Target.Address, "J", "Z")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(Target.Address, "J", "Z")).MergeArea.Locked = False
            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(Target.Address, "J", "Z")).Value = "Not Applicable"
            Sheet1.Range(Replace(Target.Address, "J", "Z")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(Target.Address, "J", "Z")).MergeArea.Locked = True
            Sheet1.Protect Password:=getmsgstate
        End If
'End If

Dim nOfCount1  As Long
                nOfCount1 = CountDropdown1(Trim(Target.Value), Sheet1.Range("Others.NOI_2").Cells, Sheet1.Range("Others.NOI_2").count)
                If currCnt22 > 1 Or currCnt33 > 1 Or currCnt44 > 1 Or currCnt66 > 1 Or currCnt67 > 1 Or currCnt68 > 1 Or currCnt69 > 1 Or currCnt70 > 1 Then
                    fmsgbox "* Drop down has already been selected in Sheet Income Details"
                    Target.Value = "(Select)"
                End If
           End If  'Ankita_28/01/2026======
'     If Not Application.Intersect(Sheet1.Range("Others.NOI"), Sheet1.Range(Target.Address)) Is Nothing Then
'     Dim nOfCount  As Long
'                nOfCount = CountDropdown(Trim(Target.Value), Sheet1.Range("Others.NOI").Cells, Sheet1.Range("Others.NOI").count)
'                If currCnt1 > 1 Or currCnt2 > 1 Or currCnt3 > 1 Or currCnt4 > 1 Or currCnt5 > 1 Or currCnt6 > 1 Or currCnt7 > 1 Or currCnt8 > 1 Or currCnt9 > 1 Or currCnt10 > 1 Or currCnt11 > 1 Or currCnt12 > 1 Or currCnt13 > 1 Or currCnt14 > 1 Or currCnt15 > 1 Or currCnt16 > 1 Or currCnt17 > 1 Then
'                   fmsgbox "* Drop down has already been selected in Sheet Income Details"
'                    Target.Value = "(Select)"
'                End If
'
'        If Not checkfieldSuperSpecialcharactername(Range(Target.Address).Value) Then
'                fmsgbox ("* Nature of income should not contain Special characters < , >" & ",& in Sheet Income Details" & Chr(13))
'               Sheet1.Range(Target.Address).Value = ""
'        End If
'          If Sheet1.Range(Target.Address).Value = "Any Other" Then
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range(Replace(Target.Address, "H", "T")).Value = ""
'            Sheet1.Range(Replace(Target.Address, "H", "T")).Interior.Color = (&HCCFFCC)  ', "AU"
'            Sheet1.Range(Replace(Target.Address, "H", "T")).MergeArea.Locked = False
'            Sheet1.Protect Password:=getmsgstate
'           Else
'            Sheet1.Unprotect Password:=getmsgstate
'            Sheet1.Range(Replace(Target.Address, "H", "T")).Value = "Not Applicable"
'            Sheet1.Range(Replace(Target.Address, "H", "T")).Interior.Color = (&HD8D8D8)
'            Sheet1.Range(Replace(Target.Address, "H", "T")).MergeArea.Locked = True
'            Sheet1.Protect Password:=getmsgstate
'          End If
          
'          If Mid(Target.Value, 1, 13) = "Sec 10(26AAA)" Then
'            MsgBox ("The exemption 10(26AAA) is available only to certain categories of Sikkimese tax payers")
'          End If
'
'          If Mid(Target.Value, 1, 10) = "Sec 10(26)" Then
'            MsgBox ("The exemption 10(26) is available only to certain categories of tax payers in NER and Ladakh")
'          End If
'    End If
    
If Sheet5.Range("bacvalue").Value = 2 Then
        If Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) <> "Not" Then
            ThisWorkbook.Unprotect Password:=getmsgstate
            Sheet18.Visible = xlSheetVisible
            ThisWorkbook.Protect Password:=getmsgstate
        Else
            ThisWorkbook.Unprotect Password:=getmsgstate
            Sheet18.Visible = xlSheetHidden
            ThisWorkbook.Protect Password:=getmsgstate
        End If
End If
'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("sheet1.ReturnFileSec"), Sheet1.Range(Target.Address)) Is Nothing Then
        sPassword = EfilingCommon.getmsgstate
        Sheet1.Unprotect Password:=sPassword
If Range("sheet1.ReturnFileSec2").Value <> "(Select)" And Range("sheet1.ReturnFileSec2").Value <> "" Then
     fmsgbox "* Please select ""Filed u/s"" or ""Filed in response to notice u/s"" in Sheet Income Details"
      Range("sheet1.ReturnFileSec").Value = "(Select)"
      Range("sheet1.ReturnFileSec2").Value = "(Select)"
End If
If Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date" Or Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
       Dim vdate1 As Variant
        vdate1 = Sheet3.Range("Ver.Date").Value
        If Not ChkMaxInclusiveDate(Dformat(vdate1, "yyyy-mm-dd"), "2025-09-15") Then 'Ankita_15/07/2025
          Sheet1.Unprotect Password:=getmsgstate
           Range("IncD.IntrstPayUs234A").MergeArea.Locked = False
           Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = RGB(255, 255, 204)
           Range("IncD.IntrstPayUs234B").MergeArea.Locked = False
           Range("IncD.IntrstPayUs234B").MergeArea.Interior.Color = RGB(255, 255, 204)
           Range("IncD.IntrstPayUs234F").MergeArea.Locked = False
           Range("IncD.IntrstPayUs234F").MergeArea.Interior.Color = RGB(255, 255, 204)
        Else
           Sheet1.Unprotect Password:=getmsgstate
            Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
            Range("IncD.IntrstPayUs234A").MergeArea.Interior.ColorIndex = 2
            Range("IncD.IntrstPayUs234B").MergeArea.Locked = True
            Range("IncD.IntrstPayUs234B").MergeArea.Interior.ColorIndex = 2
            Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
            Range("IncD.IntrstPayUs234F").MergeArea.Interior.ColorIndex = 2
        End If
ElseIf Range("sheet1.ReturnFileSec").Value <> "139(1)-On or before due date" Then
          Sheet1.Unprotect Password:=getmsgstate
          Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
          Range("IncD.IntrstPayUs234A").MergeArea.Interior.ColorIndex = 2
          Range("IncD.IntrstPayUs234B").MergeArea.Locked = True
          Range("IncD.IntrstPayUs234B").MergeArea.Interior.ColorIndex = 2
          Range("IncD.IntrstPayUs234F").MergeArea.Locked = True
          Range("IncD.IntrstPayUs234F").MergeArea.Interior.ColorIndex = 2
End If

'Ankita_07/11/2025====================================
    If Range("sheet1.ReturnFileSec").Value = "139(4)-Belated" Then

           Sheet1.Unprotect Password:=getmsgstate
           Range("IncD.IntrstPayUs234A").MergeArea.Locked = False
           Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = RGB(255, 255, 204)

           Range("IncD.IntrstPayUs234B").MergeArea.Locked = False
           Range("IncD.IntrstPayUs234B").MergeArea.Interior.Color = RGB(255, 255, 204)

           Range("IncD.IntrstPayUs234F").MergeArea.Locked = False
           Range("IncD.IntrstPayUs234F").MergeArea.Interior.Color = RGB(255, 255, 204)
           
     End If
     
If Range("sheet1.ReturnFileSec").Value = "139(4)-Belated" Then
    BacYesValueChange
    Sheet1.Unprotect Password:=sPassword
    Sheet1.Range("AN31").MergeArea.Locked = True
    Sheet1.Range("AN31").MergeArea.Interior.Color = (&HD8D8D8)
    Sheet1.Protect Password:=sPassword
    ActiveSheet.Shapes("BacYes").ControlFormat.Enabled = False
   Else
     Sheet1.Unprotect Password:=sPassword
     Sheet1.Unprotect Password:=sPassword
     Sheet1.Range("AN31").MergeArea.Locked = False
     Sheet1.Range("AN31").MergeArea.Interior.Color = (&HCCFFCC)
     Sheet1.Protect Password:=sPassword
     ActiveSheet.Shapes("BacYes").ControlFormat.Enabled = True
     ActiveSheet.Shapes("BacNO").ControlFormat.Enabled = True
End If
        If Range("sheet1.ReturnFileSec").Value = "139(4)-Belated" Then
            If Sheet5.Range("BacValue").Value = 1 Then
            Dim vdate As Variant
            vdate = Sheet3.Range("Ver.Date").Value
            If Not ChkMinInclusiveDate(Dformat(vdate, "yyyy-mm-dd"), "2025-09-16") Then 'Ankita_22/07/2025 'Year changed by Ankita on 29/11/2024
'               fmsgbox "* Filing section u/s. 139(4) cannot be selected till 31st July, 2024 or extended due date "
                fmsgbox "* Filing section u/s. 139(4) cannot be selected till 31st July, 2025 or extended due date " 'Year changed by Ankita on 29/11/2024
                Range("sheet1.ReturnFileSec").Value = "(Select)"
            End If
            End If
        End If
        
        'Ankita_27/02/2026
    If Range("sheet1.ReturnFileSec").Value = "139(5)-Revised" Then
            If (Dformat(Sheet3.Range("Ver.Date").Value, "yyyy-mm-dd")) > "2027-03-31" Then
                fmsgbox "* ""Selected Filing section 139(5) cannot be selected post 31st March, 2027"""
                Range("sheet1.ReturnFileSec").Value = "(Select)"
            End If
        End If
        
    PartAGen_FilingStatusCode1
    PartAGen_FilingStatusCode2
        If Range("sheet1.ReturnFileSec").Value = "(Select)" Then
           Range("sheet1.ReturnFileSec1").Value = "(Select)"
        ElseIf Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date" Then
           Range("sheet1.ReturnFileSec1").Value = "11 - 139(1)-On or before due date"
        ElseIf Range("sheet1.ReturnFileSec").Value = "139(4)-Belated" Then
           Range("sheet1.ReturnFileSec1").Value = "12 - 139(4)-Belated"

        ElseIf Range("sheet1.ReturnFileSec").Value = "142(1)" Then
           Range("sheet1.ReturnFileSec1").Value = "13 - 142(1)"
          ElseIf Range("sheet1.ReturnFileSec").Value = "148" Then
           Range("sheet1.ReturnFileSec1").Value = "14 - 148"
        ElseIf Range("sheet1.ReturnFileSec").Value = "153A" Then
           Range("sheet1.ReturnFileSec1").Value = "15 - 153A"
        ElseIf Range("sheet1.ReturnFileSec").Value = "153C" Then
           Range("sheet1.ReturnFileSec1").Value = "16 - 153C"
        ElseIf Range("sheet1.ReturnFileSec").Value = "139(9)" Then
           Range("sheet1.ReturnFileSec1").Value = "18 - 139(9)"
        ElseIf Range("sheet1.ReturnFileSec").Value = "139(5)-Revised" Then
           Range("sheet1.ReturnFileSec1").Value = "17 - 139(5)-Revised"
        ElseIf Range("sheet1.ReturnFileSec").Value = "119(2)(b)- After condonation of delay" Then
           Range("sheet1.ReturnFileSec1").Value = "20 -119(2)(b)- After condonation of delay"
        ElseIf Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
           Range("sheet1.ReturnFileSec1").Value = "21 - 139(8A)"
        End If
        
            If Mid(Range("sheet1.ReturnFileSec1").Value, 1, 2) = "20" Then
                fmsgbox "* ""119(2)(b)- After condonation of delay should be filed after expiry of due date mentioned in 139(1)/139(4)""."
             End If
            LockUnlockReturnfilesec
            Sheet1.Protect Password:=sPassword
End If


'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("sheet1.EmployerCategory1"), Sheet1.Range(Target.Address)) Is Nothing Then
'If (Target.name.name) = "sheet1.EmployerCategory1" Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Cen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Sta" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Pub" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 16) = "Pensioners - Oth") Then
        Range("IncD.Deduction16").MergeArea.ClearContents
        Sheet1.Unprotect Password:=sPassword
        Range("IncD.Deduction16").MergeArea.Interior.Color = (&HD8D8D8)
        Range("IncD.Deduction16").MergeArea.Locked = True
        Sheet1.Protect Password:=sPassword
ElseIf (Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Cen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Sta" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
         Sheet1.Unprotect Password:=sPassword
  If Sheet5.Range("BacValue").Value <> 1 Then    '03-02-2026  Malli SIT-109371
        Range("IncD.Deduction16").MergeArea.Locked = False
        Range("IncD.Deduction16").MergeArea.Interior.Color = (&HCCFFCC)
  End If
 End If
If (Range("Sheet1.EmployerCategory1").Value = "(Select)") Then
        Sheet1.Unprotect Password:=sPassword
   If Sheet5.Range("BacValue").Value <> 1 Then    '03-02-2026  Malli SIT-109371
       Range("IncD.Deduction16").MergeArea.Locked = False
       Range("IncD.Deduction16").MergeArea.Interior.Color = (&HCCFFCC)
   End If
End If
If (Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen") Then
    Range("IncD.Section80CCD").MergeArea.ClearContents
     Sheet1.Unprotect Password:=sPassword
    Range("IncD.Section80CCD").MergeArea.Locked = True
    Range("IncD.Section80CCD").MergeArea.Interior.Color = (&HD8D8D8)
Else
     Sheet1.Unprotect Password:=sPassword
    Range("IncD.Section80CCD").MergeArea.Locked = False
    Range("IncD.Section80CCD").MergeArea.Interior.Color = (&HCCFFCC)
End If
 Sheet1.Unprotect Password:=sPassword
If (Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen") Then
ElseIf (Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Cen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Sta" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
End If
If (Range("Sheet1.EmployerCategory1").Value = "(Select)") Then
End If
If (Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Not" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pen") Then
End If
If (Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Cen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Sta" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Oth") Then
End If
If (Range("Sheet1.EmployerCategory1").Value = "(Select)") Then
End If
Sheet1.Protect Password:=getmsgstate
End If
'End If 'Ankita_28/01/2026==========

'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("sheet1.SeventhProvisoFlag"), Sheet1.Range(Target.Address)) Is Nothing Then

'If (Target.name.name) = "sheet1.SeventhProvisoFlag" Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("sheet1.SeventhProvisoFlag").Value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").Value = "(Select)" Then
        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)
        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)
        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)
        Range("clauseiv7provisio139iFlg").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iFlg").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HD8D8D8)
        Range("clauseiv7provisio139iFlg").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("clauseiv7provisio139iFlg").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").Value, 1, 1) = "Y") Then
        Range("clauseiv7provisio139iFlg").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg").MergeArea.Value = "(Select)"
        Range("clauseiv7provisio139iFlg").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg").MergeArea.Interior.Color = (&HCCFFCC)
End If
If (Mid(Range("sheet1.SeventhProvisoFlag").Value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").Value = "(Select)" Then
        Range("Sheet1.DepositAmountFlag").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("Sheet1.DepositAmountFlag").MergeArea.Locked = True
        Range("Sheet1.DepositAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)
        Range("Sheet1.DepositAmount").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("Sheet1.DepositAmount").MergeArea.Locked = True
        Range("Sheet1.DepositAmount").MergeArea.Interior.Color = (&HD8D8D8)
End If
If (Mid(Range("sheet1.SeventhProvisoFlag").Value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").Value = "(Select)" Then
        Range("Sheet1.AggrigateAmountFlag").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HD8D8D8)
        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").Value, 1, 1) = "Y") Then
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.AggrigateAmountFlag").MergeArea.Value = "(Select)"
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If
If (Mid(Range("sheet1.SeventhProvisoFlag").Value, 1, 1) = "N") Or Range("sheet1.SeventhProvisoFlag").Value = "(Select)" Then
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("sheet1.SeventhProvisoFlag").Value, 1, 1) = "Y") Then
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Interior.Color = (&HCCFFCC)
        Range("Sheet1.AggrigateAmountFlag1").MergeArea.Value = "(Select)"
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If
'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("Sheet1.AggrigateAmountFlag"), Sheet1.Range(Target.Address)) Is Nothing Then

'If (Target.name.name) = "Sheet1.AggrigateAmountFlag" Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("Sheet1.AggrigateAmountFlag").Value, 1, 1) = "N") Then
        Range("Sheet1.AggrigateAmount").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("Sheet1.AggrigateAmount").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag").Value, 1, 1) = "Y") Then

        Range("Sheet1.AggrigateAmount").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If

'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("Sheet1.AggrigateAmountFlag1"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("Sheet1.AggrigateAmountFlag1").Value, 1, 1) = "N") Then
        Range("Sheet1.AggrigateAmount1").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = True
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("Sheet1.AggrigateAmountFlag1").Value, 1, 1) = "Y") Then
        Range("Sheet1.AggrigateAmount1").MergeArea.Locked = False
        Range("Sheet1.AggrigateAmount1").MergeArea.Interior.Color = (&HCCFFCC)
End If

Sheet1.Protect Password:=getmsgstate
End If

'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("Sheet1.DepositAmount"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("Sheet1.DepositAmount").Value < 10000000 Or Sheet1.Range("Sheet1.DepositAmount").Value = 10000000 Or Sheet1.Range("Sheet1.DepositAmount").Value = "") Then
            fmsgbox "* Amount entered cannot be less than or equal to Rs 1 Crore in Sheet Income Details." & Chr(10) & Chr(13)
                    Sheet1.Range("Sheet1.DepositAmount").Value = " "
        End If
End If
'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("Sheet1.AggrigateAmount"), Sheet1.Range(Target.Address)) Is Nothing Then
'If (Target.name.name) = "Sheet1.AggrigateAmount" Then
        If (Sheet1.Range("Sheet1.AggrigateAmount").Value < 200000 Or Sheet1.Range("Sheet1.AggrigateAmount").Value = 200000) Then
            fmsgbox "* Amount entered cannot be less than or equal to Rs 2 Lakhs in Sheet Income Details." & Chr(10) & Chr(13)
                    Sheet1.Range("Sheet1.AggrigateAmount").Value = " "
        End If
End If

'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("Sheet1.AggrigateAmount1"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Sheet1.Range("Sheet1.AggrigateAmount1").Value < 100000 Or Sheet1.Range("Sheet1.AggrigateAmount1").Value = 100000) Then
            fmsgbox "* Amount entered cannot be less than or equal to Rs 1 Lakhs in Sheet Income Details." & Chr(10) & Chr(13)
                    Sheet1.Range("Sheet1.AggrigateAmount1").Value = " "
        End If
End If

'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("IncD.Section80EE"), Sheet1.Range(Target.Address)) Is Nothing Then
            Sheet1.Unprotect Password:=getmsgstate
    If (Range("IncD.Section80EE").Value) <> "" And (Range("IncD.Section80EE").Value) <> 0 Then
        Sheet1.Range("IncD.Section80EEA").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
        Sheet1.Range("IncD.Section80EEA").MergeArea.Interior.Color = "&HD8D8D8"
    Else
        Sheet1.Range("IncD.Section80EEA").MergeArea.ClearContents
         Sheet1.Unprotect Password:=sPassword
        Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = False
        Sheet1.Range("IncD.Section80EEA").MergeArea.Interior.Color = "&HCCFFCC"

    End If
    Sheet1.Protect Password:=getmsgstate
    End If
If Target.Validation.Type = 3 Then
'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
            If UCase(Range("sheet1.StateCode1").Value) <> "99-FOREIGN" Then
                If Range("sheet1.StateCode1").Value = "(Select)" Then
                   Range("sheet1.Country").Value = "(Select)"
                Else
                    Range("sheet1.Country").Value = ""
                    Range("sheet1.Country").Value = "91-INDIA"
                    Sheet1.Range("sheet1.mobileCountryCode").Value = "91"
                End If

                Sheet1.Unprotect Password:=getmsgstate

                Sheet1.Range("sheet1.PinCode").MergeArea.Value = ""
                Sheet1.Range("sheet1.PinCode").MergeArea.Locked = False
                Sheet1.Range("sheet1.PinCode").MergeArea.Interior.Color = (&HCCFFCC)

                Sheet1.Range("HASZIP").MergeArea.Value = ""
                Sheet1.Unprotect Password:=sPassword
                Sheet1.Range("HASZIP").MergeArea.Locked = True
                Sheet1.Range("HASZIP").MergeArea.Interior.Color = (&HD8D8D8)

                Sheet1.Range("sheet1.ZipCode").MergeArea.Value = ""
                Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
                Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Protect Password:=getmsgstate
            End If
        End If
If Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
         If UCase(Range("sheet1.StateCode1").Value) = "99-FOREIGN" Then
            Range("sheet1.Country").Value = "(Select)"
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.mobileCountryCode").Value = ""
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
End If
End If  'Ankita_28/01/2026

If Target.Validation.Type = 3 Then
'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
            If UCase(Range("sheet1.StateCode2").Value) <> "99-FOREIGN" Then
                If Range("sheet1.StateCode2").Value = "(Select)" Then
                   Range("sheet1.Country1").Value = "(Select)"
                Else
                    Range("sheet1.Country1").Value = ""
                    Range("sheet1.Country1").Value = "91-INDIA"
                    Sheet1.Range("sheet1.mobileCountryCode").Value = "91"
                End If

                Sheet1.Unprotect Password:=getmsgstate

                Sheet1.Range("sheet1.PinCode1").MergeArea.Value = ""
                Sheet1.Range("sheet1.PinCode1").MergeArea.Locked = False
                Sheet1.Range("sheet1.PinCode1").MergeArea.Interior.Color = (&HCCFFCC)

                Sheet1.Range("HASZIP1").MergeArea.Value = ""
                Sheet1.Unprotect Password:=sPassword
                Sheet1.Range("HASZIP1").MergeArea.Locked = True
                Sheet1.Range("HASZIP1").MergeArea.Interior.Color = (&HD8D8D8)

                Sheet1.Range("sheet1.ZipCode1").MergeArea.Value = ""
                Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = True
                Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Protect Password:=getmsgstate
            End If
        End If
If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
       If UCase(Range("sheet1.StateCode2").Value) = "99-FOREIGN" Then
            Range("sheet1.Country1").Value = "(Select)"
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range("sheet1.mobileCountryCode").Value = ""
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
End If
End If 'Ankita_28/01/2026==========


'=====================================
'Ankita_28/01/2026==========
If Not Application.Intersect(Sheet1.Range("sheet1.Country"), Sheet1.Range(Target.Address)) Is Nothing Then
            PartAGen_Country
        End If

If Not Application.Intersect(Sheet1.Range("HASZIP"), Sheet1.Range(Target.Address)) Is Nothing Then
            Target.Value = UCase(Target.Value)
            If Mid(Trim(Sheet1.Range("HASZIP").Value), 1, 1) = "Y" Then
                Sheet1.Unprotect Password:=getmsgstate
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Value = "XXXXXX"
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = True
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Protect Password:=getmsgstate
            Else
                Sheet1.Unprotect Password:=getmsgstate
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Value = ""
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Locked = False
                   Sheet1.Range("sheet1.ZipCode").MergeArea.Interior.Color = (&HCCFFCC)
                Sheet1.Protect Password:=getmsgstate
            End If
        End If

'Ankita_24/01/2026================
If Not Application.Intersect(Sheet1.Range("sheet1.Country1"), Sheet1.Range(Target.Address)) Is Nothing Then
            PartAGen_Country1
        End If

If Not Application.Intersect(Sheet1.Range("HASZIP1"), Sheet1.Range(Target.Address)) Is Nothing Then
            Target.Value = UCase(Target.Value)
            If Mid(Trim(Sheet1.Range("HASZIP1").Value), 1, 1) = "Y" Then
                Sheet1.Unprotect Password:=getmsgstate
                   Sheet1.Range("sheet1.ZipCode1").MergeArea.Value = "XXXXXX"
                   Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = True
                   Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = (&HD8D8D8)
                Sheet1.Protect Password:=getmsgstate
            Else
                Sheet1.Unprotect Password:=getmsgstate
                   Sheet1.Range("sheet1.ZipCode1").MergeArea.Value = ""
                   Sheet1.Range("sheet1.ZipCode1").MergeArea.Locked = False
                   Sheet1.Range("sheet1.ZipCode1").MergeArea.Interior.Color = (&HCCFFCC)
                Sheet1.Protect Password:=getmsgstate
            End If
        End If
'=================================

If Target.Validation.Type = 6 Then
'Ankita_28/01/2026=======
If Not Application.Intersect(Sheet1.Range("sheet1.PAN"), Sheet1.Range(Target.Address)) Is Nothing Then
'    If (Target.name.name) = "sheet1.PAN" Then
        Dim TargetPAN As String
        TargetPAN = ""
        TargetPAN = Target.name.name
        If Not mIncmDtls.CheckPAN(UCase(Trim(Range(TargetPAN).Value))) Then
            fmsgbox "* Invalid PAN. PAN format should be First 5 Alphabets, next 4 digits, then 1 Alphabet in Sheet Income Details." & Chr(10) & Chr(13)
            Sheet1.Range(TargetPAN).Select
            Sheet1.Range(TargetPAN).Value = ""
        End If
        'Ankita_03/02/2025
        Application.EnableEvents = False
        Sheet1.Range(TargetPAN).Value = UCase(Range(TargetPAN).Value)
        
        'Ankita_28/01/2026=======
        If Not Application.Intersect(Sheet1.Range("sheet1.PAN"), Sheet1.Range(Target.Address)) Is Nothing Then
            If (Sheet1.Range("sheet1.DOB") <> "") And calculateAge(Trim(Sheet1.Range("sheet1.DOB").Value)) < 18 Then
                If Sheet3.Range("Ver.pan").Value <> "" Then
                    Sheet3.Range("Ver.pan").Value = ""
                End If
                If Sheet3.Range("Ver.capacity").Value = "Self" Then
                    Sheet3.Range("Ver.capacity").Value = ""
                End If
            Else
                Sheet3.Range("Ver.pan").Value = UCase(Range(TargetPAN).Value)
            End If
        End If
    End If
    
       'Ankita_28/01/2026=======
        If Not Application.Intersect(Sheet1.Range("sheet1.PinCode"), Sheet1.Range(Target.Address)) Is Nothing Then

'               If (Target.name.name) = "sheet1.PinCode" Then
                If Not (mIncmDtls.ValidatePinCode()) Then
                 Sheet1.Range("sheet1.PinCode").Value = ""
                End If
                If Sheet1.Range("sheet1.Country").Value = "(Select)" Then
                 Sheet1.Range("sheet1.Country") = "91-INDIA"
                 End If
                  'Dim PinCode As Range
                 For Each PinCode In Sheet5.Range("PincodeList")
                    Dim code As String
                    code = Left(CStr(Sheet1.Range("sheet1.PinCode").Value), 3)
                    If CStr(PinCode.Value) = CStr(code) Then
                        Sheet1.Range("sheet1.StateCode1").Value = Sheet5.Range("V" & PinCode.row).Value
                        Sheet1.Range("sheet1.MobileCountryCode").Value = "91"
                        Exit For
                    End If
                Next PinCode
            End If
'Ankita_16/10/2025===========
If Not Application.Intersect(Sheet1.Range("sheet1.Mobileno"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Range("sheet1.Mobileno")) <> "" Then
            If Sheet1.Range("sheet1.MobileCountryCode").Value = "91" Then
                If Len(Range("sheet1.Mobileno")) <> 10 Then
                    fmsgbox "* ""Primary Mobile No. should be 10 digits in sheet Income Details"""
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                End If
            Else
     If Len(Range("sheet1.Mobileno")) < 5 Or Len(Range("sheet1.Mobileno")) > 10 Then
                    fmsgbox "* ""Primary Mobile No. is invalid in sheet Income Details"""
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                End If
            End If
            If Len(Range("sheet1.Mobileno")) > 0 Then
                If Not IsNumeric(Trim(Range("sheet1.Mobileno").Value)) Then
                    fmsgbox "* ""Primary Mobile No. is Invalid in sheet Income Details"""
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                End If
            End If
        If Len(Range("sheet1.Mobileno")) > 0 Then
              If Not (ValidateMobileNumber()) Then
                 Sheet1.Range("sheet1.Mobileno").Value = ""
                End If
            End If
            If Len(Range("sheet1.Mobileno")) > 0 Then
                If Not CheckSpecialCharacter(Trim(Range("sheet1.Mobileno").Value)) Then
                    fmsgbox "* ""Primary Mobile No. is Invalid in sheet Income Details"""
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                End If
            End If
            
           If Len(Range("sheet1.Mobileno")) > 0 Then
                If Mid(Trim(Range("sheet1.Mobileno").Value), 1, 1) = 0 Then
                    fmsgbox "* ""Primary Mobile No. should not begin With '0' in sheet Income Details"""
                    Range("sheet1.Mobileno").MergeArea.ClearContents
                End If
            End If
            
            If Len(Range("sheet1.Mobileno")) + Len(Range("sheet1.MobileCountryCode")) > 15 Then
                fmsgbox "* ""Primary Mobile No. And country code combination should not be more than 15 digits in Sheet Income Details"""
                Range("sheet1.Mobileno").MergeArea.ClearContents
                Range("sheet1.mobileCountryCode").Value = ""
             End If
        End If
    End If
    If Not Application.Intersect(Sheet1.Range("sheet1.MobileCountryCode"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Len(Range("sheet1.MobileCountryCode")) <= 5 Then
            If Not IsNumeric(Trim(Range("sheet1.MobileCountryCode").Value)) Then
                fmsgbox "* ""Country code for Primary Mobile No. is Invalid in sheet Income Details"""
                Range("sheet1.MobileCountryCode").Value = ""
            End If
        End If
        If Len(Range("sheet1.MobileCountryCode")) <= 5 Then
            If Not CheckSpecialCharacter(Trim(Range("sheet1.MobileCountryCode").Value)) Then
                fmsgbox "* ""Country code for Primary Mobile No. is Invalid in sheet Income Details"""
                Range("sheet1.MobileCountryCode").Value = ""
            End If
        End If
    If Len(Range("sheet1.MobileCountryCode")) <= 5 Then
            If Not chkNumeric(Trim(Range("sheet1.MobileCountryCode").Value)) Then
                fmsgbox "* Country code for Primary Mobile No. is Invalid in sheet Income Details"
                Range("sheet1.MobileCountryCode").Value = ""
            End If
        End If
    End If
    
'Ankita_13/01/2026_V0.2==========

If Not Intersect(Range("sheet1.Mobileno1"), Range(Target.Address)) Is Nothing Then
        If (Range("sheet1.Mobileno1")) <> "" Then
            If Sheet1.Range("sheet1.MobileCountryCode1").Value = "91" Then
                If Len(Range("sheet1.Mobileno1")) <> 10 Then
                    fmsgbox "* ""Secondary mobile number should be 10 digits in sheet Income Details"""
                    Range("sheet1.Mobileno1").MergeArea.ClearContents
                End If
            Else
     If Len(Range("sheet1.Mobileno1")) < 5 Or Len(Range("sheet1.Mobileno1")) > 10 Then
                    fmsgbox "* ""Secondary mobile number is invalid in sheet Income Details"""
                    Range("sheet1.Mobileno1").MergeArea.ClearContents
                End If
            End If
            If Len(Range("sheet1.Mobileno1")) > 0 Then
                If Not IsNumeric(Trim(Range("sheet1.Mobileno1").Value)) Then
                    fmsgbox "* ""Secondary mobile number is Invalid in sheet Income Details"""
                    Range("sheet1.Mobileno1").MergeArea.ClearContents
                End If
            End If
        If Len(Range("sheet1.Mobileno1")) > 0 Then
              If Not (ValidateMobileNumber()) Then
                 Sheet1.Range("sheet1.Mobileno1").Value = ""
                End If
            End If
            If Len(Range("sheet1.Mobileno1")) > 0 Then
                If Not CheckSpecialCharacter(Trim(Range("sheet1.Mobileno1").Value)) Then
                    fmsgbox "* Secondary mobile number is Invalid in sheet Income Details"
                    Range("sheet1.Mobileno1").MergeArea.ClearContents
                End If
            End If
           If Len(Range("sheet1.Mobileno1")) > 0 Then
                If Mid(Trim(Range("sheet1.Mobileno1").Value), 1, 1) = 0 Then
                    fmsgbox "* ""Secondary mobile number should not begin With '0' in sheet Income Details"""
                    Range("sheet1.Mobileno1").MergeArea.ClearContents
                End If
            End If
            If Len(Range("sheet1.Mobileno1")) + Len(Range("sheet1.MobileCountryCode1")) > 15 Then
                fmsgbox "* ""Secondary mobile number And country code combination should not be more than 15 digits in Sheet Income Details"""
                Range("sheet1.Mobileno1").MergeArea.ClearContents
                Range("sheet1.mobileCountryCode1").Value = ""
             End If
        End If
    End If
    
    If Not Application.Intersect(Sheet1.Range("sheet1.MobileCountryCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Len(Range("sheet1.MobileCountryCode1")) <= 5 Then
            If Not IsNumeric(Trim(Range("sheet1.MobileCountryCode1").Value)) Then
                fmsgbox "* ""Country code for Secondary mobile number is Invalid in sheet Income Details"""
                Range("sheet1.MobileCountryCode1").Value = ""
            End If
        End If
        If Len(Range("sheet1.MobileCountryCode1")) <= 5 Then
            If Not CheckSpecialCharacter(Trim(Range("sheet1.MobileCountryCode1").Value)) Then
                fmsgbox "* ""Country code for Secondary mobile number is Invalid in sheet Income Details"""
                Range("sheet1.MobileCountryCode1").Value = ""
            End If
        End If
    If Len(Range("sheet1.MobileCountryCode1")) <= 5 Then
            If Not chkNumeric(Trim(Range("sheet1.MobileCountryCode1").Value)) Then
                fmsgbox "* ""Country code for Secondary mobile number is Invalid in sheet Income Details"""
                Range("sheet1.MobileCountryCode1").Value = ""
            End If
        End If
    End If
    
'Ankita_13/01/2026_V0.2==========
    If Not Application.Intersect(Sheet1.Range("sheet1.ContactRepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then
        If (Range("sheet1.ContactRepAssessee")) <> "" Then
            If Sheet1.Range("sheet1.CountryCodeRepAssessee").Value = "91" Then
                If Len(Range("sheet1.ContactRepAssessee")) <> 10 Then
                    fmsgbox "* ""Contact No. of Representative assessee should be 10 digits in sheet Income Details"""
                    Range("sheet1.ContactRepAssessee").MergeArea.ClearContents
                End If
            Else
     If Len(Range("sheet1.ContactRepAssessee")) < 5 Or Len(Range("sheet1.ContactRepAssessee")) > 10 Then
                    fmsgbox "* ""Contact No. of Representative assessee is invalid in sheet Income Details"""
                    Range("sheet1.ContactRepAssessee").MergeArea.ClearContents
                End If
            End If
            If Len(Range("sheet1.ContactRepAssessee")) > 0 Then
                If Not IsNumeric(Trim(Range("sheet1.ContactRepAssessee").Value)) Then
                    fmsgbox "* ""Contact No. of Representative assessee is Invalid in sheet Income Details"""
                    Range("sheet1.ContactRepAssessee").MergeArea.ClearContents
                End If
            End If
        If Len(Range("sheet1.ContactRepAssessee")) > 0 Then
              If Not (ValidateMobileNumber()) Then
                 Sheet1.Range("sheet1.ContactRepAssessee").Value = ""
                End If
            End If
            If Len(Range("sheet1.ContactRepAssessee")) > 0 Then
                If Not CheckSpecialCharacter(Trim(Range("sheet1.ContactRepAssessee").Value)) Then
                    fmsgbox "* ""Contact No. of Representative assessee is Invalid in sheet Income Details"""
                    Range("sheet1.ContactRepAssessee").MergeArea.ClearContents
                End If
            End If
           If Len(Range("sheet1.ContactRepAssessee")) > 0 Then
                If Mid(Trim(Range("sheet1.ContactRepAssessee").Value), 1, 1) = 0 Then
                    fmsgbox "* ""Contact No. of Representative assessee should not begin With '0' in Sheet Income Details"""
                    Range("sheet1.ContactRepAssessee").MergeArea.ClearContents
                End If
            End If
            If Len(Range("sheet1.ContactRepAssessee")) + Len(Range("sheet1.CountryCodeRepAssessee")) > 15 Then
                fmsgbox "* ""Mobile No. And Country Code in contact No. of representative assessee combination should not be more than 15 digits in Sheet Income Details"""
                Range("sheet1.ContactRepAssessee").MergeArea.ClearContents
                Range("sheet1.CountryCodeRepAssessee").Value = ""
             End If
        End If
    End If
    
      If Not Application.Intersect(Sheet1.Range("sheet1.CountryCodeRepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Len(Range("sheet1.CountryCodeRepAssessee")) <= 5 Then
            If Not IsNumeric(Trim(Range("sheet1.CountryCodeRepAssessee").Value)) Then
                fmsgbox "* ""Country code in contact No. of Representative assessee is Invalid in sheet Income Details"""
                Range("sheet1.CountryCodeRepAssessee").Value = ""
            End If
        End If
        If Len(Range("sheet1.CountryCodeRepAssessee")) <= 5 Then
            If Not CheckSpecialCharacter(Trim(Range("sheet1.CountryCodeRepAssessee").Value)) Then
                fmsgbox "* ""Country code in contact No. of Representative assessee is Invalid in sheet Income Details"""
                Range("sheet1.CountryCodeRepAssessee").Value = ""
            End If
        End If
    If Len(Range("sheet1.CountryCodeRepAssessee")) <= 5 Then
            If Not chkNumeric(Trim(Range("sheet1.CountryCodeRepAssessee").Value)) Then
                fmsgbox "* ""Country code in contact No. of Representative assessee is Invalid in sheet Income Details"""
                Range("sheet1.CountryCodeRepAssessee").Value = ""
            End If
        End If
    End If
    '==================
          If Not Application.Intersect(Sheet1.Range("sheet1.FirstName"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.FirstName").Value) > 0 Then
                If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.FirstName").Value) Then
                    fmsgbox ("* First name should not Contain <, > characters in Sheet Income Details" & Chr(13))
                    Sheet1.Range("sheet1.FirstName").Value = ""
                Else
                 Sheet1.Range("sheet1.FirstName").Value = (Sheet1.Range("sheet1.FirstName").Value)
                End If
                End If
            End If
          If Not Application.Intersect(Sheet1.Range("sheet1.MiddleName"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.MiddleName")) > 0 Then
                 If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.MiddleName").Value) Then
                    fmsgbox ("* Middle name should not Contain <, >  characters in Sheet Income Details" & Chr(13))
                    Sheet1.Range("sheet1.MiddleName").Value = ""
                Else
                 Sheet1.Range("sheet1.MiddleName").Value = (Sheet1.Range("sheet1.MiddleName").Value)
                 End If
                End If
            End If
         If Not Application.Intersect(Sheet1.Range("sheet1.SurNameOrOrgName"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.SurNameOrOrgName")) > 0 Then
                 If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.SurNameOrOrgName").Value) Then
                    fmsgbox ("* Last name should not Contain <, > characters in Sheet Income Details" & Chr(13))
                    Sheet1.Range("sheet1.SurNameOrOrgName").Value = ""
                Else
                 Sheet1.Range("sheet1.SurNameOrOrgName").Value = (Sheet1.Range("sheet1.SurNameOrOrgName").Value)
                End If
                End If
            End If
         If Not Application.Intersect(Sheet1.Range("sheet1.ResidenceNo"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.ResidenceNo")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.ResidenceNo").Value) Then
                         fmsgbox ("* Flat/Door/Block No. under primary address should not Contain <, > characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.ResidenceNo").Value = ""
                    Else
                    Sheet1.Range("sheet1.ResidenceNo").Value = (Sheet1.Range("sheet1.ResidenceNo").Value)
                    End If
                End If
            End If

  If Not Application.Intersect(Sheet1.Range("sheet1.NoticeNo"), Sheet1.Range(Target.Address)) Is Nothing Then
                    If Len(Sheet1.Range("sheet1.NoticeNo")) > 0 Then
                        If Not checkfieldspecialcharacter2(Sheet1.Range("sheet1.NoticeNo").Value) Then
                            fmsgbox "* Notice Number in Sheet Income Details. Only "" / "" special character is allowed in Sheet Income Details." & Chr(13)
                            Sheet1.Range("sheet1.NoticeNo").Value = ""
                        End If
                    End If
                End If
  If Not Application.Intersect(Sheet1.Range("sheet1.DOB"), Sheet1.Range(Target.Address)) Is Nothing Then
                If (Sheet1.Range("sheet1.DOB") <> "") And calculateAge(Trim(Range("sheet1.DOB").Value)) < 18 Then
                        If Sheet3.Range("Ver.pan").Value <> "" Then
                        Sheet3.Range("Ver.pan").Value = ""
                        End If
                        If Sheet3.Range("Ver.capacity").Value = "Self" Then
                        
                        Sheet3.Range("Ver.capacity").Value = ""
                        End If
                End If
             sPassword = EfilingCommon.getmsgstate
            Sheet1.Unprotect Password:=sPassword
                If Trim(Range("sheet1.DOB").Value) <> "00/00/0000" Then
                    Range("sheet1.DOB").Value = Replace(Range("sheet1.DOB").Value, "-", "/")
                    If Not CheckDOB23_24(Trim(Range("sheet1.DOB").Value)) Then
                        Sheet1.Range("sheet1.DOB").Select
                        Sheet1.Range("sheet1.DOB").Value = ""
                        Range("IncD.Section80TTB").MergeArea.Locked = False
                        Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)
                        Range("IncD.Section80TTA").MergeArea.Locked = False
                        Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
                    Else
                        If Sheet5.Range("BacValue").Value = 2 Then
                            Sheet1.Range("sheet1.DOB").Value = AssesseeDob
                            Range("IncD.Section80TTB").MergeArea.Locked = False
                            Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)
                            'Malli----
                            Range("IncD.Section80TTA").MergeArea.Locked = False
                            Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
                            '------------
                        End If
                    End If
                    AnyOtherDeductions80CCH Target
                End If

                If (Range("sheet1.DOB").Value) <> "" Then
                If calculateAge(Trim(Range("sheet1.DOB").Value)) >= 60 Then
                 Sheet1.Unprotect Password:=getmsgstate
                 Range("IncD.Section80TTA").Value = ""
                 Sheet1.Unprotect Password:=getmsgstate
                 Range("IncD.Section80TTA").MergeArea.Locked = True
                 Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HD8D8D8)
                 Sheet1.Unprotect Password:=getmsgstate
                 Sheet1.Unprotect Password:=getmsgstate
                 If Sheet5.Range("BacValue").Value = 2 Then
                   Sheet1.Unprotect Password:=getmsgstate
                   Range("IncD.Section80TTB").MergeArea.Locked = False
                   Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)
                 End If
                 Else
                 Sheet1.Unprotect Password:=getmsgstate
                 If Sheet5.Range("BacValue").Value = 2 Then
                 Range("IncD.Section80TTA").MergeArea.Locked = False
                 Sheet1.Unprotect Password:=getmsgstate
                 Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
                 End If
                 Sheet1.Unprotect Password:=getmsgstate
                 Range("IncD.Section80TTB").Value = ""
                 Sheet1.Unprotect Password:=getmsgstate
                 Range("IncD.Section80TTB").MergeArea.Locked = True
                 Sheet1.Unprotect Password:=getmsgstate
                 Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)
                End If
                End If
                End If

If Not Application.Intersect(Sheet1.Range("sheet1.OrigRetFiledDate"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Trim(Range("sheet1.OrigRetFiledDate").Value) <> "00/00/0000" And Range("sheet1.OrigRetFiledDate").Value <> "" Then
            If Not FormatNCheckDate(Trim(Range("sheet1.OrigRetFiledDate").Value)) Then
                fmsgbox "* Date Of filing Orignal Return Is Invalid in Sheet Income Details. Please enter a valid Date in dd/mm/yyyy format."        ', vbOKOnly, mdCommon.vbMessageCaption
                Sheet1.Range("sheet1.OrigRetFiledDate").Select
                Sheet1.Range("sheet1.OrigRetFiledDate").Value = ""
            Else
            
'Ankita_17/12/2025==============
'SIT-109150
Dim cutoff2 As Date
cutoff2 = CDate(Sheet5.Range("DOB_1").Value)
                 If Dformat(Trim(Range("sheet1.OrigRetFiledDate").Value), "yyyy-mm-dd") < Dformat(Trim(Sheet5.Range("DOB_1").Value), "yyyy-mm-dd") Then
                    fmsgbox "* ""Date of filing of original return cannot be prior to " & Dformat1(cutoff2, "dd/mm/yyyy") & " of AY 26-27"""
                    Sheet1.Range("sheet1.OrigRetFiledDate").Select
                    Sheet1.Range("sheet1.OrigRetFiledDate").Value = ""
                End If
            End If
            End If
        End If
                 If Dformat(Trim(Range("sheet1.OrigRetFiledDate").Value), "yyyy-mm-dd") > Dformat(Trim(Sheet3.Range("Ver.Date").Value), "yyyy-mm-dd") Then
                 fmsgbox "* Date of filing of original return can not be more than system date in Income Details"
                        Range("sheet1.OrigRetFiledDate").Select
                        Range("sheet1.OrigRetFiledDate").Value = ""
                End If
            End If
        If Not Application.Intersect(Sheet1.Range("sheet1.NoticeDate"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Trim(Range("sheet1.NoticeDate").Value) <> "00/00/0000" Then
                    If Not CheckNoticeDateBefore23_24(Trim(Range("sheet1.NoticeDate").Value)) Then
                        Range("sheet1.NoticeDate").Value = ""
                    Else
                    If Range("sheet1.NoticeDate").Value <> "" Then
                        Range("sheet1.NoticeDate").Value = NoticeDate
                    End If
                End If
            End If
            If Dformat(Trim(Range("sheet1.NoticeDate").Value), "yyyy-mm-dd") > Dformat(Trim(Sheet3.Range("Ver.Date").Value), "yyyy-mm-dd") Then
                fmsgbox "* Date of such Notice/Order can not be more than system date in Income Details"
                Range("sheet1.NoticeDate").Select
                Range("sheet1.NoticeDate").Value = ""
            End If
    End If
   If Not Application.Intersect(Sheet1.Range("sheet1.NoticeDate"), Sheet1.Range(Target.Address)) Is Nothing Then
        If Trim(Range("sheet1.NoticeDate").Value) <> "00/00/0000" Then
            If Not FormatNCheckDate(Range("sheet1.NoticeDate").Value) Then
                'SIT-66367 'Malli
                fmsgbox "*Please enter a date in Date of notice/order field in dd/mm/yyyy format"
                Sheet1.Range("sheet1.NoticeDate").Select
                Sheet1.Range("sheet1.NoticeDate").Value = ""
            Else
                If Range("sheet1.NoticeDate").Value <> "" Then
                    If ChkMaxDate(Range("sheet1.NoticeDate").Value, "") Then
                        fmsgbox "*""Date of notice/Order cannot be prior to 01/04/2025"""
                        Sheet1.Range("sheet1.NoticeDate").Select
                        Sheet1.Range("sheet1.NoticeDate").Value = ""
                    End If
                End If
            End If
            Dim result
            result = ""
        End If
    End If
       
        If Not Application.Intersect(Sheet1.Range("sheet1.STDcode"), Sheet1.Range(Target.Address)) Is Nothing Then
                 If Not (ValidateSTDcode()) Then
                     Sheet1.Range("sheet1.STDcode").Value = ""
                     Sheet1.Range("sheet1.STDcode").Select
                 End If
                 stDcode = Sheet1.Range("sheet1.STDcode").Value
             End If
             
'Ankita_30/01/2026=====
         If Not Application.Intersect(Sheet1.Range("sheet1.MiddleName"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("sheet1.FirstName"), Sheet1.Range(Target.Address)) Is Nothing Or _
             Not Application.Intersect(Sheet1.Range("sheet1.SurNameOrOrgName"), Sheet1.Range(Target.Address)) Is Nothing Then
                sPassword = EfilingCommon.getmsgstate
                Sheet3.Unprotect Password:=sPassword
                Sheet3.Range("Ver.AssesseeVerName").Value = Trim(UCase(Sheet1.Range("sheet1.FirstName").Value + " " + Sheet1.Range("sheet1.MiddleName").Value + " " + Sheet1.Range("sheet1.SurNameOrOrgName").Value))
                If Sheet1.Range("sheet1.MiddleName").Value = "" Then
                Sheet3.Range("Ver.AssesseeVerName").Value = Trim(UCase(Sheet1.Range("sheet1.FirstName").Value + "" + Sheet1.Range("sheet1.MiddleName").Value + " " + Sheet1.Range("sheet1.SurNameOrOrgName").Value))
                End If
                Sheet3.Protect Password:=sPassword
                GoTo endline
          End If
  'Ankita_30/01/2026===
  If Not Application.Intersect(Sheet1.Range("IncD.Section80C"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80CCC"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80CCD_SE"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80CCD1B_SE"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80CCD"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80DD"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80DDB"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80E"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80EE"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80EEA"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.Section80EEB"), Sheet1.Range(Target.Address)) Is Nothing Or _
             Not Application.Intersect(Sheet1.Range("IncD.Section80GG"), Sheet1.Range(Target.Address)) Is Nothing Then
   End If
'Ankita_16/04/2026=========
If Range("IncD.Section80CCD_SE").Value > 0 Or Range("IncD.Section80CCD1B_SE").Value > 0 Then
Sheet1.Unprotect Password:=sPassword
Sheet1.Range("pran_new").Locked = False
Sheet1.Range("pran_new").Interior.Color = (&HCCFFCC)
Else
Sheet1.Unprotect Password:=sPassword
Sheet1.Range("pran_new").ClearContents
Sheet1.Range("pran_new").Interior.Color = (&HD8D8D8)
Sheet1.Range("pran_new").Locked = True
End If
'Ankita_21/04/2026
If Not Application.Intersect(Sheet1.Range("Type_80CCC"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Target.Value = "PRAN" Then
        DPMAddComment_PRAN Target
    Else
        DPMADeleteComment_PRAN Target
    End If
End If
If Not Intersect(Range("sheet1.DOB"), Range(Target.Address)) Is Nothing Then
    If isdropdownblank(Sheet1.Range("IncD.Section80D_Calc").Value) Then
        Sheet1.Range("IncD.Section80D_Calc").Value = Sheet1.Range("IncD.Section80DAutop").Value
    End If
End If
If Not Application.Intersect(Sheet1.Range("IncD.Deduction16"), Sheet1.Range(Target.Address)) Is Nothing Or _
    Not Application.Intersect(Sheet1.Range("IncD.Allowances"), Sheet1.Range(Target.Address)) Is Nothing Then

If (Range("IncD.Deduction16").Value) > 5000 Or (Range("IncD.Deduction16").Value) > (Range("IncD.Allowances").Value) / 5 Then
fmsgbox "* Deduction of Entertainment allowance u/s 16(ii) can not exceed 1/5th of Salary as per section 17(1) or Rs. 5,000 whichever is lower in Income Details"
Range("IncD.Deduction16").Value = ""
Range("IncD.Deduction16").Select
End If
End If
If Not Application.Intersect(Sheet1.Range("IncD.Section89"), Sheet1.Range(Target.Address)) Is Nothing Then
If Range("IncD.Section89").Value > 0 Then
fmsgbox "* Please ensure to submit Form 10E."
End If
If Not Application.Intersect(Sheet1.Range("IncD.Section89"), Sheet1.Range(Target.Address)) Is Nothing Then
If Range("IncD.Section89").Value = "" Then  'Ankita_06/03/2026====
Range("IncD.Section89").Value = 0
End If
End If
End If
If Not Application.Intersect(Sheet1.Range("IncD.Section89A"), Sheet1.Range(Target.Address)) Is Nothing Then
End If
If Not Application.Intersect(Sheet1.Range("IncD.IntrstPayUs234A"), Sheet1.Range(Target.Address)) Is Nothing Or _
            Not Application.Intersect(Sheet1.Range("IncD.IntrstPayUs234B"), Sheet1.Range(Target.Address)) Is Nothing Or _
             Not Application.Intersect(Sheet1.Range("IncD.IntrstPayUs234C"), Sheet1.Range(Target.Address)) Is Nothing Or _
             Not Application.Intersect(Sheet1.Range("IncD.IntrstPayUs234F"), Sheet1.Range(Target.Address)) Is Nothing Then
Sheet1.Range("taxcheck").Value = "N"
Else
Sheet1.Range("taxcheck").Value = ""
End If
If Not Application.Intersect(Sheet1.Range("sheet1.ResidenceNo"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.ResidenceName"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.RoadOrStreet"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.LocalityOrArea"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.CityOrTownOrDistrict"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.StateCode1"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("HASZIP"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.Country"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.ZipCode"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("sheet1.PinCode"), Sheet1.Range(Target.Address)) Is Nothing Or _
Not Application.Intersect(Sheet1.Range("Secondary_Address"), Sheet1.Range(Target.Address)) Is Nothing Then
Call mIncmDtls.AutoPopulateSecondaryAddress
End If
endline:
sPassword = EfilingCommon.getmsgstate
Sheet1.Protect Password:=sPassword
Application.EnableEvents = True
Application.ScreenUpdating = True
 Sheet1.Unprotect Password:=getmsgstate
 Sheet1.Protect Password:=sPassword
End Sub

Function CountDropdown(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
    Dim currCnt  As Long
    Dim rangeValue As Variant
    Dim i As Long
    currCnt1 = 0
     currCnt2 = 0
      currCnt3 = 0
      currCnt4 = 0
      currCnt5 = 0
      currCnt6 = 0
      currCnt7 = 0

      currCnt8 = 0
      currCnt9 = 0
      currCnt10 = 0
      currCnt11 = 0
      currCnt12 = 0
      currCnt13 = 0
      currCnt14 = 0
      currCnt15 = 0
      currCnt16 = 0
      currCnt17 = 0


    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).Value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).Value
        End If
      Next

    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If Trim(rangeValue(i)) = Trim(sDropdownValue) And (IIf(Trim(rangeValue(i)) = "", "(Select)", Trim(rangeValue(i))) <> "(Select)") Then

       If Mid(Trim(sDropdownValue), 1, 4) = "Agri" Then
       currCnt1 = currCnt1 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 12) = "Sec 10(10BC)" Then
       currCnt2 = currCnt2 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 11) = "Sec 10(10D)" Then
       currCnt3 = currCnt3 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(19)" Then
       currCnt4 = currCnt4 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(26)" Then
       currCnt5 = currCnt5 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 13) = "Sec 10(26AAA)" Then
       currCnt6 = currCnt6 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(34)" Then
       currCnt7 = currCnt7 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(11)" Then
       currCnt8 = currCnt8 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(13)" Then
       currCnt9 = currCnt9 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(16)" Then
       currCnt10 = currCnt10 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(17)" Then
       currCnt11 = currCnt11 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 11) = "Sec 10(17A)" Then
       currCnt12 = currCnt12 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(18)" Then
       currCnt13 = currCnt13 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 10) = "Sec 10(12)" Then
       currCnt14 = currCnt14 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 11) = "Sec 10(12C)" Then
       currCnt17 = currCnt17 + 1
       End If
       
       If Mid(Trim(sDropdownValue), 1, 7) = "Defense" Then
       currCnt15 = currCnt15 + 1
       End If
        
       If Mid(Trim(sDropdownValue), 1, 17) = "Any exempt income" Then
       currCnt16 = currCnt16 + 1
       End If



       End If
      Next
    End If
    CountDropdown = currCnt1
End Function
Function CountDropdown1(ByVal sDropdownValue As String, rangecells As Range, rangeCount As Long) As Long
    Dim currCnt  As Long
    Dim rangeValue As Variant
    Dim i As Long

     currCnt22 = 0
      currCnt33 = 0
       currCnt44 = 0
        currCnt66 = 0
        currCnt67 = 0
        currCnt68 = 0
        currCnt69 = 0
        currCnt70 = 0

    ReDim rangeValue(rangeCount)
      For i = 1 To rangeCount
        If Not rangecells.item(i).Value = "(Select)" Then
            rangeValue(i) = rangecells.item(i).Value
        End If
      Next

    If Not IsEmpty(rangeValue) And UBound(rangeValue) > 0 Then
      For i = 1 To UBound(rangeValue)
       If Trim(rangeValue(i)) = Trim(sDropdownValue) And (IIf(Trim(rangeValue(i)) = "", "(Select)", Trim(rangeValue(i))) <> "(Select)") Then

       If Mid(Trim(sDropdownValue), 1, 15) = "Interest from S" Then
       currCnt22 = currCnt22 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 15) = "Interest from D" Then
       currCnt33 = currCnt33 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 15) = "Interest from I" Then
       currCnt44 = currCnt44 + 1
       End If
               
       If Mid(Trim(sDropdownValue), 1, 1) = "F" Then
       currCnt66 = currCnt66 + 1
       End If
       
       'Uncommented by Ankita on 10/01/2025__SIT-85413
       If Mid(Trim(sDropdownValue), 1, 112) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(11)" Then
       currCnt67 = currCnt67 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 113) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(11)" Then
       currCnt68 = currCnt68 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 112) = "Interest accrued on contributions to provident fund to the extent taxable as per first proviso to section 10(12)" Then
       currCnt69 = currCnt69 + 1
       End If

       If Mid(Trim(sDropdownValue), 1, 113) = "Interest accrued on contributions to provident fund to the extent taxable as per second proviso to section 10(12)" Then
       currCnt70 = currCnt70 + 1
       End If
        
       End If
      Next
    End If
    
    CountDropdown1 = currCnt22
End Function
 Sub pp()
Application.EnableEvents = True
 End Sub
Sub LockUnlock80DU()
On Error Resume Next
Application.EnableEvents = False
Sheet14.Unprotect Password:=getmsgstate
If Sheet8.Range("oldbacValue").Value = 2 Then
            Sheet14.Range("NatureDisability_80U").ClearContents
            Sheet14.Range("NatureDisability_80U").Locked = True
            Sheet14.Range("NatureDisability_80U").Interior.Color = "&HD8D8D8"
            Sheet14.Range("DateFiling10IA_80U").ClearContents
            Sheet14.Range("DateFiling10IA_80U").Locked = True
            Sheet14.Range("DateFiling10IA_80U").Interior.Color = "&HD8D8D8"
            Sheet14.Range("DateFiling10IA_80U").Value = ""
            Sheet14.Range("AckNumber10IA_80U").ClearContents
            Sheet14.Range("AckNumber10IA_80U").Locked = True
            Sheet14.Range("AckNumber10IA_80U").Interior.Color = "&HD8D8D8"
            Sheet14.Range("AckNumber10IA_80U").Value = ""
            Sheet14.Range("UDIDNum_80U").ClearContents
            Sheet14.Range("UDIDNum_80U").Locked = True
            Sheet14.Range("UDIDNum_80U").Interior.Color = "&HD8D8D8"
            Sheet14.Range("UDIDNum_80U").Value = ""
    Else
            Sheet14.Range("NatureDisability_80U").Locked = False
            Sheet14.Range("NatureDisability_80U").Interior.Color = "&HCCFFCC"
            Sheet14.Range("DateFiling10IA_80U").Locked = False
            Sheet14.Range("DateFiling10IA_80U").Interior.Color = "&HCCFFCC"
            Sheet14.Range("AckNumber10IA_80U").Locked = False
            Sheet14.Range("AckNumber10IA_80U").Interior.Color = "&HCCFFCC"
            Sheet14.Range("UDIDNum_80U").Locked = False
            Sheet14.Range("UDIDNum_80U").Interior.Color = "&HCCFFCC"
    End If
Sheet14.Protect Password:=getmsgstate
Application.EnableEvents = True
End Sub


Sub PartAGenChangeSub1(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate
Dim PinCode4 As Range
        If Not Application.Intersect(Sheet1.Range("sheet1.StateCode2"), Sheet1.Range(Target.Address)) Is Nothing Then
            If Sheet1.Range("sheet1.PinCode1").Value <> "" Then
                Dim state4 As Variant
                state4 = Application.VLookup(Sheet1.Range("sheet1.PinCode1").Value, Sheet5.Range("All_Pincode_V"), 2, False)
                If Not (state4 = UCase(Mid(Sheet1.Range("sheet1.StateCode2").Value, 4)) Or StateMatchesPin1(Target.Value)) Then
                        Sheet1.Range("sheet1.PinCode1").Value = ""
                End If
            End If
         End If
        If Not Application.Intersect(Sheet1.Range("sheet1.PinCode1"), Sheet1.Range(Target.Address)) Is Nothing Then
         Dim DupFlag1 As Boolean
         DupFlag1 = False
                 For Each PinCode4 In Sheet5.Range("Duplicate_Pincode_List")
                    If Sheet1.Range("sheet1.PinCode1").Value = PinCode4 Then
                        DupFlag1 = True
                        Exit For
                    End If
                 Next PinCode4
                If DupFlag1 = False Then
                    Dim StateName2
                    Dim StateCode2
                    StateName2 = Application.VLookup(Sheet1.Range("sheet1.PinCode1").Value, Sheet5.Range("All_Pincode_V"), 2, False)
                    If IsError(StateName2) Then
                       Application.EnableEvents = False
                     Application.EnableEvents = True
                    Else
                        StateCode2 = Application.VLookup(StateName2, Sheet5.Range("CS:CT"), 2, False)
                        Sheet1.Range("sheet1.StateCode2").Value = StateCode2
                        Sheet1.Range("sheet1.CountryCode1").Value = "91-INDIA"
                    End If
                 End If
                 If DupFlag1 = True Then
                 Application.EnableEvents = False
                    If Sheet1.Range("sheet1.StateCode2").Value <> "" Then
                        If Not StateMatchesPin(Sheet1.Range("sheet1.StateCode2").Value) Then
                            Sheet1.Range("sheet1.StateCode2").Value = "(Select)"
                        End If
                    End If
                 End If
         End If

Sheet1.Protect Password:=getmsgstate
endline:
End Sub


Sub PartAGenChangeSub2(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate
'Ankita_14/01/2026_V0.2========
         If Not Application.Intersect(Sheet1.Range("sheet1.ResidenceNo1"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.ResidenceNo1")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.ResidenceNo1").Value) Then
                         fmsgbox ("* Flat/Door/Block No. under updated address should not Contain <, > characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.ResidenceNo1").Value = ""
                    Else
                    Sheet1.Range("sheet1.ResidenceNo1").Value = (Sheet1.Range("sheet1.ResidenceNo1").Value)
                    End If
                End If
            End If
         If Not Application.Intersect(Sheet1.Range("sheet1.ResidenceName"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.ResidenceName")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.ResidenceName").Value) Then
                         fmsgbox ("* Name of Premises/Building/Village under primary address should not Contain <, > characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.ResidenceName").Value = ""
                    Else
                    Sheet1.Range("sheet1.ResidenceName").Value = (Sheet1.Range("sheet1.ResidenceName").Value)
                    End If
                End If
            End If
'Ankita_14/01/2026_V0.2========
         If Not Application.Intersect(Sheet1.Range("sheet1.ResidenceName1"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.ResidenceName1")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.ResidenceName1").Value) Then
                         fmsgbox ("* Name of Premises/Building/Village under updated address should not Contain <, > characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.ResidenceName1").Value = ""
                    Else
                    Sheet1.Range("sheet1.ResidenceName1").Value = (Sheet1.Range("sheet1.ResidenceName1").Value)
                    End If
                End If
            End If
         If Not Application.Intersect(Sheet1.Range("sheet1.RoadOrStreet"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.RoadOrStreet")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.RoadOrStreet").Value) Then
                         fmsgbox ("* Road/Street/Post Office under primary address should not Contain <, >  characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.RoadOrStreet").Value = ""
                    Else
                    Sheet1.Range("sheet1.RoadOrStreet").Value = (Sheet1.Range("sheet1.RoadOrStreet").Value)
                    End If
                End If
            End If
'Ankita_14/01/2026_V0.2========
         If Not Application.Intersect(Sheet1.Range("sheet1.RoadOrStreet1"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.RoadOrStreet1")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.RoadOrStreet1").Value) Then
                         fmsgbox ("* Road/Street/Post Office under updated address should not Contain <, >  characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.RoadOrStreet1").Value = ""
                    Else
                    Sheet1.Range("sheet1.RoadOrStreet1").Value = (Sheet1.Range("sheet1.RoadOrStreet1").Value)
                    End If
                End If
            End If
         If Not Application.Intersect(Sheet1.Range("sheet1.LocalityOrArea"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.LocalityOrArea")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.LocalityOrArea").Value) Then
                         fmsgbox ("* Area/Locality under primary address should not Contain <, >  characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.LocalityOrArea").Value = ""
                    Else
                    Sheet1.Range("sheet1.LocalityOrArea").Value = (Sheet1.Range("sheet1.LocalityOrArea").Value)
                    End If
                End If
            End If
'Ankita_14/01/2026_V0.2========
         If Not Application.Intersect(Sheet1.Range("sheet1.LocalityOrArea1"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.LocalityOrArea1")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.LocalityOrArea1").Value) Then
                         fmsgbox ("* Area/Locality under updated address should not Contain <, >  characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.LocalityOrArea1").Value = ""
                    Else
                    Sheet1.Range("sheet1.LocalityOrArea1").Value = (Sheet1.Range("sheet1.LocalityOrArea1").Value)
                    End If
                End If
            End If
         If Not Application.Intersect(Sheet1.Range("Sheet1.Aadhaar"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("Sheet1.Aadhaar").Value) >= 12 Then
                    If IsNumeric(Sheet1.Range("Sheet1.Aadhaar").Value) Then
                    If Not checkfieldSuperSpecialcharacterDot(Sheet1.Range("Sheet1.Aadhaar").Value) Then
                      fmsgbox "* Please enter the Aadhaar number in valid format"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range("Sheet1.Aadhaar").Value = ""
                    Else
                       Sheet1.Unprotect Password:=getmsgstate
                    End If
                       If Sheet1.Range("Sheet1.Aadhaar").Value <> "000000000000" Then
                         Sheet1.Unprotect Password:=getmsgstate
                         Sheet1.Range("Sheet1.Aadhaar").Value = Sheet1.Range("Sheet1.Aadhaar").Value
                       Else
                         fmsgbox "* Please enter the Aadhaar number in valid format"
                         Sheet1.Unprotect Password:=getmsgstate
                         Sheet1.Range("Sheet1.Aadhaar").Value = ""
                       End If
                       If Sheet1.Range("Sheet1.Aadhaar").Value < 0 Then
                            fmsgbox "* Please enter the Aadhaar number in valid format"
                            Sheet1.Unprotect Password:=getmsgstate
                            Sheet1.Range("Sheet1.Aadhaar").Value = ""
                        End If
                       If Sheet1.Range("Sheet1.Aadhaar").Value <> "111111111111" Then
                         Sheet1.Unprotect Password:=getmsgstate
                         Sheet1.Range("Sheet1.Aadhaar").Value = Sheet1.Range("Sheet1.Aadhaar").Value
                       Else
                         fmsgbox "* Please enter the Aadhaar number in valid format"
                         Sheet1.Unprotect Password:=getmsgstate
                         Sheet1.Range("Sheet1.Aadhaar").Value = ""
                       End If
                    Else
                       fmsgbox "* Please enter the Aadhaar number in valid format"
                      Sheet1.Unprotect Password:=getmsgstate
                      Sheet1.Range("Sheet1.Aadhaar").Value = ""
                    End If
                Else
                    Sheet1.Unprotect Password:=getmsgstate
                    Sheet1.Range("Sheet1.Aadhaar").Value = ""
                End If
            End If
                 Sheet1.Unprotect Password:=getmsgstate
  If Not Application.Intersect(Sheet1.Range("sheet1.ReceiptNo"), Sheet1.Range(Target.Address)) Is Nothing Then
            Dim recieptno As Variant
            recieptno = Sheet1.Range("sheet1.ReceiptNo").Value
            If Not chkNumeric(recieptno) Then
            fmsgbox "* Invalid Receipt Number. Please retry in Sheet Income Details"
            Sheet1.Range("sheet1.ReceiptNo").Value = ""
            Sheet1.Range("sheet1.ReceiptNo").Select
            End If
            If Sheet1.Range("sheet1.ReceiptNo").Value <> "000000000000000" Then
                         Sheet1.Range("sheet1.ReceiptNo").Value = Sheet1.Range("sheet1.ReceiptNo").Value
            Else
                         fmsgbox "* Invalid Receipt Number. Please retry in Sheet Income Details"
                         Sheet1.Range("sheet1.ReceiptNo").Value = ""
                         Sheet1.Range("sheet1.ReceiptNo").Select
            End If
            End If
  If Not Application.Intersect(Sheet1.Range("sheet1.ReceiptNo1"), Sheet1.Range(Target.Address)) Is Nothing Then
            Dim recieptno1 As Variant
            recieptno1 = Sheet1.Range("sheet1.ReceiptNo1").Value
            If Not chkNumeric(recieptno1) Then
            fmsgbox "* Invalid Receipt Number. Please retry in Sheet Income Details"
            Sheet1.Range("sheet1.ReceiptNo1").Value = ""
            Sheet1.Range("sheet1.ReceiptNo1").Select
            End If

            If Sheet1.Range("sheet1.ReceiptNo1").Value <> "000000000000000" Then
                         Sheet1.Range("sheet1.ReceiptNo1").Value = Sheet1.Range("sheet1.ReceiptNo1").Value
            Else
                         fmsgbox "* Invalid Receipt Number. Please retry in Sheet Income Details"
                         Sheet1.Range("sheet1.ReceiptNo1").Value = ""
                         Sheet1.Range("sheet1.ReceiptNo1").Select
            End If
            End If
  If Not Application.Intersect(Sheet1.Range("sheet1.CityOrTownOrDistrict"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.CityOrTownOrDistrict")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.CityOrTownOrDistrict").Value) Then
                         fmsgbox ("* Town/City/District should not Contain <, > characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.CityOrTownOrDistrict").Value = ""
                    Else
                    Sheet1.Range("sheet1.CityOrTownOrDistrict").Value = (Sheet1.Range("sheet1.CityOrTownOrDistrict").Value)
                    End If
                End If
            End If
'Ankita_14/01/2026_V0.2===============
  If Not Application.Intersect(Sheet1.Range("sheet1.CityOrTownOrDistrict1"), Sheet1.Range(Target.Address)) Is Nothing Then
                If Len(Sheet1.Range("sheet1.CityOrTownOrDistrict1")) > 0 Then
                    If Not checkfieldSuperSpecialcharactername(Sheet1.Range("sheet1.CityOrTownOrDistrict1").Value) Then
                         fmsgbox ("* Town/City/District should not Contain <, > characters in Sheet Income Details" & Chr(13))
                        Sheet1.Range("sheet1.CityOrTownOrDistrict1").Value = ""
                    Else
                    Sheet1.Range("sheet1.CityOrTownOrDistrict1").Value = (Sheet1.Range("sheet1.CityOrTownOrDistrict1").Value)
                    End If
                End If
            End If
            
            '========================================================

Sheet1.Protect Password:=getmsgstate
endline:
End Sub
Sub PartAGenChangeSub3(ByVal Target As Range)
On Error GoTo endline
    Dim validateType As String

Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg").Value, 1, 1) = "N") Or Range("clauseiv7provisio139iFlg").Value = "(Select)" Then
        Range("clauseiv7provisio139iFlg_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iFlg_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_1").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_1").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_1").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_2").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_2").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_2").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)

        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)

ElseIf (Mid(Range("clauseiv7provisio139iFlg").Value, 1, 1) = "Y") Then
        Range("clauseiv7provisio139iFlg_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_3").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_3").MergeArea.Value = "(Select)"

        Range("clauseiv7provisio139iFlg_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iFlg_4").MergeArea.Interior.Color = (&HCCFFCC)
        Range("clauseiv7provisio139iFlg_4").MergeArea.Value = "(Select)"
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)

        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HCCFFCC)
End If
End If
'Ankita_12/01/2026===========================
If Not Application.Intersect(Sheet1.Range("sheet1.RepAssessee"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
        If Range("sheet1.RepAssessee").Value = "No" Then
            Sheet1.Unprotect Password:=sPassword
            Sheet1.Range("sheet1.NameRepAssessee").Value = ""
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Locked = True
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
            
            Sheet1.Range("sheet1.EmailRepAssessee").Value = ""
            Sheet1.Range("sheet1.EmailRepAssessee").MergeArea.Locked = True
            Sheet1.Range("sheet1.EmailRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Range("sheet1.CountryCodeRepAssessee").Value = ""
            Sheet1.Range("sheet1.CountryCodeRepAssessee").MergeArea.Locked = True
            Sheet1.Range("sheet1.CountryCodeRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Range("sheet1.ContactRepAssessee").Value = ""
            Sheet1.Range("sheet1.ContactRepAssessee").MergeArea.Locked = True
            Sheet1.Range("sheet1.ContactRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)

            Sheet1.Range("sheet1.AadhaarRepAssessee").Value = ""
            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Locked = True
            Sheet1.Range("sheet1.AadhaarRepAssessee").MergeArea.Interior.Color = (&HD8D8D8)
        Else
            Sheet1.Range("sheet1.NameRepAssessee").Value = ""
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Locked = False
            Sheet1.Range("sheet1.NameRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Range("sheet1.EmailRepAssessee").Value = ""
            Sheet1.Range("sheet1.EmailRepAssessee").MergeArea.Locked = False
            Sheet1.Range("sheet1.EmailRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Range("sheet1.CountryCodeRepAssessee").Value = ""
            Sheet1.Range("sheet1.CountryCodeRepAssessee").MergeArea.Locked = False
            Sheet1.Range("sheet1.CountryCodeRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)

            Sheet1.Range("sheet1.ContactRepAssessee").Value = ""
            Sheet1.Range("sheet1.ContactRepAssessee").MergeArea.Locked = False
            Sheet1.Range("sheet1.ContactRepAssessee").MergeArea.Interior.Color = (&HCCFFCC)

        End If
        Sheet1.Protect Password:=getmsgstate

    End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_3"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_3").Value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_3").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_3").Value, 1, 1) = "Y") Then
        Range("clauseiv7provisio139iAmount_3").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_3").MergeArea.Interior.Color = (&HCCFFCC)
End If
Sheet1.Protect Password:=getmsgstate
End If
If Not Application.Intersect(Sheet1.Range("clauseiv7provisio139iFlg_4"), Sheet1.Range(Target.Address)) Is Nothing Then
   Sheet1.Unprotect Password:=getmsgstate
If (Mid(Range("clauseiv7provisio139iFlg_4").Value, 1, 1) = "N") Then
        Range("clauseiv7provisio139iAmount_4").MergeArea.ClearContents
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = True
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HD8D8D8)
ElseIf (Mid(Range("clauseiv7provisio139iFlg_4").Value, 1, 1) = "Y") Then
        Range("clauseiv7provisio139iAmount_4").MergeArea.Locked = False
        Range("clauseiv7provisio139iAmount_4").MergeArea.Interior.Color = (&HCCFFCC)
End If

Sheet1.Protect Password:=getmsgstate
End If

Sheet1.Protect Password:=getmsgstate


endline:
End Sub



 'Added by Ankita on 05/03/2026

Sub PartAGenChangeSub4(ByVal Target As Range)
On Error GoTo endline
Dim validateType As String
Application.EnableEvents = False
Sheet1.Unprotect Password:=getmsgstate

If Not Application.Intersect(Sheet1.Range("Others.NOI"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Sheet5.Range("bacValue").Value = 1 Then
                Sheet1.Unprotect Password:=getmsgstate
                Application.EnableEvents = False
                
                Dim Sub_Cat As Variant
                Sub_Cat = Target.Address
                Sub_Cat = Replace(Sub_Cat, "$", "")
                Sub_Cat = Replace(Sub_Cat, "H", "T")
                
            If Target.Value = "Agricultural & related incomes" Then
                Sheet1.Range(Sub_Cat).Value = ""
                Sheet1.Range(Sub_Cat).Validation.Delete
                Sheet1.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Agri_dropdown"
            ElseIf Target.Value = "Compensation/other sums received by government or other approved entities" Then
                Sheet1.Range(Sub_Cat).Value = ""
                Sheet1.Range(Sub_Cat).Validation.Delete
                Sheet1.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Comp_dropdown"
            ElseIf Target.Value = "Income from specified Investments" Then
                Sheet1.Range(Sub_Cat).Value = ""
                Sheet1.Range(Sub_Cat).Validation.Delete
                Sheet1.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Inco_dropdown"
            ElseIf Target.Value = "Specified sums received by armed forces personnel" Then
                Sheet1.Range(Sub_Cat).Value = ""
                Sheet1.Range(Sub_Cat).Validation.Delete
                Sheet1.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Spec__dropdown"
            ElseIf Target.Value = "Sums received by Senior Citizens/Minors " Then
                Sheet1.Range(Sub_Cat).Value = ""
                Sheet1.Range(Sub_Cat).Validation.Delete
                Sheet1.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Seminew_dropdown"
            ElseIf Target.Value = "Sums received by specified Category of Taxpayers" Then
                Sheet1.Range(Sub_Cat).Value = ""
                Sheet1.Range(Sub_Cat).Validation.Delete
                Sheet1.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Taxp_dropdown"
            ElseIf Target.Value = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana" Then
                Sheet1.Range(Sub_Cat).Value = ""
                Sheet1.Range(Sub_Cat).Validation.Delete
                Sheet1.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Yojana_dropdown"
            ElseIf Target.Value = "Other Incomes" Then
                Sheet1.Range(Sub_Cat).Value = ""
                Sheet1.Range(Sub_Cat).Validation.Delete
                Sheet1.Range(Sub_Cat).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown"
            ElseIf Target.Value = "(Select)" Or Target.Value = "" Then
                Sheet1.Range(Sub_Cat).Value = ""
                Sheet1.Range(Sub_Cat).Validation.Delete
            End If
        Else
            Sheet1.Unprotect Password:=getmsgstate
            Application.EnableEvents = False
            
            Dim Sub_Cat1 As Variant
            Sub_Cat1 = Target.Address
            Sub_Cat1 = Replace(Sub_Cat1, "$", "")
            Sub_Cat1 = Replace(Sub_Cat1, "H", "T")
            
        If Target.Value = "Agricultural & related incomes" Then
            Sheet1.Range(Sub_Cat1).Value = ""
            Sheet1.Range(Sub_Cat1).Validation.Delete
            Sheet1.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Agri_dropdown"
        ElseIf Target.Value = "Compensation/other sums received by government or other approved entities" Then
            Sheet1.Range(Sub_Cat1).Value = ""
            Sheet1.Range(Sub_Cat1).Validation.Delete
            Sheet1.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Comp_dropdown"
        ElseIf Target.Value = "Income from specified Investments" Then
            Sheet1.Range(Sub_Cat1).Value = ""
            Sheet1.Range(Sub_Cat1).Validation.Delete
            Sheet1.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Inco_dropdown"
        ElseIf Target.Value = "Specified sums received by armed forces personnel" Then
            Sheet1.Range(Sub_Cat1).Value = ""
            Sheet1.Range(Sub_Cat1).Validation.Delete
            Sheet1.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Spec__dropdown"
        ElseIf Target.Value = "Sums received by Senior Citizens/Minors " Then
            Sheet1.Range(Sub_Cat1).Value = ""
            Sheet1.Range(Sub_Cat1).Validation.Delete
            Sheet1.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Semiold_dropdown"
        ElseIf Target.Value = "Sums received by specified Category of Taxpayers" Then
            Sheet1.Range(Sub_Cat1).Value = ""
            Sheet1.Range(Sub_Cat1).Validation.Delete
            Sheet1.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Taxp_dropdown"
        ElseIf Target.Value = "Sums received from policies/contributions such as LIC/NPS/PF/Sukanya Samriddhi Yojana" Then
            Sheet1.Range(Sub_Cat1).Value = ""
            Sheet1.Range(Sub_Cat1).Validation.Delete
            Sheet1.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Yojana_dropdown"
        ElseIf Target.Value = "Other Incomes" Then
            Sheet1.Range(Sub_Cat1).Value = ""
            Sheet1.Range(Sub_Cat1).Validation.Delete
            Sheet1.Range(Sub_Cat1).Validation.add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=Othe_dropdown"
        ElseIf Target.Value = "(Select)" Or Target.Value = "" Then
            Sheet1.Range(Sub_Cat1).Value = ""
            Sheet1.Range(Sub_Cat1).Validation.Delete
        End If
            Sheet1.Protect Password:=getmsgstate
            Application.EnableEvents = True
    End If
End If

' If Not Application.Intersect(Sheet1.Range("Nature_Others"), Sheet1.Range(Target.Address)) Is Nothing Then
'        If 1 < Application.CountIf(Sheet1.Range("Nature_Others"), Target.Value) Then
'             fmsgbox "* Same drop downs cannot be selected twice"
'             Range(Target.Address).Value = "(Select)"
'        End If
'End If


'Ankita_15/06/2026=======
If Not Application.Intersect(Sheet1.Range("Nature_Others"), Sheet1.Range(Target.Address)) Is Nothing Then
    Dim subcategory As String
    If Target.Value = "Income exempt as per CBDT Circular" Or Target.Value = "Income exempt as per CBDT Notification" Or Target.Value = "Receipts not in the nature of income" Then
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range(Replace(Target.Address, "T", "AH")).MergeArea.ClearContents
        Sheet1.Range(Replace(Target.Address, "T", "AH")).MergeArea.Interior.Color = "&HCCFFCC"
        Sheet1.Range(Replace(Target.Address, "T", "AH")).MergeArea.Locked = False
        Sheet1.Range(Replace(Target.Address, "T", "AH")).MergeArea.Value = ""
        Sheet1.Protect Password:=getmsgstate
        
Else
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Range(Replace(Target.Address, "T", "AH")).MergeArea.ClearContents
        Sheet1.Range(Replace(Target.Address, "T", "AH")).MergeArea.Interior.Color = "&HD8D8D8"
        Sheet1.Range(Replace(Target.Address, "T", "AH")).MergeArea.Locked = True
        Sheet1.Range(Replace(Target.Address, "T", "AH")).MergeArea.Value = ""
        Sheet1.Protect Password:=getmsgstate
        
    End If
        Sheet1.Unprotect Password:=getmsgstate
        Sheet1.Protect Password:=getmsgstate
End If
If Not Application.Intersect(Sheet1.Range("Nature_Others"), Sheet1.Range(Target.Address)) Is Nothing Then
    If Not Target.Value = "(Select)" Or Target.Value = "" Then
    If Target.Value = "Income exempt as per CBDT Circular" Or Target.Value = "Income exempt as per CBDT Notification" Or Target.Value = "Receipts not in the nature of income" Then
         Else
        If WorksheetFunction.CountIf(Sheet1.Range("Nature_Others"), Target.Value) > 1 Then
        Application.EnableEvents = False
            MsgBox "You cannot select twice"
            Target.Value = ""
        Application.EnableEvents = True
        End If
      End If
    End If
End If


Sheet1.Protect Password:=getmsgstate
endline:
End Sub



Sub ttt()
Application.EnableEvents = True

End Sub
Sub DPMAddComment_PRAN(Target As Range)
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
.text text:="PRAN (Permanent Retirement Account Number) is a unique 12-digit number allotted to every subscriber of the National Pension System (NPS)."
.Shape.ScaleHeight 1.26, msoFalse, msoScaleFromTopLeft
.Shape.ScaleWidth 1.87, msoFalse, msoScaleFromTopLeft
End With
End With
Sheet1.Protect Password:=getmsgstate
End Sub

Sub DPMADeleteComment_PRAN(Target As Range)
Dim TargtVal, targetadd As Variant
On Error Resume Next

Sheet1.Unprotect Password:=getmsgstate
    TargtVal = Target.Value
    targetadd = Target.Address
    targetadd = Replace(targetadd, "$", "")
Sheet1.Range(targetadd).Comment.Delete
Sheet1.Protect Password:=getmsgstate
End Sub



