Attribute VB_Name = "mdCalInterst234B"
Option Explicit
Public MonthsAfterDueDate As Long
Dim sPassword As String
Dim Date_9 As Variant
Dim MsgCalcInterest As Variant
Dim IntrstPayUs234A As Double
Dim delayedInMonths As Double
Dim taxbase234A As Double
Public calcInterestPayable234A As Double
Const CONST_IntrstPay234A_Percentage As Double = 1
Dim taxbase234B As Double
Public IntrstPayUs234B As Double
Dim calcInterestPayable234B As Double
Public CONST_IntrstPay234B_Percentage_1, April_1, may_1, June_1 As Variant
Const CONST_IntrstPay234B_Percentage As Double = 1
Const CONST_NET_Limit As Double = 10000
Const CONST_ATP_Limit As Double = 90
Dim IntrstPayUs234C As Double
Dim shortFall As Double
Dim OrigRetFiledDate_1 As Variant


Sub ComputeInterest()
On Error GoTo endline
    
    Dim SysCalculatedNetTaxLiability As Double
    Dim matchedAdvanceTax As Double
    Dim tdsamtused As Double
    Dim tcsamtused As Double
    Dim IncSal, IncHp, IncOS, Inc44AD, Inc44AE As Variant
    Dim IsSeniorCitizen, IsSuperSeniorCitizen, IsNRI As Variant
    
    Application.EnableEvents = False
    sPassword = getmsgstate
    Sheet5.Unprotect sPassword
    
    IntrstPayUs234B = 0
   Sheet1.Unprotect sPassword
        
    Sheet1.Range("IncD.IntrstPayUs234B").Value = IntrstPayUs234B
    
    filingdate1
        
    SysCalculatedNetTaxLiability = Sheet1.Range("IncD.NetTaxLiability").Value

    matchedAdvanceTax = Range("IncD.AdvanceTax").Value
    tdsamtused = Range("IncD.TDS").Value
    tcsamtused = Range("IncD.TCS").Value
    
                                      
'''''''''''''''''''''''''''''''''''''''''''''Interest234B - Start''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    IntrstPayUs234B = Calculate_InterestPayable234B(SysCalculatedNetTaxLiability, _
                                                     matchedAdvanceTax, _
                                                     tdsamtused, _
                                                     tcsamtused, _
                                                     intrst234A, _
                                                     intrst234C, _
                                                     intrst234F)
                                                        
'''''''''''''''''''''''''''''''''''''''''''''Interest234B - Done''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' If (age > 59) And (Mid(Range("sheet1.ResidentialStatus1"), 1, 2) <> "NR") Then
    If (bacage > 59) Then
            IntrstPayUs234B = 0
    End If
  
'Added by Riyaz on 07/01/2026

Dim i As Long
    
    calcInterestPayable234A = 0
    
    If MonthsAfterDueDate = 0 Then
        calcInterestPayable234A = 0
    Else
        For i = 1 To MonthsAfterDueDate
            calcInterestPayable234A = calcInterestPayable234A + Sheet5.Range("GH" & i + 8).Value
        Next i
    End If

    Sheet5.Unprotect sPassword
    Sheet5.Range("section234A").Value = calcInterestPayable234A
    Sheet1.Range("IncD.IntrstPayUs234A").Value = calcInterestPayable234A
    Sheet5.Protect sPassword

'End -------------------

    'Sheet1.Range("IncD.IntrstPayUs234B").Value = Round(intrstPayUs234B)
    
endline:
    Application.EnableEvents = True
    Sheet5.Protect sPassword
    Sheet1.Protect sPassword
End Sub

Function Calculate_InterestPayable234B(SysCalculatedNetTaxLiability As Double, _
                                        matchedAdvanceTax As Double, _
                                        tdsamtused As Double, _
                                        tcsamtused As Double, _
                                        intrstPayUs234A_1 As Double, _
                                        intrstPayUs234C_1 As Double, _
                                        intrstPayUs234F_1 As Double) As Double
  
        Dim balancePrincipal As Double
        Dim balanceInterest As Double
        Dim adjustedPrincipal As Double
        Dim adjustedInterest As Double
        Dim SATPaidAtPeriod As Double
        Dim carryForwardPrinicipal As Double
        Dim carryForwardInterest As Double
        Dim calcIntrst234BOnPeriod As Double
        Dim balancePrincipalR As Double
        Dim balanceInterestR As Double
        Dim calcIntrst234BOnPeriodR As Double
        Dim calcIntrst234BUptoPeriod As Double
        Dim DueDate_1 As String
        Dim i As Long
        Dim AssYear As Long
        Dim dateOfProcessing As String
        Dim yrdop As Long
        Dim mthdop As Long
        Dim j As Long
        Dim SATPaidAtlooper As Double
        Dim looperint, looper As Long
        
        looperint = 0
        looper = 0
        calcInterestPayable234B = 0

        Dim AdditionalTax As Variant
        AdditionalTax = 0
        
        If Sheet202.Range("U_Refund").Value <> "" And (Sheet202.Range("U_TotRefund").Value = "" Or Sheet202.Range("U_TotRefund").Value = 0) Then
            AdditionalTax = Sheet202.Range("U_Refund").Value
        End If
 
        If Sheet202.Range("U_TotRefund") <> "" Then
            AdditionalTax = AdditionalTax + Sheet202.Range("U_TotRefund").Value
        End If
 
        If Mid(Range("sheet1.ReturnFileSec"), 1, 7) = "139(8A)" Then
           SysCalculatedNetTaxLiability = SysCalculatedNetTaxLiability + AdditionalTax
        End If

   If ((SysCalculatedNetTaxLiability - (tdsamtused + tcsamtused)) >= CONST_NET_Limit And matchedAdvanceTax < CONST_ATP_Limit / 100 * (SysCalculatedNetTaxLiability - (tdsamtused + tcsamtused))) Then

        shortFall = WorksheetFunction.Max(0, SysCalculatedNetTaxLiability - (matchedAdvanceTax + tdsamtused + tcsamtused))

                If (shortFall > 100) Then
                    shortFall = WorksheetFunction.Floor((shortFall / 100), 1) * 100
                End If

         If (shortFall = 0) Then
             Calculate_InterestPayable234B = 0
          Else

             balancePrincipal = 0
             balanceInterest = 0
             adjustedPrincipal = 0
             adjustedInterest = 0
             SATPaidAtPeriod = 0
             carryForwardPrinicipal = 0
             carryForwardInterest = 0
             calcIntrst234BOnPeriod = 0
             balancePrincipalR = 0
             balanceInterestR = 0
             calcIntrst234BOnPeriodR = 0
             calcIntrst234BUptoPeriod = 0

        End If

        DueDate_1 = Sheet5.Range("DueDate1").Value

'Commented by Riyaz on 14/10/2025 for testing purpose

'        dateOfProcessing = Sheet5.Range("DateOfFiling").Value

        dateOfProcessing = Sheet3.Range("Ver.Date").Value
        AssYear = 2026 'Ankita_02/01/2026===============
        dateOfProcessing = Dformat(dateOfProcessing, "yyyy-mm-dd")
        
        yrdop = CInt(Mid(dateOfProcessing, 1, 4))
        mthdop = CInt(Mid(dateOfProcessing, 6, 2))

        If (yrdop >= AssYear) Then
            calcIntrst234BUptoPeriod = mthdop - 4 + (yrdop - AssYear) * 12
        End If

        calcIntrst234BUptoPeriod = calcIntrst234BUptoPeriod + 1
        carryForwardPrinicipal = shortFall
    
        Application.EnableEvents = False
        Sheet5.Unprotect Password:=getmsgstate
        Sheet5.Range("Month_Wise_234Bvalues").ClearContents
        Sheet5.Range("Balance_Interest").ClearContents
        Sheet5.Range("IT_Interest").ClearContents
        Sheet5.Protect Password:=getmsgstate
        Application.EnableEvents = True

        Dim month_count_1 As Variant

        carryForwardPrinicipal = WorksheetFunction.RoundDown(carryForwardPrinicipal, -2)

        month_count_1 = (CLng(Mid(Dformat(DueDate_1, "yyyy-mm-dd"), 6, 2)) - 2)


        For i = 1 To calcIntrst234BUptoPeriod

                SATPaidAtPeriod = 0
                SATPaidAtlooper = 0
                balancePrincipal = carryForwardPrinicipal
                calcIntrst234BOnPeriod = Math.Round(CONST_IntrstPay234B_Percentage / 100 * WorksheetFunction.RoundDown(balancePrincipal, -2))

'Changed by Riyaz on 22/04/2026

                If (bacage > 59) Then
                    calcIntrst234BOnPeriod = 0
                    Sheet5.Unprotect Password:=getmsgstate
                    Sheet5.Range("GC" & i + 4).Value = calcIntrst234BOnPeriod
                    Sheet5.Protect Password:=getmsgstate

                Else

                    Sheet5.Unprotect Password:=getmsgstate
                    Sheet5.Range("GC" & i + 4).Value = calcIntrst234BOnPeriod
                    Sheet5.Protect Password:=getmsgstate
                    calcInterestPayable234B = calcInterestPayable234B + calcIntrst234BOnPeriod

                End If

'-------end--
 
'Added by Riyaz on 22/04/2026

        If Mid(Sheet1.Range("sheet1.ReturnFileSec1").Value, 1, 2) = 17 Or Mid(Sheet1.Range("sheet1.ReturnFileSec1").Value, 1, 2) = 18 Or _
        (((Trim(Mid(Range("sheet1.ReturnFileSec1"), 1, 2))) = "21") And (Mid(Range("U_PreviouslyFiledForThisAY"), 1, 1) = "Y")) Then
                
'Added by Riyaz on 24/04/2024 for original return date diffrence month calc
                Dim dateOfProcessing1 As String
                Dim yrdop1 As Long
                Dim mthdop1 As Long
                Dim Revisedmonthdiff As Double
                Dim AssYear1 As Long
                
                dateOfProcessing1 = Sheet1.Range("sheet1.OrigRetFiledDate").Value
        
                AssYear1 = 2026
                dateOfProcessing1 = Dformat(dateOfProcessing1, "yyyy-mm-dd")
        
                yrdop1 = CInt(Mid(dateOfProcessing1, 1, 4))
                mthdop1 = CInt(Mid(dateOfProcessing1, 6, 2))
        
                If (yrdop1 >= AssYear) Then
                    Revisedmonthdiff = mthdop1 - 4 + (yrdop1 - AssYear1) * 12
                End If
        
                    Revisedmonthdiff = Revisedmonthdiff + 1
        
                If MonthsAfterDueDate = 0 Then
                    If i = 1 Then
                        balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
                    Else
                        balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                    End If
                Else
                    If i = 1 Then
                        balanceInterest = Sheet5.Range("GH" & i + 4).Value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
                    ElseIf i = month_count_1 And DueDate_1 = Sheet5.Range("DueDate1").Value Then
                        balanceInterest = Sheet5.Range("GH" & i + 4).Value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                    ElseIf i <= Revisedmonthdiff Then
                        balanceInterest = Sheet5.Range("GH" & i + 4).Value + carryForwardInterest + calcIntrst234BOnPeriod
                    Else
                        balanceInterest = carryForwardInterest + calcIntrst234BOnPeriod
                    End If
                End If
        Else
                If i = 1 Then
                    balanceInterest = Sheet5.Range("GH" & i + 4).Value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234C_1
                ElseIf i = month_count_1 And DueDate_1 = Sheet5.Range("DueDate1").Value Then
                    balanceInterest = Sheet5.Range("GH" & i + 4).Value + carryForwardInterest + calcIntrst234BOnPeriod + intrstPayUs234F_1
                Else
                    balanceInterest = Sheet5.Range("GH" & i + 4).Value + carryForwardInterest + calcIntrst234BOnPeriod
                End If
        End If
'-------end--
                Sheet5.Unprotect Password:=getmsgstate
                Sheet5.Range("GE" & i + 5).Value = balanceInterest
                Sheet5.Protect Password:=getmsgstate

                'Ayush_22/09

                SATPaidAtPeriod = Application.WorksheetFunction.SumIf(Range("FormulaOfSAT1"), "=" & i, Range("TaxP.Amt"))
                Sheet5.Unprotect Password:=getmsgstate
                Sheet5.Range("GF" & i + 5).Value = SATPaidAtPeriod
                Sheet5.Protect Password:=getmsgstate

                adjustedInterest = WorksheetFunction.Min(SATPaidAtPeriod, balanceInterest)
                adjustedPrincipal = WorksheetFunction.Max(0, WorksheetFunction.Min(SATPaidAtPeriod - adjustedInterest, balancePrincipal))
                carryForwardPrinicipal = WorksheetFunction.Max(0, balancePrincipal - adjustedPrincipal)
                carryForwardInterest = WorksheetFunction.Max(0, balanceInterest - adjustedInterest)
        Next

    End If

        Calculate_InterestPayable234B = WorksheetFunction.Round(calcInterestPayable234B, 0)

End Function

 


Function Dformat(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(dt) > 0 Then

Year = Mid(dt, 7, 4)
month = Mid(dt, 4, 2)
day = Mid(dt, 1, 2)
formateddate = Year & "-" & month & "-" & day
Dformat = formateddate

Else
Dformat = ""
End If
End Function

Function Dformat1(dt As Variant, timepass As String) As String
'yyyy-mm-dd'
Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(Sheet5.Range("DOB_extended").Value) > 0 Then
Year = Mid(Sheet5.Range("DOB_extended").Value, 7, 4)
month = Mid(Sheet5.Range("DOB_extended").Value, 4, 2)
day = Mid(Sheet5.Range("DOB_extended").Value, 1, 2)
formateddate = day & "/" & month & "/" & Year
Dformat1 = formateddate
Else
Dformat1 = ""
End If
End Function

'Ankita_29/12/2025=========

Function Dformat2(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(Sheet5.Range("IT_DOD").Value) > 0 Then

Year = Mid(Sheet5.Range("IT_DOD").Value, 7, 4)
month = Mid(Sheet5.Range("IT_DOD").Value, 4, 2)
day = Mid(Sheet5.Range("IT_DOD").Value, 1, 2)
formateddate = day & "-" & month & "-" & Year
Dformat2 = formateddate

Else
Dformat2 = ""
End If
End Function

'added by Chetan C M on 30/12/2025 for AY 2026-27
'start--
Function Dformat3(dt As Variant, timepass As String) As String
'dd-mm-yyyy'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(dt) > 0 Then

day = Mid(dt, 1, 2)
month = Mid(dt, 4, 2)
Year = Mid(dt, 7, 4)

formateddate = day & "/" & month & "/" & Year
Dformat3 = formateddate

Else
Dformat3 = ""
End If
End Function
'--end

'Ankita_07/01/2026=========

Function Dformat4(dt As Variant, timepass As String) As String
'yyyy-mm-dd'
Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(Sheet5.Range("IT_DOD1").Value) > 0 Then
Year = Mid(Sheet5.Range("IT_DOD1").Value, 7, 4)
month = Mid(Sheet5.Range("IT_DOD1").Value, 4, 2)
day = Mid(Sheet5.Range("IT_DOD1").Value, 1, 2)
formateddate = day & "-" & month & "-" & Year
Dformat4 = formateddate
Else
Dformat4 = ""
End If
End Function

 'Ankita_29/12/2025=========

Function Dformat5(dt As Variant, timepass As String) As String
'yyyy-mm-dd'

Dim formateddate As String
Dim day As String
Dim month As String
Dim Year As String
If Len(Sheet5.Range("Date_8A").Value) > 0 Then

Year = Mid(Sheet5.Range("Date_8A").Value, 7, 4)
month = Mid(Sheet5.Range("Date_8A").Value, 4, 2)
day = Mid(Sheet5.Range("Date_8A").Value, 1, 2)
formateddate = day & "-" & month & "-" & Year
Dformat5 = formateddate

Else
Dformat5 = ""
End If
End Function

'==========================


'==========================

Sub filingdate1()
    Dim todaysdate As String
    Dim newfilingdate As String
    Dim originalfilingdate As String
    Dim origrevised As Boolean
    Dim origrevisedstatus As String
    Dim retfilestatus As String
    
    
'    origrevisedstatus = Mid(Sheet1.Range("sheet1.ReturnType1").Value, 1, 1)
    retfilestatus = Mid(Sheet1.Range("sheet1.ReturnFileSec1"), 1, 2)
'    If retfilestatus = "18" Then
'        origrevisedstatus = "R"
'    End If
        
    todaysdate = Range("DateOfProcessing").Value
'    If (origrevisedstatus = "R") Then
        If Not ValidateOrigRetFiledDate_1() Then
            fmsgbox "Please enter Original Return Date "
        Else
            newfilingdate = OrigRetFiledDate_1
        End If
        newfilingdate = Mid(newfilingdate, 9, 2) + "/" + Mid(newfilingdate, 6, 2) + "/" + Mid(newfilingdate, 1, 4)
        Range("DateOfFiling234A").Value = newfilingdate
'    End If
    
    'If (origrevisedstatus <> "R") Then
        If Not ValidateDate_9() Then
            fmsgbox "Please enter Verification Date"
            CloseMsg
        Else
            newfilingdate = Date_9
            If Not ChkMinInclusiveDate(Date_9, todaysdate) Then
                newfilingdate = todaysdate
            End If
            newfilingdate = Mid(newfilingdate, 9, 2) + "/" + Mid(newfilingdate, 6, 2) + "/" + Mid(newfilingdate, 1, 4)
            Range("DateOfFiling").Value = newfilingdate
        End If
    'End If
End Sub

Function ValidateDate_9() As Boolean
    ValidateDate_9 = True
    Date_9 = Sheet3.Range("Ver.Date").Value
    
    If Date_9 = "" Or IsEmpty(Date_9) Then
        MsgCalcInterest = MsgCalcInterest & "* Verification Date in Sheet TaxPaid and Verification is Compulsory" & Chr(13)
        ValidateDate_9 = False
        Exit Function
    End If
    
    If Not CheckDateddmmyyyy(Date_9) Then
        ValidateDate_9 = False
        MsgCalcInterest = MsgCalcInterest & "* Date in Sheet  TaxPaid and Verification must be a valid dd/mm/yyyy format" & Chr(13)
        Exit Function
    Else
        Date_9 = Dformat(Sheet3.Range("Ver.Date"), "yyyy-mm-dd")
    End If
    
'    If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2024, "* Verification date cannot be less than 01/04/2024") Then
'        ValidateDate_9 = False
'        MsgCalcInterest = MsgCalcInterest & "* Date in Verification , Sheet TaxPaid and Verification must not be less than 01/04/2024" & Chr(13)
                    
    'Changed year from 2024 to 2025 by Ankita on 14/12/2023
     If Not CheckDateMinDDMMYYYY(Date_9, 1, 4, 2025, "* Verification date cannot be less than 01/04/2025") Then
        ValidateDate_9 = False
        MsgCalcInterest = MsgCalcInterest & "* Date in Verification , Sheet TaxPaid and Verification must not be less than 01/04/2025" & Chr(13)
 
        Exit Function
    Else
        Date_9 = Dformat(Sheet3.Range("Ver.Date"), "yyyy-mm-dd")
    End If
End Function

Function CheckDateMinDDMMYYYY(ByVal dt As String, ByVal minday As Long, ByVal minmonth As Long, ByVal minyear As Long, ByVal errormsg As String) As Boolean
    CheckDateMinDDMMYYYY = True
        
    If (val(Mid(dt, 1, 4)) < minyear) Then
        CheckDateMinDDMMYYYY = False
        fmsgbox "* Invalid Date, " & errormsg
        
        GoTo exit1
    End If
    
    If (val(Mid(dt, 1, 4)) = minyear) And (val(Mid(dt, 6, 2)) < minmonth) Then
        CheckDateMinDDMMYYYY = False
        fmsgbox "* Invalid Date, " & errormsg
        GoTo exit1
    End If
    If (val(Mid(dt, 1, 4)) = minyear) And (val(Mid(dt, 6, 2)) < minmonth) And (val(Mid(dt, 9, 2)) < minday) Then
        CheckDateMinDDMMYYYY = False
        fmsgbox "* Invalid Date, " & errormsg
        GoTo exit1
    End If
exit1:
End Function

Function MonthDiff(startDate As String, endDate As String) As Long
    Dim NoofDaysinmonth As Long
    Dim endyear, startyear, startmonth, endmonth, startDay As Long
    Dim noofdaysinendmonth As Long
    
    NoofDaysinmonth = 31
    MonthDiff = 0
    startDate = CStr(Dformat(startDate, "yyyy-mm-dd"))
    endDate = CStr(Dformat(endDate, "yyyy-mm-dd"))
    startyear = Mid(startDate, 1, 4)
    endyear = Mid(endDate, 1, 4)
    startmonth = Mid(startDate, 6, 2)
    If (startmonth = 2) Then
        NoofDaysinmonth = 28
    End If
    
    startDay = Mid(startDate, 9, 2)
    If ((startmonth = 4) Or (startmonth = 6) Or (startmonth = 9) Or (startmonth = 11)) Then
        NoofDaysinmonth = 30
    End If
    
    
    noofdaysinendmonth = 31
    endmonth = Mid(endDate, 6, 2)
    If (endmonth = 2) Then
        noofdaysinendmonth = 28
    End If
    
    If ((endmonth = 4) Or (endmonth = 6) Or (endmonth = 9) Or (endmonth = 11)) Then
        noofdaysinendmonth = 30
    End If
    
    MonthDiff = (CInt(endyear) - CInt(startyear)) * 12
    MonthDiff = MonthDiff + (CInt(endmonth) - CInt(startmonth))
    
    If (startDate < endDate) Then
        If (CInt(startDay) < Mid(endDate, 9, 2)) Then
            MonthDiff = MonthDiff + 1
            If CInt(startDay) = 30 And NoofDaysinmonth = 30 Then
                MonthDiff = MonthDiff - 1
            End If
        End If
    End If
End Function

Function MonthDiffPrev(startDate As String, endDate As String) As Long
    Dim NoofDaysinmonth As Long
    Dim endyear, startyear, startmonth, endmonth, startDay As Long
    
    NoofDaysinmonth = 31
    MonthDiffPrev = 0
    startDate = CStr(Dformat(startDate, "yyyy-mm-dd"))
    endDate = CStr(Dformat(endDate, "yyyy-mm-dd"))
    startyear = Mid(startDate, 1, 4)
    endyear = Mid(endDate, 1, 4)
    startmonth = Mid(startDate, 6, 2)
    If (startmonth = 2) Then
         NoofDaysinmonth = 28
    End If
    
    startDay = Mid(startDate, 9, 2)
    If ((startmonth = 4) Or (startmonth = 6) Or (startmonth = 9) Or (startmonth = 11)) Then
        NoofDaysinmonth = 30
    End If
    
    endmonth = Mid(endDate, 6, 2)
    MonthDiffPrev = (CInt(endyear) - CInt(startyear)) * 12
    MonthDiffPrev = MonthDiffPrev + (CInt(endmonth) - CInt(startmonth))
    
    If (startDate < endDate) Then
        If (CInt(startDay) < Mid(endDate, 9, 2)) Then
                MonthDiffPrev = MonthDiffPrev + 1
        End If
    End If
End Function

Function CalculateDelayedInMonths(startDate As Date, endDate As Date) As Double
      Dim noOfMonths As Long
      Dim startDateTotDayMonths As Long
      Dim endDateTotDayMonths As Long
      
      noOfMonths = 0
      If (Year(endDate) >= Year(startDate)) Then
          noOfMonths = Math.Abs(month(endDate) - month(startDate)) + (Year(endDate) - Year(startDate)) * 12
          endDateTotDayMonths = month(endDate) + day(endDate)
          startDateTotDayMonths = month(startDate) + day(startDate)

          If ((endDateTotDayMonths > startDateTotDayMonths) And ((endDateTotDayMonths - startDateTotDayMonths) >= 2) And day(endDate) <> day(startDate)) Then
                noOfMonths = noOfMonths + 1
          End If
      End If
   CalculateDelayedInMonths = noOfMonths
End Function


Function ValidateOrigRetFiledDate_1() As Boolean
    ValidateOrigRetFiledDate_1 = True
    OrigRetFiledDate_1 = Sheet1.Range("sheet1.OrigRetFiledDate").Value
    
    If Not CheckDateddmmyyyy(OrigRetFiledDate_1) Then
        ValidateOrigRetFiledDate_1 = False
        fmsgbox ("* OrigRetFiledDate in Sheet Income Details  must be a valid dd/mm/yyyy format")
        Exit Function
    Else
        OrigRetFiledDate_1 = Dformat(Sheet1.Range("sheet1.OrigRetFiledDate"), "yyyy-mm-dd")
    
        If Len(OrigRetFiledDate_1) > 0 Then
        
            'PAG_E10
            'If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_1, 1, 4, 2023, "Original return filing date cannot be less than 01/04/2023") Then
             '   fmsgbox ("* Original return filing date to be corrected as it should not be less than 01/04/2023")
                 
            'PAG_C10 2024-25 Bindu
            
'            If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_1, 1, 4, 2024, "Original return filing date cannot be less than 01/04/2024") Then
'                fmsgbox ("* Original return filing date to be corrected as it should not be less than 01/04/2024")
               
            'Changed year from 2024 to 2025 by Ankita on 14/12/2023
             If Not CheckDateMinDDMMYYYY(OrigRetFiledDate_1, 1, 4, 2025, "Original return filing date cannot be less than 01/04/2025") Then
                fmsgbox ("* Original return filing date to be corrected as it should not be less than 01/04/2025")
           
                
                ValidateOrigRetFiledDate_1 = False
                Exit Function
            Else
                OrigRetFiledDate_1 = Dformat(Sheet1.Range("sheet1.OrigRetFiledDate"), "yyyy-mm-dd")
            End If
        End If
    End If
End Function

Function checkFirstDateBefore1(firstDate As Variant, secondDate As Variant) As Boolean
On Error Resume Next

If val(Mid(firstDate, 7, 4)) < val(Mid(secondDate, 7, 4)) Then
        checkFirstDateBefore1 = True
ElseIf val(Mid(firstDate, 7, 4)) = val(Mid(secondDate, 7, 4)) Then

        If val(Mid(firstDate, 4, 2)) < val(Mid(secondDate, 4, 2)) Then
        
             checkFirstDateBefore1 = True
             
        ElseIf val(Mid(firstDate, 4, 2)) = val(Mid(secondDate, 4, 2)) Then
        
                If val(Mid(firstDate, 1, 2)) < val(Mid(secondDate, 1, 2)) Then
                
                        checkFirstDateBefore1 = True
                        
                ElseIf val(Mid(firstDate, 1, 2)) = val(Mid(secondDate, 1, 2)) Then
                
                        checkFirstDateBefore1 = True
            
                Else
                      checkFirstDateBefore1 = False
                End If
        Else
            checkFirstDateBefore1 = False
        End If
Else

    checkFirstDateBefore1 = False
End If
  
End Function




