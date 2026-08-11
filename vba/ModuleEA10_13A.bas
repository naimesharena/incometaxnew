Attribute VB_Name = "ModuleEA10_13A"
Option Explicit
Public MsgBox_EA10_13A As String

Public EA_Sch10of13A_2Plus4, endEA10_13A, endEA10_13A_1, endEA10_13A_2, endEA10_13A_3, endEA10_13A_4, endEA10_13A_5, endEA10_13A_6 As Variant  'Ankita_14/05/2025
Public EA_Sch10of13A_PlaceofWrk, EA_Sch10of13A_ActlHRArecivedA, EA_Sch10of13A_ActlRentpaid, EA_Sch10of13A_DetlsofSalpersec17of1, EA_Sch10of13A_BasicSalary, EA_Sch10of13A_DearAllowance, EA_Sch10of13A_Actlrentpaid10persalaryB, EA_Sch10of13A_50Por40Pofsalary, EA_Sch10of13A_ElgiblExmptAllwnce10of13A As Variant



Sub ValidateSheetEA10_13A_Click()
 Dim vbMessgaeCaption As String
 vbMessgaeCaption = "ITR 1: AY: 2026-27"
    ValidateEA10_13A
 fmsgboxoK "Schedule EA 10(13A) is OK"
End Sub
Sub ValidateEA10_13A()
    If Not ValidateEA10_13A_1 Then
        Sheet18.Activate
        fmsgbox (MsgBox_EA10_13A)
        CloseMsg
    End If
End Sub
Function ValidateEA10_13A_1() As Boolean
    ValidateEA10_13A_1 = True
    MsgBox_EA10_13A = "Schedule EA 10(13A) : " & Chr(10)
    
'    Sheet18.Activate
    
    endEA10_13A = 0
    endEA10_13A_1 = 0
    endEA10_13A_2 = 0
    endEA10_13A_3 = 0
    endEA10_13A_4 = 0
    endEA10_13A_5 = 0
    endEA10_13A_6 = 0               'Ankita_14/05/2025_Changed as per clarififcation tracker updated by Bhoomija


            EA_Sch10of13A_PlaceofWrk = Sheet18.Range("Sch10of13A_PlaceofWrk")
            EA_Sch10of13A_ActlHRArecivedA = Sheet18.Range("Sch10of13A_ActlHRArecivedA")
            EA_Sch10of13A_ActlRentpaid = Sheet18.Range("Sch10of13A_ActlRentpaid")
            EA_Sch10of13A_DetlsofSalpersec17of1 = Sheet18.Range("Sch10of13A_DetlsofSalpersec17of1")
            EA_Sch10of13A_BasicSalary = Sheet18.Range("Sch10of13A_BasicSalary")
            EA_Sch10of13A_DearAllowance = Sheet18.Range("Sch10of13A_DearAllowance")
            EA_Sch10of13A_Actlrentpaid10persalaryB = Sheet18.Range("Sch10of13A_Actlrentpaid10persalaryB")
            EA_Sch10of13A_50Por40Pofsalary = Sheet18.Range("Sch10of13A_50Por40Pofsalary")
            EA_Sch10of13A_ElgiblExmptAllwnce10of13A = Sheet18.Range("Sch10of13A_ElgiblExmptAllwnce10of13A")
            
    
     If EA_Sch10of13A_PlaceofWrk <> "" Then
        If Not isdropdownblank(EA_Sch10of13A_PlaceofWrk) And Not UCase(EA_Sch10of13A_PlaceofWrk) = UCase("Select") Then
        endEA10_13A_1 = endEA10_13A_1 + 1
        End If
     End If
     
     If EA_Sch10of13A_ActlHRArecivedA <> "" Then
        If EA_Sch10of13A_ActlHRArecivedA <> "" Then
        endEA10_13A_2 = endEA10_13A_2 + 1
        End If
     End If
      
      If EA_Sch10of13A_ActlRentpaid <> "" Then
        If EA_Sch10of13A_ActlRentpaid <> "" Then
        endEA10_13A_3 = endEA10_13A_3 + 1
        End If
     End If
     
     If EA_Sch10of13A_BasicSalary <> "" Then
        If EA_Sch10of13A_BasicSalary <> "" Then
        endEA10_13A_4 = endEA10_13A_4 + 1
        End If
     End If
     
     If EA_Sch10of13A_DearAllowance <> "" Then
        If EA_Sch10of13A_DearAllowance <> "" Then
        endEA10_13A_5 = endEA10_13A_5 + 1
        End If
     End If
     
          
  'Ankita_14/05/2025_Changed as per clarififcation tracker updated by Bhoomija

     If EA_Sch10of13A_50Por40Pofsalary <> "" Then
        If EA_Sch10of13A_50Por40Pofsalary <> "" Then
        endEA10_13A_6 = endEA10_13A_6 + 1
        End If
     End If
     
      'Ankita_14/05/2025_Changed as per clarififcation tracker updated by Bhoomija
       endEA10_13A = WorksheetFunction.Max(0, endEA10_13A_1, endEA10_13A_2, endEA10_13A_3, endEA10_13A_4, endEA10_13A_5, endEA10_13A_6)
        

 If endEA10_13A > 0 Then
  ' MsgBox_EA10_13A = MsgBox_EA10_13A + "     at Sr. No " & i & "" & Chr(13)
  
        'Place of Work
        If isdropdownblank(EA_Sch10of13A_PlaceofWrk) Then
        MsgBox_EA10_13A = MsgBox_EA10_13A + "* Please select dropdown from ""Place of residence"" in 10(13A) schedule." & Chr(13)
        ValidateEA10_13A_1 = False
        Exit Function
        End If
        
        'Actual HRA received (A)
        If Not chkCompulsory(EA_Sch10of13A_ActlHRArecivedA) Or EA_Sch10of13A_ActlHRArecivedA = 0 Then
            MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Please provide actual HRA received in schedule 10(13A)""" & Chr(13)
            ValidateEA10_13A_1 = False
            Exit Function
        Else
            If Len(EA_Sch10of13A_ActlHRArecivedA) > 14 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Actual HRA received (A) cannot be greater than 14 digits in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
            
            'Ankita_14/05/2025_Changed as per clarififcation tracker updated by Bhoomija
'            If Not IsNumeric(EA_Sch10of13A_ActlHRArecivedA) Then
             If Not IsNumeric(EA_Sch10of13A_ActlHRArecivedA) Or EA_Sch10of13A_ActlHRArecivedA < 0 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Actual HRA received (A) in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
               Exit Function
             End If
         End If
        
        'Actual Rent paid
        If Not chkCompulsory(EA_Sch10of13A_ActlRentpaid) Or EA_Sch10of13A_ActlRentpaid = 0 Then
            MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Please provide actual rent paid in schedule 10(13A)""" & Chr(13)
            ValidateEA10_13A_1 = False
            Exit Function
        Else
            If Len(EA_Sch10of13A_ActlRentpaid) > 14 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Actual Rent paid cannot be greater than 14 digits in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
            If Not IsNumeric(EA_Sch10of13A_ActlRentpaid) Or EA_Sch10of13A_ActlRentpaid < 0 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Actual Rent paid in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
        End If
        
        'Details of Salary as per section 17(1)
        If Not chkCompulsory(EA_Sch10of13A_DetlsofSalpersec17of1) Or EA_Sch10of13A_DetlsofSalpersec17of1 = 0 Then
            MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Please provide salary as per section 17(1) in schedule 10(13A)""" & Chr(13)
            ValidateEA10_13A_1 = False
            Exit Function
        Else
            If Len(EA_Sch10of13A_DetlsofSalpersec17of1) > 14 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Details of Salary as per section 17(1) cannot be greater than 14 digits in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
            If Not IsNumeric(EA_Sch10of13A_DetlsofSalpersec17of1) Or EA_Sch10of13A_DetlsofSalpersec17of1 < 0 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Details of Salary as per section 17(1) in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
        End If
        
        
        ''Note:-Actual HRA received(2) + Details of salary (4) shall not be more than salary as per 17(1) in field i(a) as per part A gen
        If EA_Sch10of13A_ActlHRArecivedA > 0 Or EA_Sch10of13A_DetlsofSalpersec17of1 > 0 Then
        EA_Sch10of13A_2Plus4 = EA_Sch10of13A_ActlHRArecivedA + EA_Sch10of13A_DetlsofSalpersec17of1
                If EA_Sch10of13A_2Plus4 > Sheet1.Range("IncD.Allowances").Value Then
                        MsgBox_EA10_13A = MsgBox_EA10_13A + "* Salary for the purpose of this computation shall not be more than the basic salary and dearness allowance" & Chr(13)
                        ValidateEA10_13A_1 = False
                        Exit Function
                
                End If
        End If
        
        
         'Basic Salary
         If Not chkCompulsory(EA_Sch10of13A_BasicSalary) Or EA_Sch10of13A_BasicSalary = 0 Then
            MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Please provide Basic salary in schedule 10(13A)""" & Chr(13)
            ValidateEA10_13A_1 = False
            Exit Function
        Else
            If Len(EA_Sch10of13A_BasicSalary) > 14 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Basic salary cannot be greater than 14 digits in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
            If Not IsNumeric(EA_Sch10of13A_BasicSalary) Or EA_Sch10of13A_BasicSalary < 0 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Basic salary in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
               Exit Function
            End If
        End If
            
        'Dearness Allowance
        If EA_Sch10of13A_DearAllowance <> "" Then
            If Len(EA_Sch10of13A_DearAllowance) > 14 Then
            MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Dearness Allowance cannot be greater than 14 digits in schedule 10(13A)""" & Chr(13)
            ValidateEA10_13A_1 = False
            Exit Function
            End If
            If Not IsNumeric(EA_Sch10of13A_DearAllowance) Or EA_Sch10of13A_DearAllowance < 0 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Dearness Allowance in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
        End If
            
        'Actual rent paid-10% of salary (B)
        If EA_Sch10of13A_Actlrentpaid10persalaryB <> "" Then
            If Len(EA_Sch10of13A_Actlrentpaid10persalaryB) > 14 Then
            MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Actual rent paid-10% of salary (B) cannot be greater than 14 digits in schedule 10(13A)""" & Chr(13)
            ValidateEA10_13A_1 = False
            Exit Function
            End If
            If Not IsNumeric(EA_Sch10of13A_Actlrentpaid10persalaryB) Or EA_Sch10of13A_Actlrentpaid10persalaryB < 0 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Actual rent paid-10% of salary (B) in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
               Exit Function
            End If
        End If
            
        '50% /40% of salary (C)
    'Ankita_14/05/2025_Changed as per clarififcation tracker updated by Bhoomija

'       If EA_Sch10of13A_50Por40Pofsalary <> "" Then
        If Not chkCompulsory(EA_Sch10of13A_50Por40Pofsalary) Or EA_Sch10of13A_50Por40Pofsalary = 0 Then
            MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Please provide 50% /40% of salary (C) in schedule 10(13A)""" & Chr(13)
            ValidateEA10_13A_1 = False
            Exit Function
         Else
            
           Dim Per50Salaryaspersection17 As Variant
              Per50Salaryaspersection17 = EA_Sch10of13A_DetlsofSalpersec17of1 * 0.5
            If Per50Salaryaspersection17 <> "" Then
                If EA_Sch10of13A_50Por40Pofsalary > Per50Salaryaspersection17 Then
                  MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""The value in this Field 6 ""50% /40% of salary (C)"" shall not be more than 50% of Field 4 ""Details of Salary as per Section 17(1)""""" & Chr(13)
                  ValidateEA10_13A_1 = False
                  Exit Function
                End If
            End If
            
            If Len(EA_Sch10of13A_50Por40Pofsalary) > 14 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""50% /40% of salary (C) cannot be greater than 14 digits in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
            
            If Not IsNumeric(EA_Sch10of13A_50Por40Pofsalary) Or EA_Sch10of13A_50Por40Pofsalary < 0 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at 50% /40% of salary (C) in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
        End If
            
        'Eligible Exempt Allowance u/s 10(13A)
        If EA_Sch10of13A_ElgiblExmptAllwnce10of13A <> "" Then
            If Len(EA_Sch10of13A_ElgiblExmptAllwnce10of13A) > 14 Then
            MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Eligible Exempt Allowance u/s 10(13A) cannot be greater than 14 digits in schedule 10(13A)""" & Chr(13)
            ValidateEA10_13A_1 = False
            Exit Function
            End If
            If Not IsNumeric(EA_Sch10of13A_ElgiblExmptAllwnce10of13A) Or EA_Sch10of13A_ElgiblExmptAllwnce10of13A < 0 Then
                MsgBox_EA10_13A = MsgBox_EA10_13A + "* ""Amount entered Should be Numeric, Non negative, no decimal, upto 99,999,999,999,999 at Eligible Exempt Allowance u/s 10(13A) in schedule 10(13A)""" & Chr(13)
                ValidateEA10_13A_1 = False
                Exit Function
            End If
        End If
            
            
  
  End If
        
End Function


Sub setTableInfoSch10of13A_PlaceofWrk()
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
    Set rangecells = Sheet18.Range("Sch10of13A_PlaceofWrk").Cells
    mIntCells = Sheet18.Range("Sch10of13A_PlaceofWrk").count
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) And Not UCase(rangecells.item(mIntCtr).Value) = UCase("Select") Then
        ccount = ccount + 1
        End If
    Next
    endEA10_13A_1 = ccount
     
End Sub

Sub setTableInfoSch10of13A_ActlHRArecivedA()

    Sheet18.Activate
    Dim rangecells As Range
    Dim mIntCells  As Long
    Dim mIntCtr  As Long
    Dim ccount  As Long
    ccount = 0
   
    Set rangecells = Sheet18.Range("Sch10of13A_ActlHRArecivedA").Cells
    mIntCells = Sheet18.Range("Sch10of13A_ActlHRArecivedA").count
    For mIntCtr = 1 To mIntCells
    Debug.Print Trim(rangecells.item(mIntCtr).Value)
        If Not rangecells.item(mIntCtr).Value <> 0 Then
        ccount = ccount + 1
        End If
    Next
    endEA10_13A_2 = ccount
     
End Sub

 'Ankita_08/05/2025
 
Sub cmd_NextEA100f13A_click()
'Sheet16.Activate
Sheet2.Activate

End Sub


'Ankita_08/05/2025

Sub cmd_PrevEA100f13A_click()
'Sheet1.Activate
Sheet19.Activate
End Sub


