Attribute VB_Name = "EfilingCommon"
Option Explicit

Public AssesseeDob As String
Public MsgPISheet, MsgAuditSheet, MsgPartBSheet, MsgSchIT, MsgTDS1, MsgTDS2, MsgTCS, Msgbox_FSI, MsgPISheet1 As String

Public DefinedgridNameRange As Variant
Public ActiveCellRange As String
Public strCurrActiveCellRange As String
Public strNewActiveCellRange As String
Public Msgbox_I As String
Const ProCell As String = "ITD_ITR_1_"
Dim PWD As String
Dim rowcount As Long
Public sPassword As String


Public errmsgVerification As String
Public ScheduleName() As Variant
Public Total_SheetCount As Long
Public EnabledSchedule() As Variant
Public PrintYN As Variant

Public RangeAddress As Variant



Public TotalTaxPayable As Variant
Function getSWVersionNo() As String
'...Ankita_ Updating version name from R2 to R3
  '  getSWVersionNo = "R2"
'    getSWVersionNo = "R3"
'...Ankita_ Updating version name from R3 to R4
'   getSWVersionNo = "R4"
'   getSWVersionNo = "R1"   ' For AY25-26 testing purpose Ankita 16/12/2024
'    getSWVersionNo = "R2"   ' For AY25-26 testing purpose Ankita 26/06/2025
'    getSWVersionNo = "R3"   'Changed by Ayush on 29/07/2025
'    getSWVersionNo = "R4"   'Changed by Ayush on 25/08/2025
'    getSWVersionNo = "R5"    'Changed by Ankita on 15/09/2025
 '    getSWVersionNo = "R6"    'Changed by Jyoti on 18/12/2025
'      getSWVersionNo = "R7"    'Changed by Jyoti on 05/01/2025
'      getSWVersionNo = "R1"    'Changed by Ankita_29/01/2026
'      getSWVersionNo = "R2"    'Changed by Ankita_29/06/2026
      getSWVersionNo = "R3"    'Changed by Ankita_30/07/2026

End Function

Function getSWCreatedBy() As String
'Newly changed from SW90002324
'   getSWCreatedBy = "SW90002425"
'    getSWCreatedBy = "SW90002526"  'Changed from SW90002425 to SW90002526 by Ankita on 02/01/2025
    getSWCreatedBy = "SW90002627"  'Changed by Ankita_29/01/2026
End Function

Function getJSONCreatedBy() As String
'Newly changed from SW90002324
'    getJSONCreatedBy = "SW90002425"
'    getJSONCreatedBy = "SW90002526" 'Changed from SW90002425 to SW90002526 by Ankita on 02/01/2025
    getJSONCreatedBy = "SW90002627" 'Changed by Ankita_11/02/2026
End Function

Function getIntermediaryCity() As String
    getIntermediaryCity = "Delhi"
End Function

Function getFormName() As String
    getFormName = "ITR-1"
End Function

Function getFormDescription() As String
    getFormDescription = "For Indls having Income from Salary, Pension, family pension and Interest"
End Function

Function getAssessmentYear() As String
    'Changed the year from 2023 to 2024
    'Chetan C M Changed Year from 2024 to 2025
    'getAssessmentYear = "2025"
    'New Schema updated by Konda as AY-2026-27 on 26-12-2025
    getAssessmentYear = "2026"
End Function

Function getSchemaVer() As String
    getSchemaVer = "Ver1.0"
End Function

Function getFormVer() As String
    getFormVer = "Ver1.0"
End Function

Function isdropdownblank(dropdown As Variant) As Boolean
    isdropdownblank = False
    If Mid(dropdown, 1, 1) = "(" Then
        isdropdownblank = True
    End If
    If Mid(dropdown, 1, 1) = "" Or IsEmpty(dropdown) Then
        isdropdownblank = True
    End If
End Function

Function searchLastRow(ByVal gridRangeName As String) As String
On Error Resume Next
 strCurrActiveCellRange = Replace(ActiveSheet.Range(gridRangeName).AddressLocal, "$", "")
 strNewActiveCellRange = Mid(strCurrActiveCellRange, InStr(1, strCurrActiveCellRange, ":") + 1, Len(strCurrActiveCellRange))
 ActiveSheet.Range(strNewActiveCellRange).Select
 searchLastRow = strCurrActiveCellRange
End Function
Function insertRowUnderSectionWithFormulaOne(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        'nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         nRows = 1
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    
        'Newly added by bindu 21/01/2025
    LinkCheckBoxes
    Dim chk As CheckBox
    For Each chk In ActiveSheet.CheckBoxes
        If EfilingCommon.onlyDigits(chk.LinkedCell, "I") = EfilingCommon.onlyDigits(sNewCellValue, "I") Then
           chk.Value = False
        End If
    Next
    '...............

    insertRowUnderSectionWithFormulaOne = nRows
'----------------Lock Password-------------------START---
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
    ActiveSheet.Protect Password:=sPassword
End Function
Function insertRowUnderSectionWithFormula(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula = nRows
'----------------Lock Password-------------------START---
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
    ActiveSheet.Protect Password:=sPassword
End Function



'Ankita_31/12/2025========================
'80CCC_80CCD1_80CCD1B

Function insertRowUnderSectionWithFormula_80CCC(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            'Ankita_31/12/2026=================
            If sTempFirstCellValue = "" Then
            sRangeValue = sTempCellValue + ":" + sNewCellValue
            End If
            '==================================
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            'Ankita_31/12/2026==================
            If sTempFirstCellValue = "" Then
            sRangeValue = sTempCellValue + ":" + sNewCellValue
            End If
            '===================================
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula_80CCC = nRows
'----------------Lock Password-------------------START---
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
    ActiveSheet.Protect Password:=sPassword
End Function


'=========================================
Function insertRowUnderSectionWithFormulaBA(nOfRows As Long, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False

    ActiveCell.EntireRow.Select
  nRows = nOfRows
    '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormulaBA = nRows
'----------------Lock Password-------------------START---
   
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
    ActiveSheet.Protect Password:=sPassword
    
End Function


Function getmsgstate() As String
    PWD = Sheet5.Range("pwd")
    getmsgstate = PWD
End Function


Function onlyDigits(ByVal s As String, ExtracType As String) As String
On Error Resume Next
    Dim retval As String    ' This is the return string.
    Dim i As Long        ' Counter for character position.

    retval = ""
    For i = 1 To Len(s)
        If (Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9") Then
         If ExtracType = "I" Then
            retval = retval + Mid(s, i, 1)
         End If
        Else
         If ExtracType = "S" Then
          retval = retval + Mid(s, i, 1)
         End If
        End If
    Next
    onlyDigits = retval
End Function


Function checkfieldspecialcharacter(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter = True
    Dim arr As Variant
    arr = Array("@", "|", "_", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter = False
            Exit Function
        End If
        Next
    Next
End Function
'Ankita_13/03/2026====
Function checkfieldspecialcharacter4(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter4 = True
    Dim arr As Variant
    arr = Array("@", "|", "_", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter4 = False
            Exit Function
        End If
        Next
    Next
End Function

'Ankita_23/02/2026=======
Function checkfieldspecialcharacter1(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter1 = True
    Dim arr As Variant
    'arr = Array("@", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<")
    arr = Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        'For j = 0 To UBound(arr)
        If Mid(field, i, 1) = "&" Or Mid(field, i, 1) = """" Or Mid(field, i, 1) = "'" Or Mid(field, i, 1) = ">" Or Mid(field, i, 1) = "<" Then
            checkfieldspecialcharacter1 = False
            Exit Function
        End If
        'Next
    Next
End Function



'Ankita_03/06/2025_UDIDnumber
Function checkfieldspecialcharacterUDID(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacterUDID = True
    Dim arr As Variant
    arr = Array("@", "|", "_", "*", "!", "-", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<", ".")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacterUDID = False
            Exit Function
        End If
        Next
    Next
End Function

Function checkfieldspecialcharacter_TDS_TCS(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_TDS_TCS = True
    Dim arr As Variant
    arr = Array("@", "|", "_", "*", "!", "&", "#", "~", ";", "?", "/", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_TDS_TCS = False
            Exit Function
        End If
        Next
    Next
End Function

Function checkfieldspecialcharacter3(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter3 = True
    Dim arr As Variant
    arr = Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter3 = False
            Exit Function
        End If
        Next
    Next
End Function

Function chkCompulsory(field As Variant) As Boolean
    chkCompulsory = True
    If Len(Trim(field)) <= 0 Then
        chkCompulsory = False
    End If
    If IsEmpty(field) Then
    chkCompulsory = False
    End If
End Function

Function getRangeName(ByVal Target As Range) As String
Dim start As Long
Dim nm As name
Dim rng As name


For Each nm In ThisWorkbook.Names
    Set rng = Intersect(ActiveCell, Range(nm.name))
    If Not rng Is Nothing Then
        'MsgBox ActiveCell.Address & " is in the named range " & nm.name
        fmsgbox ActiveCell.Address & " is in the named range " & nm.name
        Exit For
    End If
Next nm
getRangeName = nm.name

End Function


Sub CloseMsg()
End
End Sub
Function InitializeAllPublicVariablePI()
    
     LastName = ""
     PAN = ""
     Flat = ""
     Area = ""
     roadOrStreet = ""
     City = ""
     State = ""
     PinCode = ""
     stDcode = ""
     phoneNo = ""
     mobileNo = ""
   
     Email = ""
     dob = ""
     Gender = ""
     ReturnFileSec = ""
     ReturnFurSec = ""
     ReturnType = ""
     'ResidentialStatus = ""
     status = ""
     residenceName = ""
     
    
End Function
Sub InitializeAllPublicVariables()
    
    InitializeAllPublicVariablePI
 
    End Sub

'method to calculate the age from dateOfBirth
Function calculateAge(dob As Variant) As Long

On Error Resume Next
'Changed year from 2023 to 2024
'Ankita_25-26
'calculateAge = (2024) - val(Mid(dob, 7, 4))
' calculateAge = (2025) - val(Mid(dob, 7, 4))
 calculateAge = Sheet5.Range("DOB_YEAR").Value - val(Mid(dob, 7, 4))  'Ankita_27/12/2025=============
    If 4 < val(Mid(dob, 4, 2)) Then
        calculateAge = calculateAge - 1
    ElseIf val(Mid(dob, 4, 2)) = 4 And 1 < val(Mid(dob, 1, 2)) Then
                calculateAge = calculateAge - 1
    End If

End Function

Function checkFirstDateBefore(firstDate As Variant, secondDate As Variant) As Boolean
On Error Resume Next

If val(Mid(firstDate, 7, 4)) < val(Mid(secondDate, 7, 4)) Then
        checkFirstDateBefore = True
ElseIf val(Mid(firstDate, 7, 4)) = val(Mid(secondDate, 7, 4)) Then

        If val(Mid(firstDate, 4, 2)) < val(Mid(secondDate, 4, 2)) Then
        
             checkFirstDateBefore = True
             
        ElseIf val(Mid(firstDate, 4, 2)) = val(Mid(secondDate, 4, 2)) Then
        
                If val(Mid(firstDate, 1, 2)) < val(Mid(secondDate, 1, 2)) Then
                
                        checkFirstDateBefore = True
                        
                ElseIf val(Mid(firstDate, 1, 2)) = val(Mid(secondDate, 1, 2)) Then
                
                        checkFirstDateBefore = True
            
                Else
                      checkFirstDateBefore = False
                End If
        Else
            checkFirstDateBefore = False
        End If
Else

    checkFirstDateBefore = False
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

'Ayush_DueDate_08/09/2025

Sub rrrr()
Dim mon
mon = calcNoOfMonths("16/10/2025", "16/09/2025")
Debug.Print mon
End Sub

Function calcNoOfMonths(currentdate As Variant, startDate As Variant) As Long
On Error Resume Next
 
Dim currentYear As Variant
Dim startyear As Variant
Dim currentMonth As Variant
Dim startmonth As Variant
 
currentYear = val(Mid(currentdate, 7, 4))
startyear = val(Mid(startDate, 7, 4))
'Added by Shrutika 15/07/2025
'Changed by Riyaz on 29/08/2025
'If getDueDate = "15/09/2025" Then
 
'If getDueDate = Sheet52.Range("Duedate1").Value Then
    If Enddateofthemonth(startDate) Then
        currentMonth = val(Mid(currentdate, 4, 2)) + 1
Else
    currentMonth = val(Mid(currentdate, 4, 2))
End If
'End If
startmonth = val(Mid(startDate, 4, 2))
calcNoOfMonths = 0

If currentYear = startyear And currentMonth = startmonth And _
        val(Mid(currentdate, 1, 2)) = val(Mid(startDate, 1, 2)) Then
        calcNoOfMonths = 0
ElseIf (checkFirstDateBefore(currentdate, startDate)) Then
        calcNoOfMonths = 0
Else
    If currentYear = (startyear + 1) Then
        If currentMonth < startmonth Then
            'calcNoOfMonths = 12 - (startmonth - currentMonth) + 1
            calcNoOfMonths = 12 - (startmonth - currentMonth)
        Else
        'tobetested
            'calcNoOfMonths = 12 + (currentMonth - startmonth) + 1
            calcNoOfMonths = 12 + (currentMonth - startmonth)
        End If
   ElseIf (currentYear = startyear) Then
        'calcNoOfMonths = currentMonth - startmonth + 1
        calcNoOfMonths = currentMonth - startmonth
   If (val(Mid(currentdate, 1, 2)) > val(Mid(startDate, 1, 2))) And (currentMonth = startmonth) Then
        calcNoOfMonths = calcNoOfMonths + 1
   'ElseIf val(Mid(startDate, 1, 2)) = 15 And (currentMonth >= startmonth) Then
   ElseIf val(Mid(startDate, 1, 2)) = 7 And (currentMonth >= startmonth) Then '15/07/2025
        calcNoOfMonths = calcNoOfMonths + 1
     End If
   ElseIf currentMonth < startmonth Then
        'calcNoOfMonths = Round((currentYear - startyear - 1) * 12) + 12 - startmonth + currentMonth + 1
        calcNoOfMonths = Round((currentYear - startyear - 1) * 12) + 12 - startmonth + currentMonth
   ElseIf (currentMonth > startmonth) Then
        'calcNoOfMonths = Round(((currentYear - startyear) * 12)) + currentMonth - startmonth + 1
        calcNoOfMonths = Round(((currentYear - startyear) * 12)) + currentMonth - startmonth
    Else
        'calcNoOfMonths = Round(((currentYear - startyear) * 12)) + 1
        calcNoOfMonths = Round(((currentYear - startyear) * 12))
    End If
End If
'Added by Shrutika 15/07/2025
If currentYear = (startyear + 1) Then
 
If val(Mid(startDate, 1, 2)) = 7 And (currentYear > startyear) Then
        calcNoOfMonths = calcNoOfMonths + 1
End If
End If
 
End Function

Function calcNoOfMonths_old(currentdate As Variant, startDate As Variant) As Long
On Error Resume Next

Dim currentYear As Variant
Dim startyear As Variant
Dim currentMonth As Variant
Dim startmonth As Variant
Dim currentDay As Variant
Dim startDay As Variant

currentYear = val(Mid(currentdate, 7, 4))
startyear = val(Mid(startDate, 7, 4))

'Ayush_Due_Date_08/09/2025
If Enddateofthemonth(startDate) Then
    currentMonth = val(Mid(currentdate, 4, 2))
Else
    startmonth = val(Mid(startDate, 4, 2))
End If
'------------

'startmonth = val(Mid(startDate, 4, 2))
currentDay = val(Mid(currentdate, 1, 2))
startDay = val(Mid(startDate, 1, 2))

 calcNoOfMonths = 0
 
 
 If currentYear = startyear And currentMonth = startmonth And _
        val(Mid(currentdate, 1, 2)) = val(Mid(startDate, 1, 2)) Then
'        calcNoOfMonths = 1    'Ankita_22/07/2025
         calcNoOfMonths = 0
ElseIf (checkFirstDateBefore(currentdate, startDate)) Then
        calcNoOfMonths = 0
Else
    If currentYear = (startyear + 1) Then
        If currentMonth < startmonth Then
            calcNoOfMonths = 12 - (startmonth - currentMonth) + 1
        Else
            calcNoOfMonths = 12 + (currentMonth - startmonth) + 1
        End If
   ElseIf (currentYear = startyear) Then
        calcNoOfMonths = currentMonth - startmonth + 1
   ElseIf currentMonth < startmonth Then
        calcNoOfMonths = Round((currentYear - startyear - 1) * 12) + 12 - startmonth + currentMonth + 1
   ElseIf (currentMonth > startmonth) Then
        calcNoOfMonths = Round(((currentYear - startyear) * 12)) + currentMonth - startmonth + 1
    Else
        calcNoOfMonths = Round(((currentYear - startyear) * 12)) + 1
    End If
End If

'If calcNoOfMonths > 0 Then
'calcNoOfMonths = calcNoOfMonths - 1
'End If

' commented by Ankita on 28/07/2025****************************************************

'If calcNoOfMonths > 0 And startDay <> 10 Then
'calcNoOfMonths = calcNoOfMonths - 1
'End If
'
'If calcNoOfMonths > 0 And startDay = 10 And (currentDay <= startDay) Then
'calcNoOfMonths = calcNoOfMonths - 1
'End If
'***************************************************************
End Function

'Ayush_DueDate_08/09/2025
Function Enddateofthemonth(startDate As Variant) As Boolean
Enddateofthemonth = True
 
If Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 1 Then
    Enddateofthemonth = False
ElseIf (Mid(startDate, 1, 2) = 28 Or Mid(startDate, 1, 2) = 29) And Mid(startDate, 4, 2) = 2 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 3 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 4 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 5 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 6 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 7 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 8 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 9 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 10 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 30 And Mid(startDate, 4, 2) = 11 Then
    Enddateofthemonth = False
ElseIf Mid(startDate, 1, 2) = 31 And Mid(startDate, 4, 2) = 12 Then
    Enddateofthemonth = False
Else
    Enddateofthemonth = True
End If
End Function
'------------------




Function ValidateIFSCList(IFSC As String) As Boolean
    ValidateIFSCList = False
    Dim IFSCCodeSet1 As Range
    Dim IFSCCodeSet2 As Range
    Dim IFSCCodeSet3 As Range
    Dim IFSCCodeSet4 As Range
    Dim IFSCCodeSet5 As Range
    
    Dim IFSC1 As Range
    Dim IFSC2 As Range
    Dim IFSC3 As Range
    Dim IFSC4 As Range
    Dim IFSC5 As Range

'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   Sheet8.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    
    Set IFSCCodeSet1 = Range("IFSCCodeSet1")
    Set IFSCCodeSet2 = Range("IFSCCodeSet2")
    Set IFSCCodeSet3 = Range("IFSCCodeSet3")
    Set IFSCCodeSet4 = Range("IFSCCodeSet4")
    Set IFSCCodeSet5 = Range("IFSCCodeSet5")
    
    For Each IFSC1 In IFSCCodeSet1
        If IFSC1.Value = IFSC Then
        ValidateIFSCList = True
        Exit For
        End If
    Next

    If Not ValidateIFSCList Then
        For Each IFSC2 In IFSCCodeSet2
            If IFSC2.Value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    
    If Not ValidateIFSCList Then
        For Each IFSC3 In IFSCCodeSet3
            If IFSC3.Value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    
    If Not ValidateIFSCList Then
        For Each IFSC4 In IFSCCodeSet4
            If IFSC4.Value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    
    
    If Not ValidateIFSCList Then
        For Each IFSC5 In IFSCCodeSet5
            If IFSC5.Value = IFSC Then
                ValidateIFSCList = True
                Exit For
            End If
        Next
    End If
    ValidateIFSCList = True
   '----------------Lock Password-------------------START---
   Sheet8.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
 
End Function


Function printWorkSheet()
    Dim PrintMsgOP As Long
Sheet1.Activate
    
    PrintMsgOP = MsgBox("Do you want to preview the workbook for printing?", vbOKCancel, "Print preview")
    If PrintMsgOP = 1 Then
           With Sheet1.PageSetup 'PART A GENERAL
'                .PrintArea = Sheet1.Range("Income_Details").Address
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3 'Newly Updated by Bindu
                .FitToPagesWide = 1 'Newly added by Bindu
                '.FitToPagesTall = False 'Newly added by Bindu
                .FitToPagesTall = False
                .Zoom = False 'Newly added by Bindu
                .Orientation = xlLandscape
            End With
            With Sheet2.PageSetup 'TDS
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            With Sheet201.PageSetup 'Part A General 139(8A)
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            With Sheet202.PageSetup 'Part B ATI
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With

             With Sheet3.PageSetup 'TAX PAID & VERIFICATION
               .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
            With Sheet4.PageSetup '80G
            .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            
'            With Sheet9.PageSetup
'            .BlackAndWhite = True
'                .CenterHorizontally = True
'                .CenterVertically = False
'                .LeftMargin = 0.4
'                .RightMargin = 0.4
'                .PaperSize = xlPaperA4
'                .FitToPagesWide = 1
'                .Orientation = xlLandscape
'            End With
            
            With Sheet11.PageSetup 'TCS
            .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
             With Sheet12.PageSetup '80GGA
            .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
             With Sheet13.PageSetup '80GGC
            .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
             With Sheet14.PageSetup '80U-80DD
            .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
             With Sheet15.PageSetup '80C
            .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
             With Sheet16.PageSetup '24_B
            .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
             With Sheet17.PageSetup '80E_SERIES
            .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = 70
                .Orientation = xlLandscape
            End With
             With Sheet18.PageSetup '10(13A)
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
            With Sheet9.PageSetup '80D
                .BlackAndWhite = True
                .CenterHorizontally = True
                .CenterVertically = False
                .LeftMargin = 0.4
                .RightMargin = 0.4
                .PaperSize = xlPaperA3
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .Zoom = False
                .Orientation = xlLandscape
            End With
         
        
           
        
     'rowcount = ThisWorkbook.Worksheets.count
    ThisWorkbook.PrintOut , PREVIEW:=True
    Sheet1.Activate
 Else
 
 End If

End Function

Sub SaveXML_Click()
'fmsgbox ("Pending")
'Sheet7.Visible = xlSheetHidden
 Sheet1.Activate
'Exit Sub
' Application.ScreenUpdating = False
        InitProgBar
        UserForm1.Show vbModeless
        ProgressFrameCaption = "Generating JSON"
        mainProcCaption = "Calculating Taxes"
        noOfProcessMain = 6
        
        UpdateProgressBar
         ValidateSchedulePI
         
 '        Application.ScreenUpdating = True
         
         UpdateProgressBar
         ValidateTDS_All
         validateSchTCS
         
         UpdateProgressBar
         If Sheet5.Range("BacValue").Value = 2 Then
         Validate80G_All
         End If
         
         UpdateProgressBar
         If Sheet5.Range("BacValue").Value = 2 Then
         Validate80D_All
         End If
         
         UpdateProgressBar
         ValidatePartBTI_BTTI_Verification
         
         'mIncmDtls.calcItr1
         UpdateProgressBar
         ValidateBA
         
        ' Sheet1.LockUnlockInterest
         
         UpdateProgressBar
        ' GenarateXML.Generate_XML
        ' ProgressBarHide
        Sheet1.Activate
         Generate_JSON
'         Application.ScreenUpdating = True
End Sub

Sub CalculateTax_Click_Click()
         'Application.ScreenUpdating = False
'----------------Unlock Password-------------------START---
    sPassword = EfilingCommon.getmsgstate
    Sheet3.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

        Sheet3.Range("IncD.TDS").Value = 0
        Sheet3.Range("IncD.TCS").Value = 0
        Sheet3.Range("IncD.AdvanceTax").Value = 0
        Sheet3.Range("IncD.SelfAssessmentTax").Value = 0
        
'----------------Lock Password-------------------START---
  Sheet3.Protect Password:=sPassword
 '----------------Lock Password-------------------END-----
 InitProgBar
   ProgressFrameCaption = "Tax Calculation"
   mainProcCaption = "Calculating"
   noOfProcessMain = 7
   UserForm1.Show vbModeless
        UpdateProgressBar
        ValidateSchedulePI
        
        ValidateSheetHouseProperty  'Konda Added on 23-01-2026
        
        If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
            Validate_Gen_1398A
        End If
         'mainProcCaption = "Validating TDS"
         UpdateProgressBar
         ValidateTDS_All
         
         UpdateProgressBar
         validateSchTCS
         UpdateProgressBar
        If Sheet5.Range("BacValue").Value = 2 Then
        Validate80D_All
        End If
        ' mainProcCaption = "Validating 80G"
         UpdateProgressBar
        If Sheet5.Range("BacValue").Value = 2 Then
         Validate80G_All
        End If
         ' mainProcCaption = "Validating 80G"
         UpdateProgressBar
        If Sheet5.Range("BacValue").Value = 2 Then
            Validate80GGA
        End If
        '80GGC_C1 2024-25 Bindu
         UpdateProgressBar
        If Sheet5.Range("BacValue").Value = 2 Then
            Validate80GGC
            
             'Malli--------AY_2025_26
'                Validate80C_80CCC_All
'                Validate80C_80CCC2_All
                 Validate80C_All
                Validate80E_All
                Validate80EE_All
                Validate80EEA_All
                Validate80EEB_All
                'Validate24b_All
                
       '-----------------------
            
        End If
        
'        If Sheet16.Visible = True Then
'
'        End If
        
        'If Sheet5.Range("BacValue").Value = 2 And Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "S" Then
        '25/04/2025
        'Ankita_21/01/2026====
'        If Sheet5.Range("BacValue").Value = 2 Then
'                If Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" Then
'                       Validate24b_All
'                End If
'        ElseIf Sheet5.Range("BacValue").Value = 1 Then 'Ankita_15/05/2025
'                If Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" And Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "S" Then
'                       Validate24b_All
'                End If
'        End If
        
        If Sheet5.Range("BacValue").Value = 2 And Mid(Range("sheet1.EmployerCategory1").Value, 1, 1) <> "N" Then
          ValidateEA10_13A
        End If
        
                 'mainProcCaption = "Validating PartB-TTI"
        UpdateProgressBar
        ValidatePartBTI_BTTI_Verification
         
'        If Sheet1.Range("sheet1.ReturnFileSec").Value = "139(8A)" Then
'            ValidateATI
'        End If
        
         'mainProcCaption = "Validating BA"
         UpdateProgressBar
         ValidateBA
         
         UpdateProgressBar
         'ValidateSchAL
         'mainProcCaption = "Calculating Interest"
         UpdateProgressBar
         
        'Sheet1.LockUnlockInterest
         
         mIncmDtls.calcItr1
         
         If Sheet3.Range("IncD.BalTaxPayable").Value > 0 Then
         Sheet3.Shapes("Button 131").Visible = True
          Else
           Sheet3.Shapes("Button 131").Visible = False
        End If
        
            'Ankita_UR
    If UCase(Sheet1.Range("sheet1.ReturnFileSec").Value) = UCase("139(8A)") Then
        If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
            If Sheet202.Range("U_AmtPayable") <= 0 Then
                MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then updated return can't be filed if there is no payment u/s 140B", vbCritical
                Sheet1.Range("taxcheck").Value = "Y"
            Exit Sub
            End If
        End If
        
        'Ankita_UR
       If UCase(Sheet201.Range("U_PreviouslyFiledForThisAY").Value) = UCase("No") Then
            If Sheet202.Range("U_AmtRefundable") > 0 Then
                MsgBox "If A5 in Part A gen 139(8A)  is ""No"",then updated return can't be filed with Refund" & Chr(13)
                Sheet1.Range("taxcheck").Value = "Y"
           Exit Sub
           End If
        End If
        
        'Ankita_UR
        If Sheet202.Range("U_NetPayable") <= 0 Then
        MsgBox "Since you are filing Updated return, you are expected to have ""Net amount payable"" more than 0 to be able to pay tax u/s 140B." & Chr(13)
'       Sheet1.Range("taxcheck").Value = "Y"
'       Exit Sub
        End If
      
             'Ankita_UR
       If Sheet202.Range("U_TaxUS140B") = 0 Then
       MsgBox """Tax Paid u/s 140B at sl.no.12 should be more than ""0"" to file ITR u/s 139(8A)""" & Chr(13)
       End If
       

    'Ankita_UR
        If Sheet202.Range("U_TaxUS140B") > Sheet202.Range("U_NetPayable").Value Then
            MsgBox "Tax paid u/s 140B is more than ""sl.no.11.Net amount payable"". Hence your updated return may be invalid as it appears to result in refund as provided in clause (c) to first proviso of 139(8A).", vbInformation
        End If
End If


        ProgressBarHide
        Sheet1.Activate
        Application.EnableEvents = False
        Sheet1.Range("taxcheck").Value = "N"
        Application.EnableEvents = True
      ' Application.ScreenUpdating = True
End Sub
Sub IDHelpClick_Click()
sPassword = EfilingCommon.getmsgstate
ActiveWorkbook.Unprotect Password:=sPassword
Sheet6.Activate
Sheet6.Visible = xlSheetVisible

ActiveWorkbook.Protect Password:=sPassword
End Sub
Sub GenearteXML_Click()
'fmsgbox ("Pending")
'Sheet7.Visible = xlSheetHidden
 Sheet1.Activate
'Exit Sub
       Dim sourceSheet As Worksheet
If Sheet1.Range("taxcheck").Value <> "N" Then
'MsgBox "Please ensure to click on 'Calculate Tax' button before generating XML.", vbCritical, "Error"
fmsgbox "Please ensure to click on 'Calculate Tax' button before generating JSON."
End
End If
sPassword = EfilingCommon.getmsgstate
ActiveWorkbook.Unprotect Password:=sPassword
    Set sourceSheet = ThisWorkbook.Sheets("SUMMARY")
    InitProgBar
    UserForm1.Show vbModeless
       ProgressFrameCaption = "Generating XML"
       noOfProcessMain = 6
       mainProcCaption = "Calculating Tax"
       UpdateProgressBar
       ValidateSchedulePI
       'Changed by Ankita 30/12/2024
       'Correctly updated by Bindu on 4th Feb 2025 as per DE V3
        'If (Sheet1.Range("IncD.TotalIncome").Value > 5000000) Then
        If ((Sheet1.Range("IncD.TotalIncome_New").Value - Sheet1.Range("IncD.CG_LTCG")) > 5000000) Then
            'MsgBox "ITR 1 is for individuals being a resident other than not ordinarily resident having Income from Salaries, one house property, other sources (Interest etc.), agricultural income upto Rs.5 thousand and having total income upto Rs.50 lakh. Please file other ITR.", vbOKOnly, "Error"
'           fmsgbox "* ITR 1 is for individuals being a resident other than not ordinarily resident having Income from Salaries, one house property, other sources (Interest etc.), agricultural income upto Rs.5 thousand and having total income upto Rs.50 lakh. Please file other ITR."
            fmsgbox "* ""ITR 1 is for individuals being a resident (other than not ordinarily resident) having total income upto Rs.50 lakh, having Income from Salaries, one house property, interest income, Family pension income etc. and agricultural income upto Rs.5 thousand. Please file another ITR as applicable""" & Chr(13)

 
        End If
        
       UpdateProgressBar
       ValidateTDS_All
       UpdateProgressBar
       If Sheet5.Range("BacValue").Value = 2 Then
       Validate80D_All
       End If
       UpdateProgressBar
       UpdateProgressBar
       If Sheet5.Range("BacValue").Value = 2 Then
       Validate80G_All
       End If
       
       UpdateProgressBar
       If Sheet5.Range("BacValue").Value = 2 Then
       Validate80GGA
       End If
       
       '80GGC_C1 2024-25 Bindu
       UpdateProgressBar
       If Sheet5.Range("BacValue").Value = 2 Then
       Validate80GGC
       Validate80U
       Validate80DD
       
       'Malli--------AY_2025_26
'                Validate80C_80CCC_All
'                Validate80C_80CCC2_All
                 Validate80C_All
                Validate80E_All
                Validate80EE_All
                Validate80EEA_All
                Validate80EEB_All
                'Validate24b_All
                
       '-----------------------
       
       End If
       
       'Malli-25/04/2025
       'Ankita_21/01/2026====
'         If Sheet5.Range("BacValue").Value = 2 Then
'                If Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" Then
'                       Validate24b_All
'                End If
'        ElseIf Sheet5.Range("BacValue").Value = 1 Then 'Ankita_15/05/2025
'                If Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" And Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "S" Then
'                       Validate24b_All
'                End If
'        End If
        
        
        If Sheet5.Range("BacValue").Value = 2 And Mid(Range("sheet1.EmployerCategory1").Value, 1, 1) <> "N" Then
            ValidateEA10_13A
            End If
       '-----------------------
       
       
       UpdateProgressBar
       If Mid(Sheet1.Range("sheet1.ReturnFileSec").Value, 1, 7) = "139(8A)" Then
            Validate_Gen_1398A
        End If
       UpdateProgressBar
       'ValidateSchAL
       UpdateProgressBar
       validateSchTCS
       UpdateProgressBar
       
       
       ValidatePartBTI_BTTI_Verification
       UpdateProgressBar

        If Mid(Sheet1.Range("sheet1.ReturnFileSec").Value, 1, 7) = "139(8A)" Then
            ValidateATI
        End If
        UpdateProgressBar
       ValidateBA
           
       UpdateProgressBar
       'Malli Validate 80U_DD
    
       
'       If Sheet2.Range("TDS.IncSum").Value > 0 Then
'            If IIf(Sheet1.Range("IncD.TotalHeadSalaries").Value = "", 0, Sheet1.Range("IncD.TotalHeadSalaries").Value) < (Sheet2.Range("TDS.IncSum").Value - Round(Sheet2.Range("TDS.IncSum").Value * 0.1, 0)) Then
'             '  IIf(Sheet1.Range("IncD.IncomeFromSal").Value = "", 0, Sheet1.Range("IncD.IncomeFromSanl").Value) > (Sheet2.Range("TDS.IncSum").Value + Round(Sheet2.Range("TDS.IncSum").Value * 0.1, 0)) Then
'            If Not end_OthersNOI > 0 Then
'               MsgBox "Since the amount disclosed in Income chargeable under the Head Salaries is less than 90% of Salary reported in TDS1, please ensure to fill details in " & Chr(34) & "Others" & Chr(34) & " in " & """Exempt Income""" & Chr(13), vbCritical, "Exempt Income"
'
'
'           End If
'        End If
'    End If

    If Sheet2.Range("TDS.IncSum").Value > 0 Then
            If IIf(Sheet1.Range("IncD.IncomeFromSal").Value = "", 0, (Sheet1.Range("IncD.IncomeFromSal").Value)) < (Sheet2.Range("TDS.IncSum").Value - 10) Then
             '  IIf(Sheet1.Range("IncD.IncomeFromSal").Value = "", 0, Sheet1.Range("IncD.IncomeFromSanl").Value) > (Sheet2.Range("TDS.IncSum").Value + Round(Sheet2.Range("TDS.IncSum").Value * 0.1, 0)) Then

                'MsgBox "Amount of gross salary disclosed in Income details is less than (100% of Salary reported in Schedule TDS1 - Rs 10).", vbExclamation, "Warning"
                fmsgbox "* Amount of gross salary disclosed in Income details is less than (100% of Salary reported in Schedule TDS1 - Rs 10)."
            End If
        End If
       
       
       If (((Sheet3.Range("IncD.TotalTaxesPaid").Value - Sheet3.Range("IncD.TCS").Value) > 0) And _
            (IIf(Sheet1.Range("IncD.TotalHeadSalaries").Value = "", 0, Sheet1.Range("IncD.TotalHeadSalaries").Value) = 0 And _
            IIf(Sheet1.Range("IncD.IncomeFromSal").Value = "", 0, Sheet1.Range("IncD.IncomeFromSal").Value) = 0 And _
            IIf(Sheet1.Range("IncD.Allowances").Value = "", 0, Sheet1.Range("IncD.Allowances").Value) = 0 And _
            IIf(Sheet1.Range("IncD.Perquisites").Value = "", 0, Sheet1.Range("IncD.Perquisites").Value) = 0 And _
            IIf(Sheet1.Range("IncD.Profits").Value = "", 0, Sheet1.Range("IncD.Profits").Value) = 0 And _
            IIf(Sheet1.Range("IncD.Deduction16").Value = "", 0, Sheet1.Range("IncD.Deduction16").Value) = 0 And _
            IIf(Sheet1.Range("IncD.IncomeHeadHouseProperty").Value = "", 0, Sheet1.Range("IncD.IncomeHeadHouseProperty").Value) = 0 And _
            IIf(Sheet1.Range("IncD.GrossRentRecieved").Value = "", 0, Sheet1.Range("IncD.GrossRentRecieved").Value) = 0 And _
            IIf(Sheet1.Range("IncD.TaxPaidLocalAuthorities").Value = "", 0, Sheet1.Range("IncD.TaxPaidLocalAuthorities").Value) = 0 And _
            IIf(Sheet1.Range("IncD.AnnualValue").Value = "", 0, Sheet1.Range("IncD.AnnualValue").Value) = 0 And _
            IIf(Sheet1.Range("IncD.StandardDeduction").Value = "", 0, Sheet1.Range("IncD.StandardDeduction").Value) = 0 And _
            IIf(Sheet1.Range("IncD.InterestBorrowedCapital").Value = "", 0, Sheet1.Range("IncD.InterestBorrowedCapital").Value) = 0 And _
            IIf(Sheet1.Range("IncD.IncomeFromOS").Value = "", 0, Sheet1.Range("IncD.IncomeFromOS").Value) = 0)) Then
            
            'MsgBox "No Income details or tax computation is provided in ITR but details regarding taxes paid is provided", vbOKOnly, "ITR-1 AY 2020-21"
            fmsgbox "* No Income details or tax computation is provided in ITR but details regarding taxes paid is provided"
            CloseMsg
        End If
        
         
        
'         If Sheet1.Range("IncD.TypeOfHP").Value <> "Self Occupied" Then
'         If (Sheet2.Range("TDS26QB.Sum").Value > 0) Then
'            If Sheet1.Range("IncD.TypeOfHP").Value = "Let Out" Then
'                If Not Sheet1.Range("IncD.GrossRentRecieved").Value > 0 Then
'                    MsgBox "If TDS3(16C) is filled, then please ensure that the 'Type of house property' should be 'Let out'  & 'Gross rent received /receivable /letable value' should be greater than zero", vbCritical
'                    Sheet1.Activate
'                    CloseMsg
'                End If
'            Else
'                    MsgBox "If TDS3(16C) is filled, then please ensure that the 'Type of house property' should be 'Let out'  & 'Gross rent received /receivable /letable value' should be greater than zero", vbCritical
'                    Sheet1.Activate
'                    CloseMsg
'            End If
'        End If
       ' End If
        
        'If Sheet3.Range("IncD.BalTaxPayable").Value > 0 Then
        '    MsgBox ("Please ensure that the taxes are paid before the submission of the return, else return shall be treated as defective."), vbExclamation, "Warning"
        'End If
        ProgressBarHide
        sPassword = EfilingCommon.getmsgstate
        sourceSheet.Activate
        sourceSheet.Visible = xlSheetVisible
        sourceSheet.Unprotect Password:=sPassword
        
        If end_I > 0 Then
            sourceSheet.Range("D14").Value = end_I
        Else
            sourceSheet.Range("D14").Value = 0
        End If
        
        If ColCount1 > 0 Then
            sourceSheet.Range("D15").Value = ColCount1
        Else
            sourceSheet.Range("D15").Value = 0
        End If
        
        If ColCount3 > 0 Then
            sourceSheet.Range("D16").Value = ColCount3
        Else
            sourceSheet.Range("D16").Value = 0
        End If
        
        If ColCount2 > 0 Then
            sourceSheet.Range("D17").Value = ColCount2
        Else
            sourceSheet.Range("D17").Value = 0
        End If
        If end_TCS > 0 Then
            sourceSheet.Range("D18").Value = end_TCS
        Else
            sourceSheet.Range("D18").Value = 0
        End If
        sourceSheet.Protect Password:=sPassword
        ActiveWorkbook.Protect Password:=sPassword
        'MsgBox "To Compute Tax and Interest using this utility,you must click on 'Calculate Tax' button and verify the figures before saving the XML,if you  have not done so,please do it and then Generate XML", vbInformation, "Alert"
        fmsgbox "Please verify that the TDS statement is submitted for complete financial year"
        fmsgbox "To Compute Tax and Interest using this utility,you must click on 'Calculate Tax' button and verify the figures before saving the JSON,if you  have not done so,please do it and then Generate JSON"
        fmsgbox "To Save JSON Click on the 'Save JSON' Button on this Sheet"
       ''Newly added for Summary Gray off issue by Bindu 10/01/2025
        Sheet1.Activate
        sourceSheet.Activate
        ThisWorkbook.Unprotect Password:=sPassword
        
        sourceSheet.Visible = xlSheetVisible
        ThisWorkbook.Protect Password:=sPassword
        
        End Sub
Sub cmdNext_Click()
    Dim sourceSheet As Worksheet

    Set sourceSheet = ThisWorkbook.Sheets("TDS")
    sourceSheet.Activate

End Sub
Sub Print_Click()

printWorkSheet
End Sub
Sub CmdValidate_Click()
    
    If (Sheet1.Range("HRA").Value > 0 And Sheet1.Range("IncD.Section80GG").Value > 0) Then
    'MsgBox "Deduction u/s 10(13A) & 80GG cannot be claimed for the same period "
'    MsgBox "Person receiving HRA cannot claim Deduction u/s 80GG "
    MsgBox """Deduction u/s 10(13A) & 80GG cannot be claimed for the same period""" 'Changed by Ankita on 02/01/2025
    End If
    ValidateSchedulePI
    ValidatePropertyType
    ValidateSIncmFrmHP
    'Ankita_04/02/2025
     If Not validateLTCG() Then
        Sheet1.Activate
        fmsgboxsmall_LTCG msgError_LTCG
        CloseMsg
     End If
     
   If MsgPISheet = "" Then fmsgboxoK ("Sheet Income Details is OK")

End Sub



Function chkNumeric(field As Variant) As Boolean
    Dim k As Long
    Dim chkchar As String

    chkNumeric = True
    For k = 1 To Len(field)
    chkchar = Mid(field, k, 1)
        If Not IsNumeric(chkchar) Then
            chkNumeric = False
            Exit Function
        End If
    Next
End Function

Function CheckDateddmmyyyy(dt As Variant) As Boolean
    CheckDateddmmyyyy = True
    
    If InStr(1, dt, ".") > 0 Then
        CheckDateddmmyyyy = False
    End If
    
    If Len(dt) > 0 Then
        
        If Mid(dt, 3, 1) <> "/" Then
            If Mid(dt, 3, 1) <> "\" Then
                If Mid(dt, 3, 1) <> "-" Then
                    If Mid(dt, 3, 1) <> "." Then
                        CheckDateddmmyyyy = False
                    Else
                        dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
                        CheckDateddmmyyyy = False
                    End If
                Else
                    dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
                    CheckDateddmmyyyy = False
                End If
            Else
                dt = Mid(dt, 1, 2) & "/" & Mid(dt, 4, 7)
                CheckDateddmmyyyy = False
            End If
        End If
        
        If Mid(dt, 6, 1) <> "/" Then
            If Mid(dt, 6, 1) <> "-" Then
                If Mid(dt, 6, 1) <> "\" Then
                    If Mid(dt, 6, 1) <> "." Then
                        CheckDateddmmyyyy = False
                    Else
                        dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
                        CheckDateddmmyyyy = False
                    End If
                Else
                    dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
                    CheckDateddmmyyyy = False
                End If
            Else
                dt = Mid(dt, 1, 5) & "/" & Mid(dt, 7, 4)
                CheckDateddmmyyyy = False
            End If
        End If
        
        If Not IsDate(dt) Then CheckDateddmmyyyy = False
        If val(Mid(dt, 1, 2)) < 0 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 1, 2)) > 31 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 4, 2)) < 0 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 4, 2)) > 12 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 7, 4)) < 1 Then CheckDateddmmyyyy = False
        If val(Mid(dt, 7, 4)) > 3000 Then CheckDateddmmyyyy = False
    End If
End Function

Function ChkMinInclusiveDate(Mininclusive As Variant, Mininclusivedate As Variant) As Boolean
    ChkMinInclusiveDate = True
    If Len(Mininclusive) > 0 Then
        If Mid(Mininclusive, 1, 4) < Mid(Mininclusivedate, 1, 4) Then
            ChkMinInclusiveDate = False
            Exit Function
        Else
            If Mid(Mininclusive, 1, 4) = Mid(Mininclusivedate, 1, 4) Then
                If (Mid(Mininclusive, 6, 2) < Mid(Mininclusivedate, 6, 2)) Then
                    ChkMinInclusiveDate = False
                    Exit Function
                ElseIf ((Mid(Mininclusive, 6, 2) = Mid(Mininclusivedate, 6, 2))) Then
                    If (Mid(Mininclusive, 9, 2) < Mid(Mininclusivedate, 9, 2)) Then
                        ChkMinInclusiveDate = False
                        Exit Function
                   End If
                End If
            End If
        End If
    End If
  
End Function
Function checkfieldSuperSpecialcharacter(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldSuperSpecialcharacter = True
    Dim arr As Variant
    arr = Array(">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldSuperSpecialcharacter = False
            Exit Function
        End If
        Next
    Next
End Function
Function checkfieldSuperSpecialcharactermax(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldSuperSpecialcharactermax = True
    Dim arr As Variant
    arr = Array(">", "<", "~", "!", "@", "#", "$", "%", "^", "&", "*", "+", "=", "{", "}", "[", "]", "'", "`", "?", ".", """")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldSuperSpecialcharactermax = False
            Exit Function
        End If
        Next
    Next
End Function
Function checkfieldSuperSpecialcharactername(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldSuperSpecialcharactername = True
    Dim arr As Variant
    arr = Array(">", "<", "&")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldSuperSpecialcharactername = False
            Exit Function
        End If
        Next
    Next
End Function



Sub CreateArray()
    Dim i, s As Long

    Total_SheetCount = 0

    For i = 3 To 7
        s = i
        If Not Sheet5.Range("O" & s).Value = "" Then
           Total_SheetCount = Total_SheetCount + 1
        End If
    Next
    Total_SheetCount = Total_SheetCount + 1

    ReDim ScheduleName(Total_SheetCount)
        For i = 3 To Total_SheetCount
        s = i
        ScheduleName(i) = Sheet5.Range("P" & s)
    Next

    ReDim EnabledSchedule(Total_SheetCount)
    For i = 3 To Total_SheetCount
        s = i
        EnabledSchedule(i) = Sheet5.Range("Q" & s)
    Next
    
    ReDim PrintYN(Total_SheetCount)
        For i = 3 To Total_SheetCount
            s = i
            PrintYN(i) = Sheet5.Range("R" & s)
        Next
End Sub

Private Sub ButtonHyperlink()
On Error Resume Next
fmsgbox "After e-payment of Tax, the details of amount paid, Challan No., BSR Code, etc. should be filled in Schedule IT of Income Tax Return before submission of the return to claim the challan"
ActiveWorkbook.FollowHyperlink _
Address:="https://eportal.incometax.gov.in/iec/foservices/#/e-pay-tax-prelogin/user-details"
'Address:="https://onlineservices.tin.egov-nsdl.com/etaxnew/tdsnontds.jsp"   'Ankita_04/06/2025
End Sub

Sub BacYesValueChange()
On Error Resume Next 'Ankita_16/06/2025
Dim answer As Integer
 
 
Dim sPassword As Variant
sPassword = EfilingCommon.getmsgstate

'PAG_C4 2024-25 Bindu
'answer = MsgBox("Since you are opting for New tax Regime u/s 115BAC. You will not be eligible to set-off House property loss and claim following deduction/allowances." & vbCrLf & "1) Certain allowances u/s section 10 (LTA, HRA, allowances granted to meet expenses in performance of duties of office, Allowances granted to meet personal expenses in performance of duties of office, Allowance received by MP/MLA/MLC) " & vbCrLf & "2) Deductions u/s 16 (Standard Deduction ,Entertainment allowance and Professional tax) " & vbCrLf & "3) Interest payable on borrowed capital for self occupied property " & vbCrLf & "4) Standard Deduction in case of family pension " & vbCrLf & "5) Chapter VIA Deduction (life insurance, health insurance premium, pension funds, provident fund,donation etc except Contribution made by employer to notified pension scheme u/s 80CCD(2))", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
'If answer = vbYes Then
    ThisWorkbook.Unprotect Password:=sPassword
    'PAG_C3 Need to Comment below code 2024-25 Bindu
    
    'Ankita_25/07/2025==================================
'    If Range("sheet1.ReturnFileSec").Value = "139(4)-Belated" Then
'        'fmsgbox "* Please select ""Are you opting for new tax regime u/s 115BAC?""  as n since you have filed return later."
'        fmsgbox "* In order to avail the option under Section 115BAC, the return should be filed on or before due date."
'        sPassword = EfilingCommon.getmsgstate
'        Sheet5.Unprotect Password:=sPassword
''        Sheet5.Range("BacValue").Clear
''        Sheet5.Protect Password:=sPassword
'        ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 0
''         ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 0
'          Exit Sub
'    End If

'=======================================================
  Sheet4.Unprotect Password:=sPassword
  Sheet4.Visible = xlSheetHidden
  Sheet4.Protect Password:=sPassword
  Sheet12.Unprotect Password:=sPassword
  Sheet12.Visible = xlSheetHidden
  Sheet12.Protect Password:=sPassword
  Sheet9.Unprotect Password:=sPassword
  Sheet9.Visible = xlSheetHidden
  Sheet9.Protect Password:=sPassword
  
  '80GGC_C1 2024-25 Bindu
  Sheet13.Unprotect Password:=sPassword
  Sheet13.Visible = xlSheetHidden
  Sheet13.Protect Password:=sPassword
  
  ' Ankita_21/04/2025
  Sheet17.Unprotect Password:=sPassword
  Sheet17.Visible = xlSheetHidden
  Sheet17.Protect Password:=sPassword
  
 ' Ankita_21/04/2025
  Sheet15.Unprotect Password:=sPassword
  Sheet15.Visible = xlSheetHidden
  Sheet15.Protect Password:=sPassword
  
  '80U-DD_C1 2024-25 Malli
  Sheet14.Unprotect Password:=sPassword
  Sheet14.Visible = xlSheetHidden
  Sheet14.Protect Password:=sPassword
  
  'Malli------23/04/2025
  
  
 '25/05/2025
  
  'If Mid(Range("IncD.TypeOfHP").Value, 1, 1) = "S"  Then
'If Mid(Range("IncD.TypeOfHP").Value, 1, 1) = "S" Or Mid(Range("IncD.TypeOfHP").Value, 1, 1) = "(" Then
'Ankita_21/01/2026
'If Mid(Range("IncD.TypeOfHP").Value, 1, 1) = "S" Or Mid(Range("IncD.TypeOfHP").Value, 1, 1) = "(" Or Range("IncD.TypeOfHP").Value = "" Then
'  Sheet16.Unprotect Password:=sPassword
'  Sheet16.Visible = xlSheetHidden
'  Sheet16.Protect Password:=sPassword
'Else
'   ThisWorkbook.Unprotect Password:=getmsgstate
'   Sheet16.Visible = xlSheetVisible
'   Sheet16.Protect Password:=sPassword
''   ThisWorkbook.Protect Password:=getmsgstate
'End If
'----------------------
  'Malli_05/05/2025
  ThisWorkbook.Unprotect Password:=sPassword
' Sheet18.Unprotect Password:=sPassword
  Sheet18.Visible = xlSheetHidden
  Sheet18.Protect Password:=sPassword
'----------------------------
  
  sPassword = EfilingCommon.getmsgstate
  Sheet5.Unprotect Password:=sPassword
  Sheet5.Range("BacValue").Value = 1
  Sheet5.Protect Password:=sPassword
  Sheet1.Unprotect Password:=sPassword
  Sheet1.Activate
  Sheet1.Unprotect Password:=sPassword          'Malli
  resetBacYes
  
'  ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 1
  Sheet1.Unprotect Password:=sPassword
  Sheet1.Activate
        ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 0  'Ankita_29/01/2026
        ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 1
  Sheet1.Protect Password:=sPassword
  
  Sheet1.Unprotect Password:=sPassword
  ThisWorkbook.Protect Password:=sPassword
  
'Else
'    sPassword = EfilingCommon.getmsgstate
'    Sheet5.Unprotect Password:=sPassword
'    Sheet5.Range("BacValue").Clear
'    Sheet5.Protect Password:=sPassword
'    ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 0
'     ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 0
'
'    'INC_E9 - INC_C9 - Need to remove below logic
'    'Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HCCFFCC)  ', "AU"
'    'Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = False
'
'
'
'
'      resetIncomeDetails
'      Sheet1.Protect Password:=sPassword
'      Sheet1.Activate
'    Exit Sub
'End If

'Ankita_02/06/2025
If Sheet5.Range("bacValue").Value = 1 Then

Dim Sheet1_Nature_cellscnt, i As Variant
Sheet1_Nature_cellscnt = Sheet1.Range("Others.NOI").Rows.count

        For i = 1 To Sheet1_Nature_cellscnt

                If Sheet1.Range("Others.NOI").item(i, 1).Value = "Sec 10(17)-Allowance MP/MLA/MLC" Then  'Ankita

                     Sheet1.Range("Others.NOI").item(i, 1).Value = "(Select)"

                     Sheet1.Range("Others.NOI").item(i, 1).Offset(0, 1).Value = ""

                     Sheet1.Range("Others.NOI").item(i, 1).Offset(0, 2).Value = ""

                 End If

        Next

End If
'.................................................................
    Sheet1.Protect Password:=sPassword
    Sheet1.Activate
    'Change-24.11.2022.102.10B
    Nature_ExemptDropdown
    '---end
End Sub
Sub BacNoValueChange()
On Error Resume Next
Dim answer As Integer
 
'Ankita_30/07/2026-----SIT-125961
If Range("sheet1.ReturnFileSec").Value = "139(4)-Belated" Then
    Exit Sub
End If

 'PAG_C4 2024-25 Bindu
'answer = MsgBox("Do you want to change value of Are you opting for new tax regime u/s 115BAC?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmation")
'If answer = vbYes Then
  ThisWorkbook.Unprotect Password:=sPassword
  Sheet4.Unprotect Password:=sPassword
  Sheet4.Visible = xlSheetVisible
  Sheet4.Protect Password:=sPassword
  Sheet12.Unprotect Password:=sPassword
  Sheet12.Visible = xlSheetVisible
  Sheet12.Protect Password:=sPassword
  Sheet9.Unprotect Password:=sPassword
  Sheet9.Visible = xlSheetVisible
  Sheet9.Protect Password:=sPassword
  sPassword = EfilingCommon.getmsgstate
  Sheet5.Unprotect Password:=sPassword
  Sheet5.Range("BacValue").Value = 2
  Sheet5.Protect Password:=sPassword
   
'   Sheet1.Unprotect Password:=sPassword
'   ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 1  'Ankita_29/01/2026
'   ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 0
'   Sheet1.Protect Password:=sPassword

   '80GGC_C1 2024-25 Bindu
  Sheet13.Unprotect Password:=sPassword
  Sheet13.Visible = xlSheetVisible
  Sheet13.Protect Password:=sPassword
  
  '80U-DD_C1 2024-25 Malli
  Sheet14.Unprotect Password:=sPassword
  Sheet14.Visible = xlSheetVisible
  Sheet14.Protect Password:=sPassword
  
  
  'Ankita_21/04/2025
  Sheet17.Unprotect Password:=sPassword
  Sheet17.Visible = xlSheetVisible
  Sheet17.Protect Password:=sPassword
  
  '80U-DD_C1 2024-25 Malli
  Sheet15.Unprotect Password:=sPassword
  Sheet15.Visible = xlSheetVisible
  Sheet15.Protect Password:=sPassword
  
  'If Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" Then
  'Ankita_21/01/2026
'  If Mid(Range("IncD.TypeOfHP").Value, 1, 1) <> "(" And Range("IncD.TypeOfHP").Value <> "" Then
' 'Malli------23/04/2025
'  Sheet16.Unprotect Password:=sPassword
'  Sheet16.Visible = xlSheetVisible
'  Sheet16.Protect Password:=sPassword
'
'  Else
'  Sheet16.Unprotect Password:=sPassword
'  Sheet16.Visible = xlSheetHidden
'  Sheet16.Protect Password:=sPassword
'  End If
  '-----------------------
  
  'Malli---------05/05/2025
 ' If Mid(Sheet1.Range(Sheet1.EmployerCategory1), 1, 3) <> "Not" Then
  If Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) <> "Not" Then
  Sheet18.Unprotect Password:=sPassword
  Sheet18.Visible = xlSheetVisible
  Sheet18.Protect Password:=sPassword
  Else
  Sheet18.Unprotect Password:=sPassword
  Sheet18.Visible = xlSheetHidden
  Sheet18.Protect Password:=sPassword
  End If
  
  '-------------------------
   
  Sheet1.Unprotect Password:=sPassword
  Sheet1.Activate
  resetBacNo
  Sheet1.Protect Password:=sPassword
 
   Sheet1.Unprotect Password:=sPassword
   Sheet1.Activate
   ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 1  'Ankita_29/01/2026
   ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 0
   Sheet1.Protect Password:=sPassword

   Sheet1.Unprotect Password:=sPassword 'Ankita_29/01/2026

  Sheet1.Activate
  ThisWorkbook.Protect Password:=sPassword
'Else
'sPassword = EfilingCommon.getmsgstate
'Sheet5.Unprotect Password:=sPassword
'Sheet5.Range("BacValue").Clear
'Sheet5.Protect Password:=sPassword
'    ActiveSheet.Shapes("BacYes").OLEFormat.Object.Value = 0
'     ActiveSheet.Shapes("BacNO").OLEFormat.Object.Value = 0
'
'     'INC_E9 INC_C9 2024-25 Bindu Need to comment below code
'    ' Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HCCFFCC)  ', "AU"
'    'Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = False
'
'      resetIncomeDetails
'      Sheet1.Activate
'
'  Exit Sub
'End If
'Change-24.11.2022.102.10C
Nature_ExemptDropdown
'---end

End Sub
Sub resetIncomeDetails()
Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet1.Range("IncD.InterestBorrowedCapital").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.InterestBorrowedCapital").MergeArea.Locked = False
    
    If Sheet5.Range("BacValue").Value <> 1 Then    '03-02-2026  Malli SIT-109371
    Range("IncD.Deduction16ic").MergeArea.Locked = False
    Range("IncD.Deduction16ic").MergeArea.Interior.Color = (&HCCFFCC)
    If (Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Cen" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Sta" Or Mid(Range("Sheet1.EmployerCategory1").Value, 1, 3) = "Pub") Then
    Range("IncD.Deduction16").MergeArea.Locked = False
    Range("IncD.Deduction16").MergeArea.Interior.Color = (&HCCFFCC)
    End If
    End If
    
     'Commented by Shrutika(19/04/2025)NewDev
'     Sheet1.Range("IncD.Section80C").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80C").MergeArea.Locked = False

   'Ankita_26/01/2026=======
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Type_80CCC").Interior.Color = (&HCCFFCC)
    Sheet1.Range("Type_80CCC").ClearContents
    Sheet1.Range("Type_80CCC").Locked = False
    
    
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Name_80CCC").Interior.Color = (&HCCFFCC)
    Sheet1.Range("Name_80CCC").ClearContents
    Sheet1.Range("Name_80CCC").Locked = False
     
    
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Amount_80CCC").Interior.Color = (&HCCFFCC)
    Sheet1.Range("Amount_80CCC").ClearContents
    Sheet1.Range("Amount_80CCC").Locked = False
     
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Type_80CCD1").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Type_80CCD1").ClearContents
'    Sheet1.Range("Type_80CCD1").Locked = False
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Name_80CCD1").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Name_80CCD1").ClearContents
'    Sheet1.Range("Name_80CCD1").Locked = False
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Amount_80CCD1").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Amount_80CCD1").ClearContents
'    Sheet1.Range("Amount_80CCD1").Locked = False
'
'    Sheet1.Range("Type_80CCD1b").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Type_80CCD1b").ClearContents
'    Sheet1.Range("Type_80CCD1b").Locked = False
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Name_80CCD1b").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Name_80CCD1b").ClearContents
'    Sheet1.Range("Name_80CCD1b").Locked = False
'
'    Sheet1.Unprotect Password:=getmsgstate
'    Sheet1.Range("Amount_80CCD1b").Interior.Color = (&HCCFFCC)
'    Sheet1.Range("Amount_80CCD1b").ClearContents
'    Sheet1.Range("Amount_80CCD1b").Locked = False
    '----------------

'    Sheet1.Range("IncD.Section80CCC").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = False
''    Sheet1.Range("IncD.Section80CCC").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = True
'    Range("IncD.Section80CCC").Value = 0

'-------------------------------------------------------------------------
    Sheet1.Range("IncD.Section80C").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80C").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80C").Formula = "=IF(BacValue=1,0,TotAmount.80C)"
'     Sheet1.Range("IncD.Section80CCC").Interior.Color = vbWhite  ', "AU"
'    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = True
'    Sheet1.Range("IncD.Section80CCC").Formula = "=IF(BacValue=1,0,TotAmount.80CC)"
'-------------------------------------------------------------------------------------
    'Added by Ankita_13/06/2025
    Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = False
    Sheet1.Range("Sheet1.AckNum").Interior.Color = (&HCCFFCC)
    Sheet1.Range("Sheet1.AckNum").Font.Color = vbBlack
    
    
    '-----------------------------------------------------------------
      'Ankita_26/01
     Sheet1.Range("IncD.Section80CCD_SE").Interior.Color = (&HCCFFCC)  ', "AU"
     Sheet1.Range("IncD.Section80CCD_SE").MergeArea.Locked = False
     Sheet1.Range("IncD.Section80CCD1B_SE").Interior.Color = (&HCCFFCC)  ', "AU"
     Sheet1.Range("IncD.Section80CCD1B_SE").MergeArea.Locked = False
    'Sheet1.Range("IncD.Section80DD").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80DDB").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = False
'Commented by Shrutika(19/04/2025)NewDev
'     Sheet1.Range("IncD.Section80E").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80E").MergeArea.Locked = False
'     Sheet1.Range("IncD.Section80EE").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80EE").MergeArea.Locked = False
'     Sheet1.Range("IncD.Section80EEA").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = False
'     Sheet1.Range("IncD.Section80EEB").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = False

'------------------------------------------------------------------------------

    Sheet1.Range("IncD.Section80E").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80E").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80E").Formula = "=IF(BacValue=1,0,TotAmt.80E)"
    Sheet1.Range("IncD.Section80EE").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80EE").Formula = "=IF(BacValue=1,0,TotAmt.80EE)"
    Sheet1.Range("IncD.Section80EEA").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80EEA").Formula = "=IF(BacValue=1,0,TotAmt.80EEA)"
    Sheet1.Range("IncD.Section80EEB").Interior.Color = vbWhite  ', "AU"
    Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = True
    Sheet1.Range("IncD.Section80EEB").Formula = "=IF(BacValue=1,0,TotAmt.80EEB)"
'-----------------------------------------------------------------------------
     Sheet1.Range("IncD.Section80GG").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80GG").MergeArea.Locked = False
    'Ayush_23/05 Ankita_13/06/2025
'    Sheet1.Range("Sheet1.AckNum").Interior.Color = (&HCCFFCC)  ', "AU"
'    Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = False
    
     'Sheet1.Range("IncD.Section80GGC").Interior.Color = (&HCCFFCC)  ', "AU"
    'Sheet1.Range("IncD.Section80GGC").MergeArea.Locked = False
     Sheet1.Range("IncD.Section80TTA").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = False
     Sheet1.Range("IncD.Section80TTB").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = False
    'Inserted
    If Sheet1.Range("FP_Value").Value > 0 Then
            Sheet1.Range("IncD.LessDeduction57").MergeArea.Interior.Color = (&HCCFFCC)
            Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = False
    End If
                If (Range("sheet1.DOB").Value) <> "" Then
                If calculateAge(Trim(Range("sheet1.DOB").Value)) >= 60 Then
               
                 Range("IncD.Section80TTA").Value = ""
                 Range("IncD.Section80TTA").MergeArea.Locked = True
                 Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HD8D8D8)
                If Sheet5.Range("BacValue").Value = 2 Then
                 Range("IncD.Section80TTB").MergeArea.Locked = False
                 Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HCCFFCC)
                 End If
                 Else
                 If Sheet5.Range("BacValue").Value = 2 Then
                 Range("IncD.Section80TTA").MergeArea.Locked = False
                 Range("IncD.Section80TTA").MergeArea.Interior.Color = (&HCCFFCC)
                 End If
                 Range("IncD.Section80TTB").Value = ""
                 Range("IncD.Section80TTB").MergeArea.Locked = True
                 Range("IncD.Section80TTB").MergeArea.Interior.Color = (&HD8D8D8)
                End If
                End If
    
    
    'closed
    
    
     'Sheet1.Range("IncD.Section80U").Interior.Color = (&HCCFFCC)  ', "AU"
    Sheet1.Range("IncD.Section80U").MergeArea.Locked = True
    '-------------------Malli
     'Sheet1.Range("SELECT80DD").Interior.Color = (&HCCFFCC) ', "AU"   'Malli
    'Sheet1.Range("SELECT80DD").MergeArea.Locked = False
    '-----------------------------
'    Sheet1.Range("SELECT80DD").Value = "(Select)"
     Sheet1.Range("SELECT80DDS").Interior.Color = (&HCCFFCC) ', "AU"
    Sheet1.Range("SELECT80DDS").MergeArea.Locked = False
    
    'Added by Shrutika(21/04/2025)NewDev
    Sheet1.Range("Sheet1.Specified_Disease").Interior.Color = (&HCCFFCC) ', "AU"
    Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Locked = False
    '-----------------------------------------------------------------
    
'    Sheet1.Range("SELECT80DDB").Value = "(Select)"
'MAlli--------------
'    Sheet1.Range("SELECT80U").Interior.Color = (&HCCFFCC) ', "AU"
'    Sheet1.Range("SELECT80U").MergeArea.Locked = False
'----------------------
'    Sheet1.Range("SELECT80U").Value = "(Select)"


    Sheet1.Protect Password:=EfilingCommon.getmsgstate
    Sheet1.Activate
End Sub

Sub resetBacNo()
Application.EnableEvents = False
'Sheet1.Range("IncD.LessDeduction57").Interior.Color = (&HCCFFCC)  ', "AU"
'Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = False

'Ankita
' Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'    ccount = 0
'    mIntCells = Range("Others.NOI_1").count
'    Set rangecells = Range("Others.NOI_1").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).Value = "(Select)"
'    Next
'
'     mIntCells = Range("Others.Amount_1").count
'    Set rangecells = Range("Others.Amount_1").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).Value = ""
'    Next
     resetNatureexempt
     
     Sheet19.LockUnlock_24bHP  'Ankita_29/01/2026
'      mIntCells = Range("Nature_Others_1").count
'    Set rangecells = Range("Nature_Others_1").Cells
'    For mIntCtr = 1 To mIntCells
'    Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
'    rangecells.item(mIntCtr).Value = "Not Applicable"
'    Next
     Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
     
    Range("IncD.Deduction16").Value = 0
    
    Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
    Range("IncD.Deduction16ic").Value = 0
    
    'Ankita_22/02/2026=====
    Sheet19.Unprotect Password:=sPassword
    Sheet19.Range("HP.IntOnBorwCap1").Value = 0
    
    Sheet19.Unprotect Password:=sPassword
    Sheet19.Range("HP.IntOnBorwCap2").Value = 0

     resetIncomeDetails
    
     
     Application.EnableEvents = True
End Sub


Sub resetBacYes()
Application.EnableEvents = False
' Dim rangecells As Range
'    Dim mIntCells As Long
'    Dim mIntCtr As Long
'    Dim ccount As Long
'
' Added by Ankita on 12/12/2024
        resetNatureexempt
        Sheet19.LockUnlock_24bHP      'Ankita_29/01/2026==

    'INC_E9 - INC_C9 2024-25 Bindu Need to comment the below lines
'    Sheet1.Range("IncD.LessDeduction57").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("IncD.LessDeduction57").MergeArea.Locked = True
'    Range("IncD.LessDeduction57").Value = 0
    
  '   Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
'     If Range("IncD.TypeOfHP").Value = "Self Occupied" Then
'    Sheet1.Range("IncD.InterestBorrowedCapital").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("IncD.InterestBorrowedCapital").MergeArea.Locked = True
'    Sheet1.Range("IncD.InterestBorrowedCapital").Value = ""
'    Else
'    Sheet1.Range("IncD.InterestBorrowedCapital").Interior.Color = (&HCCFFCC)   ', "AU"
'    Sheet1.Range("IncD.InterestBorrowedCapital").MergeArea.Locked = False
'       End If
   ' Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
   
     Sheet1.Range("IncD.Deduction16").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Deduction16").MergeArea.Locked = True
    Range("IncD.Deduction16").Value = 0
    
   '  Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
     Sheet1.Range("IncD.Deduction16ic").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Deduction16ic").MergeArea.Locked = True
    Range("IncD.Deduction16ic").Value = 0
    
    ' Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
     Sheet1.Range("IncD.Section80C").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80C").MergeArea.Locked = True
     Range("IncD.Section80C").Value = 0
   '  Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
   'Ankita_26/01
'     Sheet1.Range("IncD.Section80CCC").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("IncD.Section80CCC").MergeArea.Locked = True
'    Range("IncD.Section80CCC").Value = 0
    
    
    'Ankita_05/01/2026=====================
    
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Type_80CCC").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("Type_80CCC").ClearContents
    Sheet1.Range("Type_80CCC").Locked = True
        
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Name_80CCC").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("Name_80CCC").ClearContents
    Sheet1.Range("Name_80CCC").Locked = True
     
    Sheet1.Unprotect Password:=getmsgstate
    Sheet1.Range("Amount_80CCC").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("Amount_80CCC").ClearContents
    Sheet1.Range("Amount_80CCC").Locked = True
    
'    Sheet1.Range("Type_80CCD1").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Type_80CCD1").Locked = True
'    Sheet1.Range("Type_80CCD1").ClearContents
''    Sheet1.Unprotect Password:=getmsgstate
'
'    Sheet1.Range("Name_80CCD1").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Name_80CCD1").Locked = True
'    Sheet1.Range("Name_80CCD1").ClearContents
''    Sheet1.Unprotect Password:=getmsgstate
'
'    Sheet1.Range("Amount_80CCD1").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Amount_80CCD1").Locked = True
'    Sheet1.Range("Amount_80CCD1").ClearContents
''    Sheet1.Unprotect Password:=getmsgstate
'
'    Sheet1.Range("Type_80CCD1b").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Type_80CCD1b").Locked = True
'    Sheet1.Range("Type_80CCD1b").ClearContents
''    Sheet1.Unprotect Password:=getmsgstate
'
'    Sheet1.Range("Name_80CCD1b").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Name_80CCD1b").Locked = True
'    Sheet1.Range("Name_80CCD1b").ClearContents
''    Sheet1.Unprotect Password:=getmsgstate
'
'    Sheet1.Range("Amount_80CCD1b").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Amount_80CCD1b").Locked = True
'    Sheet1.Range("Amount_80CCD1b").ClearContents
''   Sheet1.Unprotect Password:=getmsgstate
    
'         'Ankita_16/04/2026========
     Sheet1.Unprotect Password:=getmsgstate
     Sheet1.Range("pran_new").Interior.ColorIndex = 15   ', "AU"
     Sheet1.Range("pran_new").Locked = True
     Sheet1.Range("pran_new").ClearContents
    
    If Sheet1.Range("Type_80CCC").Locked = True Then
    Sheet1.Range("Type_80CCC").ClearNotes
    End If

    '-----------------------------------------
    'Ankita_16/05/2025
    'Commented by Ankita as per AY26-27 DE sheet v0.1

'    Sheet1.Range("sheet1.PRAN").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("sheet1.PRAN").MergeArea.Locked = True
'    Sheet1.Range("sheet1.PRAN").MergeArea.ClearContents
    
'    Range("sheet1.PRAN").Value = 0

  '  Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
  'Ankita_26/01
      Sheet1.Range("IncD.Section80CCD_SE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCD_SE").MergeArea.Locked = True
    Range("IncD.Section80CCD_SE").Value = 0
    Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
  'Ankita_26/01
  'Ankita_17/04/2026
    Sheet1.Range("IncD.Section80CCD1B_SE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80CCD1B_SE").MergeArea.Locked = True
    Range("IncD.Section80CCD1B_SE").Value = 0
    Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
    'Sheet1.Range("IncD.Section80DD").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80DD").MergeArea.Locked = True
    'Range("IncD.Section80DD").Value = 0
  'Malli-------------
'    Sheet1.Range("SELECT80DD").Interior.ColorIndex = 15
'    Sheet1.Range("SELECT80DD").MergeArea.Locked = True
'   Sheet1.Range("SELECT80DD").Value = "(Select)"
    'Malli-----------------
    
   ' Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
     Sheet1.Range("IncD.Section80DDB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80DDB").MergeArea.Locked = True
    Range("IncD.Section80DDB").Value = 0
    
    Sheet1.Range("SELECT80DDS").Interior.ColorIndex = 15 ', "AU"
    Sheet1.Range("SELECT80DDS").MergeArea.Locked = True
    Sheet1.Range("SELECT80DDS").Value = "(Select)"
    
    'Added by Shrutika(21/04/2025)NewDev
    Sheet1.Range("Sheet1.Specified_Disease").Interior.ColorIndex = 15 ', "AU"
    Sheet1.Range("Sheet1.Specified_Disease").MergeArea.Locked = True
    Sheet1.Range("Sheet1.Specified_Disease").Value = "(Select)"
    '-------------------------------------------------------------------
    
  '  Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet1.Range("IncD.Section80E").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80E").MergeArea.Locked = True
    Range("IncD.Section80E").Value = 0
   ' Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet1.Range("IncD.Section80EE").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EE").MergeArea.Locked = True
    Range("IncD.Section80EE").Value = 0
    'Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet1.Range("IncD.Section80EEA").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EEA").MergeArea.Locked = True
    Range("IncD.Section80EEA").Value = 0
   ' Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet1.Range("IncD.Section80EEB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80EEB").MergeArea.Locked = True
    Range("IncD.Section80EEB").Value = 0
   ' Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
    Range("IncD.Section80GG").Value = 0
  '  Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
    Sheet1.Range("IncD.Section80GG").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80GG").MergeArea.Locked = True
    
    'newly added by Ankita on 19/05/2025
    'Ankita_13/06/2025
    'Range("Sheet1.AckNum").Value = ""
     Sheet1.Unprotect Password:=sPassword
     Sheet1.Range("Sheet1.AckNum").Interior.Color = (&HD8D8D8)
     Sheet1.Range("Sheet1.AckNum").Font.Color = (&HD8D8D8)
     Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = True
     
'    Sheet1.Range("Sheet1.AckNum").Interior.Color = (&HD8D8D8)
'    Sheet1.Range("Sheet1.AckNum").Font.Color = (&HD8D8D8)
'    Sheet1.Range("Sheet1.AckNum").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("Sheet1.AckNum").MergeArea.Locked = True
   
    
    'INC_C26 NEED TO COMMENT BELOW LINE 2024-25 Bindu
   ' Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
'    Sheet1.Range("IncD.Section80GGC").Interior.ColorIndex = 15   ', "AU"
'    Sheet1.Range("IncD.Section80GGC").MergeArea.Locked = True
'    Range("IncD.Section80GGC").Value = 0
    'Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
    '--end
    
    Sheet1.Range("IncD.Section80TTA").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80TTA").MergeArea.Locked = True
    Range("IncD.Section80TTA").Value = 0
    
    
   ' Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
     Sheet1.Range("IncD.Section80TTB").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80TTB").MergeArea.Locked = True
    Range("IncD.Section80TTB").Value = 0
    'Sheet1.Unprotect Password:=EfilingCommon.getmsgstate
     ' Sheet1.Range("IncD.Section80U").Interior.ColorIndex = 15   ', "AU"
    Sheet1.Range("IncD.Section80U").MergeArea.Locked = True
    'Range("IncD.Section80U").Value = 0
    '#D
    
    'Malli-------------
'    Sheet1.Range("SELECT80U").Interior.ColorIndex = 15 ', "AU"
'    Sheet1.Range("SELECT80U").MergeArea.Locked = True
'    Sheet1.Range("SELECT80U").Value = "(Select)"
'
     'Malli---------------------------
     
'     'Ankita_22/02/2026=======
'     Sheet19.Unprotect Password:=sPassword
'     Sheet19.Range("HP.IntOnBorwCap1").Value = 0
     
    Sheet1.Unprotect Password:=sPassword
    Dim rangecellsnatureofincome As Range
    Dim rangecellsamount As Range
    Dim mIntCellsnatureofincome As Long
    Dim mIntCtrnatureofincome As Long
    Dim ccountnatureofincome As Long
    ccountnatureofincome = 0
    mIntCellsnatureofincome = Range("Others.NOI").count
     Set rangecellsnatureofincome = Range("Others.NOI").Cells
    Set rangecellsamount = Range("Others.Amount").Cells
    For mIntCtrnatureofincome = 1 To mIntCellsnatureofincome
        If rangecellsnatureofincome.item(mIntCtrnatureofincome).Value = "Sec 10(17)-Allowance MP/MLA/MLC" Then 'Ankita
            rangecellsamount.item(mIntCtrnatureofincome).Value = "0"
        End If
    Next
    
         'Ankita_22/02/2026=======
     Sheet19.Unprotect Password:=sPassword
     Sheet19.Range("HP.IntOnBorwCap1").Value = 0
     
     Sheet19.Unprotect Password:=sPassword
     Sheet19.Range("HP.IntOnBorwCap2").Value = 0
    
    'Commented by Ankita on 12/12/2024
'    ccount = 0
'    mIntCells = Range("Others.NOI_1").count
'    Set rangecells = Range("Others.NOI_1").Cells
'    Dim countrycd As Variant
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).Value = "(Select)"
'    Next
'     mIntCells = Range("Others.Amount_1").count
'    Set rangecells = Range("Others.Amount_1").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).Value = ""
'    Next
'      mIntCells = Range("Nature_Others_1").count
'    Set rangecells = Range("Nature_Others_1").Cells
'    For mIntCtr = 1 To mIntCells
'    rangecells.item(mIntCtr).Value = "Not Applicable"
'    Next
    Application.EnableEvents = True
End Sub

Sub LinkCheckBoxes()
Dim chk As CheckBox
Dim lCol As Long

'INC_E43
'lCol = 7 'number of columns to the right for link

'INC_C43
lCol = 6
For Each chk In ActiveSheet.CheckBoxes
   With chk
      .LinkedCell = _
         .TopLeftCell.Offset(0, lCol).Address
   End With
Next chk
End Sub


Function setName()
    Dim fname As String
                Dim mname As String
                Dim sname As String
                fname = Trim(UCase(Sheet1.Range("sheet1.FirstName").Value))
                mname = Trim(UCase(Sheet1.Range("sheet1.MiddleName").Value))
                sname = Trim(UCase(Sheet1.Range("sheet1.SurNameOrOrgName").Value))
                Dim name As String
                If Trim(fname) <> "" Then
                    name = name & fname
                End If
                 If Trim(mname) <> "" Then
                    If Trim(name) <> "" Then
                        name = name & " "
                    End If
                    name = name & mname
                End If
                 If Trim(sname) <> "" Then
                     If Trim(name) <> "" Then
                        name = name + " "
                    End If
                    name = name + sname
                End If
setName = name
End Function




Function ChkMaxInclusiveDate(Maxinclusive As Variant, MaxInclusiveDate As Variant) As Boolean
'Note:
'' both date must be in format yyyy-mm-dd
    ChkMaxInclusiveDate = True
    If Len(Maxinclusive) > 0 Then
        If Mid(Maxinclusive, 1, 4) > Mid(MaxInclusiveDate, 1, 4) Then
            ChkMaxInclusiveDate = False
            Exit Function
        Else
            If Mid(Maxinclusive, 1, 4) = Mid(MaxInclusiveDate, 1, 4) Then
                If (Mid(Maxinclusive, 6, 2) > Mid(MaxInclusiveDate, 6, 2)) Then
                    ChkMaxInclusiveDate = False
                    Exit Function
                ElseIf ((Mid(Maxinclusive, 6, 2) = Mid(MaxInclusiveDate, 6, 2))) Then
                    If (Mid(Maxinclusive, 9, 2) > Mid(MaxInclusiveDate, 9, 2)) Then
                        ChkMaxInclusiveDate = False
                        Exit Function
                   End If
                End If
            End If
        End If
    End If
End Function
'Malli
Function insertRowUnderSectionWithFormula80GGC(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
'Malli
If HOIflag = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
        'Set RngHOI = Selection.Offset(1).Resize(rowsize:=1)
        rrCount = 0
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
                Range("E" & rr.row).Locked = False
                Range("E" & rr.row).Interior.Color = (&HCCFFCC)
                Range("F" & rr.row).Locked = False
                Range("F" & rr.row).Interior.Color = (&HCCFFCC)
                Range("G" & rr.row).Locked = False
                Range("G" & rr.row).Interior.Color = (&HCCFFCC)
                Range("J" & rr.row).Locked = False
                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
                Range("K" & rr.row).Locked = False
                Range("K" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("K" & rr.row).Locked = True
'                Range("K" & rr.row).Interior.Color = "&HD8D8D8"
            End If
            rrCount = rrCount + 1
        Next
    End If
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula80GGC = nRows
'----------------Lock Password-------------------START---
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
    ActiveSheet.Protect Password:=sPassword
End Function

'Ankita_09/04
Function insertRowUnderSectionWithFormula24b(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim sPassword As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue, sAddingRangeValue, sNewCellValueNext As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    ''''''''''malli
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If HOIflag = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        'Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
        Set RngHOI = Selection.Offset(1).Resize(rowsize:=1)
        rrCount = 1
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
                Range("D" & rr.row).Locked = False
                Range("D" & rr.row).Interior.Color = (&HCCFFCC)
                Range("E" & rr.row).Locked = False
                Range("E" & rr.row).Interior.Color = (&HCCFFCC)
                Range("F" & rr.row).Locked = False
                Range("F" & rr.row).Interior.Color = (&HCCFFCC)
                Range("G" & rr.row).Locked = False
                Range("G" & rr.row).Interior.Color = (&HCCFFCC)
                Range("H" & rr.row).Locked = False
                Range("H" & rr.row).Interior.Color = (&HCCFFCC)
                Range("I" & rr.row).Locked = False
                Range("I" & rr.row).Interior.Color = (&HCCFFCC)
                Range("J" & rr.row).Locked = False
                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
                Range("K" & rr.row).Locked = False
                Range("K" & rr.row).Interior.Color = (&HCCFFCC)
                Range("L" & rr.row).Locked = False
                Range("L" & rr.row).Interior.Color = (&HCCFFCC)
                
            End If
            rrCount = rrCount + 1
        Next
        
    End If

    
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    sNewCellValueNext = ""
    sAddingRangeValue = ""
    
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula24b = nRows

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function


'Ankita_09/04
Function insertRowUnderSectionWithFormula80C(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim sPassword As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue, sAddingRangeValue, sNewCellValueNext As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    ''''''''''malli
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If HOIflag = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        'Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
        Set RngHOI = Selection.Offset(1).Resize(rowsize:=1)
        rrCount = 1
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
                Range("D" & rr.row).Locked = False
                Range("D" & rr.row).Interior.Color = (&HCCFFCC)
                Range("E" & rr.row).Locked = False
                Range("E" & rr.row).Interior.Color = (&HCCFFCC)
                Range("F" & rr.row).Locked = False
                Range("F" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("G" & rr.row).Locked = False
'                Range("G" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("H" & rr.row).Locked = False
'                Range("H" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("I" & rr.row).Locked = False
'                Range("I" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("J" & rr.row).Locked = False
'                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("K" & rr.row).Locked = False
'                Range("K" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("L" & rr.row).Locked = False
'                Range("L" & rr.row).Interior.Color = (&HCCFFCC)
'
            End If
            rrCount = rrCount + 1
        Next
        
    End If

    
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    sNewCellValueNext = ""
    sAddingRangeValue = ""
    
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula80C = nRows

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function


'Ankita_09/04
Function insertRowUnderSectionWithFormula80D(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0, Optional HOIflag As Variant = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim sPassword As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue, sAddingRangeValue, sNewCellValueNext As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    ''''''''''malli
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
     
'----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----

    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
    If HOIflag = 1 Then
        Dim rr, rrCount As Variant
        Dim RngHOI As Range
        'Set RngHOI = Selection.Resize(rowsize:=nRows + 1)
        Set RngHOI = Selection.Offset(1).Resize(rowsize:=1)
        rrCount = 1
        For Each rr In RngHOI.Rows
            If Not rrCount = 0 Then
'                Range("D" & rr.row).Locked = False
'                Range("D" & rr.row).Interior.Color = (&HCCFFCC)
                Range("E" & rr.row).Locked = False
                Range("E" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("F" & rr.row).Locked = False
'                Range("F" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("G" & rr.row).Locked = False
'                Range("G" & rr.row).Interior.Color = (&HCCFFCC)
                Range("H" & rr.row).Locked = False
                Range("H" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("I" & rr.row).Locked = False
'                Range("I" & rr.row).Interior.Color = (&HCCFFCC)
                Range("J" & rr.row).Locked = False
                Range("J" & rr.row).Interior.Color = (&HCCFFCC)
'                Range("K" & rr.row).Locked = False
'                Range("K" & rr.row).Interior.Color = (&HCCFFCC)
                Range("L" & rr.row).Locked = False
                Range("L" & rr.row).Interior.Color = (&HCCFFCC)

            End If
            rrCount = rrCount + 1
        Next
        
    End If

    
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    sNewCellValueNext = ""
    sAddingRangeValue = ""
    
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            
            sNewCellValueNext = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + 1) 'C6 + 1 => C7
            If nRows > 1 Then
            sAddingRangeValue = sNewCellValueNext + ":" + sNewCellValue
            ElseIf nRows = 1 Then
            sAddingRangeValue = sNewCellValue
            End If
            
'            If gridRange(iCount) = "TDS2.TdsCredit" Or _
'                gridRange(iCount) = "TDS3.TdsCredit" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
'
'            If gridRange(iCount) = "TDSoth.DeductedYear" Or _
'                gridRange(iCount) = "TDSoth2.DeductedYear" Then
'                Range(sAddingRangeValue).value = "(Select)"
'            End If
            
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula80D = nRows

'----------------Lock Password-------------------START---
  ActiveSheet.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
    Application.EnableEvents = True
End Function















Function checkfieldspecialcharacter_Trans(field As Variant) As Boolean
    Dim i, j As Long
    checkfieldspecialcharacter_Trans = True
    Dim arr As Variant
    arr = Array("@", "|", "_", "*", "!", "&", "#", "~", ";", "?", "\", ":", "(", ")", "+", "=", "{", "}", "[", "]", "^", "%", "$", """", "'", ">", "<") 'Array("&", """", "'", ">", "<")
    For i = 1 To Len(field)
        For j = 0 To UBound(arr)
        If Mid(field, i, 1) = arr(j) Then
            checkfieldspecialcharacter_Trans = False
            Exit Function
        End If
        Next
    Next
End Function

'SAI--------------
'Function StateMatchesPin(StateCode As Variant, Pincode_pag2) As Boolean
'    StateMatchesPin = False
'    Dim State1, State2 As Variant
'
'
'    State1 = Application.VLookup(Pincode_pag2, Sheet5.Range("Duplicate_V"), 2, False)
'
'    State2 = Application.VLookup(Pincode_pag2, Sheet5.Range("Duplicate_V"), 3, False)
'
'    If IsError(State1) Or IsError(State2) Then
'    StateMatchesPin = False
'    Exit Function
'    End If
'
'    If UCase(Mid(StateCode, 4)) = State1 Or UCase(Mid(StateCode, 4)) = State2 Then
'    'MsgBox UCase(Mid(StateCode, 4))
'    StateMatchesPin = True
'    End If
'
'End Function


'Sai>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'SAI--------------

Function StateMatchesPin(StateCode As Variant, Pincode_pag2) As Boolean
    StateMatchesPin = False
    Dim state1, state2 As Variant


    state1 = Application.VLookup(Pincode_pag2, Sheet5.Range("Duplicate_V"), 2, False)

    state2 = Application.VLookup(Pincode_pag2, Sheet5.Range("Duplicate_V"), 3, False)

    If IsError(state1) Or IsError(state2) Then
    StateMatchesPin = False
    Exit Function
    End If

    If UCase(Mid(StateCode, 4)) = state1 Or UCase(Mid(StateCode, 4)) = state2 Then
    'MsgBox UCase(Mid(StateCode, 4))
    StateMatchesPin = True
    End If

End Function
Function state_Validation(ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String) As Boolean
    state_Validation = True
    Dim PinCode As Range
            If Range(pin_trgt_adrs).Value <> "" Then
              Dim state1 As String
                state1 = Application.VLookup(Range(pin_trgt_adrs).Value, Sheet5.Range("All_PinCode_V"), 2, False)
                If Not (state1 = UCase(Mid(Range(state_trgt_adrs).Value, 4)) Or EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).Value, Range(pin_trgt_adrs).Value)) Then
                    state_Validation = False
                End If
            End If

End Function

''SAI---------------
' Function state_Validation(ByVal pin_trgt_adrs As Variant, ByVal state_trgt_adrs As Variant)
'    state_Validation = True
'    Application.EnableEvents = False
'    Dim PinCode As Range
'
'            If Range(pin_trgt_adrs).Value <> "" Then
'              Dim State1 As Variant
'                State1 = Application.VLookup(Range(pin_trgt_adrs).Value, Sheet5.Range("All_Pincode_V"), 2, False)
'                If Not (State1 = UCase(Mid(Range(state_trgt_adrs).Value, 4)) Or EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).Value, Range(pin_trgt_adrs).Value)) Then
'                Range(pin_trgt_adrs).Value = ""
'                End If
'            End If
'
'
'End Function

   'ramya1234--------------
'Function PinState_codes_validation(ByVal pin_trgt_adrs As Variant, ByVal state_trgt_adrs As Variant, ByVal country_trgt_adrs As Variant, ByVal Zip_trgt_adrs As Variant)
'    PinState_codes_validation = True
'    Application.EnableEvents = False
'    Dim PinCode As Range
'    Dim DupFlag As Boolean
'    DupFlag = False
'
'                 For Each PinCode In Sheet5.Range("Duplicate_Pincode_List")
'                    If Range(pin_trgt_adrs).Value = PinCode Then
'                        DupFlag = True
'                        Exit For
'                    End If
'                 Next PinCode
'
'                 If DupFlag = True Then
'                    If Range(state_trgt_adrs).Value <> "" Then
'                        If Not EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).Value, Range(pin_trgt_adrs).Value) Then
'                            Range(state_trgt_adrs).Value = "(Select)"
'                        End If
'                    End If
'                 End If
'
'                 If DupFlag = False Then
'                    Dim StateName
'                    Dim StateCode
'
'                    StateName = Application.VLookup(Range(pin_trgt_adrs).Value, Sheet5.Range("All_Pincode_V"), 2, False)
'                    If IsError(StateName) Then
'                        If Range(pin_trgt_adrs).Value <> "" Then
'                        MsgBox ("Invalid Pincode")
'                     End If
'                        Application.EnableEvents = False
'                        pin_trgt_adrs.Value = ""
'                        Application.EnableEvents = True
'
'                    Else
'                        StateCode = Application.VLookup(StateName, Sheet5.Range("CS:CT"), 2, False)
'                        Range(state_trgt_adrs).Value = StateCode
'                        Range(country_trgt_adrs).Value = "91-INDIA"
'
'
'                         ActiveSheet.Unprotect Password:=EfilingCommon.getmsgstate
'                        If Mid(Range(state_trgt_adrs).Value, 1, 2) = "99" Then
'                                Range(Zip_trgt_adrs).Interior.Color = (&HCCFFCC)
'                                Range(Zip_trgt_adrs).Value = ""
'                                Range(Zip_trgt_adrs).Locked = False
'                        ElseIf Mid(Range(state_trgt_adrs).Value, 1, 2) <> "99" Then
'                                Range(Zip_trgt_adrs).Interior.Color = (&HD8D8D8)
'                                Range(Zip_trgt_adrs).Value = ""
'                                Range(Zip_trgt_adrs).Locked = True
'                        End If
'                    End If
'                 End If
'
'
'
'End Function

   'ramya1234--------------
   
   
Function PinState_codes_validation(sheetname1 As Worksheet, ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String, ByVal country_trgt_adrs As String, ByVal Zip_trgt_adrs As String)
Dim DupFlag As Boolean
Dim PinCode1 As String
DupFlag = False
PinCode1 = Application.IsError(Application.VLookup(sheetname1.Range(pin_trgt_adrs).Value, Sheet5.Range("Duplicate_Pincode_List"), 1, False))
If PinCode1 = "False" Then DupFlag = True

If DupFlag = True Then
   If sheetname1.Range(state_trgt_adrs).Value <> "" Then
       If Not EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).Value, sheetname1.Range(pin_trgt_adrs).Value) Then
           sheetname1.Range(state_trgt_adrs).Value = "(Select)"
       End If
   End If
End If
                 
If DupFlag = False Then
   Dim StateName
   Dim StateCode

   StateName = Application.VLookup(sheetname1.Range(pin_trgt_adrs).Value, Sheet5.Range("All_PinCode_V"), 2, False)
   If IsError(StateName) Then
       If sheetname1.Range(pin_trgt_adrs).Value <> "" Then
           
           MsgBox ("Invalid Pincode")
           sheetname1.Range(pin_trgt_adrs).ClearContents
       End If
   Else
       StateCode = Application.VLookup(StateName, Sheet5.Range("CS:CT"), 2, False)
       sheetname1.Range(state_trgt_adrs).Value = StateCode
       sheetname1.Range(country_trgt_adrs).Value = "91-INDIA"
       sheetname1.Unprotect Password:=EfilingCommon.getmsgstate
   End If
End If


End Function

'Ankita_20/01/2026=

Function PINstate_ModualValidation(sheetname2 As Worksheet, ByVal pin_trgt_adrs As String, ByVal state_trgt_adrs As String) As Boolean
    PINstate_ModualValidation = True
    
    Dim PinCode As Range
    
    sheetname2.Activate
        
            If Range(pin_trgt_adrs).Value <> "" Then
                Dim state1 As String
                
                state1 = Application.IsError(Application.VLookup(Range(pin_trgt_adrs).Value, Sheet5.Range("All_PinCode_V"), 2, False))
                
                If state1 = True Then
                    PINstate_ModualValidation = False
                Else
                    state1 = Application.VLookup(Range(pin_trgt_adrs).Value, Sheet5.Range("All_PinCode_V"), 2, False)
                    If Not (state1 = UCase(Mid(Range(state_trgt_adrs).Value, 4)) Or EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).Value, Range(pin_trgt_adrs).Value)) Then
                    PINstate_ModualValidation = False
                    End If
                 End If
                End If
                   
End Function

 

Function PinState_codes_validation_1234(ByVal pin_trgt_adrs As Variant, ByVal state_trgt_adrs As Variant, ByVal country_trgt_adrs As Variant, ByVal Zip_trgt_adrs As Variant)
    PinState_codes_validation = True
    Application.EnableEvents = False
    Dim PinCode As Range
    Dim DupFlag As Boolean
    DupFlag = False
                 
                 For Each PinCode In Sheet5.Range("Duplicate_Pincode_List")
                    If Range(pin_trgt_adrs).Value = PinCode Then
                        DupFlag = True
                        Exit For
                    End If
                 Next PinCode

                 If DupFlag = True Then
                    If Range(state_trgt_adrs).Value <> "" Then
                        If Not EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).Value, Range(pin_trgt_adrs).Value) Then
                            Range(state_trgt_adrs).Value = "(Select)"
                        End If
                    End If
                 End If
                 
                 If DupFlag = False Then
                    Dim StateName
                    Dim StateCode

                    StateName = Application.VLookup(Range(pin_trgt_adrs).Value, Sheet5.Range("All_Pincode_V"), 2, False)
                    If IsError(StateName) Then
                        If Range(pin_trgt_adrs).Value <> "" Then
                        MsgBox ("Invalid Pincode")
                     End If
                        Application.EnableEvents = False
                        pin_trgt_adrs.Value = ""
                        Application.EnableEvents = True

                    Else
                        StateCode = Application.VLookup(StateName, Sheet5.Range("CS:CT"), 2, False)
                        Range(state_trgt_adrs).Value = StateCode
                        Range(country_trgt_adrs).Value = "91-INDIA"
                        
                      
                         ActiveSheet.Unprotect Password:=EfilingCommon.getmsgstate
                        If Mid(Range(state_trgt_adrs).Value, 1, 2) = "99" Then
                                Range(Zip_trgt_adrs).Interior.Color = (&HCCFFCC)
                                Range(Zip_trgt_adrs).Value = ""
                                Range(Zip_trgt_adrs).Locked = False
                        ElseIf Mid(Range(state_trgt_adrs).Value, 1, 2) <> "99" Then
                                Range(Zip_trgt_adrs).Interior.Color = (&HD8D8D8)
                                Range(Zip_trgt_adrs).Value = ""
                                Range(Zip_trgt_adrs).Locked = True
                        End If
                    End If
                 End If
     
    
     
End Function
 'sai---------------
Function SetStateCountryImmovableAsset_M(Tval As Variant, Tadd As Variant, CallType As Variant)
On Error Resume Next
Application.EnableEvents = False


'------------------Un Protect--------------------
   sPassword = EfilingCommon.getmsgstate
   Sheet43.Unprotect Password:=sPassword
'------------------Un Protect Ends--------------------
    
    Dim rangecellsState, rangecellsCountry, rangecellsPin, rangecellsZip As Range
    Dim mIntCells, q  As Long
    Dim PrevVal As Variant

    Set rangecellsState = Range("SchAL.A.Address_State").Cells
    Set rangecellsCountry = Range("SchAL.A.Address_Country").Cells
    Set rangecellsPin = Range("SchAL.A.Address_Pin").Cells
    Set rangecellsZip = Range("SchAL.A.Address_Zip").Cells

     'MsgBox rangecellsZip.Address

    mIntCells = Range("SchAL.A.Description").count

If CallType = "S" Then
    'If UCase(Tval) = "99-STATE OUTSIDE INDIA" Then
    'Mid(UCase(Tval), 1, 2) = "99"
      If Mid(UCase(Tval), 1, 2) = "99" Then
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, 1) = "(Select)"
       Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = True
        Sheet43.Range(Tadd).Offset(0, 3).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 3).Locked = False
        GoTo endfd
    ElseIf UCase(Tval) = "(SELECT)" Then
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, 1) = "(Select)"
        Sheet43.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 1).Locked = False
        
        Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = False
        
        Sheet43.Range(Tadd).Offset(0, 3) = ""
        Sheet43.Range(Tadd).Offset(0, 3).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 3).Locked = False
        GoTo endfd
    Else
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, 1) = "91-INDIA"
        
      If Sheet43.Range(Tadd).Offset(0, 2) = "" Then
        Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = False
        End If
        
        Sheet43.Range(Tadd).Offset(0, 3) = ""
        Sheet43.Range(Tadd).Offset(0, 3).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 3).Locked = True
        GoTo endfd
    End If
Else

    If UCase(Tval) = "(SELECT)" Then
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, -1) = "(Select)"
        Sheet43.Range(Tadd).Offset(0, 1) = ""
        Sheet43.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 1).Locked = False
        Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = True
        GoTo endfd
    ElseIf UCase(Tval) <> "91-INDIA" Then
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, -1) = "99-STATE OUTSIDE INDIA"
        Sheet43.Range(Tadd).Offset(0, 1) = ""
        Sheet43.Range(Tadd).Offset(0, 1).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 1).Locked = True
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = False
        GoTo endfd
    Else
        Sheet43.Range(Tadd).Select
        Sheet43.Range(Tadd).Offset(0, -1) = "(Select)"
        Sheet43.Range(Tadd).Offset(0, 1) = ""
        Sheet43.Range(Tadd).Offset(0, 1).Interior.Color = (&HCCFFCC)
        Sheet43.Range(Tadd).Offset(0, 1).Locked = False
        Sheet43.Range(Tadd).Offset(0, 2) = ""
        Sheet43.Range(Tadd).Offset(0, 2).Interior.Color = (&HD8D8D8)
        Sheet43.Range(Tadd).Offset(0, 2).Locked = True
        GoTo endfd
    End If

End If

    For q = 1 To mIntCells
        If UCase(rangecellsState.item(q).Value) = "99-STATE OUTSIDE INDIA" Then
            rangecellsCountry.item(q).Value = "(Select)"
            rangecellsPin.item(q).Value = ""
            rangecellsZip.item(q).Value = ""
        ElseIf UCase(rangecellsState.item(q).Value) = "(SELECT)" Then
            rangecellsCountry.item(q).Value = "(Select)"
            rangecellsPin.item(q).Value = ""
            rangecellsZip.item(q).Value = ""
        Else
            rangecellsCountry.item(q).Value = "91-INDIA"
            rangecellsPin.item(q).Value = ""
            rangecellsZip.item(q).Value = ""
        End If
        
        If UCase(rangecellsCountry.item(q).Value) <> "91-INDIA" Then
            rangecellsState.item(q).Value = "99-STATE OUTSIDE INDIA"
            rangecellsPin.item(q).Value = ""
            rangecellsZip.item(q).Value = ""
        Else
            rangecellsCountry.item(q).Value = "(Select)"
            rangecellsPin.item(q).Value = ""
            rangecellsZip.item(q).Value = ""
        End If
        
    Next
endfd:
'----------------Lock Password-------------------START---
  Sheet43.Protect Password:=sPassword
'----------------Lock Password-------------------END-----
Application.EnableEvents = True
End Function
 'sai---------------
Function PinState_codes_validation_80G(ByVal pin_trgt_adrs As Variant, ByVal state_trgt_adrs As Variant)
    PinState_codes_validation_80G = True
    Application.EnableEvents = False
    Dim PinCode As Range
    Dim DupFlag As Boolean
    DupFlag = False
                 
                 For Each PinCode In Sheet5.Range("Duplicate_Pincode_List")
                    If Range(pin_trgt_adrs).Value = PinCode Then
                        DupFlag = True
                        Exit For
                    End If
                 Next PinCode

                 If DupFlag = True Then
                    If Range(state_trgt_adrs).Value <> "" Then
                        If Not EfilingCommon.StateMatchesPin(Range(state_trgt_adrs).Value, Range(pin_trgt_adrs).Value) Then
                            Range(state_trgt_adrs).Value = "(Select)"
                        End If
                    End If
                 End If
                 
                 If DupFlag = False Then
                    Dim StateName
                    Dim StateCode

                    StateName = Application.VLookup(Range(pin_trgt_adrs).Value, Sheet5.Range("All_Pincode_V"), 2, False)
                     If IsError(StateName) Then
                        If Range(pin_trgt_adrs).Value <> "" Then
                        MsgBox ("Invalid Pincode")
                        End If
                        Application.EnableEvents = False
                        pin_trgt_adrs.Value = ""
                        Application.EnableEvents = True

                     Else
                        StateCode = Application.VLookup(StateName, Sheet5.Range("CS:CT"), 2, False)
                        Range(state_trgt_adrs).Value = StateCode
                        
                     End If
                 End If
     
    
     
End Function


'Ankita_16/04/2025
Function CheckPAN(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckPAN = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPAN = False
            Exit Function
        End If
                
        If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckPAN = False
            Exit Function
        End If
            
        
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPAN = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckPAN = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPAN = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacterdot(PAN) Then
        CheckPAN = False
            Exit Function
        End If
               
    End If
End Function
  
'Ankita_16/03/2026
Function CheckPAN1(PAN As Variant) As Boolean
On Error Resume Next
'PAN : Consist of 10 characters
'PAN format: First Five Alphabets, next 4 digits, then Alphabet.

    CheckPAN1 = True
    If Len(PAN) > 0 Then
        If Not ChkAlphabet(Mid(PAN, 1, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 2, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 3, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
                
        If Not ChkAlphabet(Mid(PAN, 4, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
            
        
        If Not ChkAlphabet(Mid(PAN, 5, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
        If Not IsNumeric(Mid(PAN, 6, 4)) Then
            CheckPAN1 = False
            Exit Function
        End If
        If Not ChkAlphabet(Mid(PAN, 10, 1)) Then
            CheckPAN1 = False
            Exit Function
        End If
        
        If Not CheckSpecialCharacterdot(PAN) Then
        CheckPAN1 = False
            Exit Function
        End If
               
    End If
End Function
  

Public Function CheckSpecialCharacterdot(emailid As Variant) As Boolean
On Error Resume Next
    Dim specialCharArray As Variant
    Dim iCharCount, iSpecialChar As Long
    
    CheckSpecialCharacterdot = True
    specialCharArray = Array(".")
    For iCharCount = 1 To Len(emailid)
        For iSpecialChar = 0 To UBound(specialCharArray)
        If Mid(emailid, iCharCount, 1) = specialCharArray(iSpecialChar) Then
            CheckSpecialCharacterdot = False
            Exit Function
        End If
        Next
    Next
End Function

'Newly added by sai on 22/04/2025
Function insertRowUnderSectionWithFormula_80C(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula_80C = nRows
'----------------Lock Password-------------------START---
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
    ActiveSheet.Protect Password:=sPassword
End Function

'Newly added by sai on 22/04/2025
Function insertRowUnderSectionWithFormula_80CCC_IncomeDetails(Optional nOfRows As Long = 0, Optional isExtension As Boolean = False, Optional index As Long = 0) As Long
On Error Resume Next
    Dim nRows As Long
    Dim gridRange() As String
    Dim srange As Range
    Dim sRangeAddress, sTempCellValue, sTempFirstCellValue, sTempLastCellValue, sNewCellValue, sRangeValue As String
    Dim iCount, sTempCellValueInt As Long
    Dim x As Long
    Application.EnableEvents = False
    nRows = 0
    ActiveCell.EntireRow.Select
    If nOfRows = 0 Then
        nRows = Application.InputBox( _
                            prompt:="Enter the number of rows you want to add below selected cell", _
                            Title:="Add Rows below the selected cell", _
                            Default:=1, _
                            Type:=1)
         
        If nRows = 0 Then
         Exit Function
        End If
    Else
        nRows = nOfRows
    End If
    '----------------Unlock Password-------------------START---
   sPassword = EfilingCommon.getmsgstate
   ActiveSheet.Unprotect Password:=sPassword
'----------------Unlock Password-------------------END-----
    x = Sheets(ActiveSheet.name).UsedRange.Rows.count 'lastcell fixup
    
    Selection.Resize(rowsize:=2).Rows(2).EntireRow.Resize(rowsize:=nRows).Insert Shift:=xlDown
    Selection.AutoFill Selection.Resize(rowsize:=nRows + 1), xlFillDefault
    Selection.Offset(1).Resize(nRows).EntireRow.SpecialCells(xlConstants).ClearContents
    
'Method: DefineName after no of rows insert Method
    sRangeAddress = ""
    sTempCellValue = ""
    sTempFirstCellValue = ""
    sTempLastCellValue = ""
    sNewCellValue = ""
    sRangeValue = ""
    iCount = 0
    sTempCellValueInt = 0
    gridRange = Split(EfilingCommon.DefinedgridNameRange, "||")
            
    If Not isExtension Then
        
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount)).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount), RefersTo:=srange, Visible:=True
        Next
    Else
    
        For iCount = 0 To UBound(gridRange)
            sTempCellValue = Replace(Range(gridRange(iCount) & index).AddressLocal, "$", "") '$C$4:$C$6 => C4:C6
            
            sTempFirstCellValue = Mid(sTempCellValue, 1, InStr(1, sTempCellValue, ":") - 1) 'C4:C6 => C4
            sTempLastCellValue = Mid(sTempCellValue, InStr(1, sTempCellValue, ":") + 1, Len(sTempCellValue)) 'C4:C6 => C6
                   
            sTempCellValueInt = EfilingCommon.onlyDigits(sTempLastCellValue, "I") 'C6 => 6
            sNewCellValue = Replace(sTempLastCellValue, sTempCellValueInt, sTempCellValueInt + nRows) 'C6 + nRows(eg:1) => C7
            
            sRangeValue = sTempFirstCellValue + ":" + sNewCellValue 'C4:C7
            Set srange = Range(sRangeValue)  '$C$4:$C$7
            ThisWorkbook.Names.add name:=gridRange(iCount) & index, RefersTo:=srange, Visible:=True
        Next
    
    End If
    insertRowUnderSectionWithFormula_80CCC_IncomeDetails = nRows
'----------------Lock Password-------------------START---
'----------------Lock Password-------------------END-----

    Application.EnableEvents = True
    ActiveSheet.Protect Password:=sPassword
End Function


Sub tttttttttttttttttttt()
Application.EnableEvents = True
End Sub


