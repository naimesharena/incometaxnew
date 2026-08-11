Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Public ws As Variant

Private Sub Workbook_BeforeClose(Cancel As Boolean)
 sPassword = EfilingCommon.getmsgstate
ActiveWorkbook.Protect Password:=sPassword
End Sub

Private Sub Workbook_Deactivate()
Application.CommandBars("Ply").Enabled = True
End Sub

Private Sub Workbook_Open()
'Sheet1.Shapes("Button 47").Visible = False
'Sheet1.Shapes("Button 50").Visible = False
'Sheet1.Shapes("Button 51").Visible = False
'Sheet2.Shapes("Button 675").Visible = False
'Sheet11.Shapes("Button 35").Visible = False
'Sheet3.Shapes("Button 127").Visible = False
'Sheet9.Shapes("Button 7").Visible = False
'Sheet4.Shapes("Button 17").Visible = False
'Sheet12.Shapes("Button 7").Visible = False
Sheet1.Shapes("Button 47").Visible = True
Sheet1.Shapes("Button 50").Visible = True
Sheet1.Shapes("Button 51").Visible = True
Sheet2.Shapes("Button 675").Visible = True
Sheet11.Shapes("Button 35").Visible = True
Sheet3.Shapes("Button 127").Visible = True
Sheet9.Shapes("Button 7").Visible = True
Sheet4.Shapes("Button 17").Visible = True
Sheet12.Shapes("Button 33").Visible = True
Sheet1.Range("taxcheck").Value = ""

    mdInitializer.Intialize
    For Each ws In Worksheets
    Application.DisplayFormulaBar = False
    Next
    
    
'        'Ankita_25/07/2025========
'
'Dim due1 As String
'Dim sheet1_ReturnFileSec_duplicate As Variant
'
'due1 = "2025-09-15"
'
'
''If Dformat(Sheet3.Range("Ver.Date"), "") > due1 Then
''Sheet1.Range("sheet1.ReturnFileSec").Value = "139(4)-Belated"
''Else
''Sheet1.Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date"
''End If  '==============================
'
''For AY 2025-26_Sadineni_28_07_2025
''*************************************************************************************************
'sheet1_ReturnFileSec_duplicate = Sheet1.Range("sheet1.ReturnFileSec").Value
'
''If Dformat(Sheet3.Range("Ver.Date"), "") <= due1 Then
''    Sheet1.Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date"
''
''    If Sheet1.Range("sheet1.ReturnFileSec").Value <> "139(5)-Revised" Then
''        ElseIf Dformat(Sheet3.Range("Ver.Date"), "") > due1 Then
''        Sheet1.Range("sheet1.ReturnFileSec").Value = sheet1_ReturnFileSec_duplicate
''    End If
''
''End If
''*************************************************************************************************
'
'

    sPassword = EfilingCommon.getmsgstate
    ThisWorkbook.Protect Password:=sPassword
    Application.CommandBars("Ply").Enabled = False
    'ActiveSheet.Shapes("Button 47").ControlFormat.Enabled = False
'ActiveSheet.Shapes("Button 51").ControlFormat.Enabled = False
End Sub
