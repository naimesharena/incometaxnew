Attribute VB_Name = "mdInitializer"
Public imported
Option Explicit

Sub Intialize()
On Error Resume Next
Dim strMsgOnStart, strMsgOnStart1, strMsgOnStart2, vbMessgaeCaption As String

    Application.EnableEvents = True
    With Application
        .Caption = ""
        
        'Cut
        .CommandBars("Worksheet Menu Bar").Controls("Cut").Controls("Rows").Enabled = False
        .CommandBars("Row").Controls("Cut").Enabled = False
        
        .CommandBars("Column").Controls("Cut").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Cut").Controls("Columns").Enabled = False
        
        .CommandBars("Standard").Controls.item("Cut").Enabled = False
        .CommandBars("Edit").Controls.item("Cut").Enabled = False
        .CommandBars("Cell").Controls.item("Cut").Enabled = False

        'Insert
        .CommandBars("Row").Controls("Insert").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Insert").Controls("Rows").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Insert...").Controls("Rows").Enabled = False
        
        .CommandBars("Column").Controls("Insert").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Insert...").Controls("Columns").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Insert").Controls("Columns").Enabled = False

        .CommandBars("Cell").Controls.item("Insert...").Enabled = False
        
        'Delete
        .CommandBars("Row").Controls("Delete").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Delete").Controls("Rows").Enabled = False
        .CommandBars("Worksheet Menu Bar").Controls("Delete...").Controls("Rows").Enabled = False

        .CommandBars("Column").Controls("Delete").Enabled = True
        .CommandBars("Worksheet Menu Bar").Controls("Delete").Controls("Columns").Enabled = True
        .CommandBars("Worksheet Menu Bar").Controls("Delete...").Controls("Columns").Enabled = True
        .CommandBars("Cell").Controls.item("Delete...").Enabled = True
         
        'Other Properties Set
        .OnKey "^x", "DisableCut"
        .WindowState = xlMaximized
        .DisplayFormulaBar = False
        .Calculation = xlAutomatic
        .CellDragAndDrop = False
        .EnableEvents = True
        .MaxChange = 0.001
    End With
    
     UserForm2.Show
'     If Sheet1.Range("sheet1.PAN").Value <> "" Then
'        imported = 1
'     Else
'        imported = 0
'     End If
'
'
'
'     While imported = 0
'        imported = -1
'        UserForm3.Show
'        If imported = -1 Then
'            ThisWorkbook.Close
'            imported = 0
'        End If
'
'     Wend
     
    'Messages shown to user on Workbook opening
'    vbMessgaeCaption = "ITR 1: AY 2020-21 - Important Note"
'
'    strMsgOnStart = "1. This utility is for Assessment Year 2020-21." + Chr(10) + Chr(13) + Chr(10) + Chr(13) & "2. Green cells are for data entry." + Chr(10) + Chr(13) + Chr(10) + Chr(13) & "3. Red labels indicate mandatory fields." & Chr(13) & vbNewLine & "4. Caution: Do not use CTRL X or cut paste while entering the data.. "
'    MsgBox strMsgOnStart, vbInformation, vbMessgaeCaption
'    Sheet1.Activate
    
'Intial level setting for Workbook---------------------------------------------------------End---
End Sub

Sub DisableCut()
'MsgBox "You are not allowed to cut a cell, instead use copy function.", vbInformation, "Error"
fmsgbox ("You are not allowed to cut a cell, instead use copy function.")
End Sub






