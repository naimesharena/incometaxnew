Attribute VB_Name = "ImportExcel"
Option Explicit
Public retrievedSheetName As Variant
Dim DestBook As Workbook, SrcBook As Workbook
Dim end_Per10080G, end_PerNO5080G, end_PerYES10080G, end_Per5080G, end_ITimport, end_BAimport, end_OtherEI, end_OtherEI1, end_OtherEI2, end_TDS1import, end_TCSimport, end_TDS2import, end_TDS3import As Variant
Dim rngname_Per10080G, rngname_PerNO5080G, rngname_PerYES10080G, rngname_Per5080G, rngname_ITimport, rngname_OtherEI, rngname_OtherEI1, rngname_OtherEI2, rngname_BAimport, rngname_TDS1import, rngname_TCSimport, rngname_TDS2import, rngname_TDS3import As Variant
Public rngname_80GGA As Variant
Sub ImportPreviousVersion()
On Error Resume Next
    
    Dim rangenamearr() As Variant
    Dim Filename As Variant
    Dim dfilename, ndfilename, newfilename As Variant
    Dim flag As Boolean
    Dim add As Variant
    Dim rname, ws As Variant
    Dim a As Long
    Dim destadd As Variant
    
    Dim cnt, dcnt As Long
    Dim newrname As String
    Dim sfirstbound, supperbound, sTEMP, dfirstbound, dTemp, dupperbound, ddTemp As Variant
    
    InitProgBar
    ProgressFrameCaption = "Importing Excel"
    mainProcCaption = "Importing"
    noOfProcessMain = 2
    
    flag = True
    'MsgBox "Please use this functionality only to 'Import from previous version' of current assessment year.", vbOKOnly, "Alert"
    fmsgbox "Please use this functionality only to 'Import from previous version' of current assessment year."
    
    'MsgBox "After Importing, Please check the  utility, to ensure all rows are imported.", vbOKOnly, "Alert"
    fmsgbox "After Importing, Please check the  utility, to ensure all rows are imported."
    Filename = cmdFileDialog()
    If Filename <> "" Then UserForm1.Show vbModeless
    If Not Filename = "" Then
        Filename = Split(Filename, "\")
        newfilename = Filename(UBound(Filename))

        cnt = 0
        Application.ScreenUpdating = False
        Set SrcBook = Workbooks.Open(newfilename)
        Err.Clear
        Set DestBook = ThisWorkbook
        
        If Err = "1004" Then
          ProgressBarHide
        End If
        
Dim targetSheet As Worksheet
Set targetSheet = DestBook.Worksheets(1)
Dim sourceSheet As Worksheet
Set sourceSheet = SrcBook.Worksheets(1)



     If (targetSheet.Range("AN1").Value) <> (sourceSheet.Range("AN1").Value) Then
        fmsgbox "Invalid ITR or Assessment Year"
        SrcBook.Save
        SrcBook.Close
          Exit Sub
          
        Else
     
       End If

        dfilename = Split((DestBook.FullName), "\")
        ndfilename = dfilename(UBound(dfilename))
        noOfProcessSub = 10
        UpdateProgressBar
        UpdateProgressBar
        If newfilename <> ndfilename Then
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            For Each rname In Workbooks(newfilename).Names
                newrname = rname.name
                newrname = Mid(newrname, InStr(1, newrname, "!") + 1)
                
                
                
            Application.ScreenUpdating = False
                
          If newrname = "TDSal.TAN" Then
                    subProcCaption = "Importing TDS1"
                    sfirstbound = SrcBook.Sheets("TDS").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(rname.name).count
                    dcnt = DestBook.Sheets("TDS").Range(rname.name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_TDS1import
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_TDS1import)
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                   
                End If
                If newrname = "TCS.TAN" Then
                    subProcCaption = "Importing TCS"
                    sfirstbound = SrcBook.Sheets("TCS").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TCS").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TCS").Range(rname.name).count
                    dcnt = DestBook.Sheets("TCS").Range(rname.name).count
                    DestBook.Sheets("TCS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TCS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_TCSimport
'                       If SrcBook.Sheets("TCS").Range("TCS.AmtClaimedBySpouse").Locked = False Then
'                            DestBook.Sheets("TCS").Unprotect Password:=EfilingCommon.getmsgstate
'                            DestBook.Sheets("TCS").Range("TCS.amtClaimedBySpouse").Locked = False
'                            DestBook.Sheets("TCS").Range("TCS.amtClaimedBySpouse").Interior.Color = "&HCCFFCC"
''                           DestBook.Sheets("TDS").Range("TCS.6panspouse").Locked = False
''                           DestBook.Sheets("TDS").Range("TCS.6panspouse").Interior.Color = "&HCCFFCC"
'                            DestBook.Sheets("Income Details").Protect Password:=EfilingCommon.getmsgstate
'                            InsertRowsToImport (cnt - dcnt)
'                            Call ExendRangeNameToTable(cnt - dcnt, rngname_TCSimport)
'
'                            SrcBook.Sheets("TCS").Range(rname.name).Copy
'                            DestBook.Sheets("TCS").Range(rname.name).PasteSpecial xlValues
                       
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_TCSimport)
                        SrcBook.Sheets("TCS").Range(rname.name).Copy
                        DestBook.Sheets("TCS").Range(rname.name).PasteSpecial xlValues
                     'End If
                    Else
                        SrcBook.Sheets("TCS").Range(rname.name).Copy
                        DestBook.Sheets("TCS").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                   
                End If
'Sch TDS2
                If newrname = "TDSoth.TAN" Then
                    subProcCaption = "Importing TDS2"
                    sfirstbound = SrcBook.Sheets("TDS").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(rname.name).count
                    dcnt = DestBook.Sheets("TDS").Range(rname.name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_TDS2import
                        
'                        If SrcBook.Sheets("TDS").Range("TDSoth.6panspouse").Locked = False And SrcBook.Sheets("TDS").Range("TDSoth.AmtClaimedBySpouse").Locked = False Then
'
'                            DestBook.Sheets("TDS").Unprotect Password:=EfilingCommon.getmsgstate
'                            DestBook.Sheets("TDS").Range("TDSoth.amtClaimedBySpouse").Locked = False
'                            DestBook.Sheets("TDS").Range("TDSoth.amtClaimedBySpouse").Interior.Color = "&HCCFFCC"
'                            DestBook.Sheets("TDS").Range("TDSoth.6panspouse").Locked = False
'                            DestBook.Sheets("TDS").Range("TDSoth.6panspouse").Interior.Color = "&HCCFFCC"
'                            DestBook.Sheets("Income Details").Protect Password:=EfilingCommon.getmsgstate
'                            InsertRowsToImport (cnt - dcnt)
'                            Call ExendRangeNameToTable(cnt - dcnt, rngname_TDS2import)
'                            SrcBook.Sheets("TDS").Range(rname.name).Copy
'                            DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
'
'                        Else
'
                            InsertRowsToImport (cnt - dcnt)
                            Call ExendRangeNameToTable(cnt - dcnt, rngname_TDS2import)
                            SrcBook.Sheets("TDS").Range(rname.name).Copy
                            DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                        
                        'End If
                        
                    Else
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
               
                End If
                
                
                
  'NEW Sch TDS3 2018-19
                If newrname = "TDS26QB.PAN" Then
                    subProcCaption = "Importing TDS3"
                    sfirstbound = SrcBook.Sheets("TDS").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(rname.name).count
                    dcnt = DestBook.Sheets("TDS").Range(rname.name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_TDS3import
'                        InsertRowsToImport (cnt - dcnt)
'                        Call ExendRangeNameToTable(cnt - dcnt, rngname_TDS3import)
'                        SrcBook.Sheets("TDS").Range(rname.name).Copy
'                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
'                         If SrcBook.Sheets("TDS").Range("TDS26QB.6panspouse").Locked = False And SrcBook.Sheets("TDS").Range("TDS26QB.AmtClaimedBySpouse").Locked = False Then
'
'                            DestBook.Sheets("TDS").Unprotect Password:=EfilingCommon.getmsgstate
'                            DestBook.Sheets("TDS").Range("TDS26QB.amtClaimedBySpouse").Locked = False
'                            DestBook.Sheets("TDS").Range("TDS26QB.amtClaimedBySpouse").Interior.Color = "&HCCFFCC"
'                            DestBook.Sheets("TDS").Range("TDS26QB.6panspouse").Locked = False
'                            DestBook.Sheets("TDS").Range("TDS26QB.6panspouse").Interior.Color = "&HCCFFCC"
'                            DestBook.Sheets("Income Details").Protect Password:=EfilingCommon.getmsgstate
'                            InsertRowsToImport (cnt - dcnt)
'                            Call ExendRangeNameToTable(cnt - dcnt, rngname_TDS3import)
'                            SrcBook.Sheets("TDS").Range(rname.name).Copy
'                            DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
'
'                        Else
                        
                            InsertRowsToImport (cnt - dcnt)
                            Call ExendRangeNameToTable(cnt - dcnt, rngname_TDS3import)
                            SrcBook.Sheets("TDS").Range(rname.name).Copy
                            DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                        
                        'End If
                    Else
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
               
                End If
                
                
                
                
                
                
                
                
                
                
                
    'Sch IT
                If newrname = "TaxP.BSRCode" Then
                subProcCaption = "Importing IT"
                    sfirstbound = SrcBook.Sheets("TDS").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("TDS").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("TDS").Range(rname.name).count
                    dcnt = DestBook.Sheets("TDS").Range(rname.name).count
                    DestBook.Sheets("TDS").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("TDS").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_ITimport
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_ITimport)
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("TDS").Range(rname.name).Copy
                        DestBook.Sheets("TDS").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                   
                End If
                
'Sch 80G- Section A
                If newrname = "Per10080G.DoneeName" Then
                subProcCaption = "Importing 80GA"
                    sfirstbound = SrcBook.Sheets("80G").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
        
                    dfirstbound = DestBook.Sheets("80G").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
        
                    cnt = SrcBook.Sheets("80G").Range(rname.name).count
                    dcnt = DestBook.Sheets("80G").Range(rname.name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_Per10080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_Per10080G)
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                    
                End If
        
'Sch 80G- Section B
                If newrname = "PerNO5080G.DoneeName" Then
                subProcCaption = "Importing 80GB"
                    sfirstbound = SrcBook.Sheets("80G").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80G").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                
                    cnt = SrcBook.Sheets("80G").Range(rname.name).count
                    dcnt = DestBook.Sheets("80G").Range(rname.name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_PerNO5080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_PerNO5080G)
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                   
                End If
        
        
'Sch 80G- Section C

                If newrname = "PerYES10080G.DoneeWithPanName" Then
                subProcCaption = "Importing 80GC"
                    sfirstbound = SrcBook.Sheets("80G").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80G").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("80G").Range(rname.name).count
                    dcnt = DestBook.Sheets("80G").Range(rname.name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_PerYES10080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_PerYES10080G)
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                   
                End If
        
        
'Sch 80G- Section D
                If newrname = "Per5080G.DoneeWithPanName" Then
                noOfProcessSub = "Importing 80GD"
                    sfirstbound = SrcBook.Sheets("80G").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80G").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("80G").Range(rname.name).count
                    dcnt = DestBook.Sheets("80G").Range(rname.name).count
                    DestBook.Sheets("80G").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80G").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_Per5080G
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_Per5080G)
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80G").Range(rname.name).Copy
                        DestBook.Sheets("80G").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                   
                End If
                
                
                'Sch 80GGA-
                If newrname = "RelevantClauseClaimed_80GGA" Then
                noOfProcessSub = "Importing 80GGA"
                    sfirstbound = SrcBook.Sheets("80GGA").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("80GGA").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("80GGA").Range(rname.name).count
                    dcnt = DestBook.Sheets("80GGA").Range(rname.name).count
                    DestBook.Sheets("80GGA").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("80GGA").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_80GGA
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_80GGA)
                        SrcBook.Sheets("80GGA").Range(rname.name).Copy
                        DestBook.Sheets("80GGA").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("80GGA").Range(rname.name).Copy
                        DestBook.Sheets("80GGA").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                   
                End If
              'Sch BA
                If newrname = "SchBA.IFSC" Then
                subProcCaption = "Importing BA"
                    sfirstbound = SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).count
                    dcnt = DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).count
                    DestBook.Sheets("Taxes Paid and Verification").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Taxes Paid and Verification").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_BAimport
                        InsertRowsToImport (cnt - dcnt)
                        Call ExendRangeNameToTable(cnt - dcnt, rngname_BAimport)
                        SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Copy
                        DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Copy
                        DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                  
                End If
                
              'Sch BA1
               ' If newrname = "Sheet9.IBAN" Then
               '     sfirstbound = SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Address
               '     sTEMP = Split(sfirstbound, "$")
               '     supperbound = UBound(sTEMP)
               '     sTEMP = sTEMP(UBound(sTEMP))
               '
               '     dfirstbound = DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).Address
               '     dTemp = Split(dfirstbound, "$")
               '     dupperbound = UBound(dTemp)
               '     ddTemp = dTemp(UBound(dTemp))
               '
               '     cnt = SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Count
               '     dcnt = DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).Count
               '     DestBook.Sheets("Taxes Paid and Verification").Activate
               '     If (cnt - dcnt) > 0 Then
               '         DestBook.Sheets("Taxes Paid and Verification").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
               '         setTblinfo_IBAN_1
               '         InsertRowsToImport (cnt - dcnt)
               '         Call ExendRangeNameToTable(cnt - dcnt, rngname_IBAN)
               '         SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Copy
               '         DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).PasteSpecial xlValues
               '     Else
               '         SrcBook.Sheets("Taxes Paid and Verification").Range(rname.name).Copy
               '         DestBook.Sheets("Taxes Paid and Verification").Range(rname.name).PasteSpecial xlValues
               '     End If
               ' End If

              'Sch Other Income
              If newrname = "Others.NOI_1" Then
                subProcCaption = "Importing Exempt Income : Any Other"
                    sfirstbound = SrcBook.Sheets("Income Details").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Income Details").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("Income Details").Range(rname.name).count
                    dcnt = DestBook.Sheets("Income Details").Range(rname.name).count
                    DestBook.Sheets("Income Details").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Income Details").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_OthersNOI_1
                        InsertRowsToImport ((cnt - dcnt) / 16)
                        Call ExendRangeNameToTable((cnt - dcnt) / 16, rngname_OtherEI1)
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                  
                End If
                
                If newrname = "Others.NOI_2" Then
                subProcCaption = "Importing Exempt Income : Any Other"
                    sfirstbound = SrcBook.Sheets("Income Details").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Income Details").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("Income Details").Range(rname.name).count
                    dcnt = DestBook.Sheets("Income Details").Range(rname.name).count
                    DestBook.Sheets("Income Details").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Income Details").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_OthersNOI_2
                        InsertRowsToImport ((cnt - dcnt) / 16)
                        Call ExendRangeNameToTable((cnt - dcnt) / 16, rngname_OtherEI2)
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                  
                End If
                
                If newrname = "Others.NOI" Then
                subProcCaption = "Importing Exempt Income : Any Other"
                    sfirstbound = SrcBook.Sheets("Income Details").Range(rname.name).Address
                    sTEMP = Split(sfirstbound, "$")
                    supperbound = UBound(sTEMP)
                    sTEMP = sTEMP(UBound(sTEMP))
                    
                    dfirstbound = DestBook.Sheets("Income Details").Range(rname.name).Address
                    dTemp = Split(dfirstbound, "$")
                    dupperbound = UBound(dTemp)
                    ddTemp = dTemp(UBound(dTemp))
                    
                    cnt = SrcBook.Sheets("Income Details").Range(rname.name).count
                    dcnt = DestBook.Sheets("Income Details").Range(rname.name).count
                    DestBook.Sheets("Income Details").Activate
                    If (cnt - dcnt) > 0 Then
                        DestBook.Sheets("Income Details").Range(dTemp(UBound(dTemp) - 1) & dTemp(UBound(dTemp))).Select
                        setTblinfo_OtherEI
                        InsertRowsToImport ((cnt - dcnt) / 12)
                        Call ExendRangeNameToTable((cnt - dcnt) / 12, rngname_OtherEI)
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    Else
                        SrcBook.Sheets("Income Details").Range(rname.name).Copy
                        DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                    End If
                    UpdateProgressBar
                  
                End If
        
           Next
            
            Application.EnableEvents = False
            subProcCaption = "Importing Static Contents"
            noOfProcessSub = Workbooks(newfilename).Names.count
            For Each rname In Workbooks(newfilename).Names
              retrievedSheetName = getSheetName(rname)
              
            If retrievedSheetName <> "" Then
              If Range(rname.name) <> "sheet1.ReturnFileSec1" Then
              If DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True Then
                If SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = False Then
                   DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate
                    
                    
                    If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Locked = False
                    Else
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = False
                    End If
                    
                    'DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = False
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).Interior.Color = (&HCCFFCC)
                    
                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                    
                    DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate
                Else
                    
                End If
                
              ElseIf DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = False Then
'                SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
'                DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                
                If SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True Then
                   DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate


                    If DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeCells Then
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Locked = True
                    Else
                        DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True
                    End If

                    'DestBook.Worksheets(retrievedSheetName).Range(rname.name).Locked = True
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).Interior.Color = (&HD8D8D8)

                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues

                    DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate


                End If
                
                
                    SrcBook.Worksheets(retrievedSheetName).Range(rname.name).Copy
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).PasteSpecial xlValues
                    DestBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Value = SrcBook.Worksheets(retrievedSheetName).Range(rname.name).MergeArea.Value
            End If
            End If
              
              If DestBook.Worksheets(retrievedSheetName).Range(rname.name).EntireRow.Hidden = True Then
                    If SrcBook.Worksheets(retrievedSheetName).Range(rname.name).EntireRow.Hidden = False Then
                        DestBook.Worksheets(retrievedSheetName).Unprotect Password:=getmsgstate
                            DestBook.Worksheets(retrievedSheetName).Range(rname.name).EntireRow.Hidden = False
                        DestBook.Worksheets(retrievedSheetName).Protect Password:=getmsgstate
                    End If
              End If
            End If
            UpdateProgressBar
            Next
            
            
            If SrcBook.Sheets("Income Details").Range("HASZIP").MergeArea.Locked = True Then
                        DestBook.Sheets("Income Details").Unprotect Password:=EfilingCommon.getmsgstate
                        DestBook.Sheets("Income Details").Range("HASZIP").MergeArea.Locked = True
                        DestBook.Sheets("Income Details").Range("sheet1.ZipCode").MergeArea.Locked = True
                        DestBook.Sheets("Income Details").Protect Password:=EfilingCommon.getmsgstate

                        ElseIf SrcBook.Sheets("Income Details").Range("HASZIP").Value = "Y" And SrcBook.Sheets("Income Details").Range("HASZIP").MergeArea.Locked = False Then
                            DestBook.Sheets("Income Details").Unprotect Password:=EfilingCommon.getmsgstate
                            DestBook.Sheets("Income Details").Range("HASZIP").MergeArea.Locked = False
                            DestBook.Sheets("Income Details").Range("HASZIP").MergeArea.Interior.Color = "&HCCFFCC"
                            DestBook.Sheets("Income Details").Range("HASZIP").Value = "Y"

                            DestBook.Sheets("Income Details").Range("sheet1.ZipCode").MergeArea.Locked = False
                            DestBook.Sheets("Income Details").Range("sheet1.ZipCode").MergeArea.Value = "XXXXXX"
                            DestBook.Sheets("Income Details").Range("sheet1.ZipCode").MergeArea.Locked = True

                            'DestBook.Sheets("Income Details").Range(rname.name).MergeArea.Interior.Color = "&HCCFFCC"
                            'SrcBook.Sheets("Income Details").Range(rname.name).Copy
                            'DestBook.Sheets("Income Details").Range(rname.name).PasteSpecial xlValues
                            DestBook.Sheets("Income Details").Protect Password:=EfilingCommon.getmsgstate
                        End If
            
            Dim rangecells3, rangecells1, rangecells2 As Variant
                     Dim chcells3, chcells1, chcells2 As Variant
                     Set rangecells3 = SrcBook.Sheets("Income Details").Range("Others.NOI").Cells
                     Set rangecells1 = SrcBook.Sheets("Income Details").Range("Others.NOI_1").Cells
                     Set rangecells2 = SrcBook.Sheets("Income Details").Range("Others.NOI_2").Cells
                    'Dim chcells As Variant
                    
                    
                    
                     For Each chcells3 In rangecells3
                    Templock3 (chcells3.AddressLocal)
                    Next
                     
                    For Each chcells2 In rangecells2
                    Templock2 (chcells2.AddressLocal)
                    Next
                    
                    For Each chcells1 In rangecells1
                    Templock1 (chcells1.AddressLocal)
                    Next
                    
                   
                     SrcBook.Sheets("Income Details").Unprotect Password:=EfilingCommon.getmsgstate
                     DestBook.Sheets("Income Details").Unprotect Password:=EfilingCommon.getmsgstate
                     SrcBook.Sheets("Taxes Paid and Verification").Unprotect Password:=EfilingCommon.getmsgstate
                     DestBook.Sheets("Taxes Paid and Verification").Unprotect Password:=EfilingCommon.getmsgstate
                     
                     DestBook.Sheets("Income Details").Range("Sheet1.MobileCountryCode").MergeArea.Value = "91"
                    DestBook.Sheets("Income Details").Range("Sheet1.MobileCountryCode").MergeArea.Locked = True
                    DestBook.Sheets("Income Details").Range("Sheet1.MobileCountryCode").MergeArea.Interior.Color = "&HD8D8D8"
                     
                     
                     SrcBook.Worksheets("Income Details").Range("Nature_Others_1").Copy
                    DestBook.Worksheets("Income Details").Range("Nature_Others_1").PasteSpecial xlValues
                    
                    SrcBook.Worksheets("Income Details").Range("Nature_Others_2").Copy
                    DestBook.Worksheets("Income Details").Range("Nature_Others_2").PasteSpecial xlValues
                    
                    SrcBook.Worksheets("Income Details").Range("Nature_Others").Copy
                    DestBook.Worksheets("Income Details").Range("Nature_Others").PasteSpecial xlValues
                     
                     
'                     Dim rangecells4 As Variant
'                     Dim chcells4 As Variant
'                     Set rangecells4 = SrcBook.Sheets("Income Details").Range("Nature_Others").Cells
'                    'Dim chcells As Variant
'                    For Each chcells4 In rangecells4
'                    If chcells4.Locked = True Then
'                    ElseIf chcells4.Locked = False Then
'                    DestBook.Sheets("Income Details").Range("chcells4.addresslocal").Value = Range("chcells4").Value
'                    End If
'                    Next
'
'                    Set rangecells4 = SrcBook.Sheets("Income Details").Range("Nature_Others_1").Cells
'                    'Dim chcells As Variant
'                    For Each chcells4 In rangecells4
'                    If chcells4.Locked = True Then
'                    ElseIf chcells4.Locked = False Then
'                    DestBook.Sheets("Income Details").Range("chcells4.addresslocal").Value = Range("chcells4").Value
'                    End If
'                    Next
'
'                    Set rangecells4 = SrcBook.Sheets("Income Details").Range("Nature_Others_2").Cells
'                    'Dim chcells As Variant
'                    For Each chcells4 In rangecells4
'                    If chcells4.Locked = True Then
'                    ElseIf chcells4.Locked = False Then
'                    DestBook.Sheets("Income Details").Range("chcells4.addresslocal").Value = Range("chcells4").Value
'                    End If
'                    Next
                    
                    
                    
                    
                    
                    
                    
                    DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").Value = SrcBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").Value
                    Dim return1 As Variant
                    return1 = DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").Value
                    
                  returnfilesecimport

                  
'                    Sheet1.Range ("sheet1.ReturnFileSec1")
                  
'                    If Mid(SrcBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").Value, 1, 2) = "11" Or Mid(SrcBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").Value, 1, 2) = "12" Or Mid(SrcBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").Value, 1, 2) = "17" Or Mid(SrcBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").Value, 1, 2) = "18" Then
'                    DestBook.Sheets("Income Details").Range("IncD.IntrstPayUs234A").MergeArea.Interior.ColorIndex = 2
'                    DestBook.Sheets("Income Details").Range("IncD.IntrstPayUs234A").MergeArea.Locked = True
'
'                    DestBook.Sheets("Income Details").Range("IncD.intrstPayUs234B").MergeArea.Interior.ColorIndex = 2
'                    DestBook.Sheets("Income Details").Range("IncD.intrstPayUs234B").MergeArea.Locked = True
'                    Else
'                    DestBook.Sheets("Income Details").Range("IncD.IntrstPayUs234A").MergeArea.Interior.Color = RGB(255, 255, 204)
'                    DestBook.Sheets("Income Details").Range("IncD.IntrstPayUs234A").MergeArea.Locked = False
'                    DestBook.Sheets("Income Details").Range("IncD.intrstPayUs234B").MergeArea.Interior.Color = RGB(255, 255, 204)
'                    DestBook.Sheets("Income Details").Range("IncD.intrstPayUs234B").MergeArea.Locked = False
'                    End If
                    
                    DestBook.Worksheets("Income Details").Range("sheet1.Mobileno").MergeArea.Value = SrcBook.Worksheets("Income Details").Range("sheet1.Mobileno").MergeArea.Value
                    DestBook.Worksheets("Taxes Paid and Verification").Activate
                    SrcBook.Worksheets("Taxes Paid and Verification").Range("Nature_Others").Copy
                    DestBook.Worksheets("Taxes Paid and Verification").Range("Nature_Others").PasteSpecial xlValues
                    DestBook.Sheets("Taxes Paid and Verification").Protect Password:=EfilingCommon.getmsgstate
                    SrcBook.Sheets("Taxes Paid and Verification").Protect Password:=EfilingCommon.getmsgstate
                    DestBook.Sheets("Income Details").Protect Password:=EfilingCommon.getmsgstate
                    SrcBook.Sheets("Income Details").Protect Password:=EfilingCommon.getmsgstate
            
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            ProgressBarHide
            fmsgbox "Import Completed"
            'DestBook.Save
            DestBook.Worksheets("Income Details").Select
            Set SrcBook = Nothing
        Else
            fmsgbox "Source file must not have same name As destination File"
        End If
    End If
End Sub
Function InsertRowsToImport(Optional vRows As Long = 0)
    Dim x As Long
    Dim strpassword As String
    Dim sht As Worksheet, shts() As String, i As Long
    
    strpassword = EfilingCommon.getmsgstate
    ActiveSheet.Unprotect Password:=strpassword
    
    ActiveCell.EntireRow.Select
    
    ReDim shts(1 To Worksheets.Application.ActiveWorkbook. _
       Windows(1).SelectedSheets.count)
    i = 0
    For Each sht In _
        Application.ActiveWorkbook.Windows(1).SelectedSheets
        Sheets(sht.name).Select
        i = i + 1
        shts(i) = sht.name
        
        x = Sheets(sht.name).UsedRange.Rows.count
        
        Selection.Resize(rowsize:=2).Rows(2).EntireRow. _
        Resize(rowsize:=vRows).Insert Shift:=xlDown
        
        Selection.AutoFill Selection.Resize( _
        rowsize:=vRows + 1), xlFillDefault

        On Error Resume Next
            Selection.Offset(1).Resize(vRows).EntireRow. _
            SpecialCells(xlConstants).ClearContents
    Next sht
    ActiveSheet.Protect Password:=strpassword
End Function

Function cmdFileDialog() As String
    Dim fDialog As Office.FileDialog
    Dim varFile As Variant
 
    cmdFileDialog = ""
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
      .AllowMultiSelect = False
      .Filters.Clear
      .Filters.add "Microsoft Office Excel Workbook", "*.xls,*.xlsm"
        If .Show = True Then
            For Each varFile In .SelectedItems
               cmdFileDialog = varFile
            Next
        End If
    End With
End Function
Sub ExendRangeNameToTable(numberofrows As Long, rangenamestring As Variant)
    Dim i As Long
    Dim x As Long
    Dim firstbound As String
    Dim temp As Variant
    Dim upperbound As String
    Dim lastbound As String
    
    rangenamestring = Split(rangenamestring, ";")
    For i = 0 To UBound(rangenamestring) - 1
        firstbound = Range(rangenamestring(i)).Address
        temp = Split(firstbound, "$")
        upperbound = UBound(temp)
        temp = temp(UBound(temp))
        x = CLng(temp) + numberofrows
        lastbound = Replace(firstbound, temp, x)
        If upperbound < 3 Then
            RangeAddress = firstbound & ":" & lastbound
        Else
            RangeAddress = lastbound
        End If
        ThisWorkbook.Names.add name:=rangenamestring(i), _
                 RefersTo:="=" & RangeAddress, Visible:=True
    Next
End Sub
Sub setTblinfo_TCSimport()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet11.Range("TCS.TAN").count
    Set rangecells = Sheet11.Range("TCS.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TCSimport = ccount
    
rngname_TCSimport = "TCS.TAN;TCS.EmployerOrDeductorOrCollecterName;TCS.TotalTCS;TCS.AmtTCSClaimedThisYear;TCS.AmountCollected;TCS.CollectionYear;"
End Sub
Sub setTblinfo_TDS1import()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet2.Range("TDSal.TAN").count
    Set rangecells = Sheet2.Range("TDSal.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS1import = ccount
    rngname_TDS1import = "TDSal.TAN;TDSal.EmployerOrDeductorOrCollecterName;TDSal.IncChrgSalary;TDSal.TotalTDSSalary;"
End Sub

Sub setTblinfo_TDS2import()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet2.Range("TDSoth.TAN").count
    Set rangecells = Sheet2.Range("TDSoth.TAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS2import = ccount
    rngname_TDS2import = "TDSoth.TAN;TDSoth.EmployerOrDeductorOrCollecterName;TDSoth.AmountDeducted;TDSoth.DeductedYear;TDSoth.TotTDSOnAmtPaid;TDSoth.6income;"

End Sub
'NEW TABLE FOR IMPORT TDS3
Sub setTblinfo_TDS3import()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    
    ccount = 0
    mIntCells = Sheet2.Range("TDS26QB.PAN").count
    Set rangecells = Sheet2.Range("TDS26QB.PAN").Cells
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_TDS3import = ccount
    rngname_TDS3import = "TDS26QB.PAN;TDS26QB.Aadhar_Number;TDS26QB.EmployerOrDeductorName;TDS26QB.AmountDeducted;TDS26QB.DeductedYear;TDS26QB.TotTDSOnAmtPaid;TDS26QB.6income;"
End Sub
Sub setTblinfo_ITimport()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Sheet2.Range("TaxP.BSRCode").count
 Set rangecells = Sheet2.Range("TaxP.BSRCode").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_ITimport = ccount
 rngname_ITimport = "TaxP.BSRCode;TaxP.DateDep;TaxP.SrlNoOfChaln;TaxP.Amt;IT.FormulaOFS;FormulaOfQ;FormulaOfSAT;FormulaOfSAT1||FormulaOfSATNew;IT_DueDate;"  'Ayush_DueDate_08/09/2025
End Sub

Sub setTblinfo_BAimport()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Range("SchBA.IFSC").count
 Set rangecells = Range("SchBA.IFSC").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_BAimport = ccount
 rngname_BAimport = "SchBA.IFSC;SchBA.BankName;SchBA.AcntNo;SchBA.AcntType;"
End Sub

Sub setTblinfo_OtherEI()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Range("Others.NOI").count
 Set rangecells = Range("Others.NOI").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_OtherEI = ccount
 rngname_OtherEI = "Others.NOI;Nature_Others;Others.Amount;"
End Sub
Sub setTblinfo_OthersNOI_1()
Dim rangecells As Range
 Dim mIntCells As Long
 Dim mIntCtr As Long
 Dim ccount As Long
 ccount = 0
 mIntCells = Range("Others.NOI_1").count
 Set rangecells = Range("Others.NOI_1").Cells
 For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_OtherEI1 = ccount
 rngname_OtherEI1 = "Others.NOI_1;Nature_Others_1;Others.Amount_1;"
End Sub
 Sub setTblinfo_OthersNOI_2()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Range("Others.NOI_2").count
    Set rangecells = Range("Others.NOI_2").Cells
    Dim countrycd As Variant
   For mIntCtr = 1 To mIntCells
     If Not rangecells.item(mIntCtr).Value = "" Then
         ccount = ccount + 1
     End If
 Next
 end_OtherEI2 = ccount
 rngname_OtherEI2 = "Others.NOI_2;Nature_Others_2;Others.Amount_2;"
End Sub
Sub setTblinfo_Per10080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("Per10080G.DoneeName").count
    Set rangecells = Sheet4.Range("Per10080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per10080G = ccount
    rngname_Per10080G = "Per10080G.DoneeName;Per10080G.AddrDetail;Per10080G.CityOrTownOrDistrict;Per10080G.StateCode;Per10080G.PinCode;Per10080G.DoneePAN;Per10080G.DonationAmt;Per10080G.DonationAmtTotal;Per10080G.EligibleAmt;Per10080G.DonationAmtOther;"

End Sub
Sub setTblinfo_PerNO5080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("PerNO5080G.DoneeName").count
    Set rangecells = Sheet4.Range("PerNO5080G.DoneeName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerNO5080G = ccount
    rngname_PerNO5080G = "PerNO5080G.DoneeName;PerNO5080G.AddrDetail;PerNO5080G.CityOrTownOrDistrict;PerNO5080G.StateCode;PerNO5080G.PinCode;PerNO5080G.DoneePAN;PerNO5080G.DonationAmt;PerNO5080G.DonationAmtTotal;PerNO5080G.EligibleAmt;PerNO5080G.DonationAmtOther;"
End Sub
Sub setTblinfo_PerYES10080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("PerYES10080G.DoneeWithPanName").count
    Set rangecells = Sheet4.Range("PerYES10080G.DoneeWithPanName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_PerYES10080G = ccount
    rngname_PerYES10080G = "PerYES10080G.DoneeWithPanName;PerYES10080G.AddrDetail;PerYES10080G.CityOrTownOrDistrict;PerYES10080G.StateCode;PerYES10080G.PinCode;PerYES10080G.DoneePAN;PerYES10080G.DonationAmt;PerYES10080G.DonationAmtTotal;PerYES10080G.EligibleAmt;PerYES10080G.DonationAmtOther;"
End Sub
Sub setTblinfo_Per5080G()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount As Long
    ccount = 0
    mIntCells = Sheet4.Range("Per5080G.DoneeWithPanName").count
    Set rangecells = Sheet4.Range("Per5080G.DoneeWithPanName").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not rangecells.item(mIntCtr).Value = "" Then
            ccount = ccount + 1
        End If
    Next
    end_Per5080G = ccount
    rngname_Per5080G = "Per5080G.DoneeWithPanName;Per5080G.AddrDetail;Per5080G.CityOrTownOrDistrict;Per5080G.StateCode;Per5080G.PinCode;Per5080G.DoneePAN;Per5080G.DonationAmt;Per5080G.DonationAmtTotal;Per5080G.EligibleAmt;Per5080G.DonationAmtOther;"
End Sub



Sub setTblinfo_80GGA()
    Dim rangecells As Range
    Dim mIntCells As Long
    Dim mIntCtr As Long
    Dim ccount, end_GGA As Long
    ccount = 0
    mIntCells = Sheet12.Range("RelevantClauseClaimed_80GGA").count
    Set rangecells = Sheet12.Range("RelevantClauseClaimed_80GGA").Cells
    
    For mIntCtr = 1 To mIntCells
        If Not isdropdownblank(rangecells.item(mIntCtr).Value) Then
            ccount = ccount + 1
        End If
    Next
    end_GGA = ccount
    rngname_80GGA = "RelevantClauseClaimed_80GGA;Name_of_Donee_80GGA;Address_80GGA;City_Town_District_80GGA;State_Code_80GGA;Pincode_80GGA;PAN_of_donee_80GGA;Donation_cash_80GGA;Donation_other_80GGA;Donation_total_80GGA;Donation_Eligible_80GGA;"
End Sub
Sub Templock1(targetadd)


If Not Application.Intersect(Sheet1.Range("Others.NOI_1"), Sheet1.Range(targetadd)) Is Nothing Then
        
        
        If Not checkfieldSuperSpecialcharactername(Range(targetadd).Value) Then
                fmsgbox ("Nature of income should not contain Special characters < , >")
               Range(targetadd).Value = ""
        End If
        
        
        
        If Sheet1.Range(targetadd).Value = "Any Other" Then
            Sheet1.Unprotect Password:=getmsgstate
            'Sheet3.Range(Replace(targetadd, "E", "F")).Value = ""
            Sheet1.Range(Replace(targetadd, "J", "Z")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(targetadd, "J", "Z")).Locked = False
            
            'Sheet3.Range("sheet1.ReturnFurSec1").Value = "Others"
            
            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(targetadd, "J", "Z")).Value = "Not Applicable"
            Sheet1.Range(Replace(targetadd, "J", "Z")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(targetadd, "J", "Z")).Locked = True
            Sheet1.Protect Password:=getmsgstate
    End If
End If

End Sub
Sub Templock2(targetadd)


If Not Application.Intersect(Sheet1.Range("Others.NOI_2"), Sheet1.Range(targetadd)) Is Nothing Then
        
        
        If Not checkfieldSuperSpecialcharactername(Range(targetadd).Value) Then
                fmsgbox ("Nature of income should not contain Special characters < , >" & ",&" & Chr(13))
               Range(targetadd).Value = ""
        End If
        
        
        
        If Sheet1.Range(targetadd).Value = "Any Other" Then
            Sheet1.Unprotect Password:=getmsgstate
            'Sheet3.Range(Replace(targetadd, "E", "F")).Value = ""
            Sheet1.Range(Replace(targetadd, "J", "Z")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(targetadd, "J", "Z")).Locked = False
            
            'Sheet3.Range("sheet1.ReturnFurSec1").Value = "Others"
            
            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(targetadd, "J", "Z")).Value = "Not Applicable"
            Sheet1.Range(Replace(targetadd, "J", "Z")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(targetadd, "J", "Z")).Locked = True
            Sheet1.Protect Password:=getmsgstate
    End If
End If

End Sub
Sub Templock3(targetadd)


If Not Application.Intersect(Sheet1.Range("Others.NOI"), Sheet1.Range(targetadd)) Is Nothing Then
        
        
        If Not checkfieldSuperSpecialcharactername(Range(targetadd).Value) Then
                fmsgbox ("Nature of income should not contain Special characters < , >" & ",&" & Chr(13))
               Range(targetadd).Value = ""
        End If
        
        
        
        If Sheet1.Range(targetadd).Value = "Any Other" Then
            Sheet1.Unprotect Password:=getmsgstate
            'Sheet3.Range(Replace(targetadd, "E", "F")).Value = ""
            Sheet1.Range(Replace(targetadd, "H", "T")).Interior.Color = (&HCCFFCC)  ', "AU"
            Sheet1.Range(Replace(targetadd, "H", "T")).Locked = False
            
            'Sheet3.Range("sheet1.ReturnFurSec1").Value = "Others"
            
            Sheet1.Protect Password:=getmsgstate
        Else
            Sheet1.Unprotect Password:=getmsgstate
            Sheet1.Range(Replace(targetadd, "H", "T")).Value = "Not Applicable"
            Sheet1.Range(Replace(targetadd, "H", "T")).Interior.Color = (&HD8D8D8)
            Sheet1.Range(Replace(targetadd, "H", "T")).Locked = True
            Sheet1.Protect Password:=getmsgstate
    End If
End If

End Sub


Function getSheetName(rn As Variant) As String
Dim temp As Variant
temp = Replace(Mid(rn, 2, InStr(1, rn, "!") - 2), "'", "")
getSheetName = temp
End Function


Sub returnfilesecimport()
'    DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").Value = SrcBook.Sheets("Income Details").Range("sheet1.ReturnFileSec1").Value
                    Dim return1 As Variant
          
                    return1 = DestBook.Sheets("Income Details").Range("sheet1.ReturnFileSec").Value
               
         
                    
                    
        If return1 = "(Select)" Then
           Range("sheet1.ReturnFileSec").Value = "(Select)"
        ElseIf Mid(return1, 1, 2) = "11" Then
           Range("sheet1.ReturnFileSec").Value = "139(1)-On or before due date"
        ElseIf Mid(return1, 1, 2) = "12" Then
           Range("sheet1.ReturnFileSec").Value = "139(4)-Belated"
        ElseIf Mid(return1, 1, 2) = "13" Then
           Range("sheet1.ReturnFileSec").Value = "142(1)"
        ElseIf Mid(return1, 1, 2) = "14" Then
           Range("sheet1.ReturnFileSec").Value = "148"
        ElseIf Mid(return1, 1, 2) = "15" Then
           Range("sheet1.ReturnFileSec").Value = "153A"
        ElseIf Mid(return1, 1, 2) = "16" Then
           Range("sheet1.ReturnFileSec").Value = "153C"
        ElseIf Mid(return1, 1, 2) = "17" Then
           Range("sheet1.ReturnFileSec").Value = "139(5)-Revised"
        ElseIf Mid(return1, 1, 2) = "20" Then
'Konda updated on 23-12 2025
'           Range("sheet1.ReturnFileSec").Value = "139(9A) - After condonation of delay u/s 119(2)(b)"
            Range("sheet1.ReturnFileSec").Value = "119(2)(b)- After condonation of delay"
       
        End If

         
        
End Sub















