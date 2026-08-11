Attribute VB_Name = "Module5"

'Sub Addrows24b()
'Sheets("Schedule 24(b)").Activate
''    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||IFSC.24b||bankName.24b||PAN.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b||Combination_24B"
'   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.24b||bankName.24b||loanAccNum.24b||loanDate.24b||loanAmt.24b||loanOutstanding.24b||Intrst.24b||Combination_24B"
'    ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.24b")
'    EfilingCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
'End Sub

Sub Addrows80E()
Sheets("80E_80EE_80EEA_80EEB").Activate
   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'   EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80E||IFSC.80E||bankName.80E||PAN.80E||loanAccNum.80E||loanDate.80E||loanAmt.80E||loanOutstanding.80E||Intrst.80E"
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80E||bankName.80E||loanAccNum.80E||loanDate.80E||loanAmt.80E||loanOutstanding.80E||Intrst.80E"
    ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80E")
    EfilingCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
End Sub
Sub Addrows80EE()
Sheets("80E_80EE_80EEA_80EEB").Activate
   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'   EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||IFSC.80EE||bankName.80EE||PAN.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE||Combination_80EE"
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EE||bankName.80EE||loanAccNum.80EE||loanDate.80EE||loanAmt.80EE||loanOutstanding.80EE||Intrst.80EE||Combination_80EE"
    ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80EE")
    EfilingCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
End Sub
Sub Addrows80EEA()
Sheets("80E_80EE_80EEA_80EEB").Activate
   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'   EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||IFSC.80EEA||bankName.80EEA||PAN.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA||Combination_80EEA"
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEA||bankName.80EEA||loanAccNum.80EEA||loanDate.80EEA||loanAmt.80EEA||loanOutstanding.80EEA||Intrst.80EEA||Combination_80EEA"
    ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80EEA")
    EfilingCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
End Sub
Sub Addrows80EEB()
Sheets("80E_80EE_80EEA_80EEB").Activate
   'Ankita_05/05/2025_Commented as per DESheet_v0.7
'   EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||IFSC.80EEB||bankName.80EEB||PAN.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEB||loanOutstanding.80EEB||Vehicle_value.80EEB||VehicleRegNum.80EEB||Intrst.80EEB"
    EfilingCommon.DefinedgridNameRange = "LoanfrmBankOrInstitute.80EEB||bankName.80EEB||loanAccNum.80EEB||loanDate.80EEB||loanAmt.80EEB||loanOutstanding.80EEB||VehicleRegNum.80EEB||Intrst.80EEB"
    ActiveCellRange = EfilingCommon.searchLastRow("LoanfrmBankOrInstitute.80EEB")
    EfilingCommon.insertRowUnderSectionWithFormula24b HOIflag:=1
End Sub

'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Newly modified by sai on 22/04/2025
Sub Addrows80C()
Dim vRows As Long
Dim sourceSheet As Worksheet
'Sheets("80C").Activate
Set sourceSheet = ThisWorkbook.Sheets("80C")
    EfilingCommon.DefinedgridNameRange = "Identification_Number.80C||Amount.80C"  'Ankita_06/05/2025_Commented as per DESheet_v0.7
    ActiveCellRange = EfilingCommon.searchLastRow("Identification_Number.80C")
    vRows = insertRowUnderSectionWithFormula_80C
    'EfilingCommon.insertRowUnderSectionWithFormula_80C
End Sub

'Newly modified by sai on 22/04/2025
'Ankita_06/05/2025_Commented as per DESheet_v0.7

'Sub Addrows80CCC()
'Dim vRows As Long
'Dim sourceSheet As Worksheet
'Set sourceSheet = ThisWorkbook.Sheets("80C_80CCC")
''Sheets("80CCC").Activate
'     EfilingCommon.DefinedgridNameRange = "Name_of_insurer.80CCC||Policy_Document_Number.80CCC||Amount.80CCC"
'     ActiveCellRange = EfilingCommon.searchLastRow("Name_of_insurer.80CCC")
''    EfilingCommon.insertRowUnderSectionWithFormula80C HOIflag:=1
'     vRows = insertRowUnderSectionWithFormula_80CCC
'End Sub



Sub Addrows80DA1()
Sheets("80D").Activate
   EfilingCommon.DefinedgridNameRange = "NameInsurerA1.80D||PolicyNumA1.80D||AmtA1.80D" 'Ankita_06/05/2025_Commented as per DESheet_v0.7

    ActiveCellRange = EfilingCommon.searchLastRow("NameInsurerA1.80D")
    EfilingCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
End Sub
Sub Addrows80DB1()
Sheets("80D").Activate
    EfilingCommon.DefinedgridNameRange = "NameInsurerB1.80D||PolicyNumB1.80D||AmtB1.80D" 'Ankita_06/05/2025_Commented as per DESheet_v0.7

    ActiveCellRange = EfilingCommon.searchLastRow("NameInsurerB1.80D")
    EfilingCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
End Sub
Sub Addrows80DA2()
Sheets("80D").Activate
    EfilingCommon.DefinedgridNameRange = "NameInsurerA2.80D||PolicyNumA2.80D||AmtA2.80D" 'Ankita_06/05/2025_Commented as per DESheet_v0.7

    ActiveCellRange = EfilingCommon.searchLastRow("NameInsurerA2.80D")
    EfilingCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
End Sub
Sub Addrows80DB2()
Sheets("80D").Activate
    EfilingCommon.DefinedgridNameRange = "NameInsurerB2.80D||PolicyNumB2.80D||AmtB2.80D"  'Ankita_06/05/2025_Commented as per DESheet_v0.7

    ActiveCellRange = EfilingCommon.searchLastRow("NameInsurerB2.80D")
    EfilingCommon.insertRowUnderSectionWithFormula80D HOIflag:=1
End Sub

