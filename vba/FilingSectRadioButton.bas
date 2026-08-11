Attribute VB_Name = "FilingSectRadioButton"
Sub RadioButton1_Click()

Dim op As OptionButton
Set op = Sheet1.Shapes("RadioButton2").OLEFormat.Object
op.Value = False

Sheet1.Range("sheet1.ReturnFileSec").Value = "(Select)"

    Dim cellrange As Range

    Set cellrange = Sheet1.Range("sheet1.ReturnFileSec")
    Formula = "=ReturnSecList"

    sPassword = EfilingCommon.getmsgstate
    Sheet1.Unprotect Password:=sPassword

    With cellrange.Validation
        .Delete
        .add Type:=xlValidateList, _
         AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, _
         Formula1:=Formula
    End With

    Sheet1.Protect Password:=sPassword

End Sub

Sub RadioButton2_Click()

Dim op As OptionButton
Set op = Sheet1.Shapes("RadioButton1").OLEFormat.Object
op.Value = False

Sheet1.Range("sheet1.ReturnFileSec").Value = "(Select)"

    Dim cellrange As Range

    Set cellrange = Sheet1.Range("sheet1.ReturnFileSec")
'    Change.06.02.2023.102.CDS2
'    Formula = "=ReturnSecList1"
    Formula = "=Returnlist153C"
'   End Change
    sPassword = EfilingCommon.getmsgstate
    Sheet1.Unprotect Password:=sPassword

    With cellrange.Validation
        .Delete
        .add Type:=xlValidateList, _
         AlertStyle:=xlValidAlertStop, _
         Operator:=xlBetween, _
         Formula1:=Formula
    End With

    Sheet1.Protect Password:=sPassword

End Sub
