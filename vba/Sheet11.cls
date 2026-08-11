Attribute VB_Name = "Sheet11"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
''Application.OnKey "{TAB}"
''MsgBox ActiveCell.Address
'
'
'If Not Range(ActiveCell.Address).Locked = True Then
'       If ActiveCell.Validation.Type = 6 Then
'    Application.OnKey "%{DOWN}", ""
'    'Sheet1.Range(Target.Address).Value = ""
'ElseIf ActiveCell.Validation.Type = 1 Then
' Application.OnKey "%{DOWN}", ""
'ElseIf ActiveCell.Validation.Type = 3 Then
'
'    Application.OnKey "%{DOWN}"
'End If
'Else
'Application.OnKey "%{DOWN}", ""
'End If
'End Sub


Private Sub Worksheet_Activate()
On Error GoTo endline
Application.EnableEvents = False

ActiveWindow.DisplayHeadings = False

ActiveSheet.Shapes("Button 6").Left = Range("SchTCS.Button1").Left + 6
ActiveSheet.Shapes("Button 6").Top = Range("SchTCS.Button1").Top + 6

endline:
Application.EnableEvents = True

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
Cancel = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
On Error GoTo endline
Application.EnableEvents = False


sPassword = EfilingCommon.getmsgstate
Sheet1.Unprotect Password:=sPassword
Sheet1.Range("taxcheck").Value = ""
Sheet1.Protect Password:=sPassword
If Not Intersect(Range("TCS.TAN"), Range(Target.Address)) Is Nothing Then
Range(Target.Address) = UCase(Range(Target.Address))
If Not checkfieldSuperSpecialcharacterDot(Range(Target.Address).Value) Then

fmsgbox "Invalid TAN. TAN should contain First 4 alphabets, next 5 digits, then alphabet"
Range(Target.Address).Value = ""
End If
End If

'Target.Value = UCase(Target.Value)

If Not Intersect(Range("TCS.EmployerOrDeductorOrCollecterName"), Range(Target.Address)) Is Nothing Then
If Not checkfieldSuperSpecialcharactername(Range(Target.Address).Value) Then
         fmsgbox ("Please do not enter Special characters < , >" & ",&" & Chr(13))
        Range(Target.Address).Value = ""
End If
End If
endline:
Application.EnableEvents = True

End Sub
