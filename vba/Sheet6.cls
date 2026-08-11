Attribute VB_Name = "Sheet6"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'Application.OnKey "{TAB}"
'MsgBox ActiveCell.Address


If Not Range(ActiveCell.Address).Locked = True Then
       If ActiveCell.Validation.Type = 6 Then
    Application.OnKey "%{DOWN}", ""
    'Sheet1.Range(Target.Address).Value = ""
ElseIf ActiveCell.Validation.Type = 1 Then
 Application.OnKey "%{DOWN}", ""
ElseIf ActiveCell.Validation.Type = 3 Then

    Application.OnKey "%{DOWN}"
End If
Else
Application.OnKey "%{DOWN}", ""
End If
End Sub


Private Sub Worksheet_Deactivate()

    sPassword = EfilingCommon.getmsgstate
ActiveWorkbook.Unprotect Password:=sPassword
    Sheet6.Visible = xlSheetHidden
    ActiveWorkbook.Protect Password:=sPassword
End Sub

'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
''Application.OnKey "{TAB}"
'
'If Not Sheet6.Range(Target.Address).Locked = True Then
'       If Target.Validation.Type = 6 Then
'    Application.OnKey "%{DOWN}", ""
'    'Sheet1.Range(Target.Address).Value = ""
'ElseIf Target.Validation.Type = 1 Then
' Application.OnKey "%{DOWN}", ""
'ElseIf Target.Validation.Type = 3 Then
'
'    Application.OnKey "%{DOWN}"
'End If
'Else
'Application.OnKey "%{DOWN}", ""
'End If
'End Sub
