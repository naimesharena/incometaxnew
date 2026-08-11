Attribute VB_Name = "mfMessage"
Public Sub fmsgbox(IntMsg As Variant)
'With MessageBox
'    .Height = 281
'    .LMessagebox.Height = 204
'    .CommandButton1.Top = 222
'    .LMessagebox = IntMsg
'    '.LMessagebox.TextAlign = fmTextAlignLeft
'    .Show
'End With

'Chandru-27/03/2025

MsgBox IntMsg, , "Alert"

End Sub
Public Sub fmsgboxoK(iintmsg As Variant)
'With MessageBox
'    .Height = 130
'    .LMessagebox.Height = 60
'    .CommandButton1.Top = 78
'    .LMessagebox = "*" + iintmsg
'    .LMessagebox.TextAlign = fmTextAlignLeft
'    .Show
'End With
'Chandru-27/03/2025
MsgBox iintmsg & Chr(10) & Chr(10), vbOKOnly, "Alert"


End Sub
'Ankita_04/02/2025
Public Sub fmsgboxsmall_LTCG(IntMsg2 As Variant)
     MsgBox IntMsg2, , "Error(S)"
End Sub


'Ankita_15/04/2025
Public Sub fmsgboxStatus(IntMsg1 As Variant)
With MessageBox
    .Height = 250
    .LMessagebox.Height = 200
    .CommandButton1.Top = 195
    .LMessagebox = "*" + IntMsg1
    .LMessagebox.TextAlign = fmTextAlignLeft
    .Show
End With

End Sub
